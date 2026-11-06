import { createReadStream, existsSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import { createServer } from "node:http";
import { homedir, tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { CanvasError, createCanvas, joinSession } from "@github/copilot-sdk/extension";
import {
    ACCEPT_EXTENSIONS,
    KINDS,
    MAX_FILE_BYTES,
    disposeDocument,
    extractPdfText,
    isWithin,
    kindOf,
    loadDocument,
    sheetToText,
} from "./loaders.mjs";

const instances = new Map();
const extensionHome = path.join(
    process.env.COPILOT_HOME || path.join(homedir(), ".copilot"),
    "extensions",
);
const historyFile = path.join(extensionHome, "office-viewer", "artifacts", "history.json");
const legacyHistoryFiles = [
    ["excel-viewer"],
    ["ppt-viewer"],
    ["word-pdf-viewer"],
].map(([name]) => path.join(extensionHome, name, "artifacts", "history.json"));
const HISTORY_LIMIT = 30;

const indexHtml = await readFile(new URL("./public/index.html", import.meta.url), "utf8");
const excelHtml = await readFile(new URL("./public/excel.html", import.meta.url), "utf8");

const imageTypes = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
    ".svg": "image/svg+xml",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
    ".emf": "image/emf",
    ".wmf": "image/wmf",
};

async function readLegacyHistory() {
    const items = [];
    for (const file of legacyHistoryFiles) {
        try {
            const value = JSON.parse(await readFile(file, "utf8"));
            const list = Array.isArray(value) ? value : Object.values(value ?? {}).flat();
            items.push(...list.filter((item) => typeof item === "string" && item.trim()));
        } catch {
            // 旧版历史记录不存在或损坏时直接忽略
        }
    }
    return [...new Set(items)].filter((item) => kindOf(item)).slice(0, HISTORY_LIMIT);
}

async function readHistory() {
    try {
        const value = JSON.parse(await readFile(historyFile, "utf8"));
        return Array.isArray(value)
            ? value.filter((item) => typeof item === "string" && item.trim()).slice(0, HISTORY_LIMIT)
            : [];
    } catch (error) {
        if (error.code === "ENOENT") return readLegacyHistory();
        throw new Error(`读取历史记录失败：${error.message}`);
    }
}

async function addHistory(filePath) {
    const history = await readHistory();
    const next = [filePath, ...history.filter((item) => item !== filePath)].slice(0, HISTORY_LIMIT);
    await mkdir(path.dirname(historyFile), { recursive: true });
    const temporaryPath = `${historyFile}.${process.pid}.tmp`;
    await writeFile(temporaryPath, JSON.stringify(next, null, 2), "utf8");
    await rename(temporaryPath, historyFile);
    return next;
}

function withLock(entry, task) {
    const run = entry.lock.then(task, task);
    entry.lock = run.catch(() => {});
    return run;
}

async function clearUpload(entry, keep = "") {
    if (!entry.uploadDir || entry.uploadDir === keep) return;
    const directory = entry.uploadDir;
    entry.uploadDir = "";
    await rm(directory, { recursive: true, force: true });
}

function broadcast(entry) {
    for (const client of entry.clients) client.write("data: changed\n\n");
}

async function replaceDocument(entry, filePath, { uploadDir = "", displayName = "" } = {}) {
    const next = await loadDocument(filePath);
    const previous = entry.document;
    entry.document = next;
    entry.filePath = next.filePath;
    entry.displayName = displayName || next.fileName;
    entry.displayPath = uploadDir ? displayName : next.filePath;
    entry.revision++;
    await disposeDocument(previous);
    await clearUpload(entry, uploadDir);
    entry.uploadDir = uploadDir;
    broadcast(entry);
    return next;
}

function closeDocument(entry) {
    return withLock(entry, async () => {
        const previous = entry.document;
        entry.document = null;
        entry.filePath = entry.displayPath = entry.displayName = "";
        entry.revision++;
        await disposeDocument(previous);
        await clearUpload(entry);
        broadcast(entry);
    });
}

function openPath(entry, filePath) {
    return withLock(entry, async () => {
        const document = await replaceDocument(entry, filePath.trim());
        await addHistory(document.filePath);
        return document;
    });
}

function refreshDocument(entry) {
    return withLock(entry, () => {
        if (!entry.document) throw new Error("当前没有打开文件");
        return replaceDocument(entry, entry.filePath, {
            uploadDir: entry.uploadDir,
            displayName: entry.uploadDir ? entry.displayName : "",
        });
    });
}

function pageUnit(kind) {
    return { pdf: "页", word: "页", ppt: "张幻灯片", excel: "个工作表" }[kind];
}

function publicDocument(entry) {
    const document = entry.document;
    if (!document) {
        return { empty: true, accept: ACCEPT_EXTENSIONS.join(","), revision: entry.revision };
    }
    const base = {
        empty: false,
        kind: document.kind,
        label: KINDS[document.kind].label,
        filePath: entry.displayPath,
        fileName: entry.displayName,
        fileSize: document.fileSize,
        pageCount: document.pageCount,
        unit: pageUnit(document.kind),
        warning: document.warning ?? "",
        revision: entry.revision,
        accept: ACCEPT_EXTENSIONS.join(","),
    };
    if (document.kind === "pdf") return { ...base, mode: "pdf" };
    if (document.kind === "word") {
        return {
            ...base,
            mode: document.pdfPath ? "pdf" : "text",
            paragraphs: document.paragraphs.map((paragraph) => ({
                text: paragraph.text,
                style: paragraph.style,
                images: paragraph.images.map((index) => `/media/${index}?v=${entry.revision}`),
            })),
        };
    }
    if (document.kind === "ppt") {
        return {
            ...base,
            mode: document.rendered ? "slides" : "parsed-slides",
            slides: document.slides.map((slide, slideIndex) => ({
                number: slide.number,
                text: slide.text,
                images: slide.imagePaths.map((_, imageIndex) => `/media/${slideIndex}/${imageIndex}?v=${entry.revision}`),
            })),
        };
    }
    return { ...base, mode: "sheet" };
}

function documentInfo(entry) {
    const info = publicDocument(entry);
    delete info.paragraphs;
    delete info.slides;
    delete info.accept;
    return info;
}

function sendJson(response, payload, status = 200) {
    response.writeHead(status, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store",
    });
    response.end(JSON.stringify(payload));
}

function sendText(response, body, contentType) {
    response.writeHead(200, { "Content-Type": contentType, "Cache-Control": "no-store" });
    response.end(body);
}

async function sendPdf(request, response, filePath) {
    const info = await stat(filePath);
    const match = typeof request.headers.range === "string"
        ? request.headers.range.match(/^bytes=(\d*)-(\d*)$/)
        : null;
    let start = 0;
    let end = info.size - 1;
    let status = 200;
    if (match) {
        if (!match[1] && !match[2]) {
            response.writeHead(416, { "Content-Range": `bytes */${info.size}` });
            response.end();
            return;
        }
        if (!match[1]) {
            start = Math.max(0, info.size - Number(match[2]));
        } else {
            start = Number(match[1]);
            if (match[2]) end = Math.min(end, Number(match[2]));
        }
        if (start >= info.size || start > end) {
            response.writeHead(416, { "Content-Range": `bytes */${info.size}` });
            response.end();
            return;
        }
        status = 206;
    }
    response.writeHead(status, {
        "Content-Type": "application/pdf",
        "Content-Length": end - start + 1,
        "Accept-Ranges": "bytes",
        "Cache-Control": "no-store",
        "Content-Disposition": "inline",
        ...(status === 206 ? { "Content-Range": `bytes ${start}-${end}/${info.size}` } : {}),
    });
    createReadStream(filePath, { start, end }).pipe(response);
}

function sendImage(response, imagePath, root) {
    if (!imagePath || !isWithin(root, imagePath) || !existsSync(imagePath)) {
        sendJson(response, { error: "图片不存在" }, 404);
        return;
    }
    response.writeHead(200, {
        "Content-Type": imageTypes[path.extname(imagePath).toLowerCase()] ?? "application/octet-stream",
        "Cache-Control": "no-store",
    });
    createReadStream(imagePath).pipe(response);
}

function readRequestBody(request, limit) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        let failed = false;
        request.on("data", (chunk) => {
            if (failed) return;
            size += chunk.length;
            if (size > limit) {
                failed = true;
                reject(new Error(limit === MAX_FILE_BYTES ? "文件大小不能超过 100 MB" : "请求内容过大"));
                return;
            }
            chunks.push(chunk);
        });
        request.once("end", () => { if (!failed) resolve(Buffer.concat(chunks)); });
        request.once("error", reject);
    });
}

function uploadedFileName(headerValue) {
    if (typeof headerValue !== "string" || !headerValue) throw new Error("未提供文件名");
    let decoded;
    try {
        decoded = decodeURIComponent(headerValue);
    } catch {
        throw new Error("文件名无效");
    }
    const fileName = path.basename(decoded);
    if (!kindOf(fileName)) throw new Error(`不支持的文件类型，当前支持 ${ACCEPT_EXTENSIONS.join("、")}`);
    return fileName;
}

const mediaRoot = (document) => document?.mediaRoot ?? "";

const pdfjsRoot = fileURLToPath(new URL("./vendor/pdfjs/", import.meta.url));
const pdfjsTypes = {
    ".mjs": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".bcmap": "application/octet-stream",
    ".pfb": "application/octet-stream",
    ".ttf": "font/ttf",
};

function sendPdfjsAsset(response, route) {
    let relative;
    try {
        relative = decodeURIComponent(route.slice("/vendor/pdfjs/".length));
    } catch {
        return sendJson(response, { error: "Not found" }, 404);
    }
    const file = path.resolve(pdfjsRoot, relative);
    const type = pdfjsTypes[path.extname(file).toLowerCase()];
    if (!relative || !type || !isWithin(pdfjsRoot, file) || !existsSync(file)) {
        return sendJson(response, { error: "Not found" }, 404);
    }
    response.writeHead(200, { "Content-Type": type, "Cache-Control": "max-age=3600" });
    createReadStream(file).pipe(response);
}

async function handleRequest(entry, request, response) {
    const port = entry.port;
    const allowedHosts = [`127.0.0.1:${port}`, `localhost:${port}`];
    if (!allowedHosts.includes(request.headers.host ?? "")) {
        sendJson(response, { error: "Forbidden" }, 403);
        return;
    }
    if (request.method === "POST" && request.headers.origin
        && !allowedHosts.some((host) => request.headers.origin === `http://${host}`)) {
        sendJson(response, { error: "Forbidden" }, 403);
        return;
    }

    const url = new URL(request.url ?? "/", "http://127.0.0.1");
    const route = url.pathname;
    const document = entry.document;

    if (request.method === "GET") {
        if (route === "/") return sendText(response, indexHtml, "text/html; charset=utf-8");
        if (route === "/excel.html") return sendText(response, excelHtml, "text/html; charset=utf-8");
        if (route === "/vendor/excel-renderer.js" || route === "/vendor/excel-renderer.css") {
            const file = new URL(`.${route}`, import.meta.url);
            response.writeHead(200, {
                "Content-Type": route.endsWith(".js") ? "text/javascript; charset=utf-8" : "text/css; charset=utf-8",
                "Cache-Control": "max-age=3600",
            });
            createReadStream(file).pipe(response);
            return;
        }
        if (route.startsWith("/vendor/pdfjs/")) return sendPdfjsAsset(response, route);
        if (route === "/events") {
            response.writeHead(200, {
                "Content-Type": "text/event-stream",
                "Cache-Control": "no-store",
                Connection: "keep-alive",
            });
            response.write(": connected\n\n");
            entry.clients.add(response);
            response.on("close", () => entry.clients.delete(response));
            return;
        }
        if (route === "/api/document") {
            if (url.searchParams.has("refresh")) {
                try {
                    await refreshDocument(entry);
                } catch (error) {
                    return sendJson(response, { error: error.message }, 422);
                }
            }
            return sendJson(response, publicDocument(entry));
        }
        if (route === "/api/workbook") {
            if (document?.kind !== "excel") return sendJson(response, { error: "当前文件不是 Excel 工作簿" }, 409);
            return sendJson(response, document.snapshot);
        }
        if (route === "/api/history") {
            try {
                return sendJson(response, { items: await readHistory() });
            } catch (error) {
                return sendJson(response, { error: error.message }, 500);
            }
        }
        if (route === "/document.pdf" && document?.pdfPath) {
            try {
                await sendPdf(request, response, document.pdfPath);
            } catch (error) {
                if (!response.headersSent) sendJson(response, { error: error.message }, 500);
                else response.destroy(error);
            }
            return;
        }
        const wordMedia = route.match(/^\/media\/(\d+)$/);
        if (wordMedia && document?.kind === "word") {
            return sendImage(response, document.images[Number(wordMedia[1])], mediaRoot(document));
        }
        const slideMedia = route.match(/^\/media\/(\d+)\/(\d+)$/);
        if (slideMedia && document?.kind === "ppt") {
            const imagePath = document.slides[Number(slideMedia[1])]?.imagePaths[Number(slideMedia[2])];
            return sendImage(response, imagePath, mediaRoot(document));
        }
    }

    if (request.method === "POST" && route === "/api/open") {
        try {
            const input = JSON.parse((await readRequestBody(request, 16384)).toString("utf8"));
            if (typeof input.filePath !== "string" || !input.filePath.trim()) {
                return sendJson(response, { error: "filePath 必须是非空字符串" }, 400);
            }
            await openPath(entry, input.filePath);
            return sendJson(response, publicDocument(entry));
        } catch (error) {
            return sendJson(response, { error: error.message }, 422);
        }
    }

    if (request.method === "POST" && route === "/api/close") {
        try {
            await closeDocument(entry);
            return sendJson(response, publicDocument(entry));
        } catch (error) {
            return sendJson(response, { error: error.message }, 422);
        }
    }

    if (request.method === "POST" && route === "/api/upload") {
        let uploadDir = "";
        try {
            const fileName = uploadedFileName(request.headers["x-file-name"]);
            const body = await readRequestBody(request, MAX_FILE_BYTES);
            uploadDir = await mkdtemp(path.join(tmpdir(), "copilot-office-upload-"));
            const uploadPath = path.join(uploadDir, fileName);
            await writeFile(uploadPath, body);
            const finalDir = uploadDir;
            await withLock(entry, () => replaceDocument(entry, uploadPath, { uploadDir: finalDir, displayName: fileName }));
            uploadDir = "";
            return sendJson(response, publicDocument(entry));
        } catch (error) {
            if (uploadDir) await rm(uploadDir, { recursive: true, force: true });
            if (!response.destroyed && !response.writableEnded) sendJson(response, { error: error.message }, 422);
            return;
        }
    }

    sendJson(response, { error: "Not found" }, 404);
}

async function startServer(entry) {
    const server = createServer((request, response) => {
        handleRequest(entry, request, response).catch((error) => {
            if (!response.headersSent) sendJson(response, { error: error.message }, 500);
            else response.destroy(error);
        });
    });
    await new Promise((resolve, reject) => {
        server.once("error", reject);
        server.listen(0, "127.0.0.1", resolve);
    });
    entry.server = server;
    entry.port = server.address().port;
    entry.url = `http://127.0.0.1:${entry.port}/`;
}

function requireDocument(ctx) {
    const entry = instances.get(ctx.instanceId);
    if (!entry?.document) throw new CanvasError("file_not_open", "当前没有打开文件，请先使用 open_file 或在界面中选择文件");
    return entry;
}

function clip(text, maxChars) {
    return { text: text.slice(0, maxChars), truncated: text.length > maxChars };
}

async function documentText(entry, page, maxChars) {
    const document = entry.document;
    if (document.kind === "word") {
        const text = document.paragraphs.map((paragraph) => paragraph.text).filter(Boolean).join("\n");
        if (!text) throw new CanvasError("text_unavailable", "文档没有可提取的正文文字");
        return { scope: "document", ...clip(text, maxChars) };
    }
    if (document.kind === "pdf") {
        if (page && document.pageCount && page > document.pageCount) {
            throw new CanvasError("page_not_found", `PDF 只有 ${document.pageCount} 页`);
        }
        try {
            const text = await extractPdfText(document.pdfPath, page);
            return { scope: page ? `page ${page}` : "document", ...clip(text, maxChars) };
        } catch (error) {
            throw new CanvasError("pdf_text_extraction_failed", error.message);
        }
    }
    if (document.kind === "ppt") {
        if (page) {
            const slide = document.slides[page - 1];
            if (!slide) throw new CanvasError("slide_not_found", `找不到第 ${page} 张幻灯片`);
            return { scope: `slide ${page}`, ...clip(slide.text.join("\n"), maxChars) };
        }
        const text = document.slides
            .map((slide) => `--- 第 ${slide.number} 张 ---\n${slide.text.join("\n")}`)
            .join("\n");
        return { scope: "document", ...clip(text, maxChars) };
    }
    const snapshot = document.snapshot;
    if (page) {
        const sheet = snapshot.sheets[snapshot.sheetOrder[page - 1]];
        if (!sheet) throw new CanvasError("sheet_not_found", `找不到第 ${page} 个工作表`);
        return { scope: `sheet ${page}`, sheetName: sheet.name, ...clip(sheetToText(sheet), maxChars) };
    }
    const text = snapshot.sheetOrder
        .map((id) => `--- ${snapshot.sheets[id].name} ---\n${sheetToText(snapshot.sheets[id])}`)
        .join("\n");
    return { scope: "document", ...clip(text, maxChars) };
}

const emptyInput = { type: "object", properties: {}, additionalProperties: false };

const canvas = createCanvas({
    id: "office-viewer",
    displayName: "Office viewer",
    description: "统一查看 Word (.docx/.docm)、PDF、PowerPoint (.pptx/.pptm) 和 Excel (.xlsx/.xlsm/.xltx) 文件。",
    inputSchema: {
        type: "object",
        properties: { filePath: { type: "string", minLength: 1, description: "要打开的 Office / PDF 文件绝对路径" } },
        additionalProperties: false,
    },
    actions: [
        {
            name: "open_file",
            description: "在当前查看器中打开另一个文件（自动按扩展名识别 Word / PDF / PPT / Excel）。",
            inputSchema: {
                type: "object",
                properties: { filePath: { type: "string", minLength: 1 } },
                required: ["filePath"],
                additionalProperties: false,
            },
            handler: async (ctx) => {
                const entry = instances.get(ctx.instanceId);
                if (!entry) throw new CanvasError("instance_not_found", "查看器实例不存在");
                try {
                    await openPath(entry, ctx.input.filePath);
                    return documentInfo(entry);
                } catch (error) {
                    throw new CanvasError("document_open_failed", error.message);
                }
            },
        },
        {
            name: "refresh",
            description: "重新读取当前文件并返回文件信息。",
            inputSchema: emptyInput,
            handler: async (ctx) => {
                const entry = requireDocument(ctx);
                try {
                    await refreshDocument(entry);
                    return documentInfo(entry);
                } catch (error) {
                    throw new CanvasError("document_refresh_failed", error.message);
                }
            },
        },
        {
            name: "get_file",
            description: "返回当前打开文件的类型、路径、页数 / 幻灯片数 / 工作表数。",
            inputSchema: emptyInput,
            handler: async (ctx) => {
                const entry = instances.get(ctx.instanceId);
                return entry ? documentInfo(entry) : { empty: true };
            },
        },
        {
            name: "get_text",
            description: "提取当前文件文字。page 对 PDF 是页码、PPT 是幻灯片序号、Excel 是工作表序号（均从 1 开始）；省略则读取整个文件。",
            inputSchema: {
                type: "object",
                properties: {
                    page: { type: "integer", minimum: 1 },
                    maxChars: { type: "integer", minimum: 1, maximum: 100000 },
                },
                additionalProperties: false,
            },
            handler: async (ctx) => {
                const entry = requireDocument(ctx);
                const result = await documentText(entry, ctx.input.page, ctx.input.maxChars ?? 20000);
                return { kind: entry.document.kind, fileName: entry.displayName, ...result };
            },
        },
    ],
    open: async (ctx) => {
        let entry = instances.get(ctx.instanceId);
        if (!entry) {
            entry = {
                filePath: "",
                displayPath: "",
                displayName: "",
                document: null,
                uploadDir: "",
                revision: 0,
                lock: Promise.resolve(),
                clients: new Set(),
            };
            await startServer(entry);
            instances.set(ctx.instanceId, entry);
        }
        const requestedPath = typeof ctx.input?.filePath === "string" ? ctx.input.filePath.trim() : "";
        if (!requestedPath) {
            return entry.document
                ? { title: entry.displayName, url: entry.url, status: `${entry.document.pageCount ?? "?"} ${pageUnit(entry.document.kind)}` }
                : { title: "Office viewer", url: entry.url, status: "等待文件" };
        }
        try {
            const document = await openPath(entry, requestedPath);
            return {
                title: path.basename(document.filePath),
                url: entry.url,
                status: `${document.pageCount ?? "?"} ${pageUnit(document.kind)}`,
            };
        } catch (error) {
            throw new CanvasError("document_open_failed", error.message);
        }
    },
    onClose: async (ctx) => {
        const entry = instances.get(ctx.instanceId);
        if (!entry) return;
        instances.delete(ctx.instanceId);
        for (const client of entry.clients) client.end();
        await new Promise((resolve) => {
            entry.server.close(resolve);
            entry.server.closeAllConnections();
        });
        await disposeDocument(entry.document);
        await clearUpload(entry);
    },
});

await joinSession({ canvases: [canvas] });
