import { execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, open, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

export const MAX_FILE_BYTES = 100 * 1024 * 1024;

export const KINDS = {
    word: { label: "Word", extensions: [".docx", ".docm"] },
    pdf: { label: "PDF", extensions: [".pdf"] },
    ppt: { label: "PowerPoint", extensions: [".pptx", ".pptm"] },
    excel: { label: "Excel", extensions: [".xlsx", ".xlsm", ".xltx"] },
};

export const ACCEPT_EXTENSIONS = Object.values(KINDS).flatMap((kind) => kind.extensions);
export const SUPPORTED_TEXT = ACCEPT_EXTENSIONS.join("、");

export function kindOf(filePath) {
    const extension = path.extname(filePath).toLowerCase();
    return Object.keys(KINDS).find((kind) => KINDS[kind].extensions.includes(extension)) ?? "";
}

export function isWithin(directory, filePath) {
    const relative = path.relative(directory, filePath);
    return relative === ""
        || (relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
}

function decodeXml(value) {
    return value
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
        .replace(/&#([0-9]+);/g, (_, code) => String.fromCodePoint(Number(code)))
        .replace(/&amp;/g, "&");
}

function parseAttributes(tag) {
    const result = {};
    for (const match of tag.matchAll(/([:\w-]+)\s*=\s*(["'])(.*?)\2/g)) {
        result[match[1]] = decodeXml(match[3]);
    }
    return result;
}

function readRelationships(xml) {
    const result = new Map();
    for (const match of xml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
        const attributes = parseAttributes(match[0]);
        if (attributes.Id && attributes.Target) result.set(attributes.Id, attributes.Target);
    }
    return result;
}

function runPowerShell(script, env, timeout = 120000) {
    return execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ], {
        windowsHide: true,
        timeout,
        env: { ...process.env, ...env },
    });
}

async function validateFile(filePath, kind) {
    const extension = path.extname(filePath).toLowerCase();
    if (!KINDS[kind]?.extensions.includes(extension)) {
        throw new Error(`不支持的文件类型，当前支持 ${SUPPORTED_TEXT}`);
    }
    const info = await stat(filePath);
    if (!info.isFile()) throw new Error("指定路径不是文件");
    if (info.size > MAX_FILE_BYTES) throw new Error("文件大小不能超过 100 MB");

    const expected = Buffer.from(kind === "pdf" ? "%PDF-" : "PK");
    const handle = await open(filePath, "r");
    try {
        const signature = Buffer.alloc(expected.length);
        const { bytesRead } = await handle.read(signature, 0, signature.length, 0);
        if (bytesRead < expected.length || !signature.equals(expected)) {
            throw new Error("文件内容与扩展名不匹配，或文件已损坏");
        }
    } finally {
        await handle.close();
    }
    return info;
}

async function extractZip(filePath, destination) {
    const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($env:OFFICE_SOURCE)
try {
    if ($archive.Entries.Count -gt 10000) { throw '文档包含过多压缩条目' }
    $total = 0L
    foreach ($entry in $archive.Entries) {
        $total += $entry.Length
        if ($total -gt 524288000) { throw '解压后的文档不能超过 500 MB' }
    }
} finally {
    $archive.Dispose()
}
[System.IO.Compression.ZipFile]::ExtractToDirectory($env:OFFICE_SOURCE, $env:OFFICE_DEST)
`;
    await runPowerShell(script, { OFFICE_SOURCE: filePath, OFFICE_DEST: destination });
}

function countPdfPages(filePath) {
    return new Promise((resolve, reject) => {
        let child;
        try {
            child = spawn("pdftotext", ["-enc", "UTF-8", filePath, "-"], {
                windowsHide: true,
                stdio: ["ignore", "pipe", "pipe"],
            });
        } catch (error) {
            reject(error);
            return;
        }
        let pageBreaks = 0;
        let byteCount = 0;
        let lastByte = -1;
        let stderr = "";
        child.stdout.on("data", (chunk) => {
            byteCount += chunk.length;
            lastByte = chunk[chunk.length - 1];
            for (const byte of chunk) if (byte === 12) pageBreaks++;
        });
        child.stderr.on("data", (chunk) => {
            if (stderr.length < 4096) stderr += chunk.toString("utf8").slice(0, 4096 - stderr.length);
        });
        child.once("error", reject);
        child.once("close", (code) => {
            if (code !== 0) {
                reject(new Error(stderr.trim() || `pdftotext 退出码为 ${code}`));
                return;
            }
            resolve(pageBreaks + (byteCount > 0 && lastByte !== 12 ? 1 : 0));
        });
    });
}

export async function extractPdfText(filePath, page) {
    const args = page ? ["-f", String(page), "-l", String(page)] : [];
    const { stdout } = await execFileAsync("pdftotext", [...args, "-enc", "UTF-8", filePath, "-"], {
        windowsHide: true,
        timeout: 60000,
        maxBuffer: 16 * 1024 * 1024,
    });
    return stdout.replace(/\f+$/, "").trim();
}

async function loadPdf(filePath, info) {
    let pageCount = null;
    let warning = "";
    try {
        pageCount = await countPdfPages(filePath);
    } catch (error) {
        warning = `无法统计 PDF 页数或提取文字：${error.message}`;
    }
    return { pdfPath: filePath, pageCount, warning, cleanupDirs: [] };
}

async function loadWord(filePath) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "copilot-office-word-"));
    try {
        const extractDir = path.join(tempDir, "x");
        await extractZip(filePath, extractDir);
        const documentPath = path.join(extractDir, "word", "document.xml");
        if ((await stat(documentPath)).size > 25 * 1024 * 1024) {
            throw new Error("Word 文档正文 XML 超过 25 MB，无法安全解析");
        }
        const documentXml = await readFile(documentPath, "utf8");
        const relsPath = path.join(extractDir, "word", "_rels", "document.xml.rels");
        const rels = existsSync(relsPath)
            ? readRelationships(await readFile(relsPath, "utf8"))
            : new Map();
        const images = [];
        const imageIndexes = new Map();

        const paragraphs = [...documentXml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)]
            .map((match) => {
                const xml = match[1];
                const text = [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
                    .map((run) => decodeXml(run[1]))
                    .join("");
                const styleMatch = xml.match(/<w:pStyle\b[^>]*w:val=["']([^"']+)["']/);
                const paragraphImages = [];
                for (const imageMatch of xml.matchAll(/<a:blip\b[^>]*r:embed=["']([^"']+)["']/g)) {
                    const target = rels.get(imageMatch[1]);
                    if (!target || /^https?:/i.test(target)) continue;
                    const imagePath = path.resolve(path.dirname(documentPath), target.replace(/[\\/]/g, path.sep));
                    if (!isWithin(extractDir, imagePath) || !existsSync(imagePath)) continue;
                    let index = imageIndexes.get(imagePath);
                    if (index === undefined) {
                        index = images.length;
                        images.push(imagePath);
                        imageIndexes.set(imagePath, index);
                    }
                    paragraphImages.push(index);
                }
                return { text, style: styleMatch ? decodeXml(styleMatch[1]) : "", images: paragraphImages };
            })
            .filter((paragraph) => paragraph.text.trim() || paragraph.images.length > 0);

        let pdfPath = path.join(tempDir, "preview.pdf");
        let warning = "";
        try {
            await renderWordAsPdf(filePath, pdfPath);
        } catch (error) {
            pdfPath = "";
            warning = `未能使用 Microsoft Word 还原页面版式，已切换为文字预览：${error.message}`;
        }
        let pageCount = null;
        if (pdfPath) {
            try {
                pageCount = await countPdfPages(pdfPath);
            } catch (error) {
                warning = `${warning ? `${warning}；` : ""}无法统计预览页数：${error.message}`;
            }
        }
        return { pdfPath, pageCount, paragraphs, images, mediaRoot: extractDir, warning, cleanupDirs: [tempDir] };
    } catch (error) {
        await rm(tempDir, { recursive: true, force: true });
        throw new Error(`无法读取 Word 文档：${error.message}`);
    }
}

async function renderWordAsPdf(filePath, outputPath) {
    const script = `
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($env:OFFICE_SOURCE, $false, $true, $false)
    $document.ExportAsFixedFormat($env:OFFICE_DEST, 17)
} finally {
    if ($document) {
        try { $document.Close(0) } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null } catch {}
    }
    if ($word) {
        try { $word.Quit(0) } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}`;
    await runPowerShell(script, { OFFICE_SOURCE: filePath, OFFICE_DEST: outputPath });
    if (!existsSync(outputPath)) throw new Error("Microsoft Word 未生成 PDF 预览");
}

async function parsePresentation(filePath, tempDir) {
    const extractDir = path.join(tempDir, "x");
    await extractZip(filePath, extractDir);
    const pptDir = path.join(extractDir, "ppt");
    const presentationXml = await readFile(path.join(pptDir, "presentation.xml"), "utf8");
    const presentationRels = readRelationships(
        await readFile(path.join(pptDir, "_rels", "presentation.xml.rels"), "utf8"),
    );
    const slideIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*r:id\s*=\s*["']([^"']+)["']/g)]
        .map((match) => match[1]);
    const slides = [];
    for (const [index, slideId] of slideIds.entries()) {
        const target = presentationRels.get(slideId);
        if (!target) continue;
        const slidePath = path.resolve(pptDir, target.replace(/^\/+/, ""));
        if (!isWithin(extractDir, slidePath) || !existsSync(slidePath)) continue;
        const slideXml = await readFile(slidePath, "utf8");
        const relsPath = path.join(path.dirname(slidePath), "_rels", `${path.basename(slidePath)}.rels`);
        const slideRels = existsSync(relsPath) ? readRelationships(await readFile(relsPath, "utf8")) : new Map();
        const text = [...slideXml.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g)]
            .map((paragraph) => [...paragraph[1].matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)]
                .map((run) => decodeXml(run[1]))
                .join(""))
            .filter(Boolean);
        const imagePaths = [];
        for (const match of slideXml.matchAll(/<a:blip\b[^>]*r:embed\s*=\s*["']([^"']+)["']/g)) {
            const targetPath = slideRels.get(match[1]);
            if (!targetPath || /^https?:/i.test(targetPath)) continue;
            const imagePath = path.resolve(path.dirname(slidePath), targetPath.replace(/^\/+/, ""));
            if (isWithin(extractDir, imagePath) && existsSync(imagePath)) imagePaths.push(imagePath);
        }
        slides.push({ number: index + 1, text, imagePaths });
    }
    if (slides.length === 0) throw new Error("演示文稿中没有可读取的幻灯片");
    return { slides, mediaRoot: extractDir };
}

async function renderPowerPoint(filePath, outputDir) {
    const script = `
$ErrorActionPreference = 'Stop'
$powerpoint = $null
$presentation = $null
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerpoint.Presentations.Open($env:OFFICE_SOURCE, -1, 0, 0)
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $width = 1600
    $height = [Math]::Max(1, [Math]::Round($width * $slideHeight / $slideWidth))
    $index = 1
    foreach ($slide in $presentation.Slides) {
        $target = Join-Path $env:OFFICE_DEST ("slide-{0:D4}.png" -f $index)
        $slide.Export($target, "PNG", $width, $height)
        $index++
    }
} finally {
    if ($presentation) {
        try { $presentation.Close() } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
    }
    if ($powerpoint) {
        try { $powerpoint.Quit() } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}`;
    await mkdir(outputDir, { recursive: true });
    await runPowerShell(script, { OFFICE_SOURCE: filePath, OFFICE_DEST: outputDir });
    const files = (await readdir(outputDir)).filter((name) => name.toLowerCase().endsWith(".png")).sort();
    if (files.length === 0) throw new Error("PowerPoint 没有导出任何幻灯片");
    return files.map((name) => path.join(outputDir, name));
}

async function loadPpt(filePath) {
    const tempDir = await mkdtemp(path.join(tmpdir(), "copilot-office-ppt-"));
    try {
        const parsed = await parsePresentation(filePath, tempDir);
        let slides = parsed.slides;
        let rendered = false;
        let warning = "";
        let mediaRoot = parsed.mediaRoot;
        try {
            const renderDir = path.join(tempDir, "render");
            const images = await renderPowerPoint(filePath, renderDir);
            slides = images.map((imagePath, index) => ({
                number: index + 1,
                text: parsed.slides[index]?.text ?? [],
                imagePaths: [imagePath],
            }));
            rendered = true;
            mediaRoot = tempDir;
        } catch (error) {
            warning = `PowerPoint 原生渲染失败，已切换为文本解析模式：${error.message}`;
        }
        return { slides, rendered, pageCount: slides.length, mediaRoot, warning, cleanupDirs: [tempDir] };
    } catch (error) {
        await rm(tempDir, { recursive: true, force: true });
        throw new Error(`无法读取 PowerPoint 文件：${error.message}`);
    }
}

const borderStyleMap = {
    thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
    mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};

function toColor(color) {
    if (!color) return undefined;
    if (color.argb) return `#${color.argb.length === 8 ? color.argb.slice(2) : color.argb}`;
    if (color.rgb) return `#${color.rgb}`;
    return undefined;
}

function toBorderSide(side) {
    if (!side?.style) return undefined;
    return { s: borderStyleMap[side.style] ?? 1, cl: { rgb: toColor(side.color) ?? "#000000" } };
}

function toStyle(cell) {
    const style = {};
    const { font, fill, alignment, border } = cell;
    if (font?.name) style.ff = font.name;
    if (font?.size) style.fs = font.size;
    if (font?.bold) style.bl = 1;
    if (font?.italic) style.it = 1;
    if (font?.color) style.cl = { rgb: toColor(font.color) };
    if (fill?.type === "pattern" && fill.fgColor) {
        const color = toColor(fill.fgColor);
        if (color) style.bg = { rgb: color };
    }
    if (alignment?.horizontal) style.ht = { left: 1, center: 2, right: 3, justify: 4 }[alignment.horizontal] ?? 1;
    if (alignment?.vertical) style.vt = { top: 1, middle: 2, bottom: 3 }[alignment.vertical] ?? 1;
    if (alignment?.wrapText) style.tb = 3;
    if (cell.numFmt) style.n = { pattern: cell.numFmt };
    const sides = {
        t: toBorderSide(border?.top),
        r: toBorderSide(border?.right),
        b: toBorderSide(border?.bottom),
        l: toBorderSide(border?.left),
    };
    if (Object.values(sides).some(Boolean)) style.bd = sides;
    return Object.keys(style).length ? style : undefined;
}

function toCellData(cell) {
    const data = {};
    const value = cell.value;
    if (value && typeof value === "object" && "formula" in value) {
        data.f = `=${value.formula}`;
        if (value.result !== undefined && value.result !== null && typeof value.result !== "object") {
            data.v = value.result;
        }
    } else if (value instanceof Date) {
        data.v = value.toISOString().slice(0, 10);
    } else if (value !== null && value !== undefined && typeof value !== "object") {
        data.v = value;
    } else if (value && typeof value === "object" && "text" in value) {
        data.v = value.text;
    } else if (value && typeof value === "object" && Array.isArray(value.richText)) {
        data.v = value.richText.map((part) => part.text).join("");
    }
    const style = toStyle(cell);
    if (style) data.s = style;
    return Object.keys(data).length ? data : undefined;
}

function emptySheet(id, name) {
    return {
        id,
        name,
        tabColor: "",
        hidden: 0,
        freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
        rowCount: 100,
        columnCount: 26,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 73,
        defaultRowHeight: 20,
        mergeData: [],
        cellData: {},
        rowData: {},
        columnData: {},
        rowHeader: { width: 46 },
        columnHeader: { height: 24 },
        showGridlines: 1,
        rightToLeft: 0,
    };
}

async function loadExcel(filePath) {
    const { default: ExcelJS } = await import("exceljs");
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(filePath);
    } catch (error) {
        throw new Error(`无法读取 Excel 文件：${error.message}`);
    }

    const sheets = {};
    const sheetOrder = [];
    for (const worksheet of workbook.worksheets) {
        const sheetId = `sheet-${worksheet.id}`;
        const sheet = emptySheet(sheetId, worksheet.name);
        let maxColumn = 1;
        let maxRow = 1;
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            maxRow = Math.max(maxRow, rowNumber);
            if (row.height) sheet.rowData[rowNumber - 1] = { h: row.height * 1.3333 };
            row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
                maxColumn = Math.max(maxColumn, columnNumber);
                const data = toCellData(cell);
                if (data) {
                    sheet.cellData[rowNumber - 1] ??= {};
                    sheet.cellData[rowNumber - 1][columnNumber - 1] = data;
                }
            });
        });
        worksheet.columns.forEach((column, index) => {
            if (column.width) sheet.columnData[index] = { w: column.width * 7 };
        });
        sheet.mergeData = (worksheet.model.merges ?? []).map((range) => {
            const [start, end] = range.split(":");
            const startCell = worksheet.getCell(start);
            const endCell = worksheet.getCell(end);
            return {
                startRow: startCell.row - 1,
                endRow: endCell.row - 1,
                startColumn: startCell.col - 1,
                endColumn: endCell.col - 1,
            };
        });
        sheet.hidden = worksheet.state === "hidden" ? 1 : 0;
        sheet.rowCount = Math.max(maxRow, 100);
        sheet.columnCount = Math.max(maxColumn, 26);
        sheets[sheetId] = sheet;
        sheetOrder.push(sheetId);
    }
    if (!sheetOrder.length) {
        sheets["sheet-1"] = emptySheet("sheet-1", "Sheet1");
        sheetOrder.push("sheet-1");
    }
    const snapshot = {
        id: "office-viewer",
        name: path.basename(filePath),
        appVersion: "0.25.1",
        locale: "enUS",
        styles: {},
        sheetOrder,
        sheets,
    };
    return { snapshot, pageCount: sheetOrder.length, warning: "", cleanupDirs: [] };
}

export async function loadDocument(requestedPath) {
    const filePath = path.resolve(requestedPath);
    const kind = kindOf(filePath);
    if (!kind) throw new Error(`不支持的文件类型，当前支持 ${SUPPORTED_TEXT}`);
    const info = await validateFile(filePath, kind);
    const loaders = { pdf: loadPdf, word: loadWord, ppt: loadPpt, excel: loadExcel };
    const loaded = await loaders[kind](filePath, info);
    return {
        kind,
        filePath,
        fileName: path.basename(filePath),
        fileSize: info.size,
        mtimeMs: info.mtimeMs,
        ...loaded,
    };
}

export async function disposeDocument(document) {
    if (!document) return;
    for (const directory of document.cleanupDirs ?? []) {
        await rm(directory, { recursive: true, force: true });
    }
}

function cellText(cell) {
    const value = cell?.v;
    return value === undefined || value === null ? "" : String(value).replace(/[\t\r\n]+/g, " ");
}

export function sheetToText(sheet, maxRows = 2000) {
    const rows = Object.keys(sheet.cellData).map(Number).sort((a, b) => a - b).slice(0, maxRows);
    return rows.map((rowIndex) => {
        const row = sheet.cellData[rowIndex];
        const columns = Object.keys(row).map(Number);
        const width = Math.max(...columns) + 1;
        const cells = Array.from({ length: width }, (_, column) => cellText(row[column]));
        return cells.join("\t");
    }).join("\n");
}
