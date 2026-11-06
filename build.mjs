import { cp, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { copyFile } from "node:fs/promises";
import { rollup } from "@rollup/wasm-node";
import commonjs from "@rollup/plugin-commonjs";
import nodeResolve from "@rollup/plugin-node-resolve";

await mkdir(new URL("./vendor/", import.meta.url), { recursive: true });
const bundle = await rollup({
    input: fileURLToPath(new URL("./renderer.mjs", import.meta.url)),
    plugins: [nodeResolve({ browser: true }), commonjs()],
});
await bundle.write({
    file: fileURLToPath(new URL("./vendor/excel-renderer.js", import.meta.url)),
    format: "iife",
    name: "UniverSheetsViewer",
    inlineDynamicImports: true,
});
await bundle.close();
await copyFile(
    new URL("./node_modules/@univerjs/preset-sheets-core/lib/index.css", import.meta.url),
    new URL("./vendor/excel-renderer.css", import.meta.url),
);

// pdf.js 使用 legacy 构建以兼容较旧的 WebView，仅复制运行时需要的文件
const pdfjsRoot = new URL("./node_modules/pdfjs-dist/", import.meta.url);
const pdfjsOut = new URL("./vendor/pdfjs/", import.meta.url);
await mkdir(pdfjsOut, { recursive: true });
for (const file of ["pdf.min.mjs", "pdf.worker.min.mjs"]) {
    await copyFile(new URL(`./legacy/build/${file}`, pdfjsRoot), new URL(`./${file}`, pdfjsOut));
}
await cp(new URL("./cmaps/", pdfjsRoot), new URL("./cmaps/", pdfjsOut), { recursive: true });
await cp(new URL("./standard_fonts/", pdfjsRoot), new URL("./standard_fonts/", pdfjsOut), { recursive: true });
