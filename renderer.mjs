import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import presetLocale from "@univerjs/preset-sheets-core/locales/en-US";

let current = null;

function mount(containerId, data) {
    dispose();
    const { univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: { [LocaleType.EN_US]: mergeLocales(presetLocale) },
        presets: [UniverSheetsCorePreset({ container: containerId })],
    });
    univerAPI.createWorkbook(data);
    current = univerAPI;
    return univerAPI;
}

function dispose() {
    if (!current) return;
    try { current.dispose(); } catch { /* ignore */ }
    current = null;
}

function sheets() {
    const workbook = current?.getActiveWorkbook();
    if (!workbook) return { ids: [], active: -1 };
    const list = workbook.getSheets();
    const activeId = workbook.getActiveSheet()?.getSheetId();
    return {
        ids: list.map((sheet) => sheet.getSheetId()),
        names: list.map((sheet) => sheet.getSheetName()),
        active: list.findIndex((sheet) => sheet.getSheetId() === activeId),
    };
}

function activate(index) {
    const workbook = current?.getActiveWorkbook();
    const target = workbook?.getSheets()[index];
    if (target) workbook.setActiveSheet(target);
}

globalThis.OfficeExcel = { mount, dispose, sheets, activate };
