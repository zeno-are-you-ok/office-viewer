# Office Viewer

GitHub Copilot canvas 扩展，用于在本机查看 Word、PDF、PowerPoint 和 Excel 文件，并从当前文档提取文本。

## 支持的格式

| 类型 | 格式 | 预览方式 |
| --- | --- | --- |
| Word | `.docx`、`.docm` | Microsoft Word 可用时还原分页；否则显示文字和图片 |
| PDF | `.pdf` | PDF.js |
| PowerPoint | `.pptx`、`.pptm` | Microsoft PowerPoint 可用时渲染幻灯片；否则解析文字和图片 |
| Excel | `.xlsx`、`.xlsm`、`.xltx` | ExcelJS 读取，Univer 渲染 |

文件大小上限为 100 MB。查看器只读，不会保存对文档的修改。

## 安装

需要支持 Copilot 扩展的 GitHub Copilot 应用、Node.js 20 或更高版本。Windows 上还需要 PowerShell；Word 和 PowerPoint 的页面还原需要安装相应的 Microsoft Office 桌面应用。PDF 的页数统计和文字提取需要 `pdftotext`（Poppler）。

将仓库克隆到 Copilot 扩展目录并生成本地运行资源：

```powershell
git clone https://github.com/zeno-are-you-ok/office-viewer.git "$env:USERPROFILE\.copilot\extensions\office-viewer"
Set-Location "$env:USERPROFILE\.copilot\extensions\office-viewer"
npm ci
npm run build
```

完成后重启 GitHub Copilot 应用。PDF.js 和 Excel 渲染器由构建步骤生成到 `vendor/`，不会提交到仓库。

## 隐私

文档在本机处理。查看器的临时 HTTP 服务仅监听 `127.0.0.1`；通过界面上传的文件会在查看器关闭或更换文档后清理。最近打开记录保存在本机扩展目录的 `artifacts/history.json`，该目录已加入 `.gitignore`。

## 许可证

本项目采用 MIT 许可证，见 [LICENSE](LICENSE)。第三方依赖及构建生成资源受其各自许可证约束；使用或再分发时请一并遵循上游条款。
