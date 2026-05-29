# office-docs

[English](README.md) | **中文**

一个技能库，用于指导 AI 编程助手在 Windows 上**读写 Microsoft Office 文件**——包括 Word（`.docx`/`.doc`）和 Excel（`.xlsx`/`.xls`/`.csv`），主要通过 **PowerShell COM 自动化**来操控已安装的 Office 应用程序。

项目结构参考了 [obra/superpowers](https://github.com/obra/superpowers)：每个技能都是一个独立的 `SKILL.md`，附带明确的触发条件，助手按需加载。

## 为什么选择 COM 自动化？

通过 COM 直接操控 Word/Excel 应用程序，能够获得像素级精确的渲染、真正的公式重算、原生的 PDF 导出，以及对旧版 `.doc`/`.xls` 格式的完整支持——输出效果与用户在 Office 中看到的一模一样。它的局限性在于：**仅支持 Windows 平台、必须安装 Office 并在交互式桌面会话中运行，且需要谨慎管理 COM 对象的生命周期**。如果你的使用场景无法满足这些条件（如 CI 环境、无头服务器、批量数据处理），技能文档会引导你使用基于类库的替代方案。

## 技能列表

| 技能 | 适用场景 |
|------|----------|
| [`office-docs-overview`](skills/office-docs-overview/SKILL.md) | 如何选择处理 Office 文件的方式；COM 还是类库；安全前提 |
| [`word-com-powershell`](skills/word-com-powershell/SKILL.md) | 读取/写入/编辑/生成 Word 文档，或导出为 PDF |
| [`excel-com-powershell`](skills/excel-com-powershell/SKILL.md) | 读写单元格、区域、公式、工作表；Excel 转 PDF |
| [`office-com-cleanup`](skills/office-com-cleanup/SKILL.md) | 处理泄漏的 `WINWORD.EXE`/`EXCEL.EXE`（或 WPS 下的 `wps.exe`/`et.exe`）进程、被锁文件、脚本挂起、COM 释放模式 |

### 附带脚本

每个自动化技能都在 `skills/<skill>/scripts/` 下附带了开箱即用的 PowerShell 脚本，全部实现了完整的「打开 → try/finally → 释放」生命周期：

- **Word：** `Get-WordText.ps1`、`Set-WordReplace.ps1`、`Convert-WordToPdf.ps1`、`Invoke-WordTemplate.ps1`
- **Excel：** `Get-ExcelData.ps1`、`Set-ExcelData.ps1`、`Convert-ExcelToPdf.ps1`

使用示例：

```powershell
# 提取 Word 文档中的全部文本和表格
.\skills\word-com-powershell\scripts\Get-WordText.ps1 -Path .\report.docx -IncludeTables

# 填充 {{placeholder}} 模板
.\skills\word-com-powershell\scripts\Invoke-WordTemplate.ps1 `
    -TemplatePath .\invoice.docx -OutPath .\out.docx `
    -Values @{ CLIENT = 'Acme'; TOTAL = '$1,200' }

# 将 CSV 数据写入格式化后的 .xlsx
.\skills\excel-com-powershell\scripts\Set-ExcelData.ps1 -CsvPath .\people.csv -OutPath .\people.xlsx
```

## 环境要求

- 安装了 Microsoft Office（Word/Excel）**或 WPS Office** 的 Windows 系统。
  WPS 会将自身注册到 Microsoft 的 COM ProgID（`Word.Application` /
  `Excel.Application`）下，因此同样的脚本也能运行——但 WPS 只实现了 COM API
  的一个**子集**，且其后台进程是 `wps.exe` / `et.exe` 而非
  `WINWORD.EXE` / `EXCEL.EXE`。详见 [WPS 兼容性说明](#wps-兼容性说明)。
- Windows PowerShell 5.1 或 PowerShell 7+
- **交互式**桌面会话（Session 0 / 服务账户下不支持 Office COM 自动化）

## 如何使用这些技能

### 配合支持 skills/SKILL.md 约定的 AI 助手

将本仓库配置到你的 AI 助手中（或以插件形式安装）。技能会从 `skills/` 目录自动发现；当某个 `SKILL.md` 的 `description` 触发条件与当前任务匹配时，助手会自动读取相应文件。详见 [AGENTS.md](AGENTS.md)。

### Claude Code（插件）

本仓库通过 `.claude-plugin/` 打包为 Claude Code 插件。从你的 fork 注册为 marketplace 后安装：

```bash
/plugin marketplace add G-Yong/office-docs
/plugin install office-docs@office-docs
```

###WPS 兼容性说明

[WPS Office](https://www.wps.com)（金山）会将自身作为 COM 兼容层注册到
Microsoft 的 ProgID `Word.Application` 和 `Excel.Application` 下。当系统只安装
了 WPS 而没有 Microsoft Office 时，本仓库的脚本实际操控的是 WPS Writer /
Spreadsheets，而非 Word / Excel。

**可正常使用的功能：** 基本的打开/读取/写入、`SaveAs`、PDF 导出、大部分常用属性和方法。

**注意事项：**
- WPS 只实现了 Office COM API 的一个**子集**——部分高级成员、枚举值和格式功能可能缺失。
- 后台进程是 `wps.exe`（文字）、`et.exe`（表格）和 `wpp.exe`（演示），而非
  `WINWORD.EXE` / `EXCEL.EXE`。仅检查后者会漏掉 WPS 的残留进程。
- WPS 也提供了自己的原生 ProgID：`KWPS.Application`、`KET.Application`。

确认当前操控的是哪个应用：

```powershell
$app = New-Object -ComObject Word.Application
$app.Path   # 路径含 Kingsoft\WPS Office\... → 说明是 WPS
```

##  手动使用

技能文件其实就是 Markdown 文档。打开对应的 `SKILL.md`，参照其中的快速参考和模式说明直接运行附带脚本即可。

## 安全注意事项

- 使用 COM 时务必使用**绝对路径**——Office 会以自身的工作目录为准来解析相对路径。
- 务必在 `finally` 块中**退出并释放** COM 对象，否则残留的 Office 后台进程会导致资源泄漏和文件锁定。参见 `office-com-cleanup`。
- 请勿以服务账户身份或在 CI 环境中运行这些自动化脚本——请改用类库方案（概述技能中已说明替代方案）。

## 许可证

[MIT](LICENSE)
