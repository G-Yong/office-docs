# office-docs

[English](README.md) | **中文**

一个技能库，教会编码智能体如何在 Windows 上**读写 Microsoft Office 文件** —— Word（`.docx`/`.doc`）和 Excel（`.xlsx`/`.xls`/`.csv`），主要通过 **PowerShell COM 自动化**驱动已安装的 Office 应用程序。

灵感来源于 [obra/superpowers](https://github.com/obra/superpowers) 的结构设计：每个技能都是一个独立的 `SKILL.md`，具有明确的触发条件，智能体按需加载。

## 为什么选择 COM 自动化？

通过 COM 驱动真实的 Word/Excel 应用程序，可以获得精确的渲染效果、真实的公式重算、原生 PDF 导出以及对旧版 `.doc`/`.xls` 格式的支持——输出结果与用户实际看到的完全一致。代价是它**仅支持 Windows、需要安装 Office 并在交互式桌面会话中运行，且必须认真处理 COM 资源释放**。若这些限制不适用（如 CI、无头环境、批量数据处理），技能文件中会指引你使用基于库的替代方案。

## 技能列表

| 技能 | 适用场景 |
|------|----------|
| [`office-docs-overview`](skills/office-docs-overview/SKILL.md) | 决定如何处理 Office 文件；COM 还是库；安全前提条件 |
| [`word-com-powershell`](skills/word-com-powershell/SKILL.md) | 读取/写入/编辑/生成 Word 文档，或导出为 PDF |
| [`excel-com-powershell`](skills/excel-com-powershell/SKILL.md) | 读写单元格、区域、公式、工作表；Excel 转 PDF |
| [`office-com-cleanup`](skills/office-com-cleanup/SKILL.md) | 处理泄漏的 `WINWORD.EXE`/`EXCEL.EXE`、被锁文件、脚本挂起、COM 释放模式 |

### 内置脚本

每个自动化技能在 `skills/<skill>/scripts/` 下都附带开箱即用的 PowerShell 脚本，全部实现了完整的「打开 → try/finally → 释放」生命周期：

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

# 将 CSV 写入格式化的 .xlsx
.\skills\excel-com-powershell\scripts\Set-ExcelData.ps1 -CsvPath .\people.csv -OutPath .\people.xlsx
```

## 环境要求

- 已安装 Microsoft Office（Word/Excel）的 Windows 系统
- Windows PowerShell 5.1 或 PowerShell 7+
- **交互式**桌面会话（Session 0 / 服务账户下不支持 Office COM 自动化）

## 仓库结构

```
office-docs/
  .claude-plugin/
    plugin.json          # Claude Code 插件清单
    marketplace.json     # 市场入口
  skills/
    office-docs-overview/
      SKILL.md
      references/non-com-alternatives.md
    word-com-powershell/
      SKILL.md
      scripts/*.ps1
    excel-com-powershell/
      SKILL.md
      scripts/*.ps1
    office-com-cleanup/
      SKILL.md
  AGENTS.md              # 告知智能体如何发现并使用这些技能
  package.json
  LICENSE
```

## 如何使用这些技能

### 支持 skills/SKILL.md 约定的智能体

将此仓库指向你的智能体（或作为插件安装）。技能从 `skills/` 目录自动发现；当某个 `SKILL.md` 的 `description` 触发条件匹配当前任务时，智能体会读取该文件。详见 [AGENTS.md](AGENTS.md)。

### Claude Code（插件）

本仓库通过 `.claude-plugin/` 打包为 Claude Code 插件。从你的 fork 注册为 marketplace 并安装：

```bash
/plugin marketplace add G-Yong/office-docs
/plugin install office-docs@office-docs
```

### 手动使用

技能文件本质上只是 Markdown。打开对应的 `SKILL.md`，按照其快速参考和模式操作，直接运行内置脚本即可。

## 安全注意事项

- 使用 COM 时务必使用**绝对路径** —— Office 会以自身的工作目录解析相对路径。
- 务必在 `finally` 块中**退出并释放** COM 对象，否则隐藏的 Office 进程会泄漏并锁定文件。参见 `office-com-cleanup`。
- 不要以服务身份或在 CI 中运行此自动化 —— 请改用库（替代方案已在概述技能中说明）。

## 许可证

[MIT](LICENSE)
