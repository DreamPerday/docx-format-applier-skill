# DOCX Format Applier Skill

Trae CN 技能：一键修正 DOCX 文档格式，支持毕业论文排版。

## 功能

- 从模板文档自动提取页面设置、样式定义、页眉格式
- 字体字号修正（Normal 正文字号、Key words 关键词字号）
- 页边距、纸张方向修正
- 页眉/页脚格式应用
- 目录(TOC)样式联动
- OMML 公式与 MathType 公式保护
- 文本框/形状文字格式化
- 脚注/尾注格式处理
- 编号列表缩进适配
- 批注与修订标记保留

## 使用方式

### 在 Trae CN 中直接调用

本技能已上传至 GitHub，可导入 Trae CN 直接使用。在对话中描述格式修正需求，技能描述会自动匹配并调用。

### 手动运行 C# 程序

1. 安装 .NET SDK 8.0+ 和 Python 3.8+
2. 从 `SKILL.md` 中复制 `FormatFix.csproj` 和 `Program.cs` 代码
3. 在项目目录执行：

```powershell
dotnet restore
dotnet run
```

### 作为 Trae CN 技能安装

1. 将 `SKILL.md` 放入 `.trae/skills/docx-format-applier/` 目录
2. Trae CN 会自动识别并加载该技能

## 依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| .NET SDK | 8.0+ | C# 格式修正核心 |
| DocumentFormat.OpenXml | 3.2.0 | OpenXML 文档处理 |
| Python | 3.8+ | 模板提取与验证（标准库，无需 pip） |

## 许可证

MIT