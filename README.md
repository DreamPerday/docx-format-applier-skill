# DOCX Format Applier

一键修正 DOCX 文档格式，支持从模板自动提取页面设置、样式定义和页眉格式，特别适用于毕业论文排版。

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

### 1. 安装依赖

- .NET SDK 8.0+
- Python 3.8+

### 2. 创建 C# 项目

创建 `FormatFix.csproj`：

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.2.0" />
  </ItemGroup>
</Project>
```

### 3. 编写修正代码

参考 `SKILL.md` 中的 C# 代码示例，将需要的功能代码写入 `Program.cs`。

### 4. 运行

```powershell
dotnet restore
dotnet run
```

完整代码示例和详细说明见 [SKILL.md](SKILL.md)。

## 依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| .NET SDK | 8.0+ | C# 格式修正核心 |
| DocumentFormat.OpenXml | 3.2.0 | OpenXML 文档处理 |
| Python | 3.8+ | 模板提取与验证（标准库，无需 pip） |

## 许可证

MIT