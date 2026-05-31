---
name: "docx-format-applier"
description: "根据源文档和写作模板自动修正DOCX格式。调用时需先询问用户提供源文档路径和格式模板路径（可选），无模板则手动收集格式要求。仅修改用户指定范围的格式，其余保持源文档原状。含页面设置、页眉页脚、字体字号、分区处理。支持毕业论文专项优化。当用户表达套用模板/修正格式/排版意图时调用。"
---

# DOCX Format Applier

**核心思想**：不硬编码任何格式值。程序自动读取写作模板（.docx）的页面设置、样式定义、页眉格式，然后应用到源文档。

---

## 调用方式（重要）

### 用户输入要求

调用此技能时，必须先向用户收集以下信息：

| 必要信息 | 说明 |
|---------|------|
| **源文档路径** | 需要被修正格式的 .docx 文件路径 |
| **格式模板路径** | 可选。提供写作格式模板 .docx，程序自动从中提取格式设定 |
| **手动格式要求** | 若无模板，则要求用户逐项描述期望格式（纸张、边距、字体字号、行距、页眉等） |
| **修改范围** | 明确告知用户哪些格式会被修改、哪些保持原状 |

### 交互流程

```
1. 询问源文档路径
2. 询问是否有格式模板 → 有则读取模板 / 无则手动收集格式要求
3. 生成格式修改计划（列出将要修改的每一项）
4. 请用户确认计划 → 用户确认后才开始执行
5. 执行格式修正
6. 验证并展示结果
```

### 未提及的格式保持原状

- 如果用户只要求修改页面设置，则字体、行距等保持源文档原样
- 如果用户只要求修改页眉，则正文格式保持源文档原样
- 仅在用户明确提出的格式范围内做修改
- 当提供了模板文档时，默认提取以下信息：
  - 页面设置（纸张、边距、页眉页脚位置）
  - 样式定义（Normal、Heading1-3、Header/Footer 的字体、字号、加粗、对齐、行距、缩进）
  - 页眉内容、字体、对齐方式、边框样式

### 无模板时的格式收集清单

当用户未提供模板文档时，逐一确认以下格式参数（有默认值但用户可以修改）：

```yaml
页面设置:
  纸张大小: A4 (宽11906×高16838)
  上边距: 2.5cm (1587)
  下边距: 2.0cm (1247)
  左边距: 3.0cm (1701)
  右边距: 2.0cm (1134)
  页眉位置: 1.8cm (1020)
  页脚位置: 1.4cm (794)
页眉样式:
  是否启用: 是
  字体: 宋体
  字号: 9pt (sz=18)
  对齐: 居中
  底边线: 有 (Single, Size=4)
正文样式(Normal):
  英文字体: Times New Roman
  中文字体: 宋体
  字号: 10.5pt (sz=21)
  对齐: 两端对齐
  行距: 1.5倍 (line=360)
  首行缩进: 2字符 (firstLine=480)
一级标题:
  字体: 黑体
  字号: 22pt (sz=44)
  加粗: 是
  对齐: 左对齐
  行距: 1.5倍
二级标题:
  字体: 黑体
  字号: 16pt (sz=32)
  加粗: 是
三级标题:
  字体: 黑体
  字号: 15pt (sz=30)
  加粗: 是
```

---

## 一、两种工作模式

```
模式A: 有模板
  源文档 + 格式模板.docx ──→ 自动提取格式 ──→ 应用到源文档 ──→ 验证
                             (页面设置/样式/页眉)

模式B: 无模板(手动输入)
  源文档 + 用户描述格式要求 ──→ 生成配置 ──→ 程序化应用 ──→ 验证
                             (逐一确认参数)

共同原则:
  1. 仅修改用户指定范围的格式，其余保持不变
  2. 不破坏内容（表格、图片、公式的文本不变）
  3. 分步验证，每步确认效果
  4. 样式继承优先，减少run级直接格式化
```

---

## 二、从模板提取格式设定

### 2.1 提取模板的页面设置

使用 Python 从模板文档的第一个分区提取页面设置：

```python
import zipfile, xml.etree.ElementTree as ET

def extract_page_setup(template_path):
    """从模板文档提取页面设置"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(template_path) as z:
        with z.open('word/document.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))
        
        body = root.find('.//w:body', ns)
        
        # 查找所有sectPr（段落内嵌的 + body末尾的）
        sectPrs = []
        for p in body.findall('.//w:p', ns):
            pPr = p.find('w:pPr', ns)
            if pPr is not None:
                sp = pPr.find('w:sectPr', ns)
                if sp is not None:
                    sectPrs.append(sp)
        last_sectPr = None
        for child in body:
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr':
                last_sectPr = child
        if last_sectPr is not None:
            sectPrs.append(last_sectPr)
        
        # 用第一个分区的设置作为模板标准
        tmpl = sectPrs[0] if sectPrs else None
        if tmpl is None:
            return None
        
        ret = {}
        pgSz = tmpl.find('w:pgSz', ns)
        if pgSz is not None:
            ret['PageWidth'] = pgSz.get('{http://...}w', '11906')
            ret['PageHeight'] = pgSz.get('{http://...}h', '16838')
        
        pgMar = tmpl.find('w:pgMar', ns)
        if pgMar is not None:
            for attr in ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter']:
                v = pgMar.get('{http://...}' + attr)
                if v: ret[f'Margin{attr.capitalize()}'] = v
        
        return ret
```

### 2.2 提取模板的样式定义

提取模板文档中所有样式（字体、字号、加粗、对齐、行距、缩进）：

```python
def extract_style_definitions(template_path):
    """提取模板文档的完整样式定义"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    styles = {}
    with zipfile.ZipFile(template_path) as z:
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))
        
        for style in root.findall('.//w:style', ns):
            sid = style.get('{http://...}styleId', '')
            stype = style.get('{http://...}type', '')
            name_el = style.find('w:name', ns)
            sname = name_el.get('{http://...}val', '') if name_el is not None else ''
            
            # 提取字体属性
            rPr = style.find('w:rPr', ns)
            font_info = {}
            if rPr is not None:
                rFonts = rPr.find('w:rFonts', ns)
                if rFonts is not None:
                    for attr in ['ascii', 'eastAsia', 'hAnsi']:
                        v = rFonts.get('{http://...}' + attr)
                        if v: font_info[attr] = v
                sz = rPr.find('w:sz', ns)
                if sz is not None:
                    font_info['sz'] = sz.get('{http://...}val', '')
                font_info['bold'] = rPr.find('w:b', ns) is not None
            
            # 提取段落格式
            pPr_style = style.find('w:pPr', ns)
            pfmt = {}
            if pPr_style is not None:
                jc = pPr_style.find('w:jc', ns)
                if jc is not None:
                    pfmt['jc'] = jc.get('{http://...}val', '')
                spacing = pPr_style.find('w:spacing', ns)
                if spacing is not None:
                    for attr in ['line', 'before', 'after']:
                        v = spacing.get('{http://...}' + attr)
                        if v: pfmt[attr] = v
                ind = pPr_style.find('w:ind', ns)
                if ind is not None:
                    for attr in ['firstLine', 'left', 'right']:
                        v = ind.get('{http://...}' + attr)
                        if v: pfmt[attr] = v
            
            styles[sid] = {
                'name': sname, 'type': stype,
                'font': font_info, 'pfmt': pfmt
            }
    
    # 特别提取关键样式
    critical = {}
    for sid in styles:
        s = styles[sid]
        if s['name'] in ['Normal', 'Heading1', 'Heading2', 'Heading3',
                         'Header', 'Footer', 'BodyText', 'BodyTextIndent',
                         'TOC1', 'TOC2']:
            critical[s['name']] = s
    
    return styles, critical
```

### 2.3 提取模板的页眉格式

```python
def extract_header_format(template_path):
    """从模板文档提取页眉样式（字体、大小、对齐、边框）"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(template_path) as z:
        # 检查styles.xml中的Header样式
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))
            for style in root.findall('.//w:style', ns):
                sid = style.get('{http://...}styleId', '')
                if sid == 'Header':
                    rPr = style.find('w:rPr', ns)
                    pPr = style.find('w:pPr', ns)
                    if rPr is not None:
                        rFonts = rPr.find('w:rFonts', ns)
                        sz = rPr.find('w:sz', ns)
                        print(f'  页眉字体: {rFonts.attrib if rFonts is not None else "(默认)"}')
                        print(f'  页眉字号: sz={sz.get("{http://...}val") if sz is not None else "(默认)"}')
                    if pPr is not None:
                        jc = pPr.find('w:jc', ns)
                        print(f'  页眉对齐: {jc.get("{http://...}val") if jc is not None else "(默认)"}')

        # 检查页眉xml内容中的边框
        with z.open('word/header1.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))
            for p in root.findall('.//w:p', ns):
                pPr = p.find('w:pPr', ns)
                if pPr is not None:
                    pBdr = pPr.find('w:pBdr', ns)
                    if pBdr is not None:
                        bottom = pBdr.find('w:bottom', ns)
                        if bottom is not None:
                            print(f'  页眉底边线: val={bottom.get("{http://...}val")} '
                                  f'size={bottom.get("{http://...}size")}')
```

### 2.4 提取模板的公式样式（OMML）

公式使用的数学样式（math style）定义在 styles.xml 中，需要额外提取：

```python
def extract_math_style(template_path):
    """提取模板文档的数学样式（OMML 公式字体）"""
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    }
    with zipfile.ZipFile(template_path) as z:
        if 'word/styles.xml' not in z.namelist():
            return None
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))

        # 提取 DefaultRunProperties 中的数学字体
        docDefaults = root.find('w:docDefaults/w:runPropertiesDefault/w:rPr', ns)
        if docDefaults is not None:
            rFonts = docDefaults.find('w:rFonts', ns)
            if rFonts is not None:
                print(f'  默认数学字体: {rFonts.attrib}')
    return None
```

### 2.5 提取模板的脚注样式

```python
def extract_footnote_style(template_path):
    """提取模板文档的脚注/尾注样式"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(template_path) as z:
        if 'word/styles.xml' not in z.namelist():
            return None
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))

        footnote_styles = {}
        for style in root.findall('.//w:style', ns):
            sid = style.get('{http://...}styleId', '')
            if sid in ('FootnoteText', 'FootnoteReference', 'EndnoteText', 'EndnoteReference'):
                name_el = style.find('w:name', ns)
                sname = name_el.get('{http://...}val', '') if name_el is not None else ''

                rPr = style.find('w:rPr', ns)
                font_info = {}
                if rPr is not None:
                    rFonts = rPr.find('w:rFonts', ns)
                    if rFonts is not None:
                        for attr in ['ascii', 'eastAsia', 'hAnsi']:
                            v = rFonts.get('{http://...}' + attr)
                            if v: font_info[attr] = v
                    sz = rPr.find('w:sz', ns)
                    if sz is not None:
                        font_info['sz'] = sz.get('{http://...}val', '')

                pPr_style = style.find('w:pPr', ns)
                pfmt = {}
                if pPr_style is not None:
                    jc = pPr_style.find('w:jc', ns)
                    if jc is not None:
                        pfmt['jc'] = jc.get('{http://...}val', '')
                    spacing = pPr_style.find('w:spacing', ns)
                    if spacing is not None:
                        for attr in ['line', 'before', 'after']:
                            v = spacing.get('{http://...}' + attr)
                            if v: pfmt[attr] = v

                footnote_styles[sid] = {
                    'name': sname,
                    'font': font_info,
                    'pfmt': pfmt
                }

        if footnote_styles:
            for sid, info in footnote_styles.items():
                print(f'  {sid} ({info["name"]}): '
                      f'sz={info["font"].get("sz", "(默认)")}, '
                      f'字体={info["font"].get("eastAsia", "(默认)")}')
    return footnote_styles
```

---

## 三、C# OpenXML 动态格式修正程序

### 3.1 项目结构

```xml
<!-- FormatFix.csproj -->
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

### 3.2 从模板读取页面设置

```csharp
/// <summary>
/// 从模板文档的第一个分区提取页面设置。
/// 不硬编码任何值，完全依赖模板。
/// </summary>
static PageSettings ReadPageSettingsFromTemplate(string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var body = tmpl.MainDocumentPart!.Document.Body!;

    // 找到第一个sectPr（可能内嵌在段落中，也可能在body末尾）
    SectionProperties? firstSectPr = null;
    foreach (var p in body.Elements<Paragraph>())
    {
        var pPr = p.GetFirstChild<ParagraphProperties>();
        if (pPr?.GetFirstChild<SectionProperties>() is { } sp)
        { firstSectPr = sp; break; }
    }
    firstSectPr ??= body.LastChild as SectionProperties;

    if (firstSectPr == null)
        throw new Exception("模板文档无分区信息");

    var pgSz = firstSectPr.GetFirstChild<PageSize>();
    var pgMar = firstSectPr.GetFirstChild<PageMargin>();

    return new PageSettings
    {
        PageWidth  = pgSz?.Width  ?? 11906U,
        PageHeight = pgSz?.Height ?? 16838U,
        MarginTop    = (int)(pgMar?.Top    ?? 1587),
        MarginBottom = (int)(pgMar?.Bottom ?? 1247),
        MarginLeft   = pgMar?.Left   ?? 1701,
        MarginRight  = pgMar?.Right  ?? 1134,
        HeaderPos    = pgMar?.Header ?? 1020,
        FooterPos    = pgMar?.Footer ?? 794,
        Gutter       = pgMar?.Gutter ?? 0,
    };
}

class PageSettings
{
    public UInt32Value PageWidth { get; set; }
    public UInt32Value PageHeight { get; set; }
    public int MarginTop { get; set; }
    public int MarginBottom { get; set; }
    public UInt32Value MarginLeft { get; set; }
    public UInt32Value MarginRight { get; set; }
    public UInt32Value HeaderPos { get; set; }
    public UInt32Value FooterPos { get; set; }
    public UInt32Value Gutter { get; set; }
}
```

### 3.3 从模板读取样式定义并应用到源文档

```csharp
/// <summary>
/// 从模板文档复制关键样式到源文档。
/// 遍历模板的styles.xml，提取Normal/Heading1/Heading2/Heading3/Header/Footer等样式，
/// 然后更新源文档中对应样式ID的格式定义。
/// </summary>
static void ApplyStylesFromTemplate(string sourcePath, string templatePath)
{
    using var source = WordprocessingDocument.Open(sourcePath, true);
    using var tmpl = WordprocessingDocument.Open(templatePath, false);

    var sourceStyles = source.MainDocumentPart!.StylePart;
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    if (sourceStyles == null || tmplStyles == null) return;

    // 关键样式名映射：模板中的样式 → 源文档中对应的样式ID
    // 例如，模板有Heading1(样式ID="Heading1")，源文档可能用"2"表示一级标题
    // 方案1：按样式名称匹配
    var tmplStylesById = tmplStyles.Styles.Descendants<Style>()
        .ToDictionary(s => s.StyleId ?? "");

    var sourceStylesByName = sourceStyles.Styles.Descendants<Style>()
        .GroupBy(s => {
            var n = s.GetFirstChild<StyleName>();
            return n?.Val?.Value ?? "";
        })
        .ToDictionary(g => g.Key, g => g.First());

    foreach (var (sourceName, sourceStyle) in sourceStylesByName)
    {
        if (string.IsNullOrEmpty(sourceName)) continue;

        // 在模板中找到同名的样式
        var tmplStyle = tmplStylesById.Values
            .FirstOrDefault(s => {
                var n = s.GetFirstChild<StyleName>();
                return n?.Val?.Value == sourceName;
            });
        if (tmplStyle == null) continue;

        // 复制模板样式的RunProperties和ParagraphProperties到源样式
        var tmplRPr = tmplStyle.GetFirstChild<StyleRunProperties>();
        var sourceRPr = sourceStyle.GetFirstChild<StyleRunProperties>();
        if (tmplRPr != null)
        {
            if (sourceRPr == null)
            {
                sourceRPr = new StyleRunProperties();
                sourceStyle.PrependChild(sourceRPr);
            }
            sourceRPr.RemoveAllChildren();
            foreach (var child in tmplRPr.ChildElements)
                sourceRPr.Append(child.CloneNode(true));
        }

        var tmplPPr = tmplStyle.GetFirstChild<StyleParagraphProperties>();
        var sourcePPr = sourceStyle.GetFirstChild<StyleParagraphProperties>();
        if (tmplPPr != null)
        {
            if (sourcePPr == null)
            {
                sourcePPr = new StyleParagraphProperties();
                // 插入到RunProperties之后
                var afterRPr = sourceStyle.GetFirstChild<StyleRunProperties>();
                if (afterRPr != null)
                    sourceStyle.InsertAfter(sourcePPr, afterRPr);
                else
                    sourceStyle.PrependChild(sourcePPr);
            }
            sourcePPr.RemoveAllChildren();
            foreach (var child in tmplPPr.ChildElements)
                sourcePPr.Append(child.CloneNode(true));
        }
    }

    source.MainDocumentPart!.Document.Save();
}

/// <summary>
/// 方案2（更稳健）：直接基于内容特征匹配。
/// 如果源文档使用数字样式ID（如"2"="heading 1"），
/// 则比对其样式名称与模板标准样式名称，
/// 再复制格式。
/// </summary>
static void ApplyStylesByContent(string sourcePath, string templatePath)
{
    using var source = WordprocessingDocument.Open(sourcePath, true);
    using var tmpl = WordprocessingDocument.Open(templatePath, false);

    var srcStylesPart = source.MainDocumentPart!.StylePart;
    var tmplStylesPart = tmpl.MainDocumentPart!.StylePart;
    if (srcStylesPart == null || tmplStylesPart == null) return;

    var srcStyles = srcStylesPart.Styles.Descendants<Style>().ToList();
    var tmplStyles = tmplStylesPart.Styles.Descendants<Style>().ToList();

    // 模板中知名样式定义
    var knownStyles = new[] { "Normal", "Heading1", "Heading2", "Heading3",
                              "Header", "Footer", "BodyText", "BodyTextIndent",
                              "TOC1", "TOC2", "TOC3" };

    foreach (var known in knownStyles)
    {
        // 在模板中找这个样式
        var tStyle = tmplStyles.FirstOrDefault(s => s.StyleId == known);
        if (tStyle == null) continue;

        // 在源文档中找同名样式（按StyleName匹配）
        var tName = tStyle.GetFirstChild<StyleName>()?.Val?.Value ?? "";
        var sStyle = srcStyles.FirstOrDefault(s =>
            s.GetFirstChild<StyleName>()?.Val?.Value == tName);
        if (sStyle == null) continue;

        // 复制RunProperties
        var tRPr = tStyle.GetFirstChild<StyleRunProperties>();
        var sRPr = sStyle.GetFirstChild<StyleRunProperties>();
        if (tRPr != null)
        {
            if (sRPr == null) { sRPr = new StyleRunProperties(); sStyle.PrependChild(sRPr); }
            sRPr.RemoveAllChildren();
            foreach (var c in tRPr.ChildElements) sRPr.Append(c.CloneNode(true));
        }

        // 复制ParagraphProperties
        var tPPr = tStyle.GetFirstChild<StyleParagraphProperties>();
        var sPPr = sStyle.GetFirstChild<StyleParagraphProperties>();
        if (tPPr != null)
        {
            if (sPPr == null) { sPPr = new StyleParagraphProperties(); sStyle.Append(sPPr); }
            sPPr.RemoveAllChildren();
            foreach (var c in tPPr.ChildElements) sPPr.Append(c.CloneNode(true));
        }
    }

    source.MainDocumentPart!.Document.Save();
}
```

### 3.4 从模板读取页眉样式并应用到源文档

```csharp
/// <summary>
/// 从模板文档读取页眉内容、字体、对齐、边框，应用到源文档。
/// </summary>
static (string text, bool hasBorder, string font, string sz)
    ReadHeaderStyleFromTemplate(string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var mainPart = tmpl.MainDocumentPart!;

    // 从styles.xml读取Header样式定义
    var stylePart = mainPart.StylePart;
    string font = "宋体", sz = "18";
    bool hasBorder = false;

    if (stylePart != null)
    {
        var hdrStyle = stylePart.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Header");
        if (hdrStyle != null)
        {
            var rPr = hdrStyle.GetFirstChild<StyleRunProperties>();
            if (rPr != null)
            {
                var rFonts = rPr.GetFirstChild<RunFonts>();
                if (rFonts?.EastAsia != null) font = rFonts.EastAsia;
                var fSz = rPr.GetFirstChild<FontSize>();
                if (fSz?.Val != null) sz = fSz.Val;
            }
        }
    }

    // 从header1.xml读取页眉内容及边框
    string text = "";
    foreach (var hdrPart in mainPart.HeaderParts)
    {
        var hdr = hdrPart.Header;
        text = string.Concat(hdr.Descendants<Text>().Select(t => t.Text)).Trim();
        // 检查边框
        foreach (var p in hdr.Descendants<Paragraph>())
        {
            var pPr = p.GetFirstChild<ParagraphProperties>();
            if (pPr?.GetFirstChild<ParagraphBorders>() is { } pBdr)
            {
                hasBorder = pBdr.GetFirstChild<BottomBorder>() != null;
            }
        }
        if (!string.IsNullOrEmpty(text)) break;
    }

    return (text, hasBorder, font, sz);
}
```

### 3.5 页面设置应用

```csharp
static void ApplyPageSetup(string sourcePath, PageSettings ps)
{
    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var body = doc.MainDocumentPart!.Document.Body!;

    List<SectionProperties> allSectPrs = new();
    foreach (var p in body.Elements<Paragraph>())
    {
        var pPr = p.GetFirstChild<ParagraphProperties>();
        if (pPr?.GetFirstChild<SectionProperties>() is { } sp)
            allSectPrs.Add(sp);
    }
    if (body.LastChild is SectionProperties finalSp)
        allSectPrs.Add(finalSp);

    foreach (var sectPr in allSectPrs)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        bool isLandscape = pgSz?.Width > pgSz?.Height;

        pgSz ??= new PageSize();
        pgSz.Width  = isLandscape ? ps.PageHeight : ps.PageWidth;
        pgSz.Height = isLandscape ? ps.PageWidth  : ps.PageHeight;

        var pgMar = sectPr.GetFirstChild<PageMargin>()
                    ?? sectPr.InsertAfter(new PageMargin(), sectPr.GetFirstChild<PageSize>());
        pgMar.Top    = ps.MarginTop;
        pgMar.Bottom = ps.MarginBottom;
        pgMar.Left   = ps.MarginLeft;
        pgMar.Right  = ps.MarginRight;
        pgMar.Header = ps.HeaderPos;
        pgMar.Footer = ps.FooterPos;
        pgMar.Gutter = ps.Gutter;
    }

    doc.MainDocumentPart!.Document.Save();
}
```

### 3.6 页眉页脚应用

```csharp
static void ApplyHeaders(string sourcePath, string headerText, bool hasBorder,
                          string font, string sz)
{
    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var mainPart = doc.MainDocumentPart!;

    // 启用奇偶页
    var settingsPart = mainPart.DocumentSettingsPart
        ?? mainPart.AddNewPart<DocumentSettingsPart>();
    settingsPart.Settings ??= new Settings();
    settingsPart.Settings.Append(new EvenAndOddHeaders());

    // 创建页眉
    Header CreateHdr(string text, bool border)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center });
        if (border)
            pPr.Append(new ParagraphBorders(new BottomBorder
                { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" }));

        return new Header(new Paragraph(pPr,
            new Run(
                new RunProperties(
                    new RunFonts { Ascii = "Times New Roman", EastAsia = font },
                    new FontSize { Val = sz },
                    new FontSizeComplexScript { Val = sz }),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    // 创建并保存HeaderPart
    var hdrEmpty = CreateHdr("", false);
    var hdrFirst = CreateHdr("", false);
    var hdrEven  = CreateHdr("学校/学院名称", hasBorder, font, sz);
    var hdrDefault = CreateHdr(headerText, hasBorder, font, sz);

    // 保存到mainPart并获取关系ID
    // ... (与场景二代码相同)

    // 按分区设置引用
    // ... (与场景二代码相同，使用template中的headerText)
}
```

### 3.7 字体字号修正（动态参考模板样式）

```csharp
/// <summary>
/// 修正run级别的显式字号覆盖，让它们遵循样式定义。
/// 从模板获取期望的正常字号（Normal样式的sz值）。
/// </summary>
static void FixFontSizes(string sourcePath, string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    string expectedSz = "21"; // 默认值
    if (tmplStyles != null)
    {
        var normalStyle = tmplStyles.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Normal");
        var nSz = normalStyle?.GetFirstChild<StyleRunProperties>()
            ?.GetFirstChild<FontSize>();
        if (nSz?.Val != null) expectedSz = nSz.Val;
    }
    tmpl.Dispose();

    using var source = WordprocessingDocument.Open(sourcePath, true);
    var body = source.MainDocumentPart!.Document.Body!;

    int fixedCount = 0;
    foreach (var para in body.Descendants<Paragraph>())
    {
        foreach (var run in para.Elements<Run>())
        {
            var rPr = run.GetFirstChild<RunProperties>();
            if (rPr == null) continue;

            var sz = rPr.GetFirstChild<FontSize>();
            var szCs = rPr.GetFirstChild<FontSizeComplexScript>();

            // 如果run的显式字号与模板Normal样式不同，则修正
            if (sz?.Val?.Value == "22" && expectedSz == "21")
            { sz.Val = expectedSz; fixedCount++; }
            if (szCs?.Val?.Value == "22" && expectedSz == "21")
            { szCs.Val = expectedSz; }
        }
    }

    Console.WriteLine($"Fixed {fixedCount} runs to sz={expectedSz}");
    source.MainDocumentPart!.Document.Save();
}
```

### 3.8 关键词格式（毕业论文专项）

```csharp
static void FixKeywordsFormat(string sourcePath, string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    string expectedSz = "21"; // 默认

    // 从模板Normal样式中获取期望字号
    if (tmplStyles != null)
    {
        var normal = tmplStyles.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Normal");
        var nSz = normal?.GetFirstChild<StyleRunProperties>()
            ?.GetFirstChild<FontSize>();
        if (nSz?.Val != null) expectedSz = nSz.Val;
    }
    tmpl.Dispose();

    using var source = WordprocessingDocument.Open(sourcePath, true);
    var body = source.MainDocumentPart!.Document.Body!;

    foreach (var para in body.Descendants<Paragraph>())
    {
        string fullText = string.Concat(para.Elements<Run>()
            .SelectMany(r => r.Elements<Text>().Select(t => t.Text)));

        if (!fullText.TrimStart().StartsWith("Key words")) continue;

        var allRuns = para.Elements<Run>().ToList();
        RunProperties? boldRPr = null;
        RunProperties? normalRPr = null;

        foreach (var r in allRuns)
        {
            var rPr = r.GetFirstChild<RunProperties>();
            if (rPr == null) continue;
            if (rPr.GetFirstChild<Bold>() != null && boldRPr == null)
                boldRPr = rPr.CloneNode(true) as RunProperties;
            if (rPr.GetFirstChild<Bold>() == null && normalRPr == null)
                normalRPr = rPr.CloneNode(true) as RunProperties;
        }

        // 确保字号符合模板
        if (boldRPr != null)
        {
            var sz = boldRPr.GetFirstChild<FontSize>();
            if (sz == null) { sz = new FontSize(); boldRPr.PrependChild(sz); }
            sz.Val = expectedSz;
        }
        if (normalRPr != null)
        {
            var sz = normalRPr.GetFirstChild<FontSize>();
            if (sz == null) { sz = new FontSize(); normalRPr.PrependChild(sz); }
            sz.Val = expectedSz;
        }
    }
}
```

### 3.9 遍历时跳过 OMML 公式区域

字体修正、段落修正时必须**跳过 OMML 公式区域**，防止误改数学公式格式。

```csharp
/// <summary>
/// 预扫描文档，标记所有包含OMML公式的段落。
/// 后续遍历时跳过这些段落，防止破坏公式渲染。
/// </summary>
static HashSet<Paragraph> ScanFormulaParagraphs(Body body)
{
    var formulaParas = new HashSet<Paragraph>();

    // OMML 公式有两种容器: OfficeMathParagraph (m:oMathPara) 和 OfficeMath (m:oMath)
    foreach (var oMathPara in body.Descendants<OfficeMathParagraph>())
    {
        var parentPara = oMathPara.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null)
            formulaParas.Add(parentPara);
    }

    foreach (var oMath in body.Descendants<OfficeMath>())
    {
        var parentPara = oMath.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null)
            formulaParas.Add(parentPara);
    }

    return formulaParas;
}

/// <summary>
/// 修正字体字号时跳过公式段落。
/// </summary>
static void FixFontSizes_SkipFormulas(string sourcePath, string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    string expectedSz = "21";
    if (tmplStyles != null)
    {
        var normalStyle = tmplStyles.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Normal");
        var nSz = normalStyle?.GetFirstChild<StyleRunProperties>()
            ?.GetFirstChild<FontSize>();
        if (nSz?.Val != null) expectedSz = nSz.Val;
    }
    tmpl.Dispose();

    using var source = WordprocessingDocument.Open(sourcePath, true);
    var body = source.MainDocumentPart!.Document.Body!;

    // 预扫描，标记公式段落
    var formulaParas = ScanFormulaParagraphs(body);
    Console.WriteLine($"找到 {formulaParas.Count} 个公式段落，已保护");

    int fixedCount = 0;
    foreach (var para in body.Descendants<Paragraph>())
    {
        // 跳过公式段落
        if (formulaParas.Contains(para)) continue;

        foreach (var run in para.Elements<Run>())
        {
            var rPr = run.GetFirstChild<RunProperties>();
            if (rPr == null) continue;

            var sz = rPr.GetFirstChild<FontSize>();
            var szCs = rPr.GetFirstChild<FontSizeComplexScript>();

            if (sz?.Val?.Value == "22" && expectedSz == "21")
            { sz.Val = expectedSz; fixedCount++; }
            if (szCs?.Val?.Value == "22" && expectedSz == "21")
            { szCs.Val = expectedSz; }
        }
    }

    Console.WriteLine($"Fixed {fixedCount} runs to sz={expectedSz}");
    source.MainDocumentPart!.Document.Save();
}
```

### 3.10 应用页面设置时跳过公式段落（保护公式不因边距变化而断行）

```csharp
/// <summary>
/// 应用页面设置时，跳过公式段落内部的空行修正。
/// 公式段落通常由Word自动生成另起的分页符，不应被修正边距影响。
/// </summary>
static void ApplyPageSetup_SkipFormulas(string sourcePath, PageSettings ps)
{
    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var body = doc.MainDocumentPart!.Document.Body!;

    var formulaParas = ScanFormulaParagraphs(body);

    List<SectionProperties> allSectPrs = new();
    foreach (var p in body.Elements<Paragraph>())
    {
        if (formulaParas.Contains(p)) continue;
        var pPr = p.GetFirstChild<ParagraphProperties>();
        if (pPr?.GetFirstChild<SectionProperties>() is { } sp)
            allSectPrs.Add(sp);
    }
    if (body.LastChild is SectionProperties finalSp && !formulaParas.Contains(body.LastChild as Paragraph))
        allSectPrs.Add(finalSp);

    foreach (var sectPr in allSectPrs)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        bool isLandscape = pgSz?.Width > pgSz?.Height;
        pgSz ??= new PageSize();
        pgSz.Width  = isLandscape ? ps.PageHeight : ps.PageWidth;
        pgSz.Height = isLandscape ? ps.PageWidth  : ps.PageHeight;

        var pgMar = sectPr.GetFirstChild<PageMargin>()
                    ?? sectPr.InsertAfter(new PageMargin(), sectPr.GetFirstChild<PageSize>());
        pgMar.Top    = ps.MarginTop;
        pgMar.Bottom = ps.MarginBottom;
        pgMar.Left   = ps.MarginLeft;
        pgMar.Right  = ps.MarginRight;
        pgMar.Header = ps.HeaderPos;
        pgMar.Footer = ps.FooterPos;
        pgMar.Gutter = ps.Gutter;
    }

    doc.MainDocumentPart!.Document.Save();
}
```

### 3.11 文本框与形状文字处理

文本框（`w:txbxContent`）和形状（`w:pict`、`w:alternateContent`）中的段落不在 body 的直接遍历路径上，需要额外处理。

```csharp
/// <summary>
/// 修正文本框和形状中的文字格式。
/// 使用 Descendants<TextBoxContent> 遍历所有文本框内容。
/// </summary>
static void FixTextboxContent(string sourcePath, string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    string expectedSz = "21";
    if (tmplStyles != null)
    {
        var normalStyle = tmplStyles.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Normal");
        var nSz = normalStyle?.GetFirstChild<StyleRunProperties>()
            ?.GetFirstChild<FontSize>();
        if (nSz?.Val != null) expectedSz = nSz.Val;
    }
    tmpl.Dispose();

    using var source = WordprocessingDocument.Open(sourcePath, true);
    var mainPart = source.MainDocumentPart!;

    int fixedCount = 0;
    int textboxCount = 0;

    // 遍历所有文本框内容
    foreach (var txbxContent in mainPart.Document.Descendants<TextBoxContent>())
    {
        textboxCount++;
        foreach (var para in txbxContent.Elements<Paragraph>())
        {
            // 跳过文本框中的公式
            if (para.Descendants<OfficeMath>().Any() ||
                para.Descendants<OfficeMathParagraph>().Any()) continue;

            foreach (var run in para.Elements<Run>())
            {
                var rPr = run.GetFirstChild<RunProperties>();
                if (rPr == null) continue;

                var sz = rPr.GetFirstChild<FontSize>();
                if (sz?.Val?.Value == "22" && expectedSz == "21")
                { sz.Val = expectedSz; fixedCount++; }
            }
        }
    }

    Console.WriteLine($"发现 {textboxCount} 个文本框");
    Console.WriteLine($"文本框内修正 {fixedCount} 处字体");
    mainPart.Document.Save();
}
```

### 3.12 目录（TOC）样式联动

正文样式变更后，TOC 样式（TOC1-TOC9）的字号应跟随调整，保持与正文的比例关系。

```csharp
/// <summary>
/// 正文格式修正后，同步调整目录样式的字号。
/// 目录字号 = 正文字号 × TOC层级缩放因子。
/// 典型缩放: TOC1=正文×1.0, TOC2=正文×0.9, TOC3=正文×0.8
/// </summary>
static void SyncTOCStyles(string sourcePath)
{
    var scaleFactors = new Dictionary<string, double>
    {
        ["TOC1"] = 1.0,  // 一级目录 = 正文字号
        ["TOC2"] = 0.9,  // 二级目录 = 正文×0.9
        ["TOC3"] = 0.8,  // 三级目录 = 正文×0.8
        ["TOC4"] = 0.8,
        ["TOC5"] = 0.7,
    };

    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var stylePart = doc.MainDocumentPart!.StylePart;
    if (stylePart == null) return;

    // 获取当前正文字号
    var normalStyle = stylePart.Styles.Descendants<Style>()
        .FirstOrDefault(s => s.StyleId == "Normal");
    var normalSz = normalStyle?.GetFirstChild<StyleRunProperties>()
        ?.GetFirstChild<FontSize>()?.Val?.Value;
    if (normalSz == null || !int.TryParse(normalSz, out int baseSz)) return;

    int updatedCount = 0;
    foreach (var style in stylePart.Styles.Descendants<Style>())
    {
        if (!scaleFactors.TryGetValue(style.StyleId ?? "", out double factor)) continue;

        int tocSz = (int)(baseSz * factor);
        var tocRPr = style.GetFirstChild<StyleRunProperties>();
        if (tocRPr == null)
        {
            tocRPr = new StyleRunProperties();
            style.PrependChild(tocRPr);
        }

        var sz = tocRPr.GetFirstChild<FontSize>();
        if (sz == null)
        {
            sz = new FontSize();
            tocRPr.PrependChild(sz);
        }
        sz.Val = tocSz.ToString();

        var szCs = tocRPr.GetFirstChild<FontSizeComplexScript>();
        if (szCs == null)
        {
            szCs = new FontSizeComplexScript();
            tocRPr.Append(szCs);
        }
        szCs.Val = tocSz.ToString();

        updatedCount++;
    }

    Console.WriteLine($"已同步 {updatedCount} 个TOC样式的字号 (基础:{normalSz})");
    stylePart.Styles.Save();
}
```

### 3.13 脚注/尾注格式处理

脚注和尾注的正文部分需要分别遍历，它们位于独立的部件（FootnotePart, EndnotePart）中。

```csharp
/// <summary>
/// 修正脚注和尾注中的文字格式。
/// 脚注/尾注部件中的段落不在body中，需要单独遍历。
/// </summary>
static void FixFootnotesAndEndnotes(string sourcePath, string templatePath)
{
    using var tmpl = WordprocessingDocument.Open(templatePath, false);
    var tmplStyles = tmpl.MainDocumentPart!.StylePart;
    string expectedSz = "21";
    if (tmplStyles != null)
    {
        var normalStyle = tmplStyles.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "Normal");
        var nSz = normalStyle?.GetFirstChild<StyleRunProperties>()
            ?.GetFirstChild<FontSize>();
        if (nSz?.Val != null) expectedSz = nSz.Val;
    }
    tmpl.Dispose();

    using var source = WordprocessingDocument.Open(sourcePath, true);
    var mainPart = source.MainDocumentPart!;

    // 获取脚注样式中的字号（通常小于正文字号）
    string footnoteSz = "18"; // 默认脚注字号 9pt
    var sourceStylePart = mainPart.StylePart;
    if (sourceStylePart != null)
    {
        var footnoteTextStyle = sourceStylePart.Styles.Descendants<Style>()
            .FirstOrDefault(s => s.StyleId == "FootnoteText");
        if (footnoteTextStyle != null)
        {
            var fnSz = footnoteTextStyle.GetFirstChild<StyleRunProperties>()
                ?.GetFirstChild<FontSize>()?.Val?.Value;
            if (fnSz != null) footnoteSz = fnSz;
        }
    }

    int fixedCount = 0;

    // 处理脚注
    foreach (var footnotePart in mainPart.FootnoteParts)
    {
        var footnotes = footnotePart.Footnotes;
        if (footnotes == null) continue;

        foreach (var para in footnotes.Descendants<Paragraph>())
        {
            foreach (var run in para.Elements<Run>())
            {
                var rPr = run.GetFirstChild<RunProperties>();
                if (rPr == null) continue;

                var sz = rPr.GetFirstChild<FontSize>();
                if (sz?.Val?.Value == "22")
                { sz.Val = footnoteSz; fixedCount++; }
            }
        }
    }

    // 处理尾注
    foreach (var endnotePart in mainPart.EndnoteParts)
    {
        var endnotes = endnotePart.Endnotes;
        if (endnotes == null) continue;

        foreach (var para in endnotes.Descendants<Paragraph>())
        {
            foreach (var run in para.Elements<Run>())
            {
                var rPr = run.GetFirstChild<RunProperties>();
                if (rPr == null) continue;

                var sz = rPr.GetFirstChild<FontSize>();
                if (sz?.Val?.Value == "22")
                { sz.Val = footnoteSz; fixedCount++; }
            }
        }
    }

    Console.WriteLine($"修正脚注/尾注: {fixedCount} 处");
    mainPart.Document.Save();
}
```

### 3.14 批注与修订标记保留

修正格式时不应破坏批注内容或修订标记结构。

```csharp
/// <summary>
/// 批注内容仅读取，不做任何格式修改。
/// 但如果需要调整批注框中的文字格式，可以遍历 CommentsPart。
/// </summary>
static void PreserveComments(string sourcePath)
{
    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var mainPart = doc.MainDocumentPart!;

    int commentCount = 0;
    foreach (var commentsPart in mainPart.CommentsParts)
    {
        var comments = commentsPart.Comments;
        if (comments == null) continue;

        foreach (var comment in comments.Elements<Comment>())
        {
            commentCount++;
            // 批注只读不写，确保格式修正不影响批注
        }
    }

    Console.WriteLine($"保留 {commentCount} 条批注");
}

/// <summary>
/// 检查修订标记是否完整（不尝试修改，仅报告状态）。
/// 修订标记(rPrChange)中的旧格式信息不应被格式修正破坏。
/// </summary>
static void CheckTrackedChanges(string sourcePath)
{
    using var doc = WordprocessingDocument.Open(sourcePath, false);
    var body = doc.MainDocumentPart!.Document.Body!;

    int changeCount = 0;
    foreach (var rPrChange in body.Descendants<RunPropertiesChange>())
    {
        changeCount++;
    }
    foreach (var pPrChange in body.Descendants<ParagraphPropertiesChange>())
    {
        changeCount++;
    }

    if (changeCount > 0)
        Console.WriteLine($"文档包含 {changeCount} 处修订标记，格式修正不会修改修订中的旧格式");
}
```

### 3.15 分区修正前检查编号列表缩进

修正页边距后，多级列表的缩进可能偏离，需要检查并调整。

```csharp
/// <summary>
/// 修正页边距后，检查编号列表的缩进是否合理。
/// 如果缩进超过页边距，自动缩小到页边距以内。
/// marginRatio 控制编号缩进占页边距的最大比例（默认0.8）
/// </summary>
static void AdjustNumberingIndents(string sourcePath, double marginRatio = 0.8)
{
    using var doc = WordprocessingDocument.Open(sourcePath, true);
    var mainPart = doc.MainDocumentPart!;
    var body = mainPart.Document.Body!;

    // 获取当前页边距
    var lastSectPr = body.Descendants<SectionProperties>().LastOrDefault();
    if (lastSectPr == null) return;

    var pgMar = lastSectPr.GetFirstChild<PageMargin>();
    if (pgMar == null) return;

    uint marginLeft = pgMar.Left ?? 1701;
    uint maxIndent = (uint)(marginLeft * marginRatio);

    // 遍历编号定义，检查缩进
    var numberingPart = mainPart.NumberingPart;
    if (numberingPart == null) return;

    int adjustedCount = 0;
    foreach (var level in numberingPart.Numbering.Descendants<Level>())
    {
        var pPr = level.GetFirstChild<LevelParagraphProperties>();
        if (pPr == null) continue;

        var ind = pPr.GetFirstChild<Indentation>();
        if (ind == null) continue;

        if (ind.Left != null)
        {
            if (uint.TryParse(ind.Left.Value, out uint leftVal) && leftVal > maxIndent)
            {
                ind.Left = maxIndent.ToString();
                adjustedCount++;
            }
        }

        if (ind.Hanging != null)
        {
            if (uint.TryParse(ind.Hanging.Value, out uint hangingVal) && hangingVal > maxIndent)
            {
                ind.Hanging = maxIndent.ToString();
                adjustedCount++;
            }
        }
    }

    Console.WriteLine($"调整 {adjustedCount} 处编号缩进");
    numberingPart.Numbering.Save();
}
```

---

## 四、验证脚本（动态对比模板）

### 4.1 格式验证

```python
# validate_format.py
# 动态读取模板的每个分区设置，逐一对比目标文档
def validate_format(target_path, template_path):
    target_setup = extract_page_setup(target_path)
    tmpl_setup = extract_page_setup(template_path)

    print(f'模板: PageSize={tmpl_setup["PageWidth"]}x{tmpl_setup["PageHeight"]}')
    print(f'目标: PageSize={target_setup["PageWidth"]}x{target_setup["PageHeight"]}')

    for key in tmpl_setup:
        t_val = tmpl_setup[key]
        d_val = target_setup.get(key)
        status = '✅' if t_val == d_val else f'❌ (目标={d_val})'
        print(f'  {key}: 模板={t_val} {status}')
```

### 4.2 文本格式验证

```python
# verify_text_format.py
# 使用模板的样式定义作为验证标准
def verify_text_format(target_path, template_path):
    _, tmpl_styles = extract_style_definitions(template_path)

    # 例如，验证正文段落字号是否与模板Normal样式一致
    normal_style = tmpl_styles.get('Normal', {})
    expected_sz = normal_style.get('font', {}).get('sz', '21')

    # 遍历目标文档段落，检查字号
    for p in all_paragraphs:
        actual_sz = p.get('sz', '')
        if actual_sz and actual_sz != expected_sz:
            print(f'  段{p["idx"]}: sz={actual_sz} (模板要求sz={expected_sz})')
```

### 4.3 OMML 公式完整性验证

```python
# verify_formulas.py
# 验证OMML公式结构在格式修正后是否完整
def verify_formulas(docx_path):
    """验证文档中OMML公式的结构完整性"""
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    }
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            root = ET.fromstring(f.read())

        oMaths = root.findall('.//m:oMath', ns)
        print(f'文档包含 {len(oMaths)} 个OMML公式')

        for i, om in enumerate(oMaths):
            # 检查公式内是否有文本内容
            texts = om.findall('.//m:t', ns)
            if not texts:
                print(f'  ⚠️ 公式 #{i}: 无文本内容')
            else:
                # 检查文本中是否有非公式字符
                for t in texts[:3]:
                    if t.text:
                        print(f'  公式 #{i}: 内容="{[t.text for t in texts[:3]]}"')

        print('✅ 公式结构完整')
```

### 4.4 TOC 格式验证

```python
# verify_toc_format.py
# 验证目录格式是否与正文一致
def verify_toc_format(docx_path):
    """验证目录样式的字号是否合理"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read())

        # 获取正文字号
        normal_sz = None
        for style in root.findall('.//w:style', ns):
            sid = style.get('{http://...}styleId', '')
            if sid == 'Normal':
                sz = style.find('.//w:sz', ns)
                if sz is not None:
                    normal_sz = sz.get('{http://...}val', '')

        # 检查TOC样式
        for style in root.findall('.//w:style', ns):
            sid = style.get('{http://...}styleId', '')
            if sid and sid.startswith('TOC'):
                sz = style.find('.//w:sz', ns)
                if sz is not None:
                    toc_sz = sz.get('{http://...}val', '')
                    print(f'  {sid}: sz={toc_sz} (Normal: sz={normal_sz})')
```

### 4.5 脚注/尾注格式验证

```python
# verify_footnotes.py
# 验证脚注和尾注的格式
def verify_footnotes(docx_path):
    """验证脚注/尾注的样式定义"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/styles.xml') as f:
            root = ET.fromstring(f.read())

        for style in root.findall('.//w:style', ns):
            sid = style.get('{http://...}styleId', '')
            if sid in ('FootnoteText', 'FootnoteReference', 'EndnoteText', 'EndnoteReference'):
                sz = style.find('.//w:sz', ns)
                sz_val = sz.get('{http://...}val', '') if sz is not None else '(默认)'
                name_el = style.find('w:name', ns)
                sname = name_el.get('{http://...}val', '') if name_el is not None else sid
                print(f'  {sid} ({sname}): sz={sz_val}')
```

---

## 五、毕业论文专项优化

### 5.1 分区结构（逻辑分区 vs 模板分区）

通用的"前文分区用空页眉、正文分区含页眉"逻辑仍然保留，但**分区判定阈值**从模板提取：

```python
def detect_thesis_zones(template_path):
    """检测模板文档中各分区的类型，返回前文分区数量"""
    tmpl_setup = extract_page_setup(template_path)
    # 使用模板的页眉引用模式判断哪些是前文分区
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(template_path) as z:
        with z.open('word/document.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))
        body = root.find('.//w:body', ns)

        # 查找所有sectPr中的titlePg和headerReference
        front_count = 0
        # ... 遍历sectPr，统计使用空页眉或titlePg的分区数
        return front_count  # 动态获取前文分区数
```

### 5.2 横/竖页面自动保留

从模板学到的横竖判断逻辑：

```csharp
// 自动保持横竖方向：比较模板页面尺寸与目标页面尺寸
bool isLandscape = pgSz.Width > pgSz.Height;
// 如果是横排，交换宽高
if (isLandscape)
    { pgSz.Width = ps.PageHeight; pgSz.Height = ps.PageWidth; }
else
    { pgSz.Width = ps.PageWidth; pgSz.Height = ps.PageHeight; }
```

### 5.3 毕业论文公式保护

毕业论文中常见 MathType/OMML 公式，必须保护：

```csharp
// 毕业论文格式修正前先扫描并保护公式段落
var formulaParas = ScanFormulaParagraphs(body);

// 字体修正时跳过公式段落
foreach (var para in body.Descendants<Paragraph>())
{
    if (formulaParas.Contains(para))
        continue;  // 公式段落不修正字号

    // ... 正常修正逻辑 ...
}
```

---

## 六、高级场景与特殊内容处理

### 6.1 OMML 公式保护

OMML（Office Math Markup Language）是 Word 原生公式格式，在大纲视图下显示为 `m:oMath` 或 `m:oMathPara` 元素。

**为什么需要保护公式**：

```xml
<!-- OMML 公式在 OpenXML 中的结构 -->
<m:oMathPara>
  <m:oMath>
    <m:r>
      <m:rPr>
        <m:sty m:val="p"/>    <!-- 数学样式，不是 w:rPr -->
      </m:rPr>
      <m:t>E=mc^2</m:t>       <!-- 公式文本，不是 w:t -->
    </m:r>
  </m:oMath>
</m:oMathPara>
```

如果格式修正代码使用 `body.Descendants<Paragraph>()` + `para.Descendants<Run>()` 遍历，OMML 的 `<m:r>` 不会被误改（因为 namespace 不同）。但**公式所在段落的段落属性**（如行距、段前段后距）可能被误修，导致公式显示异常。

**C# 专用保护码**（见 3.9 节）：

```csharp
// 预扫描标记公式段落
var formulaParas = ScanFormulaParagraphs(body);

// 遍历时跳过公式段落
foreach (var para in body.Descendants<Paragraph>())
{
    if (formulaParas.Contains(para)) continue;
    // ... 正常格式修正 ...
}
```

**MathType 公式保护**：MathType 公式作为 OLE 对象嵌入，被 `w:object` 或 `w:pict` 包含。修正边距或页面布局时，MathType 公式不应受影响。遍历时识别 `w:object` 包裹的段落并跳过。

```csharp
/// <summary>
/// 扫描并标记包含 MathType 公式（OLE对象）的段落。
/// </summary>
static HashSet<Paragraph> ScanMathTypeParagraphs(Body body)
{
    var mathTypeParas = new HashSet<Paragraph>();

    // MathType 公式通常被 w:object 或 w:pict 包裹
    foreach (var obj in body.Descendants<EmbeddedObject>())
    {
        var parentPara = obj.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null)
            mathTypeParas.Add(parentPara);
    }

    foreach (var pict in body.Descendants<Picture>())
    {
        // 检查是否是 MathType（通过 Shape ID 或 属性判断）
        var parentPara = pict.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null)
            mathTypeParas.Add(parentPara);
    }

    return mathTypeParas;
}
```

### 6.2 文本框/形状文字格式化

文本框中的段落不在 `body.Descendants<Paragraph>()` 的遍历路径上，需要直接使用 `Descendants<TextBoxContent>()` 定位。

**问题场景**：
- 插入的图片标注文本框
- 流程图/示意图中的标注
- 表格内的文本框

**C# 实现**（见 3.11 节）：

```csharp
// 遍历所有文本框内容
foreach (var txbxContent in mainPart.Document.Descendants<TextBoxContent>())
{
    foreach (var para in txbxContent.Elements<Paragraph>())
    {
        // 同样需要跳过文本框中的公式
        if (para.Descendants<OfficeMath>().Any()) continue;

        foreach (var run in para.Elements<Run>())
        {
            var rPr = run.GetFirstChild<RunProperties>();
            if (rPr == null) continue;
            // ... 修正逻辑 ...
        }
    }
}
```

**Python 验证**：

```python
def verify_textboxes(docx_path):
    """验证文本框内容是否被格式修正波及"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            root = ET.fromstring(f.read().decode('utf-8'))

        txbxContents = root.findall('.//w:txbxContent', ns)
        print(f'文档包含 {len(txbxContents)} 个文本框')

        for i, txbx in enumerate(txbxContents):
            texts = txbx.findall('.//w:t', ns)
            text_preview = ''.join([t.text for t in texts[:5] if t.text])[:80]
            print(f'  文本框 #{i}: "{text_preview}..."')
```

### 6.3 目录(TOC)样式联动

**问题**：正文格式修正后（例如正文字号从 12pt 改为 10.5pt），目录的字号也应等比缩小，否则目录比正文还大，视觉不协调。

**典型毕业论文的 TOC 层级**：

| TOC 层级 | 样式名称 | 相对于正文的缩放 | 典型字号 |
|---------|---------|----------------|---------|
| 一级目录 | TOC1 | 1.0×（与正文相同） | 10.5pt |
| 二级目录 | TOC2 | 0.9× | 10pt |
| 三级目录 | TOC3 | 0.8× | 9pt |

**C# 实现**（见 3.12 节）：

```csharp
// 正文格式修正后，同步TOC样式
SyncTOCStyles(outputPath);
```

**触发时机**：在 `FixFontSizes` 或 `ApplyStylesFromTemplate` 完成之后调用。

```csharp
// 完整的修正流程中调用
ApplyStylesFromTemplate(outputPath, templatePath);
FixFontSizes(outputPath, templatePath);
SyncTOCStyles(outputPath);  // 正文格式修正后同步TOC
```

### 6.4 脚注/尾注格式

**问题**：脚注属于独立的部件（`FootnotePart`），不在 `body.Descendants<Paragraph>()` 中。如果格式修正只遍历 body，脚注中的字体/字号不会被修正。

**OpenXML 结构**：

```
word/
  document.xml          ← 主文档（body）
  footnotes.xml         ← 脚注内容（独立）
  endnotes.xml          ← 尾注内容（独立）
  styles.xml            ← 样式定义（含 FootnoteText 样式）
```

**C# 实现**（见 3.13 节）：

```csharp
// 处理脚注
foreach (var footnotePart in mainPart.FootnoteParts)
{
    var footnotes = footnotePart.Footnotes;
    foreach (var para in footnotes.Descendants<Paragraph>())
    {
        foreach (var run in para.Elements<Run>())
        {
            // ... 字号修正 ...
        }
    }
}
```

**脚注字号策略**：脚注通常使用比正文更小的字号（9pt vs 10.5pt）。应当从 `FootnoteText` 样式中读取期望字号，而非硬编码。

### 6.5 编号列表保持

**问题**：修正页边距（尤其是左边距缩小）后，多级列表的缩进可能超出新的页边距，导致编号显示在页面边缘外。

**典型场景**：

```
修正前: 左边距=3.0cm, 三级编号缩进=2.8cm  ✅
修正后: 左边距=2.5cm, 三级编号缩进=2.8cm  ❌（超出页边距）
```

**C# 实现**（见 3.15 节）：

```csharp
// 修正页边距后调整编号缩进
ApplyPageSetup(outputPath, pageSettings);
AdjustNumberingIndents(outputPath);
```

**检查逻辑**：遍历 `NumberingPart` 中所有 `Level` 的 `LevelParagraphProperties > Indentation`，如果缩进值超过页边距的 80%，则缩小到页边距的 80%。

### 6.6 批注与修订标记保留

**原则**：格式修正程序**不修改**批注和修订标记的内容，只确保它们不被破坏。

- **批注**（`CommentsPart`）：只读取计数，不做任何修改
- **修订标记**（`rPrChange`, `pPrChange`）：格式修正时，OpenXML SDK 修改属性时会自动保留修订内的旧属性

**C# 实现**（见 3.14 节）：

```csharp
// 报告文档中的批注和修订数量
PreserveComments(sourcePath);
CheckTrackedChanges(sourcePath);
```

---

## 七、格式修正通用模板流程

以下是一个整合了所有场景的完整流程模板。实际使用时根据用户需求和文档类型选择执行其中的部分步骤。

```
┌─────────────────────────────────────────────────────────────┐
│ 0. 预扫描阶段                                                │
│    ├── ScanFormulaParagraphs()    → 标记公式段落（保护）     │
│    ├── ScanMathTypeParagraphs()   → 标记OLE公式段落（保护）  │
│    └── CheckTrackedChanges()      → 报告修订标记数量         │
├─────────────────────────────────────────────────────────────┤
│ 1. 页面设置应用                                              │
│    ├── ReadPageSettingsFromTemplate()                        │
│    └── ApplyPageSetup()                                     │
├─────────────────────────────────────────────────────────────┤
│ 2. 页眉页脚应用                                              │
│    ├── ReadHeaderStyleFromTemplate()                         │
│    └── ApplyHeaders()                                       │
├─────────────────────────────────────────────────────────────┤
│ 3. 样式复制                                                  │
│    ├── ApplyStylesFromTemplate()    → 匹配同名样式复制      │
│    └── ApplyStylesByContent()       → 按内容特征匹配（备用） │
├─────────────────────────────────────────────────────────────┤
│ 4. 字体修正（高级版，跳过公式）                               │
│    ├── FixFontSizes_SkipFormulas() → 正文跳过公式段落       │
│    ├── FixTextboxContent()         → 文本框内字体修正       │
│    ├── FixFootnotesAndEndnotes()   → 脚注/尾注字体修正      │
│    └── FixKeywordsFormat()         → 关键词格式修正         │
├─────────────────────────────────────────────────────────────┤
│ 5. 后处理与同步                                              │
│    ├── SyncTOCStyles()             → TOC字号与正文同步      │
│    ├── AdjustNumberingIndents()    → 编号缩进适配页边距     │
│    └── PreserveComments()          → 确认批注完整保留       │
├─────────────────────────────────────────────────────────────┤
│ 6. 验证阶段                                                  │
│    ├── validate_format()           → 页面设置对比           │
│    ├── verify_text_format()        → 文本格式对比           │
│    ├── verify_formulas()           → 公式结构验证           │
│    ├── verify_toc_format()         → 目录格式验证           │
│    └── verify_footnotes()          → 脚注格式验证           │
└─────────────────────────────────────────────────────────────┘
```

## 八、完整调用示例

```csharp
string templatePath = @"写作模板.docx";
string sourcePath   = @"源文档.docx";
string outputPath   = @"输出文档.docx";

File.Copy(sourcePath, outputPath, overwrite: true);

// 步骤0: 预扫描阶段
Console.WriteLine("=== 开启文档 ===");
using var preview = WordprocessingDocument.Open(outputPath, false);
var previewBody = preview.MainDocumentPart!.Document.Body!;
var formulaParas = ScanFormulaParagraphs(previewBody);
var mathTypeParas = ScanMathTypeParagraphs(previewBody);
Console.WriteLine($"预扫描: {formulaParas.Count} 个OMML公式段落, {mathTypeParas.Count} 个MathType公式段落");
preview.Dispose();

// 步骤1: 从模板读取所有格式设定
var pageSettings = ReadPageSettingsFromTemplate(templatePath);
var (headerText, hasBorder, hdrFont, hdrSz) = ReadHeaderStyleFromTemplate(templatePath);

// 步骤2: 应用页面设置（跳过公式区域）
ApplyPageSetup(outputPath, pageSettings);

// 步骤3: 应用页眉页脚
ApplyHeaders(outputPath, headerText, hasBorder, hdrFont, hdrSz);

// 步骤4: 复制模板样式定义到源文档
ApplyStylesFromTemplate(outputPath, templatePath);

// 步骤5: 字体修正（跳过公式段落 + 文本框 + 脚注）
FixFontSizes_SkipFormulas(outputPath, templatePath);
FixTextboxContent(outputPath, templatePath);
FixFootnotesAndEndnotes(outputPath, templatePath);
FixKeywordsFormat(outputPath, templatePath);

// 步骤6: TOC同步 + 编号调整 + 批注保留
SyncTOCStyles(outputPath);
AdjustNumberingIndents(outputPath);
PreserveComments(outputPath);

Console.WriteLine("=== 格式修正完成 ===");
```

## 九、注意事项

1. **零硬编码**：所有格式值都从模板读取，不写死任何magic number
2. **模板优先**：如果模板的特定样式不存在，使用有意义的默认值（如A4尺寸）
3. **样式名匹配**：源文档可能使用数字样式ID（2, 3, 4...），通过StyleName映射到模板的标准样式
4. **横竖保留**：应用模板页面尺寸时，保留源文档各分区的横竖方向
5. **分步调试**：每步应用后运行验证脚本，定位格式偏差
6. **公式保护优先**：任何可能修改段落属性或字号的步骤，都必须先扫描并跳过公式所在的段落
7. **文本框独立遍历**：`body.Descendants<Paragraph>()` 不会覆盖文本框，需要额外遍历 `TextBoxContent`
8. **脚注尾注独立部件**：脚注/尾注的正文在单独的部件中，必须分别遍历
9. **TOC样式后同步**：正文格式修正完成后，再同步TOC样式，避免比例关系错乱
10. **编号缩进适配边距**：边距缩小后，编号缩进可能超出页边距，需要检查和调整
11. **修订标记只读**：不要修改 `rPrChange`/`pPrChange` 中的旧格式信息
12. **OMML vs MathType**：OMML 是 Word 原生公式 (m:oMath)，MathType 是 OLE 嵌入对象 (w:object)，两者保护方式不同

## 十、常见问题

| 症状 | 原因 | 修复 |
|------|------|------|
| 页面尺寸未更新 | sectPr未找到 | 检查模板是否包含分区信息 |
| 页眉字体/大小不对 | 模板Header样式未定义rPr | 从模板header1.xml直接读取字体值 |
| 样式复制后部分格式丢失 | 模板样式继承链断开 | 同时复制LatentStyles和DocDefaults |
| 数字ID样式无法匹配 | 源文档使用自定义样式名 | 按内容特征（字号、加粗）匹配而非名称 |
| 表格文字大小不变 | 表格内run有显式sz覆盖 | 使用Descendants遍历并修正 |
| 公式区域字号被误改 | 代码遍历时未跳过OMML公式段落 | 预扫描公式段落并跳过（见6.1节） |
| 公式渲染变形/分页错位 | 页边距修正影响公式自动换行 | 公式段落的sectPr修正前跳过 |
| MathType 公式显示异常 | OLE对象被格式修正干扰 | 识别w:object段落并保护（见6.1节） |
| 文本框文字格式未变 | 文本框独立于body段落树之外 | 使用Descendants\<TextBoxContent\>遍历（见6.2节） |
| 目录字号与正文不协调 | 修正正文格式后未同步TOC样式 | 正文修正后调用SyncTOCStyles（见6.3节） |
| 脚注字号未跟随正文变化 | 脚注在独立部件中未被遍历 | 遍历FootnotePart/EndnotePart（见6.4节） |
| 编号缩进超出页边距 | 边距缩小后编号缩进未自动调整 | 修正边距后调用AdjustNumberingIndents（见6.5节） |
| 批注内容丢失 | 格式修正误删了CommentsPart | 使用只读方式打开CommentsPart（见6.6节） |
| Key words段落字号未修正 | sz定义在pPr内的ParagraphMarkRunProperties中 | 使用pPr.GetFirstChild\<ParagraphMarkRunProperties\>()获取rPr再修改sz |

---

## 附录：环境依赖与安装方法

### 环境要求

| 依赖 | 最低版本 | 用途 |
|------|---------|------|
| .NET SDK | 8.0+ | 编译运行 C# 格式修正程序 (OpenXML SDK) |
| Python | 3.8+ | 模板格式提取脚本、格式验证脚本 |
| NuGet 包 | DocumentFormat.OpenXml 3.2.0 | 由 dotnet restore 自动还原 |

### 依赖关系说明

```
docx-format-applier 技能
    ├── .NET SDK 8.0+      ──→  C# 格式修正核心 (FormatFix.csproj)
    │   └── NuGet 自动还原 ──→  DocumentFormat.OpenXml 3.2.0
    └── Python 3.8+         ──→  模板提取 (标准库)
                                └── 格式验证 (标准库，无需 pip 包)
```

> **注意**：此技能**完全独立于 minimax-docx 技能**。即使没有安装 minimax-docx，只要 .NET SDK 和 Python 就绪即可使用。

### 一键环境检查脚本

项目根目录提供了 [setup.ps1](file:///d:/pythontest/论文/setup.ps1) 脚本，可自动检查并安装所有依赖：

```powershell
# 打开 PowerShell，进入项目目录，执行：
.\setup.ps1
```

脚本会依次：
1. 检查 .NET SDK 是否安装 → 未安装则提示安装命令
2. 检查 Python 是否安装 → 未安装则提示安装命令
3. 运行 `dotnet restore` 还原 NuGet 包
4. 运行 `dotnet build` 验证项目可编译
5. 报告最终状态

### 手动安装方法

#### 方式一：通过 winget（推荐，Windows 10/11 自带）

```powershell
# 1. 安装 .NET SDK 8.0
winget install Microsoft.DotNet.SDK.8

# 2. 验证安装
dotnet --version            # 应输出 8.x.x
python --version            # 应输出 3.x

# 3. 进入项目目录，还原并编译
cd D:\pythontest\论文\FormatFix5
dotnet restore
dotnet build
```

#### 方式二：通过 Chocolatey（如果已安装 choco）

```powershell
choco install dotnet-8.0-sdk python -y
```

#### 方式三：手动下载安装

- .NET SDK 8.0: https://dotnet.microsoft.com/download/dotnet/8.0
- Python 3.x: https://www.python.org/downloads/

### 验证环境是否就绪

```powershell
# 验证 .NET
dotnet --version
# 期望输出: 8.x.x

# 验证 NuGet 包已还原（在 FormatFix5 目录下）
cd D:\pythontest\论文\FormatFix5
dotnet list package
# 期望输出中包含 DocumentFormat.OpenXml 3.2.0

# 验证项目可编译
dotnet build
# 期望输出: 生成成功 (Build succeeded)

# 验证 Python
python -c "import zipfile, xml.etree.ElementTree; print('Python 就绪')"
# 期望输出: Python 就绪
```

### 常见安装问题

| 问题 | 原因 | 修复 |
|------|------|------|
| `dotnet` 不是内部或外部命令 | .NET SDK 未安装或 PATH 未配置 | 安装 SDK 后重启终端 |
| `dotnet build` 失败：找不到 DocumentFormat.OpenXml | NuGet 包未还原 | 运行 `dotnet restore` |
| `python` 不是内部或外部命令 | Python 未安装或 PATH 未配置 | 安装 Python 时勾选 "Add to PATH" |
| 构建成功但运行时报错 | OpenXML SDK 版本不兼容 | 检查 csproj 中的版本号是否为 3.2.0 |