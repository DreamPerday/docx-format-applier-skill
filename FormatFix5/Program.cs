using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string sourcePath = @"D:\pythontest\论文\20220903430216_BS_polished.docx";
string outputPath = @"D:\pythontest\论文\20220903430216_BS_final.docx";

File.Copy(sourcePath, outputPath, overwrite: true);

int sz22Fixed = 0;
int sz24Fixed = 0;
int szCs22Fixed = 0;
int szCs24Fixed = 0;

using var doc = WordprocessingDocument.Open(outputPath, true);
var mainPart = doc.MainDocumentPart!;
var body = mainPart.Document.Body!;

// 获取 Key words 段落的文本
var keywordsParas = body.Descendants<Paragraph>()
    .Where(p => {
        var text = string.Concat(p.Descendants<Text>().Select(t => t.Text));
        return text.TrimStart().StartsWith("Key words", StringComparison.OrdinalIgnoreCase) ||
               text.TrimStart().StartsWith("Keywords", StringComparison.OrdinalIgnoreCase);
    })
    .ToList();

Console.WriteLine($"Found {keywordsParas.Count} Key words paragraphs");

// 处理 Run 中的 rPr (w:r > w:rPr)
foreach (var run in body.Descendants<Run>())
{
    var rPr = run.GetFirstChild<RunProperties>();
    if (rPr == null) continue;

    ProcessFontSizeInRunProperties(rPr, false, ref sz22Fixed, ref sz24Fixed, ref szCs22Fixed, ref szCs24Fixed);
}

// 处理 pPr 中的 ParagraphMarkRunProperties (w:pPr > w:rPr)
int pPrCount = 0;
int pPrRprCount = 0;

foreach (var pPr in body.Descendants<ParagraphProperties>())
{
    pPrCount++;

    // 获取 ParagraphMarkRunProperties
    var pPrRpr = pPr.GetFirstChild<ParagraphMarkRunProperties>();
    if (pPrRpr != null)
    {
        pPrRprCount++;

        // 检查是否在 Key words 段落中
        var parentPara = pPr.Parent as Paragraph;
        bool isInKeywordsPara = parentPara != null && keywordsParas.Contains(parentPara);

        ProcessFontSizeInParagraphMarkRunProperties(pPrRpr, isInKeywordsPara, ref sz22Fixed, ref sz24Fixed, ref szCs22Fixed, ref szCs24Fixed);
    }
}

Console.WriteLine($"ParagraphProperties 数量: {pPrCount}");
Console.WriteLine($"ParagraphMarkRunProperties 数量: {pPrRprCount}");

mainPart.Document.Save();

Console.WriteLine($"=== Format Fix Complete ===");
Console.WriteLine($"w:sz sz=22 -> sz=21 fixed: {sz22Fixed}");
Console.WriteLine($"w:szCs sz=22 -> sz=21 fixed: {szCs22Fixed}");
Console.WriteLine($"w:sz Key words sz=24 -> sz=21 fixed: {sz24Fixed}");
Console.WriteLine($"w:szCs Key words sz=24 -> sz=21 fixed: {szCs24Fixed}");
Console.WriteLine($"Total sz=22 fixed: {sz22Fixed + szCs22Fixed}");
Console.WriteLine($"Total sz=24 fixed: {sz24Fixed + szCs24Fixed}");
Console.WriteLine($"Output: {outputPath}");

static void ProcessFontSizeInRunProperties(RunProperties rPr, bool isInKeywordsPara,
    ref int sz22Fixed, ref int sz24Fixed, ref int szCs22Fixed, ref int szCs24Fixed)
{
    // 处理 w:sz
    var szElements = rPr.Elements<FontSize>().ToList();
    foreach (var sz in szElements)
    {
        string val = sz.Val?.Value ?? "";
        if (val == "22")
        {
            sz.Val = "21";
            sz22Fixed++;
        }
        else if (isInKeywordsPara && val == "24")
        {
            sz.Val = "21";
            sz24Fixed++;
        }
    }

    // 处理 w:szCs
    var szCsElements = rPr.Elements<FontSizeComplexScript>().ToList();
    foreach (var szCs in szCsElements)
    {
        string val = szCs.Val?.Value ?? "";
        if (val == "22")
        {
            szCs.Val = "21";
            szCs22Fixed++;
        }
        else if (isInKeywordsPara && val == "24")
        {
            szCs.Val = "21";
            szCs24Fixed++;
        }
    }
}

static void ProcessFontSizeInParagraphMarkRunProperties(ParagraphMarkRunProperties rPr, bool isInKeywordsPara,
    ref int sz22Fixed, ref int sz24Fixed, ref int szCs22Fixed, ref int szCs24Fixed)
{
    // 处理 w:sz
    var szElements = rPr.Elements<FontSize>().ToList();
    foreach (var sz in szElements)
    {
        string val = sz.Val?.Value ?? "";
        if (val == "22")
        {
            sz.Val = "21";
            sz22Fixed++;
        }
        else if (isInKeywordsPara && val == "24")
        {
            sz.Val = "21";
            sz24Fixed++;
        }
    }

    // 处理 w:szCs
    var szCsElements = rPr.Elements<FontSizeComplexScript>().ToList();
    foreach (var szCs in szCsElements)
    {
        string val = szCs.Val?.Value ?? "";
        if (val == "22")
        {
            szCs.Val = "21";
            szCs22Fixed++;
        }
        else if (isInKeywordsPara && val == "24")
        {
            szCs.Val = "21";
            szCs24Fixed++;
        }
    }
}
