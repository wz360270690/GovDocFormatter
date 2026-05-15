using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GovDocFormatter;

public sealed partial class GovDocumentFormatter
{
    private const string FontBody = "仿宋_GB2312";
    private const string FontHeading1 = "黑体";
    private const string FontHeading2 = "楷体_GB2312";
    private const string FontTitle = "方正小标宋简体";
    private const string FontNumber = "Times New Roman";
    private const string FontPageNumber = "宋体";

    private const int A4WidthTwips = 11906;
    private const int A4HeightTwips = 16838;
    private const int MarginTopTwips = 2098;
    private const int MarginBottomTwips = 1984;
    private const int MarginLeftTwips = 1587;
    private const int MarginRightTwips = 1474;
    private const int TextWidthTwips = A4WidthTwips - MarginLeftTwips - MarginRightTwips;
    private const int FixedLineTwips = 580; // 225 mm / 22 lines, about 29 pt
    private const int OneChineseCharTwips = 320; // 16 pt
    private const int TwoChineseCharsTwips = OneChineseCharTwips * 2;
    private const int FourChineseCharsTwips = OneChineseCharTwips * 4;
    private const int TwoChineseCharsInHundredths = 200;
    private const int HeaderTwips = 850; // 15 mm
    private const int FooterTwips = 1587; // 28 mm, page number about 7 mm below text area

    private static readonly ParagraphFormat BodyFormat = new(
        "GovBody",
        FontBody,
        32,
        JustificationValues.Both,
        FirstLine: TwoChineseCharsTwips,
        FirstLineChars: TwoChineseCharsInHundredths);

    private static readonly ParagraphFormat BodyLeftFormat = BodyFormat with
    {
        StyleId = "GovBodyLeft",
        Alignment = JustificationValues.Left,
        FirstLine = 0,
        FirstLineChars = 0
    };

    private static readonly ParagraphFormat Heading1Format = new(
        "GovHeading1",
        FontHeading1,
        32,
        JustificationValues.Left,
        FirstLine: TwoChineseCharsTwips,
        FirstLineChars: TwoChineseCharsInHundredths);

    private static readonly ParagraphFormat Heading2Format = new(
        "GovHeading2",
        FontHeading2,
        32,
        JustificationValues.Left,
        FirstLine: TwoChineseCharsTwips,
        FirstLineChars: TwoChineseCharsInHundredths);

    private static readonly ParagraphFormat Heading3Format = new(
        "GovHeading3",
        FontBody,
        32,
        JustificationValues.Left,
        FirstLine: TwoChineseCharsTwips,
        FirstLineChars: TwoChineseCharsInHundredths);

    private static readonly ParagraphFormat HeaderLeftFormat = new(
        "GovHeaderLeft",
        FontHeading1,
        32,
        JustificationValues.Left);

    private static readonly ParagraphFormat TitleFormat = new(
        "GovTitle",
        FontTitle,
        44,
        JustificationValues.Center);

    private static readonly ParagraphFormat RedMarkFormat = new(
        "GovRedMark",
        FontTitle,
        84,
        JustificationValues.Center,
        Before: FixedLineTwips,
        After: FixedLineTwips,
        Color: "FF0000");

    private static readonly ParagraphFormat DispatchNumberFormat = new(
        "GovDispatchNumber",
        FontBody,
        32,
        JustificationValues.Center,
        Before: FixedLineTwips * 2,
        After: FixedLineTwips);

    private static readonly ParagraphFormat ImprintFormat = new(
        "GovImprint",
        FontBody,
        28,
        JustificationValues.Left,
        Line: 480);

    public FormatResult Format(string inputPath, string outputPath, FormattingOptions options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("找不到输入 Word 文档。", inputPath);
        }

        if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("输出文件不能覆盖原始文件，请选择一个新的 .docx 路径。");
        }

        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.Copy(inputPath, outputPath, overwrite: true);

        var result = new FormatResult { OutputPath = outputPath };

        using var document = WordprocessingDocument.Open(outputPath, true);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("文档缺少 MainDocumentPart，无法处理。");
        mainPart.Document ??= new Document(new Body());
        mainPart.Document.Body ??= new Body();

        EnsureSettings(mainPart);
        EnsureStyles(mainPart);

        var footerReferences = options.ApplyPageNumbers ? AddPageNumberFooters(mainPart) : [];
        ApplyPageSetupToAllSections(mainPart.Document.Body, footerReferences);

        if (options.FormatExistingContent)
        {
            var normalizedBreaks = NormalizeManualLineBreakParagraphs(mainPart.Document.Body);
            if (normalizedBreaks > 0)
            {
                result.Messages.Add($"已将 {normalizedBreaks} 个手动换行片段重组为可识别的公文逻辑段落。");
            }

            var removedPaginationControls = RemovePaginationControls(mainPart.Document.Body);
            if (removedPaginationControls > 0)
            {
                result.Messages.Add($"已清除 {removedPaginationControls} 处原文残留的分页控制。");
            }

            var removedBlankParagraphs = FormatExistingParagraphs(mainPart.Document.Body, options);
            if (removedBlankParagraphs > 0)
            {
                result.Messages.Add($"已删除 {removedBlankParagraphs} 个正文中的空段落。");
            }
        }

        if (options.InsertFrontMatter && options.HasFrontMatter)
        {
            InsertFrontMatter(mainPart.Document.Body, options, result.Messages);
        }

        if (options.InsertEndingMatter && options.HasEndingMatter)
        {
            InsertEndingMatter(mainPart.Document.Body, footerReferences, options, result.Messages);
        }

        mainPart.Document.Save();

        result.Messages.Add("已设置 A4 纸、上 37mm、下 35mm、左 28mm、右 26mm 页边距，版心约 156mm x 225mm。");
        result.Messages.Add("已设置正文三号仿宋_GB2312，并通过文档网格按每面 22 行、每行 28 字撑满版心。");
        result.Messages.Add("已将西文/阿拉伯数字字体设为 Times New Roman。");

        if (options.ApplyPageNumbers)
        {
            result.Messages.Add("已添加单双页外侧页码；首次打开 Word 时会自动更新 PAGE 域。");
        }

        result.Messages.Add("标题多行梯形/菱形、印章压署名日期、双面打印套正和装订眼位置仍需在打印或盖章环节复核。");
        return result;
    }

    private static void EnsureSettings(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();

        settingsPart.Settings.RemoveAllChildren<EvenAndOddHeaders>();
        settingsPart.Settings.Append(new EvenAndOddHeaders());

        settingsPart.Settings.RemoveAllChildren<UpdateFieldsOnOpen>();
        settingsPart.Settings.Append(new UpdateFieldsOnOpen { Val = new OnOffValue(true) });

        settingsPart.Settings.Save();
    }

    private static void EnsureStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        EnsureNormalStyle(stylesPart.Styles);

        AddOrReplaceStyle(stylesPart.Styles, BodyFormat, "正文-公文");
        AddOrReplaceStyle(stylesPart.Styles, BodyLeftFormat, "正文顶格-公文");
        AddOrReplaceStyle(stylesPart.Styles, Heading1Format, "一级标题-公文");
        AddOrReplaceStyle(stylesPart.Styles, Heading2Format, "二级标题-公文");
        AddOrReplaceStyle(stylesPart.Styles, Heading3Format, "三四级标题-公文");
        AddOrReplaceStyle(stylesPart.Styles, HeaderLeftFormat, "版头顶格-公文");
        AddOrReplaceStyle(stylesPart.Styles, TitleFormat, "标题-公文");
        AddOrReplaceStyle(stylesPart.Styles, RedMarkFormat, "发文机关标志-公文");
        AddOrReplaceStyle(stylesPart.Styles, DispatchNumberFormat, "发文字号-公文");
        AddOrReplaceStyle(stylesPart.Styles, ImprintFormat, "版记-公文");

        stylesPart.Styles.Save();
    }

    private static void EnsureNormalStyle(Styles styles)
    {
        var normal = styles.Elements<Style>().FirstOrDefault(style => style.StyleId?.Value == "Normal");
        if (normal is null)
        {
            normal = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            };
            normal.Append(new StyleName { Val = "Normal" });
            styles.Append(normal);
        }
    }

    private static void AddOrReplaceStyle(Styles styles, ParagraphFormat format, string name)
    {
        foreach (var oldStyle in styles.Elements<Style>().Where(style => style.StyleId?.Value == format.StyleId).ToList())
        {
            oldStyle.Remove();
        }

        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = format.StyleId,
            CustomStyle = true
        };

        style.Append(new StyleName { Val = name });
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = format.StyleId });
        style.Append(new PrimaryStyle());

        var paragraphProperties = new StyleParagraphProperties();
        paragraphProperties.Append(CreateSpacing(format));

        var indentation = CreateIndentation(format);
        if (indentation is not null)
        {
            paragraphProperties.Append(indentation);
        }

        paragraphProperties.Append(CreateJustification(format.Alignment));
        style.Append(paragraphProperties);

        var runProperties = new StyleRunProperties();
        ApplyRunProperties(runProperties, format);
        style.Append(runProperties);

        styles.Append(style);
    }

    private static List<FooterReference> AddPageNumberFooters(MainDocumentPart mainPart)
    {
        var oddFooterPart = mainPart.AddNewPart<FooterPart>();
        oddFooterPart.Footer = BuildPageNumberFooter(JustificationValues.Right, leadingFullWidthSpace: false);
        oddFooterPart.Footer.Save();

        var evenFooterPart = mainPart.AddNewPart<FooterPart>();
        evenFooterPart.Footer = BuildPageNumberFooter(JustificationValues.Left, leadingFullWidthSpace: true);
        evenFooterPart.Footer.Save();

        return
        [
            new FooterReference
            {
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(oddFooterPart)
            },
            new FooterReference
            {
                Type = HeaderFooterValues.Even,
                Id = mainPart.GetIdOfPart(evenFooterPart)
            }
        ];
    }

    private static Footer BuildPageNumberFooter(JustificationValues alignment, bool leadingFullWidthSpace)
    {
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties(
            new Justification { Val = alignment },
            new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });

        paragraph.Append(paragraphProperties);

        if (leadingFullWidthSpace)
        {
            paragraph.Append(TextRun("　", FontPageNumber, 28));
        }

        paragraph.Append(TextRun("—", FontPageNumber, 28));
        var field = new SimpleField { Instruction = " PAGE " };
        field.Append(TextRun("1", FontPageNumber, 28));
        paragraph.Append(field);
        paragraph.Append(TextRun("—", FontPageNumber, 28));

        if (!leadingFullWidthSpace)
        {
            paragraph.Append(TextRun("　", FontPageNumber, 28));
        }

        return new Footer(paragraph);
    }

    private static void ApplyPageSetupToAllSections(Body body, IReadOnlyList<FooterReference> footerReferences)
    {
        var sectionProperties = body.Descendants<SectionProperties>().ToList();
        var bodySection = body.Elements<SectionProperties>().LastOrDefault();

        if (bodySection is null)
        {
            bodySection = new SectionProperties();
            body.Append(bodySection);
            sectionProperties.Add(bodySection);
        }
        else if (!sectionProperties.Contains(bodySection))
        {
            sectionProperties.Add(bodySection);
        }

        foreach (var section in sectionProperties)
        {
            ApplyBaseSectionProperties(section, footerReferences, sectionType: null);
        }
    }

    private static void ApplyBaseSectionProperties(
        SectionProperties section,
        IReadOnlyList<FooterReference> footerReferences,
        SectionMarkValues? sectionType)
    {
        section.RemoveAllChildren<SectionType>();
        section.RemoveAllChildren<PageSize>();
        section.RemoveAllChildren<PageMargin>();
        section.RemoveAllChildren<DocGrid>();
        section.RemoveAllChildren<FooterReference>();

        if (sectionType.HasValue)
        {
            section.Append(new SectionType { Val = sectionType.Value });
        }

        foreach (var reference in footerReferences)
        {
            section.Append((FooterReference)reference.CloneNode(true));
        }

        section.Append(new PageSize
        {
            Width = UInt32Value.FromUInt32((uint)A4WidthTwips),
            Height = UInt32Value.FromUInt32((uint)A4HeightTwips)
        });

        section.Append(new PageMargin
        {
            Top = MarginTopTwips,
            Bottom = MarginBottomTwips,
            Left = UInt32Value.FromUInt32((uint)MarginLeftTwips),
            Right = UInt32Value.FromUInt32((uint)MarginRightTwips),
            Header = UInt32Value.FromUInt32((uint)HeaderTwips),
            Footer = UInt32Value.FromUInt32((uint)FooterTwips),
            Gutter = UInt32Value.FromUInt32(0U)
        });

        section.Append(new DocGrid
        {
            Type = DocGridValues.LinesAndChars,
            LinePitch = FixedLineTwips,
            CharacterSpace = 316
        });
    }

    private static int FormatExistingParagraphs(Body body, FormattingOptions options)
    {
        var titleAlreadyHandled = !string.IsNullOrWhiteSpace(options.Title);
        var mainRecipientAlreadyHandled = !string.IsNullOrWhiteSpace(options.MainRecipient);
        var numberingState = new NumberingState();
        var removedBlankParagraphs = 0;

        foreach (var paragraph in body.Elements<Paragraph>().ToList())
        {
            if (paragraph.Ancestors<Footer>().Any())
            {
                continue;
            }

            var numberingLevel = GetNumberingLevel(paragraph);
            var text = NormalizeText(paragraph.InnerText);
            if (string.IsNullOrWhiteSpace(text))
            {
                if (HasSectionProperties(paragraph))
                {
                    ApplyParagraphFormat(paragraph, BodyFormat with { FirstLine = 0 });
                    continue;
                }

                paragraph.Remove();
                removedBlankParagraphs++;
                continue;
            }

            var mayBeTitle = options.AutoDetectTitle && !titleAlreadyHandled;
            var hasLeadInHeading = TryGetLeadInHeading(text, out var leadInHeadingFormat, out var leadInLength);
            var detected = hasLeadInHeading
                ? new ParagraphClassification(BodyFormat, ParagraphKind.Body)
                : ClassifyParagraph(text, numberingLevel, mayBeTitle);
            if (numberingLevel.HasValue)
            {
                if (!StartsWithHierarchyMarker(text))
                {
                    PrefixParagraphText(paragraph, numberingState.NextPrefix(numberingLevel.Value));
                }

                RemoveParagraphNumbering(paragraph);
            }

            var format = detected.Format;
            if (format.StyleId == TitleFormat.StyleId)
            {
                titleAlreadyHandled = true;
            }
            else if (titleAlreadyHandled && !mainRecipientAlreadyHandled && IsMainRecipientCandidate(text))
            {
                format = BodyLeftFormat;
                mainRecipientAlreadyHandled = true;
            }

            ApplyParagraphFormat(paragraph, format);
            if (hasLeadInHeading)
            {
                ApplyCharacterRangeRunFormat(paragraph, 0, leadInLength, leadInHeadingFormat);
            }
        }

        return removedBlankParagraphs;
    }

    private static int NormalizeManualLineBreakParagraphs(Body body)
    {
        var changed = 0;
        foreach (var paragraph in body.Elements<Paragraph>().ToList())
        {
            var manualLines = ExtractManualBreakLines(paragraph);
            if (manualLines.Count <= 1)
            {
                continue;
            }

            var logicalParagraphs = BuildLogicalParagraphs(manualLines);
            if (logicalParagraphs.Count <= 1)
            {
                continue;
            }

            var sectionProperties = paragraph.GetFirstChild<ParagraphProperties>()?.GetFirstChild<SectionProperties>()?.CloneNode(true);
            var newParagraphs = logicalParagraphs.Select(BuildPlainParagraph).ToList();
            if (sectionProperties is not null)
            {
                var lastProperties = newParagraphs[^1].GetFirstChild<ParagraphProperties>();
                if (lastProperties is null)
                {
                    lastProperties = new ParagraphProperties();
                    newParagraphs[^1].PrependChild(lastProperties);
                }

                lastProperties.Append(sectionProperties);
            }

            foreach (var newParagraph in newParagraphs)
            {
                body.InsertBefore(newParagraph, paragraph);
            }

            paragraph.Remove();
            changed += manualLines.Count;
        }

        return changed;
    }

    private static int RemovePaginationControls(Body body)
    {
        var removed = 0;

        foreach (var paragraph in body.Elements<Paragraph>().ToList())
        {
            var paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
            if (paragraphProperties is not null)
            {
                removed += RemoveAllChildrenAndCount<KeepNext>(paragraphProperties);
                removed += RemoveAllChildrenAndCount<KeepLines>(paragraphProperties);
                removed += RemoveAllChildrenAndCount<PageBreakBefore>(paragraphProperties);
                removed += RemoveAllChildrenAndCount<WidowControl>(paragraphProperties);
            }

            foreach (var pageBreak in paragraph.Descendants<Break>()
                         .Where(IsManualPageOrColumnBreak)
                         .ToList())
            {
                pageBreak.Remove();
                removed++;
            }

            foreach (var renderedPageBreak in paragraph.Descendants<LastRenderedPageBreak>().ToList())
            {
                renderedPageBreak.Remove();
            }

            RemoveEmptyRuns(paragraph);
        }

        return removed;
    }

    private static bool IsManualPageOrColumnBreak(Break pageBreak)
    {
        var breakType = pageBreak.Type?.Value;
        return breakType is not null &&
               (breakType.Equals(BreakValues.Page) || breakType.Equals(BreakValues.Column));
    }

    private static List<string> ExtractManualBreakLines(Paragraph paragraph)
    {
        var lines = new List<string>();
        var current = new StringBuilder();
        var sawBreak = false;

        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case Text text:
                    current.Append(text.Text);
                    break;
                case TabChar:
                    current.Append('\t');
                    break;
                case Break:
                case CarriageReturn:
                    lines.Add(NormalizeText(current.ToString()));
                    current.Clear();
                    sawBreak = true;
                    break;
            }
        }

        if (!sawBreak)
        {
            return [NormalizeText(paragraph.InnerText)];
        }

        lines.Add(NormalizeText(current.ToString()));
        return lines.Where(line => !string.IsNullOrWhiteSpace(line)).ToList();
    }

    private static List<string> BuildLogicalParagraphs(IReadOnlyList<string> lines)
    {
        var paragraphs = new List<string>();
        var current = new StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = NormalizeText(rawLine);
            if (string.IsNullOrWhiteSpace(line))
            {
                FlushCurrent();
                continue;
            }

            if (ShouldStartNewLogicalParagraph(current.ToString(), line))
            {
                FlushCurrent();
            }

            current.Append(line);

            if (IsStandaloneLogicalLine(line))
            {
                FlushCurrent();
            }
        }

        FlushCurrent();
        return paragraphs;

        void FlushCurrent()
        {
            var text = NormalizeText(current.ToString());
            if (!string.IsNullOrWhiteSpace(text))
            {
                paragraphs.Add(text);
            }

            current.Clear();
        }
    }

    private static bool ShouldStartNewLogicalParagraph(string currentText, string nextLine)
    {
        var current = NormalizeText(currentText);
        if (string.IsNullOrWhiteSpace(current))
        {
            return false;
        }

        if (StartsWithHierarchyMarker(nextLine))
        {
            return true;
        }

        if (IsMainRecipientCandidate(current))
        {
            return true;
        }

        if (LooksLikeDocumentTitle(current) && IsMainRecipientCandidate(nextLine))
        {
            return true;
        }

        if (LooksLikeEndingMatter(nextLine))
        {
            return true;
        }

        return false;
    }

    private static bool IsStandaloneLogicalLine(string line)
    {
        return LooksLikeDocumentTitle(line) ||
               IsMainRecipientCandidate(line) ||
               (LevelOneHeadingRegex().IsMatch(line) && line.Length <= 24) ||
               LooksLikeEndingMatter(line);
    }

    private static bool LooksLikeDocumentTitle(string text)
    {
        return IsTitleCandidate(text) &&
               !StartsWithHierarchyMarker(text) &&
               !IsMainRecipientCandidate(text) &&
               !text.Contains('，') &&
               !text.Contains(',') &&
               !text.Contains('、') &&
               DocumentTitleSuffixRegex().IsMatch(text);
    }

    private static bool LooksLikeEndingMatter(string text)
    {
        return text.StartsWith("抄送", StringComparison.Ordinal) ||
               text.StartsWith("联系人", StringComparison.Ordinal) ||
               text.StartsWith("联系电话", StringComparison.Ordinal) ||
               text.Contains("印发", StringComparison.Ordinal) ||
               DateLineRegex().IsMatch(text);
    }

    private static Paragraph BuildPlainParagraph(string text)
    {
        return new Paragraph(TextRun(text, FontBody, 32));
    }

    private static bool HasSectionProperties(Paragraph paragraph)
    {
        return paragraph.GetFirstChild<ParagraphProperties>()?.GetFirstChild<SectionProperties>() is not null;
    }

    private static ParagraphClassification ClassifyParagraph(string text, int? numberingLevel, bool mayBeTitle)
    {
        if (numberingLevel.HasValue)
        {
            return ClassifyNumberedParagraph(text, numberingLevel.Value);
        }

        if (LevelOneHeadingRegex().IsMatch(text))
        {
            return new ParagraphClassification(Heading1Format, ParagraphKind.Heading1);
        }

        if (LevelTwoHeadingRegex().IsMatch(text))
        {
            return new ParagraphClassification(Heading2Format, ParagraphKind.Heading2);
        }

        if (LevelThreeHeadingRegex().IsMatch(text) && IsStandaloneNumberedSubheading(text))
        {
            return new ParagraphClassification(Heading3Format, ParagraphKind.Heading3);
        }

        if (LevelFourHeadingRegex().IsMatch(text) && IsStandaloneNumberedSubheading(text))
        {
            return new ParagraphClassification(Heading3Format, ParagraphKind.Heading4);
        }

        if (AttachmentRegex().IsMatch(text))
        {
            return new ParagraphClassification(BodyFormat, ParagraphKind.Body);
        }

        if (mayBeTitle && IsTitleCandidate(text))
        {
            return new ParagraphClassification(TitleFormat, ParagraphKind.Title);
        }

        return new ParagraphClassification(BodyFormat, ParagraphKind.Body);
    }

    private static ParagraphClassification ClassifyNumberedParagraph(string text, int numberingLevel)
    {
        return numberingLevel switch
        {
            0 => new ParagraphClassification(Heading1Format, ParagraphKind.Heading1),
            1 => new ParagraphClassification(Heading2Format, ParagraphKind.Heading2),
            2 => IsStandaloneNumberedSubheadingText(text)
                ? new ParagraphClassification(Heading3Format, ParagraphKind.Heading3)
                : new ParagraphClassification(BodyFormat, ParagraphKind.Body),
            3 => IsStandaloneNumberedSubheadingText(text)
                ? new ParagraphClassification(Heading3Format, ParagraphKind.Heading4)
                : new ParagraphClassification(BodyFormat, ParagraphKind.Body),
            _ => new ParagraphClassification(BodyFormat, ParagraphKind.Body)
        };
    }

    private static bool IsNumberedSubheadingMarker(string text)
    {
        return LevelThreeHeadingRegex().IsMatch(text) || LevelFourHeadingRegex().IsMatch(text);
    }

    private static bool IsStandaloneNumberedSubheading(string text)
    {
        return IsStandaloneNumberedSubheadingText(RemoveNumberedSubheadingMarker(text));
    }

    private static bool IsStandaloneNumberedSubheadingText(string text)
    {
        var headingText = NormalizeText(text);
        return headingText.Length is > 0 and <= 32 &&
               !ContainsBodyPunctuation(headingText);
    }

    private static bool IsCompactNumberedLeadInHeading(string text)
    {
        var headingText = NormalizeText(RemoveNumberedSubheadingMarker(text)).TrimEnd('。');
        return headingText.Length is > 0 and <= 24 &&
               !ContainsBodyPunctuation(headingText);
    }

    private static string RemoveNumberedSubheadingMarker(string text)
    {
        var levelThreeMatch = LevelThreeHeadingRegex().Match(text);
        if (levelThreeMatch.Success)
        {
            return text[levelThreeMatch.Length..];
        }

        var levelFourMatch = LevelFourHeadingRegex().Match(text);
        if (levelFourMatch.Success)
        {
            return text[levelFourMatch.Length..];
        }

        return text;
    }

    private static bool ContainsBodyPunctuation(string text)
    {
        return text.IndexOfAny(['，', ',', '。', '；', ';', '：', ':', '、']) >= 0;
    }

    private static bool TryGetLeadInHeading(string text, out ParagraphFormat headingFormat, out int headingLength)
    {
        headingFormat = BodyFormat;
        headingLength = 0;

        if (!TryGetHierarchyFormat(text, out headingFormat))
        {
            return false;
        }

        var stopIndex = text.IndexOf('。');
        if (stopIndex < 0 || stopIndex >= text.Length - 1)
        {
            return false;
        }

        var headingText = text[..(stopIndex + 1)];
        var remainingText = text[(stopIndex + 1)..].Trim();
        if (headingText.Length > 50 || string.IsNullOrWhiteSpace(remainingText))
        {
            return false;
        }

        if (IsNumberedSubheadingMarker(text) && !IsCompactNumberedLeadInHeading(headingText))
        {
            return false;
        }

        headingLength = stopIndex + 1;
        return true;
    }

    private static bool TryGetHierarchyFormat(string text, out ParagraphFormat format)
    {
        if (LevelOneHeadingRegex().IsMatch(text))
        {
            format = Heading1Format;
            return true;
        }

        if (LevelTwoHeadingRegex().IsMatch(text))
        {
            format = Heading2Format;
            return true;
        }

        if (LevelThreeHeadingRegex().IsMatch(text))
        {
            format = Heading3Format;
            return true;
        }

        if (LevelFourHeadingRegex().IsMatch(text))
        {
            format = Heading3Format;
            return true;
        }

        format = BodyFormat;
        return false;
    }

    private static bool IsTitleCandidate(string text)
    {
        return text.Length <= 60 &&
               !text.Contains('。') &&
               !text.EndsWith('：') &&
               !text.EndsWith(':') &&
               !LevelOneHeadingRegex().IsMatch(text) &&
               !LevelTwoHeadingRegex().IsMatch(text) &&
               !LevelThreeHeadingRegex().IsMatch(text) &&
               !LevelFourHeadingRegex().IsMatch(text);
    }

    private static bool IsMainRecipientCandidate(string text)
    {
        return text.Length <= 80 &&
               (text.EndsWith('：') || text.EndsWith(':')) &&
               !StartsWithHierarchyMarker(text);
    }

    private static void InsertFrontMatter(Body body, FormattingOptions options, ICollection<string> messages)
    {
        var elements = new List<OpenXmlElement>();

        if (!string.IsNullOrWhiteSpace(options.SerialNumber))
        {
            elements.Add(BuildParagraph(options.SerialNumber.PadLeft(6, '0'), BodyLeftFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.SecretLevel))
        {
            elements.Add(BuildParagraph(options.SecretLevel, HeaderLeftFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.Urgency))
        {
            elements.Add(BuildParagraph(options.Urgency, HeaderLeftFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.DispatchMark))
        {
            elements.Add(BuildParagraph(options.DispatchMark, RedMarkFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.DispatchNumber))
        {
            elements.Add(options.UpwardDocument && !string.IsNullOrWhiteSpace(options.SignerName)
                ? BuildDispatchAndSignerParagraph(options.DispatchNumber, options.SignerName)
                : BuildParagraph(options.DispatchNumber, DispatchNumberFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.DispatchMark) || !string.IsNullOrWhiteSpace(options.DispatchNumber))
        {
            elements.Add(BuildRedSeparatorParagraph());
        }

        if (!string.IsNullOrWhiteSpace(options.Title))
        {
            elements.Add(BuildParagraph(options.Title, TitleFormat));
        }

        if (!string.IsNullOrWhiteSpace(options.MainRecipient))
        {
            elements.Add(BuildParagraph(options.MainRecipient.TrimEnd('：', ':') + "：", BodyLeftFormat));
        }

        if (elements.Count == 0)
        {
            return;
        }

        var reference = body.ChildElements.FirstOrDefault(element => element is not SectionProperties);
        foreach (var element in elements)
        {
            if (reference is null)
            {
                body.Append(element);
            }
            else
            {
                body.InsertBefore(element, reference);
            }
        }

        messages.Add("已按填写内容插入份号/密级/紧急程度、发文机关标志、发文字号、标题和主送机关。");
    }

    private static Paragraph BuildDispatchAndSignerParagraph(string dispatchNumber, string signerName)
    {
        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(CreateSpacing(DispatchNumberFormat));
        paragraphProperties.Append(new Justification { Val = JustificationValues.Left });
        paragraphProperties.Append(new Indentation { Left = OneChineseCharTwips.ToString(), Right = OneChineseCharTwips.ToString() });
        paragraphProperties.Append(new Tabs(new TabStop
        {
            Val = TabStopValues.Right,
            Position = Int32Value.FromInt32(TextWidthTwips - OneChineseCharTwips)
        }));
        paragraph.Append(paragraphProperties);
        paragraph.Append(TextRun(dispatchNumber, FontBody, 32));
        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(TextRun("签发人：", FontBody, 32));
        paragraph.Append(TextRun(signerName, FontHeading2, 32));
        return paragraph;
    }

    private static Paragraph BuildRedSeparatorParagraph()
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        properties.Append(new SpacingBetweenLines
        {
            Before = "0",
            After = (FixedLineTwips * 2).ToString(),
            Line = "120",
            LineRule = LineSpacingRuleValues.Exact
        });
        properties.Append(new ParagraphBorders(new BottomBorder
        {
            Val = BorderValues.Single,
            Color = "FF0000",
            Size = 12,
            Space = 1
        }));
        paragraph.Append(properties);
        return paragraph;
    }

    private static void InsertEndingMatter(
        Body body,
        IReadOnlyList<FooterReference> footerReferences,
        FormattingOptions options,
        ICollection<string> messages)
    {
        var elements = new List<OpenXmlElement>();

        if (!string.IsNullOrWhiteSpace(options.Issuer))
        {
            elements.Add(BuildParagraph(options.Issuer, BodyFormat with
            {
                Alignment = JustificationValues.Right,
                FirstLine = 0,
                FirstLineChars = 0,
                Right = TwoChineseCharsTwips,
                Before = FixedLineTwips
            }));
        }

        if (!string.IsNullOrWhiteSpace(options.DocumentDate))
        {
            elements.Add(BuildParagraph(options.DocumentDate, BodyFormat with
            {
                Alignment = JustificationValues.Right,
                FirstLine = 0,
                FirstLineChars = 0,
                Right = TwoChineseCharsTwips
            }));
        }

        if (!string.IsNullOrWhiteSpace(options.CopyTo) ||
            !string.IsNullOrWhiteSpace(options.PrintingOffice) ||
            !string.IsNullOrWhiteSpace(options.PrintingDate))
        {
            elements.Add(BuildEvenPageSectionBreak(footerReferences));
            elements.Add(BuildImprintSeparator());

            if (!string.IsNullOrWhiteSpace(options.CopyTo))
            {
                var copyToText = options.CopyTo.Trim();
                if (!copyToText.EndsWith('。'))
                {
                    copyToText += "。";
                }

                elements.Add(BuildParagraph("抄送：" + copyToText, ImprintFormat with
                {
                    Left = OneChineseCharTwips,
                    Right = OneChineseCharTwips
                }));
                elements.Add(BuildImprintSeparator());
            }

            if (!string.IsNullOrWhiteSpace(options.PrintingOffice) || !string.IsNullOrWhiteSpace(options.PrintingDate))
            {
                elements.Add(BuildPrintingOfficeParagraph(options.PrintingOffice, options.PrintingDate));
                elements.Add(BuildImprintSeparator());
            }

            messages.Add("版记前已插入偶数页分节符，尽量保证版记从偶数页开始。");
        }

        var reference = body.Elements<SectionProperties>().LastOrDefault();
        foreach (var element in elements)
        {
            if (reference is null)
            {
                body.Append(element);
            }
            else
            {
                body.InsertBefore(element, reference);
            }
        }

        messages.Add("已按填写内容插入发文机关署名、成文日期和版记。");
    }

    private static Paragraph BuildEvenPageSectionBreak(IReadOnlyList<FooterReference> footerReferences)
    {
        var sectionProperties = new SectionProperties();
        ApplyBaseSectionProperties(sectionProperties, footerReferences, SectionMarkValues.EvenPage);
        return new Paragraph(new ParagraphProperties(sectionProperties));
    }

    private static Paragraph BuildImprintSeparator()
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        properties.Append(new SpacingBetweenLines
        {
            Before = "0",
            After = "0",
            Line = "80",
            LineRule = LineSpacingRuleValues.Exact
        });
        properties.Append(new ParagraphBorders(new BottomBorder
        {
            Val = BorderValues.Single,
            Color = "000000",
            Size = 4,
            Space = 1
        }));
        paragraph.Append(properties);
        return paragraph;
    }

    private static Paragraph BuildPrintingOfficeParagraph(string office, string date)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        properties.Append(CreateSpacing(ImprintFormat));
        properties.Append(new Justification { Val = JustificationValues.Left });
        properties.Append(new Indentation
        {
            Left = OneChineseCharTwips.ToString(),
            Right = OneChineseCharTwips.ToString()
        });
        properties.Append(new Tabs(new TabStop
        {
            Val = TabStopValues.Right,
            Position = Int32Value.FromInt32(TextWidthTwips - OneChineseCharTwips)
        }));
        paragraph.Append(properties);
        paragraph.Append(TextRun(office, FontBody, 28));
        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(TextRun(date, FontBody, 28));
        return paragraph;
    }

    private static Paragraph BuildParagraph(string text, ParagraphFormat format)
    {
        var paragraph = new Paragraph();
        ApplyParagraphFormat(paragraph, format);
        paragraph.Append(TextRun(text, format.EastAsiaFont, format.FontSizeHalfPoints, format.Bold, format.Color));
        return paragraph;
    }

    private static void ApplyParagraphFormat(Paragraph paragraph, ParagraphFormat format)
    {
        var paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
        if (paragraphProperties is null)
        {
            paragraphProperties = new ParagraphProperties();
            paragraph.PrependChild(paragraphProperties);
        }

        ReplaceOrAppend(paragraphProperties, new ParagraphStyleId { Val = format.StyleId });
        ReplaceOrAppend(paragraphProperties, CreateSpacing(format));

        paragraphProperties.RemoveAllChildren<Indentation>();
        var indentation = CreateIndentation(format);
        if (indentation is not null)
        {
            paragraphProperties.Append(indentation);
        }

        ReplaceOrAppend(paragraphProperties, CreateJustification(format.Alignment));

        foreach (var run in paragraph.Descendants<Run>())
        {
            ApplyRunFormat(run, format);
        }
    }

    private static void ApplyCharacterRangeRunFormat(Paragraph paragraph, int start, int length, ParagraphFormat format)
    {
        if (length <= 0)
        {
            return;
        }

        var end = start + length;
        var position = 0;

        foreach (var text in paragraph.Descendants<Text>().ToList())
        {
            var value = text.Text ?? "";
            var textStart = position;
            var textEnd = position + value.Length;
            position = textEnd;

            if (textEnd <= start || textStart >= end || value.Length == 0)
            {
                continue;
            }

            var run = text.Ancestors<Run>().FirstOrDefault();
            if (run is null)
            {
                continue;
            }

            var localStart = Math.Max(start - textStart, 0);
            var localEnd = Math.Min(end - textStart, value.Length);

            if (localStart == 0 && localEnd == value.Length)
            {
                ApplyRunFormat(run, format);
                continue;
            }

            SplitTextRun(run, text, localStart, localEnd, format);
        }
    }

    private static void SplitTextRun(Run originalRun, Text originalText, int localStart, int localEnd, ParagraphFormat highlightFormat)
    {
        var value = originalText.Text ?? "";
        if (localStart < 0 || localEnd > value.Length || localStart >= localEnd)
        {
            return;
        }

        var before = value[..localStart];
        var highlighted = value[localStart..localEnd];
        var after = value[localEnd..];

        var parent = originalRun.Parent;
        if (parent is null)
        {
            return;
        }

        OpenXmlElement insertBefore = originalRun;

        if (before.Length > 0)
        {
            parent.InsertBefore(CreateRunLike(originalRun, before), insertBefore);
        }

        var highlightedRun = CreateRunLike(originalRun, highlighted);
        ApplyRunFormat(highlightedRun, highlightFormat);
        parent.InsertBefore(highlightedRun, insertBefore);

        if (after.Length > 0)
        {
            parent.InsertBefore(CreateRunLike(originalRun, after), insertBefore);
        }

        originalRun.Remove();
    }

    private static Run CreateRunLike(Run source, string text)
    {
        var run = new Run();
        var runProperties = source.GetFirstChild<RunProperties>();
        if (runProperties is not null)
        {
            run.Append((RunProperties)runProperties.CloneNode(true));
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static Run TextRun(string text, string eastAsiaFont, int fontSizeHalfPoints, bool bold = false, string color = "000000")
    {
        var run = new Run();
        var runProperties = new RunProperties();
        ApplyRunProperties(runProperties, eastAsiaFont, fontSizeHalfPoints, bold, color);
        run.Append(runProperties);

        var parts = text.Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < parts.Length; i++)
        {
            if (i > 0)
            {
                run.Append(new Break());
            }

            run.Append(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
        }

        return run;
    }

    private static void ApplyRunFormat(Run run, ParagraphFormat format)
    {
        var runProperties = run.GetFirstChild<RunProperties>();
        if (runProperties is null)
        {
            runProperties = new RunProperties();
            run.PrependChild(runProperties);
        }

        ApplyRunProperties(runProperties, format);
    }

    private static void ApplyRunProperties(StyleRunProperties runProperties, ParagraphFormat format)
    {
        ApplyRunProperties(runProperties, format.EastAsiaFont, format.FontSizeHalfPoints, format.Bold, format.Color);
    }

    private static void ApplyRunProperties(RunProperties runProperties, ParagraphFormat format)
    {
        ApplyRunProperties(runProperties, format.EastAsiaFont, format.FontSizeHalfPoints, format.Bold, format.Color);
    }

    private static void ApplyRunProperties(OpenXmlCompositeElement runProperties, string eastAsiaFont, int fontSizeHalfPoints, bool bold, string color)
    {
        runProperties.RemoveAllChildren<RunFonts>();
        runProperties.RemoveAllChildren<FontSize>();
        runProperties.RemoveAllChildren<FontSizeComplexScript>();
        runProperties.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Color>();
        runProperties.RemoveAllChildren<Bold>();
        runProperties.RemoveAllChildren<BoldComplexScript>();
        runProperties.RemoveAllChildren<Languages>();

        runProperties.Append(new RunFonts
        {
            Ascii = FontNumber,
            HighAnsi = FontNumber,
            EastAsia = eastAsiaFont,
            ComplexScript = FontNumber
        });
        runProperties.Append(new FontSize { Val = fontSizeHalfPoints.ToString() });
        runProperties.Append(new FontSizeComplexScript { Val = fontSizeHalfPoints.ToString() });
        runProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = color });
        runProperties.Append(new Languages { Val = "zh-CN", EastAsia = "zh-CN" });

        if (bold)
        {
            runProperties.Append(new Bold());
            runProperties.Append(new BoldComplexScript());
        }
    }

    private static SpacingBetweenLines CreateSpacing(ParagraphFormat format)
    {
        return new SpacingBetweenLines
        {
            Before = format.Before.ToString(),
            After = format.After.ToString(),
            Line = format.Line.ToString(),
            LineRule = LineSpacingRuleValues.Exact
        };
    }

    private static Justification CreateJustification(JustificationValues alignment)
    {
        return new Justification { Val = alignment };
    }

    private static Indentation? CreateIndentation(ParagraphFormat format)
    {
        if (format.FirstLine == 0 && format.FirstLineChars == 0 && format.Left == 0 && format.Right == 0)
        {
            return null;
        }

        var indentation = new Indentation();

        if (format.FirstLine != 0)
        {
            indentation.FirstLine = format.FirstLine.ToString();
        }

        if (format.FirstLineChars != 0)
        {
            indentation.FirstLineChars = Int32Value.FromInt32(format.FirstLineChars);
        }

        if (format.Left != 0)
        {
            indentation.Left = format.Left.ToString();
        }

        if (format.Right != 0)
        {
            indentation.Right = format.Right.ToString();
        }

        return indentation;
    }

    private static void ReplaceOrAppend<T>(OpenXmlCompositeElement parent, T child) where T : OpenXmlElement
    {
        parent.RemoveAllChildren<T>();
        parent.Append(child);
    }

    private static int RemoveAllChildrenAndCount<T>(OpenXmlCompositeElement parent) where T : OpenXmlElement
    {
        var children = parent.Elements<T>().ToList();
        foreach (var child in children)
        {
            child.Remove();
        }

        return children.Count;
    }

    private static void RemoveEmptyRuns(Paragraph paragraph)
    {
        foreach (var run in paragraph.Descendants<Run>().ToList())
        {
            if (run.ChildElements.All(child => child is RunProperties))
            {
                run.Remove();
            }
        }
    }

    private static string NormalizeText(string text)
    {
        return LeadingWhitespaceRegex().Replace(text.Replace('\u00A0', ' '), "").Trim();
    }

    private static int? GetNumberingLevel(Paragraph paragraph)
    {
        var numberingProperties = paragraph.GetFirstChild<ParagraphProperties>()?.GetFirstChild<NumberingProperties>();
        if (numberingProperties is null)
        {
            return null;
        }

        var level = numberingProperties.NumberingLevelReference?.Val?.Value;
        return level is null ? 0 : Math.Clamp(level.Value, 0, 8);
    }

    private static void RemoveParagraphNumbering(Paragraph paragraph)
    {
        paragraph.GetFirstChild<ParagraphProperties>()?.RemoveAllChildren<NumberingProperties>();
    }

    private static bool StartsWithHierarchyMarker(string text)
    {
        return LevelOneHeadingRegex().IsMatch(text) ||
               LevelTwoHeadingRegex().IsMatch(text) ||
               LevelThreeHeadingRegex().IsMatch(text) ||
               LevelFourHeadingRegex().IsMatch(text);
    }

    private static void PrefixParagraphText(Paragraph paragraph, string prefix)
    {
        var firstText = paragraph.Descendants<Text>().FirstOrDefault();
        if (firstText is not null)
        {
            firstText.Text = prefix + NormalizeText(firstText.Text);
            firstText.Space = SpaceProcessingModeValues.Preserve;
            return;
        }

        paragraph.Append(TextRun(prefix, FontBody, 32));
    }

    [GeneratedRegex(@"^[一二三四五六七八九十百]+、")]
    private static partial Regex LevelOneHeadingRegex();

    [GeneratedRegex(@"^[（(][一二三四五六七八九十百]+[）)]")]
    private static partial Regex LevelTwoHeadingRegex();

    [GeneratedRegex(@"^[0-9０-９]+[\.．、]")]
    private static partial Regex LevelThreeHeadingRegex();

    [GeneratedRegex(@"^[（(][0-9０-９]+[）)]")]
    private static partial Regex LevelFourHeadingRegex();

    [GeneratedRegex(@"^附件[:：]")]
    private static partial Regex AttachmentRegex();

    [GeneratedRegex(@"^[\s　\t]+")]
    private static partial Regex LeadingWhitespaceRegex();

    [GeneratedRegex(@"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")]
    private static partial Regex DateLineRegex();

    [GeneratedRegex(@"(通知|报告|请示|批复|函|意见|决定|通报|方案)$")]
    private static partial Regex DocumentTitleSuffixRegex();

    private sealed class NumberingState
    {
        private readonly int[] _counts = new int[9];

        public string NextPrefix(int level)
        {
            level = Math.Clamp(level, 0, 8);
            _counts[level]++;
            for (var i = level + 1; i < _counts.Length; i++)
            {
                _counts[i] = 0;
            }

            return level switch
            {
                0 => ToChineseNumber(_counts[level]) + "、",
                1 => "（" + ToChineseNumber(_counts[level]) + "）",
                2 => _counts[level] + ".",
                3 => "（" + _counts[level] + "）",
                _ => ""
            };
        }

        private static string ToChineseNumber(int number)
        {
            string[] digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
            if (number <= 0)
            {
                return digits[0];
            }

            if (number < 10)
            {
                return digits[number];
            }

            if (number == 10)
            {
                return "十";
            }

            if (number < 20)
            {
                return "十" + digits[number % 10];
            }

            if (number < 100)
            {
                var tens = number / 10;
                var ones = number % 10;
                return digits[tens] + "十" + (ones == 0 ? "" : digits[ones]);
            }

            return number.ToString();
        }
    }

    private enum ParagraphKind
    {
        Body,
        Title,
        Heading1,
        Heading2,
        Heading3,
        Heading4
    }

    private sealed record ParagraphClassification(ParagraphFormat Format, ParagraphKind Kind);

    private sealed record ParagraphFormat(
        string StyleId,
        string EastAsiaFont,
        int FontSizeHalfPoints,
        JustificationValues Alignment,
        int FirstLine = 0,
        int FirstLineChars = 0,
        int Left = 0,
        int Right = 0,
        int Before = 0,
        int After = 0,
        int Line = FixedLineTwips,
        bool Bold = false,
        string Color = "000000");
}
