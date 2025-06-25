﻿using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

// todo need to set the HTTP "Content-Language" header, for instance:
// Content-Language: en-US
// Content-Language: fr-FR

namespace Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Converts a wordDoc to a self contained HTML
    /// </summary>
    public static class WmlToHtmlConverter
    {
        /// <summary>
        /// Converts a wordDoc to a self contained HTML
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="htmlConverterSettings"></param>
        /// <returns></returns>
        public static XElement ConvertToHtml(WmlDocument doc, WmlToHtmlConverterSettings htmlConverterSettings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
            using var document = streamDoc.GetWordprocessingDocument();
            return ConvertToHtml(document, htmlConverterSettings);
        }

        /// <summary>
        /// Converts a wordDoc to a self contained HTML
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="htmlConverterSettings"></param>
        /// <returns></returns>
        public static XElement ConvertToHtml(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings htmlConverterSettings)
        {
            RevisionAccepter.AcceptRevisions(wordDoc);
            var simplifyMarkupSettings = new SimplifyMarkupSettings
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                RemoveGoBackBookmark = true,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

            var formattingAssemblerSettings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = false,
                RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages,
                RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats,
                CreateHtmlConverterAnnotationAttributes = true,
                OrderElementsPerStandard = false,
                ListItemRetrieverSettings =
                    htmlConverterSettings.ListItemImplementations == null ?
                    new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                    } :
                    new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = htmlConverterSettings.ListItemImplementations,
                    },
            };

            FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);

            InsertAppropriateNonbreakingSpaces(wordDoc);
            CalculateSpanWidthForTabs(wordDoc);
            ReverseTableBordersForRtlTables(wordDoc);
            AdjustTableBorders(wordDoc);
            var rootElement = wordDoc?.MainDocumentPart?.GetXDocument().Root;
            FieldRetriever.AnnotateWithFieldInfo(wordDoc?.MainDocumentPart);
            AnnotateForSections(wordDoc);

            var xhtml = ConvertToHtmlTransform(wordDoc, htmlConverterSettings, rootElement, false, 0m) as XElement;

            ReifyStylesAndClasses(htmlConverterSettings, xhtml);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            return xhtml;
        }

        private static void ReverseTableBordersForRtlTables(WordprocessingDocument wordDoc)
        {
            var xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
            {
                var bidiVisual = tbl.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
                if (bidiVisual == null)
                {
                    continue;
                }

                var tblBorders = tbl.Elements(W.tblPr).Elements(W.tblBorders).FirstOrDefault();
                if (tblBorders != null)
                {
                    var left = tblBorders.Element(W.left);
                    if (left != null)
                    {
                        left = new XElement(W.right, left.Attributes());
                    }

                    var right = tblBorders.Element(W.right);
                    if (right != null)
                    {
                        right = new XElement(W.left, right.Attributes());
                    }

                    var newTblBorders = new XElement(W.tblBorders,
                        tblBorders.Element(W.top),
                        left,
                        tblBorders.Element(W.bottom),
                        right);
                    tblBorders.ReplaceWith(newTblBorders);
                }

                foreach (var tc in tbl.Elements(W.tr).Elements(W.tc))
                {
                    var tcBorders = tc.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault();
                    if (tcBorders != null)
                    {
                        var left = tcBorders.Element(W.left);
                        if (left != null)
                        {
                            left = new XElement(W.right, left.Attributes());
                        }

                        var right = tcBorders.Element(W.right);
                        if (right != null)
                        {
                            right = new XElement(W.left, right.Attributes());
                        }

                        var newTcBorders = new XElement(W.tcBorders,
                            tcBorders.Element(W.top),
                            left,
                            tcBorders.Element(W.bottom),
                            right);
                        tcBorders.ReplaceWith(newTcBorders);
                    }
                }
            }
        }

        private static void ReifyStylesAndClasses(WmlToHtmlConverterSettings htmlConverterSettings, XElement xhtml)
        {
            if (htmlConverterSettings.FabricateCssClasses)
            {
                var usedCssClassNames = new HashSet<string>();

                var elementsThatNeedClasses = xhtml.DescendantsAndSelf().Select(d => new
                {
                    Element = d,
                    Styles = d.Annotation<Dictionary<string, string>>(),
                }).Where(z => z.Styles != null);

                var augmented = elementsThatNeedClasses
                    .Select(p => new
                    {
                        p.Element,
                        p.Styles,
                        StylesString = p.Element.Name.LocalName + "|" + p.Styles.OrderBy(k => k.Key).Select(s => string.Format("{0}: {1};", s.Key, s.Value)).StringConcatenate(),
                    })
                    .GroupBy(p => p.StylesString)
                    .ToList();

                var classCounter = 1000000;
                var sb = new StringBuilder();
                sb.Append(Environment.NewLine);

                foreach (var grp in augmented)
                {
                    string classNameToUse;
                    var firstOne = grp.First();
                    var styles = firstOne.Styles;
                    if (styles.ContainsKey("PtStyleName"))
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + styles["PtStyleName"];
                        if (usedCssClassNames.Contains(classNameToUse))
                        {
                            classNameToUse = htmlConverterSettings.CssClassPrefix +
                                styles["PtStyleName"] + "-" +
                                classCounter.ToString().Substring(1);
                            classCounter++;
                        }
                    }
                    else
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix +
                            classCounter.ToString().Substring(1);
                        classCounter++;
                    }
                    usedCssClassNames.Add(classNameToUse);
                    sb.Append(firstOne.Element.Name.LocalName + "." + classNameToUse + " {" + Environment.NewLine);

                    foreach (var st in firstOne.Styles.Where(s => s.Key != "PtStyleName"))
                    {
                        var s = "    " + st.Key + ": " + st.Value + ";" + Environment.NewLine;
                        sb.Append(s);
                    }
                    sb.Append("}" + Environment.NewLine);
                    var classAtt = new XAttribute("class", classNameToUse);

                    foreach (var gc in grp)
                    {
                        gc.Element.Add(classAtt);
                    }
                }

                var styleValue = htmlConverterSettings.GeneralCss + sb + htmlConverterSettings.AdditionalCss;

                SetStyleElementValue(xhtml, styleValue);
            }
            else
            {
                // Previously, the h:style element was not added at this point. However, at least the General CSS will contain important settings.
                SetStyleElementValue(xhtml, htmlConverterSettings.GeneralCss + htmlConverterSettings.AdditionalCss);

                foreach (var d in xhtml.DescendantsAndSelf())
                {
                    var style = d.Annotation<Dictionary<string, string>>();
                    if (style == null)
                    {
                        continue;
                    }

                    var styleValue = style.Where(p => p.Key != "PtStyleName").OrderBy(p => p.Key).Select(e => $"{e.Key}: {e.Value};").StringConcatenate();
                    var st = new XAttribute("style", styleValue);
                    if (d.Attribute("style") != null)
                    {
                        d.Attribute("style").Value += styleValue;
                    }
                    else
                    {
                        d.Add(st);
                    }
                }
            }
        }

        private static void SetStyleElementValue(XElement xhtml, string styleValue)
        {
            var styleElement = xhtml
                .Descendants(Xhtml.style)
                .FirstOrDefault();
            if (styleElement != null)
            {
                styleElement.Value = styleValue;
            }
            else
            {
                styleElement = new XElement(Xhtml.style, styleValue);
                var head = xhtml.Element(Xhtml.head);
                if (head != null)
                {
                    head.Add(styleElement);
                }
            }
        }

        private static object? ListAwareConvertToHtmlTransform(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            IEnumerable<XNode> nodes,
            decimal currentMarginLeft,
            IList<string?> htmlNestedListTypes)
        {
            List<XElement> styleOrNumParagraphs = nodes.OfType<XElement>().ToList();
            var htmlNestedListItems = new List<(int Start, int CurrentNumber, bool Bidi, List<object> ListItems)>();
            int previousLevels = 0;
            int currentLevels = 0;
            int currentNumber = 0;
            int currentStart = 0;

            for (int paragraphIndex = 0; paragraphIndex < styleOrNumParagraphs.Count; paragraphIndex += 1)
            {
                XElement paragraph = styleOrNumParagraphs[paragraphIndex];
                var listType = (string?)paragraph.Attribute(PtOpenXml.HtmlStructure);
                var isBidi = IsBidi(paragraph);

                var paragraphHtml = ConvertToHtmlTransform(wordDoc, settings, paragraph, paragraphIndex != styleOrNumParagraphs.Count - 1, currentMarginLeft);
                if (!(paragraphHtml is XElement elementXml))
                {
                    continue;
                }

                // Each paragraph, contains 1 or more children with data-pt-list-item-run attribute that indicates the nested numbering.
                var paragraphChildren = elementXml.Elements().ToList();
                bool foundItemRun = false;
                List<object>? appendToList = null;
                for (int childIndex = 0; childIndex < paragraphChildren.Count; childIndex += 1)
                {
                    var child = paragraphChildren[childIndex];
                    var itemRun = (string?)child.Attribute("data-pt-list-item-run");
                    if (itemRun == null)
                    {
                        continue;
                    }

                    if (!foundItemRun)
                    {
                        foundItemRun = true;
                        string[] levels = itemRun.Split('.');
                        currentLevels = levels.Length;
                        for (int levelIndex = previousLevels - 1; levelIndex >= currentLevels && levelIndex > 0; levelIndex -= 1)
                        {
                            // We are now at a shallower nesting, add the deeper list to the shallower list as content.
                            // Note that this creates invalid HTML where a list is directly nested inside another list without a list item "li" tag,
                            // but Microsoft Word HTML export has the same behavior.
                            var (_, _, _, shallowList) = htmlNestedListItems[levelIndex - 1];
                            var (deepStart, _, _, deepList) = htmlNestedListItems[levelIndex];
                            var deepListType = htmlNestedListTypes[levelIndex];
                            if (!string.IsNullOrEmpty(deepListType))
                            {
                                var startAttribute = deepListType == "ol" ? new XAttribute("start", deepStart) : null;
                                var listHtml = new XElement(Xhtml.xhtml + deepListType, startAttribute, deepList);
                                shallowList.Add(listHtml);
                                var style = new Dictionary<string, string>
                                {
                                    ["list-style-type"] = "none",
                                    ["padding"] = "0",
                                    ["margin"] = "0",
                                };
                                listHtml.AddAnnotation(style);
                                if (isBidi)
                                {
                                    listHtml.Add(new XAttribute("dir", "rtl"));
                                }
                            }

                            htmlNestedListItems.RemoveAt(levelIndex);
                        }

                        for (int levelIndex = 0; levelIndex < currentLevels; levelIndex += 1)
                        {
                            // Create the level in htmlNestedListItems
                            int.TryParse(levels[levelIndex], out int start);

                            // Current number will be set to the one from the deepest level
                            currentNumber = start;
                            if (levelIndex >= htmlNestedListItems.Count)
                            {
                                htmlNestedListItems.Add((start, start, isBidi, new List<object>()));
                            }

                            if (levelIndex >= htmlNestedListTypes.Count)
                            {
                                htmlNestedListTypes.Add(listType);
                            }
                        }

                        (currentStart, _, _, appendToList) = htmlNestedListItems[currentLevels - 1];
                    }
                }

                if (appendToList != null)
                {
                    appendToList.Add(new XElement(Xhtml.li, paragraphHtml));
                    htmlNestedListTypes[currentLevels - 1] = listType;
                    htmlNestedListItems[currentLevels - 1] = (currentStart, currentNumber, isBidi, appendToList);
                    previousLevels = currentLevels;
                }
            }

            var rootList = new List<object>();
            for (int levelIndex = htmlNestedListItems.Count - 1; levelIndex >= 0; levelIndex -= 1)
            {
                var list = levelIndex == 0 ? rootList : htmlNestedListItems[levelIndex - 1].ListItems;
                var (deepStart, _, isBidi, deepList) = htmlNestedListItems[levelIndex];
                var deepListType = htmlNestedListTypes[levelIndex];
                if (!string.IsNullOrEmpty(deepListType))
                {
                    var startAttribute = deepListType == "ol" ? new XAttribute("start", deepStart) : null;
                    var listHtml = new XElement(Xhtml.xhtml + deepListType, startAttribute, deepList);
                    list.Add(listHtml);
                    var style = new Dictionary<string, string>
                    {
                        ["list-style-type"] = "none",
                        ["padding"] = "0",
                        ["margin"] = "0",
                    };
                    listHtml.AddAnnotation(style);
                    if (isBidi)
                    {
                        listHtml.Add(new XAttribute("dir", "rtl"));
                    }
                }
                else
                {
                    list.AddRange(deepList);
                }

                htmlNestedListItems.RemoveAt(levelIndex);
            }

            return rootList.Count > 0 ? rootList[0] : null;
        }

        private static object? ConvertToHtmlTransform(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XNode? node, bool suppressTrailingWhiteSpace, decimal currentMarginLeft, Dictionary<string, string> styleContext = default!)
        {
            if (!(node is XElement element))
            {
                return null;
            }

            // Transform the w:document element to the XHTML h:html element.
            // The h:head element is laid out based on the W3C's recommended layout, i.e.,
            // the charset (using the HTML5-compliant form), the title (which is always
            // there but possibly empty), and other meta tags.
            if (element.Name == W.document)
            {
                return new XElement(Xhtml.html,
                    new XElement(Xhtml.head,
                        new XElement(Xhtml.meta, new XAttribute("charset", "UTF-8")),
                        settings.PageTitle != null
                            ? new XElement(Xhtml.title, new XText(settings.PageTitle))
                            : new XElement(Xhtml.title, new XText(string.Empty)),
                        new XElement(Xhtml.meta,
                            new XAttribute("name", "Generator"),
                            new XAttribute("content", "PowerTools for Open XML"))),
                    element.Elements()
                        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            }

            // Transform the w:body element to the XHTML h:body element.
            if (element.Name == W.body)
            {
                return new XElement(Xhtml.body, CreateSectionDivs(wordDoc, settings, element));
            }

            // Transform the w:p element to the XHTML h:h1-h6 or h:p element (if the previous paragraph does not have a style separator).
            if (element.Name == W.p)
            {
                return ProcessParagraph(wordDoc, settings, element, suppressTrailingWhiteSpace, currentMarginLeft);
            }

            // Transform hyper links to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(R.id) != null)
            {
                try
                {
                    var href = wordDoc.MainDocumentPart
                        ?.HyperlinkRelationships
                        .First(x => x.Id == (string?)element.Attribute(R.id))
                        .Uri
                        .OriginalString
                        ?? string.Empty;
                    var anchor = element.Attribute(W.anchor);
                    if (anchor != null)
                    {
                        href += "#" + anchor.Value;
                    }

                    var a = new XElement(Xhtml.a, new XAttribute("href", href), element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run)));
                    if (!a.Nodes().Any())
                    {
                        a.Add(new XText(""));
                    }

                    return a;
                }
                catch (UriFormatException)
                {
                    return element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                }
            }

            // Transform hyperlinks to bookmarks to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(W.anchor) != null)
            {
                return ProcessHyperlinkToBookmark(wordDoc, settings, element);
            }

            // Transform contents of runs.
            if (element.Name == W.r)
            {
                return ConvertRun(wordDoc, settings, element);
            }

            // Transform w:bookmarkStart into anchor
            if (element.Name == W.bookmarkStart)
            {
                return ProcessBookmarkStart(element);
            }

            // Transform every w:t element to a text node.
            if (element.Name == W.t)
            {
                return settings.TextHandler.TransformText(element.Value, styleContext);
            }

            // Transform symbols to spans
            if (element.Name == W.sym)
            {
                return settings.SymbolHandler.TransformSymbol(element, styleContext);
            }

            // Transform tabs that have the pt:TabWidth attribute set
            if (element.Name == W.tab)
            {
                return ProcessTab(element);
            }

            // Transform w:br to h:br.
            if (element.Name == W.br || element.Name == W.cr)
            {
                return settings.BreakHandler.TransformBreak(element);
            }

            // Transform w:noBreakHyphen to '-'
            if (element.Name == W.noBreakHyphen)
            {
                return new XText("-");
            }

            // Transform w:tbl to h:tbl.
            if (element.Name == W.tbl)
            {
                return ProcessTable(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tr to h:tr.
            if (element.Name == W.tr)
            {
                return ProcessTableRow(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tc to h:td.
            if (element.Name == W.tc)
            {
                return ProcessTableCell(wordDoc, settings, element);
            }

            // Transform images
            if (element.Name == W.drawing || element.Name == W.pict || element.Name == W._object)
            {
                return ProcessImage(wordDoc, element, settings.ImageHandler);
            }

            // Transform content controls.
            if (element.Name == W.sdt)
            {
                return ProcessContentControl(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform smart tags and simple fields.
            if (element.Name == W.smartTag || element.Name == W.fldSimple)
            {
                return CreateBorderDivs(wordDoc, settings, element.Elements());
            }

            // Transform alternate content, selecting a choice that we support.
            if (element.Name == MC.AlternateContent)
            {
                return ProcessAlternateContent(wordDoc, settings, element);
            }

            // Ignore element.
            return null;
        }

        private static object ProcessHyperlinkToBookmark(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            var style = new Dictionary<string, string>();
            var a = new XElement(Xhtml.a,
                new XAttribute("href", "#" + (string)element.Attribute(W.anchor)),
                element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run)));
            if (!a.Nodes().Any())
            {
                a.Add(new XText(""));
            }

            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object? ProcessBookmarkStart(XElement element)
        {
            var name = (string)element.Attribute(W.name);
            if (name == null)
            {
                return null;
            }

            var style = new Dictionary<string, string>();
            var a = new XElement(Xhtml.a,
                new XAttribute("id", name),
                new XText(""));
            if (!a.Nodes().Any())
            {
                a.Add(new XText(""));
            }

            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object? ProcessTab(XElement element)
        {
            var tabWidthAtt = element.Attribute(PtOpenXml.TabWidth);
            if (tabWidthAtt == null)
            {
                return null;
            }

            var leader = (string)element.Attribute(PtOpenXml.Leader);
            var tabWidth = (decimal)tabWidthAtt;
            var style = new Dictionary<string, string>();
            XElement span;
            if (leader != null)
            {
                var leaderChar = ".";
                if (leader == "hyphen")
                {
                    leaderChar = "-";
                }
                else if (leader == "dot")
                {
                    leaderChar = ".";
                }
                else if (leader == "underscore")
                {
                    leaderChar = "_";
                }

                var runContainingTabToReplace = element.Ancestors(W.r).First();
                var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName") ?? runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                var dummyRun = new XElement(W.r, fontNameAtt, runContainingTabToReplace.Elements(W.rPr), new XElement(W.t, leaderChar));

                var widthOfLeaderChar = CalcWidthOfRunInTwips(dummyRun);

                var forceArial = false;

                if (widthOfLeaderChar == 0)
                {
                    dummyRun = new XElement(W.r, new XAttribute(PtOpenXml.FontName, "Arial"), runContainingTabToReplace.Elements(W.rPr), new XElement(W.t, leaderChar));
                    widthOfLeaderChar = CalcWidthOfRunInTwips(dummyRun);
                    forceArial = true;
                }

                if (widthOfLeaderChar != 0)
                {
                    var numberOfLeaderChars = (int)Math.Floor(tabWidth * 1440 / widthOfLeaderChar);
                    if (numberOfLeaderChars < 0)
                    {
                        numberOfLeaderChars = 0;
                    }

                    span = new XElement(Xhtml.span,
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        " " + "".PadRight(numberOfLeaderChars, leaderChar[0]) + " ");
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("margin-right", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", tabWidth));
                    style.Add("text-align", "center");
                    if (forceArial)
                    {
                        style.Add("font-family", "Arial");
                    }
                }
                else
                {
                    span = new XElement(Xhtml.span, new XAttribute(XNamespace.Xml + "space", "preserve"), " ");
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("margin-right", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", tabWidth));
                    style.Add("text-align", "center");
                    if (leader == "underscore")
                    {
                        style.Add("text-decoration", "underline");
                    }
                }
            }
            else
            {
                span = new XElement(Xhtml.span, new XEntity("#x00a0"));
                style.Add("margin", string.Format(NumberFormatInfo.InvariantInfo, "0 0 0 {0:0.00}in", tabWidth));
                style.Add("padding", "0 0 0 0");
            }
            span.AddAnnotation(style);
            return span;
        }

        private static object ProcessContentControl(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement element, decimal currentMarginLeft)
        {
            var relevantAncestors = element.Ancestors().TakeWhile(a => a.Name != W.txbxContent);
            var isRunLevelContentControl = relevantAncestors.Any(a => a.Name == W.p);
            if (isRunLevelContentControl)
            {
                return element.Elements(W.sdtContent).Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                    .ToList();
            }
            return CreateBorderDivs(wordDoc, settings, element.Elements(W.sdtContent).Elements());
        }

        // Transform the w:p element, including the following sibling w:p element(s)
        // in case the w:p element has a style separator. The sibling(s) will be
        // transformed to h:span elements rather than h:p elements and added to
        // the element (e.g., h:h2) created from the w:p element having the (first)
        // style separator (i.e., a w:specVanish element).
        private static object? ProcessParagraph(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement element, bool suppressTrailingWhiteSpace, decimal currentMarginLeft)
        {
            // Ignore this paragraph if the previous paragraph has a style separator.
            // We have already transformed this one together with the previous one.
            var previousParagraph = element.ElementsBeforeSelf(W.p).LastOrDefault();
            if (HasStyleSeparator(previousParagraph))
            {
                return null;
            }

            var (elementName, attributes) = GetParagraphElementNameAndAttributes(element, wordDoc);
            var isBidi = IsBidi(element);
            var paragraph = ConvertParagraph(wordDoc, settings, element, elementName,
                suppressTrailingWhiteSpace, currentMarginLeft, isBidi);
            paragraph.Add(attributes);

            // The paragraph conversion might have created empty spans. These can and should be removed because empty spans are invalid in HTML5.
            paragraph.Elements(Xhtml.span).Where(e => e.IsEmpty).Remove();

            foreach (var span in paragraph.Elements(Xhtml.span).ToList())
            {
                var v = span.Value;
                if (v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1])) && span.Attribute(XNamespace.Xml + "space") == null)
                {
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                }
            }

            while (HasStyleSeparator(element))
            {
                element = element.ElementsAfterSelf(W.p).FirstOrDefault();
                if (element == null)
                {
                    break;
                }

                elementName = Xhtml.span;
                isBidi = IsBidi(element);
                var span = ConvertParagraph(wordDoc, settings, element, elementName,
                    suppressTrailingWhiteSpace, currentMarginLeft, isBidi);
                var v = span.Value;
                if (v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1])) && span.Attribute(XNamespace.Xml + "space") == null)
                {
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                }

                paragraph.Add(span);
            }

            return paragraph;
        }

        private static object ProcessTable(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            var style = new Dictionary<string, string>();
            style.AddIfMissing("border-collapse", "collapse");
            style.AddIfMissing("border", "none");
            var bidiVisual = element.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
            var tblW = element.Elements(W.tblPr).Elements(W.tblW).FirstOrDefault();
            if (tblW != null)
            {
                var type = (string)tblW.Attribute(W.type);
                if (type != null && type == "pct")
                {
                    var w = (int)tblW.Attribute(W._w);
                    style.AddIfMissing("width", w / 50 + "%");
                }
            }
            var tblInd = element.Elements(W.tblPr).Elements(W.tblInd).FirstOrDefault();
            if (tblInd != null)
            {
                var tblIndType = (string)tblInd.Attribute(W.type);
                if (tblIndType != null && tblIndType == "dxa")
                {
                    var width = (decimal?)tblInd.Attribute(W._w);
                    if (width != null)
                    {
                        style.AddIfMissing("margin-left",
                            width > 0m
                                ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", width / 20m)
                                : "0");
                    }
                }
            }
            var tableDirection = bidiVisual != null ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            style.AddIfMissing("margin-bottom", ".001pt");
            var table = new XElement(Xhtml.table,
                // TODO: Revisit and make sure the omission is covered by appropriate CSS.
                // new XAttribute("border", "1"),
                // new XAttribute("cellspacing", 0),
                // new XAttribute("cellpadding", 0),
                tableDirection,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            table.AddAnnotation(style);
            var jc = (string)element.Elements(W.tblPr).Elements(W.jc).Attributes(W.val).FirstOrDefault() ?? "left";
            XAttribute? dir = null;
            XAttribute? jcToUse = null;
            if (bidiVisual != null)
            {
                dir = new XAttribute("dir", "rtl");
                if (jc == "left")
                {
                    jcToUse = new XAttribute("align", "right");
                }
                else if (jc == "right")
                {
                    jcToUse = new XAttribute("align", "left");
                }
                else if (jc == "center")
                {
                    jcToUse = new XAttribute("align", "center");
                }
            }
            else
            {
                jcToUse = new XAttribute("align", jc);
            }
            var tableDiv = new XElement(Xhtml.div,
                dir,
                jcToUse,
                table);
            return tableDiv;
        }

        private static object? ProcessTableCell(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            var style = new Dictionary<string, string>();
            XAttribute? colSpan = null;
            XAttribute? rowSpan = null;

            var tcPr = element.Element(W.tcPr);
            if (tcPr != null)
            {
                if ((string)tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() == "restart")
                {
                    var currentRow = element.Parent.ElementsBeforeSelf(W.tr).Count();
                    var currentCell = element.ElementsBeforeSelf(W.tc).Count();
                    var tbl = element.Parent.Parent;
                    var rowSpanCount = 1;
                    currentRow += 1;
                    while (true)
                    {
                        var row = tbl.Elements(W.tr).Skip(currentRow).FirstOrDefault();
                        if (row == null)
                        {
                            break;
                        }

                        var cell2 = row.Elements(W.tc).Skip(currentCell).FirstOrDefault();
                        if (cell2 == null)
                        {
                            break;
                        }

                        if (cell2.Elements(W.tcPr).Elements(W.vMerge).FirstOrDefault() == null)
                        {
                            break;
                        }

                        if ((string)cell2.Elements(W.tcPr).Elements(W.vMerge).Attributes(W.val).FirstOrDefault() == "restart")
                        {
                            break;
                        }

                        currentRow += 1;
                        rowSpanCount += 1;
                    }
                    rowSpan = new XAttribute("rowspan", rowSpanCount);
                }

                if (tcPr.Element(W.vMerge) != null &&
                    (string)tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() != "restart")
                {
                    return null;
                }

                if (tcPr.Element(W.vAlign) != null)
                {
                    var vAlignVal = (string)tcPr.Elements(W.vAlign).Attributes(W.val).FirstOrDefault();
                    if (vAlignVal == "top")
                    {
                        style.AddIfMissing("vertical-align", "top");
                    }
                    else if (vAlignVal == "center")
                    {
                        style.AddIfMissing("vertical-align", "middle");
                    }
                    else if (vAlignVal == "bottom")
                    {
                        style.AddIfMissing("vertical-align", "bottom");
                    }
                    else
                    {
                        style.AddIfMissing("vertical-align", "middle");
                    }
                }
                style.AddIfMissing("vertical-align", "top");

                if ((string)tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "dxa")
                {
                    decimal width = (int)tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", width / 20m));
                }
                if ((string)tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "pct")
                {
                    decimal width = (int)tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}%", width / 50m));
                }

                var tcBorders = tcPr.Element(W.tcBorders);
                GenerateBorderStyle(tcBorders, W.top, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.right, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.bottom, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.left, style, BorderType.Cell);

                CreateStyleFromShd(style, tcPr.Element(W.shd));

                var gridSpan = tcPr.Elements(W.gridSpan).Attributes(W.val).Select(a => (int?)a).FirstOrDefault();
                if (gridSpan != null)
                {
                    colSpan = new XAttribute("colspan", (int)gridSpan);
                }
            }
            style.AddIfMissing("padding-top", "0");
            style.AddIfMissing("padding-bottom", "0");

            var cell = new XElement(Xhtml.td,
                rowSpan,
                colSpan,
                CreateBorderDivs(wordDoc, settings, element.Elements()));
            cell.AddAnnotation(style);
            return cell;
        }

        private static object ProcessTableRow(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element,
            decimal currentMarginLeft)
        {
            var style = new Dictionary<string, string>();
            var trHeight = (int?)element.Elements(W.trPr).Elements(W.trHeight).Attributes(W.val).FirstOrDefault();
            if (trHeight != null)
            {
                style.AddIfMissing("height",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", (decimal)trHeight / 1440m));
            }

            var htmlRow = new XElement(Xhtml.tr,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            if (style.Any())
            {
                htmlRow.AddAnnotation(style);
            }

            return htmlRow;
        }

        private static bool HasStyleSeparator(XElement element)
        {
            return element != null && element.Elements(W.pPr).Elements(W.rPr).Any(e => GetBoolProp(e, W.specVanish));
        }

        private static bool IsBidi(XElement element)
        {
            return element
                .Elements(W.pPr)
                .Elements(W.bidi)
                .Any(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);
        }

        private static (XName ElementName, IEnumerable<XAttribute> Attributes) GetParagraphElementNameAndAttributes(XElement element, WordprocessingDocument wordDoc)
        {
            var elementName = Xhtml.p;

            var styleId = (string?)element.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
            if (styleId == null)
            {
                return (elementName, Enumerable.Empty<XAttribute>());
            }

            var style = GetStyle(styleId, wordDoc);
            if (style == null)
            {
                return (elementName, Enumerable.Empty<XAttribute>());
            }

            var outlineLevel =
                (int?)style.Elements(W.pPr).Elements(W.outlineLvl).Attributes(W.val).FirstOrDefault();
            if (outlineLevel != null && outlineLevel > 5 && outlineLevel < 9)
            {
                elementName = Xhtml.div;
                var attributes = new XAttribute[]
                {
                    new XAttribute("role", "heading"),
                    new XAttribute("aria-level", outlineLevel + 1),
                };

                return (elementName, attributes);
            }

            if (outlineLevel != null && outlineLevel <= 5)
            {
                elementName = Xhtml.xhtml + string.Format("h{0}", outlineLevel + 1);
            }

            return (elementName, Enumerable.Empty<XAttribute>());
        }

        private static XElement? GetStyle(string styleId, WordprocessingDocument wordDoc)
        {
            var stylesPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                return null;
            }

            var styles = stylesPart.GetXDocument().Root;
            return styles?.Elements(W.style).FirstOrDefault(s => (string)s.Attribute(W.styleId) == styleId);
        }

        private static object CreateSectionDivs(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            // note: when building a paging html converter, need to attend to new sections with page breaks here.
            // This code conflates adjacent sections if they have identical formatting, which is not an issue
            // for the non-paging transform.
            var groupedIntoDivs = element
                .Elements()
                .GroupAdjacent(e =>
                {
                    var sectAnnotation = e.Annotation<SectionAnnotation>();
                    return sectAnnotation != null ? sectAnnotation?.SectionElement?.ToString() : string.Empty;
                });

            // note: when creating a paging html converter, need to pay attention to w:rtlGutter element.
            var divList = groupedIntoDivs
                .Select(g =>
                {
                    var sectPr = g.First().Annotation<SectionAnnotation>();
                    XElement? bidi = null;
                    if (sectPr != null)
                    {
                        bidi = sectPr?
                            .SectionElement?
                            .Elements(W.bidi)
                            .FirstOrDefault(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);
                    }
                    if (sectPr == null || bidi == null)
                    {
                        var div = new XElement(Xhtml.div, CreateBorderDivs(wordDoc, settings, g));
                        return div;
                    }
                    else
                    {
                        var div = new XElement(Xhtml.div,
                            new XAttribute("dir", "rtl"),
                            CreateBorderDivs(wordDoc, settings, g));
                        return div;
                    }
                });
            return divList;
        }

        private enum BorderType
        {
            Paragraph,
            Cell,
        };

        /*
         * Notes on line spacing
         *
         * the w:line and w:lineRule attributes control spacing between lines - including between lines within a paragraph
         *
         * If w:spacing w:lineRule="auto" then
         *   w:spacing w:line is a percentage where 240 == 100%
         *
         *   (line value / 240) * 100 = percentage of line
         *
         * If w:spacing w:lineRule="exact" or w:lineRule="atLeast" then
         *   w:spacing w:line is in twips
         *   1440 = exactly one inch from line to line
         *
         * Handle
         * - ind
         * - jc
         * - numPr
         * - pBdr
         * - shd
         * - spacing
         * - textAlignment
         *
         * Don't Handle (yet)
         * - adjustRightInd?
         * - autoSpaceDE
         * - autoSpaceDN
         * - bidi
         * - contextualSpacing
         * - divId
         * - framePr
         * - keepLines
         * - keepNext
         * - kinsoku
         * - mirrorIndents
         * - overflowPunct
         * - pageBreakBefore
         * - snapToGrid
         * - suppressAutoHyphens
         * - suppressLineNumbers
         * - suppressOverlap
         * - tabs
         * - textBoxTightWrap
         * - textDirection
         * - topLinePunct
         * - widowControl
         * - wordWrap
         */

        private static XElement ConvertParagraph(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement paragraph, XName elementName, bool suppressTrailingWhiteSpace, decimal currentMarginLeft, bool isBidi)
        {
            var style = DefineParagraphStyle(paragraph, elementName, suppressTrailingWhiteSpace, currentMarginLeft, isBidi, settings.FontHandler);
            var rtl = isBidi ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            var firstMark = isBidi ? new XEntity("#x200f") : null;

            // Analyze initial runs to see whether we have a tab, in which case we will render a span with a defined width and ignore the tab rather than rendering the text preceding the tab and the tab as a span with a computed width.
            var firstTabRun = paragraph.Elements(W.r).FirstOrDefault(run => run.Elements(W.tab).Any());
            var elementsPrecedingTab = firstTabRun != null ? paragraph.Elements(W.r).TakeWhile(e => e != firstTabRun).Where(e => e.Elements().Any(c => c.Attributes(PtOpenXml.TabWidth).Any())).ToList() : Enumerable.Empty<XElement>().ToList();

            // TODO: Revisit
            // For the time being, if a hyperlink field precedes the tab, we'll render it as before.
            var hyperlinkPrecedesTab = elementsPrecedingTab.Elements(W.r).Elements(W.instrText).Select(e => e.Value).Any(value => value != null && value.TrimStart().ToUpper().StartsWith("HYPERLINK"));

            if (hyperlinkPrecedesTab)
            {
                var paraElement1 = new XElement(elementName, rtl, firstMark, ConvertContentThatCanContainFields(wordDoc, settings, paragraph.Elements()));
                paraElement1.AddAnnotation(style);
                return paraElement1;
            }

            var txElementsPrecedingTab = TransformElementsPrecedingTab(wordDoc, settings, elementsPrecedingTab, firstTabRun);
            var elementsSucceedingTab = TransformElementsSuceedingTab(paragraph, firstTabRun);
            var paraElement = new XElement(elementName, rtl, firstMark, txElementsPrecedingTab, ConvertContentThatCanContainFields(wordDoc, settings, elementsSucceedingTab));
            paraElement.AddAnnotation(style);

            return paraElement;
        }

        private static IEnumerable<XElement> TransformElementsSuceedingTab(XElement paragraph, XElement? firstTabRun)
        {
            var elementsSuceedingTab = firstTabRun != null ? paragraph.Elements().SkipWhile(e => e != firstTabRun).Skip(1) : paragraph.Elements();
            return elementsSuceedingTab;
        }

        private static List<object?> TransformElementsPrecedingTab(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, List<XElement> elementsPrecedingTab, XElement? firstTabRun)
        {
            var tabWidth = firstTabRun != null ? (decimal?)firstTabRun.Elements(W.tab).Attributes(PtOpenXml.TabWidth).FirstOrDefault() ?? 0m : 0m;
            var precedingElementsWidth = elementsPrecedingTab.Elements().Where(c => c.Attributes(PtOpenXml.TabWidth).Any()).Select(e => (decimal)e.Attribute(PtOpenXml.TabWidth)).Sum();
            var totalWidth = precedingElementsWidth + tabWidth;

            var txElementsPrecedingTab = elementsPrecedingTab.Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m)).ToList();
            if (txElementsPrecedingTab.Count > 1)
            {
                var span = new XElement(Xhtml.span, txElementsPrecedingTab);
                var spanStyle = new Dictionary<string, string>
                {
                    { "text-indent", "0" },
                    { "margin-right", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}in", totalWidth) }
                };
                span.AddAnnotation(spanStyle);
            }
            else if (txElementsPrecedingTab.Count == 1 && txElementsPrecedingTab.First() is XElement element)
            {
                var spanStyle = element.Annotation<Dictionary<string, string>>();
                spanStyle.AddIfMissing("text-indent", "0");
                spanStyle.AddIfMissing("margin-right", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}in", totalWidth));
            }
            else if (totalWidth > 0m)
            {
                // Leading tabs in paragraphs
                var div = new XElement(Xhtml.span);
                var spanStyle = new Dictionary<string, string>
                {
                    { "display", "inline-block" },
                    { "text-indent", "0" },
                    { "margin", string.Format(NumberFormatInfo.InvariantInfo, "0 0 0 {0:0.000}in", totalWidth) }
                };
                div.AddAnnotation(spanStyle);
                div.SetValue("\u00A0");
                return new List<object?> { div };
            }

            return txElementsPrecedingTab;
        }

        private static Dictionary<string, string> DefineParagraphStyle(XElement paragraph, XName elementName, bool suppressTrailingWhiteSpace, decimal currentMarginLeft, bool isBidi, IFontHandler fontHandler)
        {
            var style = new Dictionary<string, string>();
            var styleName = (string)paragraph.Attribute(PtOpenXml.StyleName);

            if (styleName != null)
            {
                style.Add("PtStyleName", styleName);
            }

            var pPr = paragraph.Element(W.pPr);

            if (pPr == null)
            {
                return style;
            }

            CreateStyleFromSpacing(style, pPr.Element(W.spacing), elementName, suppressTrailingWhiteSpace);
            CreateStyleFromInd(style, pPr.Element(W.ind), elementName, currentMarginLeft, isBidi);

            // todo need to handle
            // - both
            // - mediumKashida
            // - distribute
            // - numTab
            // - highKashida
            // - lowKashida
            // - thaiDistribute

            CreateStyleFromJc(style, pPr.Element(W.jc), isBidi);
            CreateStyleFromShd(style, pPr.Element(W.shd));

            // Pt.FontName
            var font = fontHandler.TranslateParagraphStyleFont(paragraph);
            if (font != null)
            {
                CreateFontCssProperty(font, style);
            }

            DefineFontSize(style, paragraph);
            DefineLineHeight(style, paragraph);

            // vertical text alignment as of December 2013 does not work in any major browsers.
            CreateStyleFromTextAlignment(style, pPr.Element(W.textAlignment));

            style.AddIfMissing("margin-top", "0");
            style.AddIfMissing("margin-left", "0");
            style.AddIfMissing("margin-right", "0");
            style.AddIfMissing("margin-bottom", ".001pt");

            return style;
        }

        private static void CreateStyleFromInd(Dictionary<string, string> style, XElement ind, XName elementName,
            decimal currentMarginLeft, bool isBidi)
        {
            if (ind == null)
            {
                return;
            }

            var left = (decimal?)ind.Attribute(W.left);
            if (left != null && elementName != Xhtml.span)
            {
                var leftInInches = (decimal)left / 1440 - currentMarginLeft;
                style.AddIfMissing(isBidi ? "margin-right" : "margin-left",
                    leftInInches > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", leftInInches)
                        : "0");
            }

            var right = (decimal?)ind.Attribute(W.right);
            if (right != null)
            {
                var rightInInches = (decimal)right / 1440;
                style.AddIfMissing(isBidi ? "margin-left" : "margin-right",
                    rightInInches > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", rightInInches)
                        : "0");
            }

            var firstLine = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.firstLine));
            if (firstLine != null && elementName != Xhtml.span)
            {
                var firstLineInInches = (decimal)firstLine / 1440m;
                style.AddIfMissing("text-indent",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", firstLineInInches));
            }

            var hanging = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.hanging));
            if (hanging != null && elementName != Xhtml.span)
            {
                var hangingInInches = (decimal)-hanging / 1440m;
                style.AddIfMissing("text-indent",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", hangingInInches));
            }
        }

        private static void CreateStyleFromJc(Dictionary<string, string> style, XElement jc, bool isBidi)
        {
            if (jc != null)
            {
                var jcVal = (string)jc.Attributes(W.val).FirstOrDefault() ?? "left";
                if (jcVal == "left")
                {
                    style.AddIfMissing("text-align", isBidi ? "right" : "left");
                }
                else if (jcVal == "right")
                {
                    style.AddIfMissing("text-align", isBidi ? "left" : "right");
                }
                else if (jcVal == "center")
                {
                    style.AddIfMissing("text-align", "center");
                }
                else if (jcVal == "both")
                {
                    style.AddIfMissing("text-align", "justify");
                }
            }
        }

        private static void CreateStyleFromSpacing(Dictionary<string, string> style, XElement spacing, XName elementName,
            bool suppressTrailingWhiteSpace)
        {
            if (spacing == null)
            {
                return;
            }

            var spacingBefore = (decimal?)spacing.Attribute(W.before);
            if (spacingBefore != null && elementName != Xhtml.span)
            {
                style.AddIfMissing("margin-top",
                    spacingBefore > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingBefore / 20.0m)
                        : "0");
            }

            var lineRule = (string)spacing.Attribute(W.lineRule);
            if (lineRule == "auto")
            {
                var line = (decimal)spacing.Attribute(W.line);
                if (line != 240m)
                {
                    var pct = line / 240m * 100m;
                    style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}%", pct));
                }
            }
            if (lineRule == "exact")
            {
                var line = (decimal)spacing.Attribute(W.line);
                var points = line / 20m;
                style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", points));
            }
            if (lineRule == "atLeast")
            {
                var line = (decimal)spacing.Attribute(W.line);
                var points = line / 20m;
                if (points >= 14m)
                {
                    style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", points));
                }
            }

            var spacingAfter = suppressTrailingWhiteSpace ? 0 : WordprocessingMLUtil.AttributeToTwips(spacing.Attribute(W.after));
            if (spacingAfter != null)
            {
                style.AddIfMissing("margin-bottom",
                    spacingAfter > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingAfter / 20.0m)
                        : "0");
            }
        }

        private static void CreateStyleFromTextAlignment(Dictionary<string, string> style, XElement textAlignment)
        {
            if (textAlignment == null)
            {
                return;
            }

            var verticalTextAlignment = (string)textAlignment.Attributes(W.val).FirstOrDefault();
            if (verticalTextAlignment == null || verticalTextAlignment == "auto")
            {
                return;
            }

            if (verticalTextAlignment == "top")
            {
                style.AddIfMissing("vertical-align", "top");
            }
            else if (verticalTextAlignment == "center")
            {
                style.AddIfMissing("vertical-align", "middle");
            }
            else if (verticalTextAlignment == "baseline")
            {
                style.AddIfMissing("vertical-align", "baseline");
            }
            else if (verticalTextAlignment == "bottom")
            {
                style.AddIfMissing("vertical-align", "bottom");
            }
        }

        private static void DefineFontSize(Dictionary<string, string> style, XElement paragraph)
        {
            var sz = paragraph.DescendantsTrimmed(W.txbxContent).Where(e => e.Name == W.r).Select(r => GetFontSize(r)).Max();

            if (sz != null)
            {
                style.AddIfMissing("font-size", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", sz / 2.0m));
            }
        }

        private static void DefineLineHeight(Dictionary<string, string> style, XElement paragraph)
        {
            var allRunsAreUniDirectional = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.r)
                .Select(run => (string)run.Attribute(PtOpenXml.LanguageType))
                .All(lt => lt != "bidi");
            if (allRunsAreUniDirectional)
            {
                style.AddIfMissing("line-height", "108%");
            }
        }

        /*
         * Handle:
         * - b
         * - bdr
         * - caps
         * - color
         * - dstrike
         * - highlight
         * - i
         * - position
         * - rFonts
         * - shd
         * - smallCaps
         * - spacing
         * - strike
         * - sz
         * - u
         * - vanish
         * - vertAlign
         *
         * Don't handle:
         * - em
         * - emboss
         * - fitText
         * - imprint
         * - kern
         * - outline
         * - shadow
         * - w
         *
         */

        private static object? ConvertRun(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement run)
        {
            var rPr = run.Element(W.rPr);
            if (rPr == null)
            {
                return run.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m));
            }

            // hide all content that contains the w:rPr/w:webHidden element
            if (rPr.Element(W.webHidden) != null)
            {
                return null;
            }

            var style = DefineRunStyle(run, settings.FontHandler);
            object content = run.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m, style));

            // Wrap content in h:sup or h:sub elements as necessary.
            if (rPr.Element(W.vertAlign) != null)
            {
                XElement? newContent = null;
                var vertAlignVal = (string)rPr.Elements(W.vertAlign).Attributes(W.val).FirstOrDefault();
                switch (vertAlignVal)
                {
                    case "superscript":
                        newContent = new XElement(Xhtml.sup, content);
                        break;

                    case "subscript":
                        newContent = new XElement(Xhtml.sub, content);
                        break;
                }
                if (newContent != null && newContent.Nodes().Any())
                {
                    content = newContent;
                }
            }

            var langAttribute = GetLangAttribute(run);

            DetermineRunMarks(run, rPr, style, out var runStartMark, out var runEndMark);

            var listItemRun = (string?)run.Attribute(PtOpenXml.ListItemRun);
            var listItemRunAttribute = !string.IsNullOrEmpty(listItemRun) ? new XAttribute("data-pt-list-item-run", listItemRun) : null;
            if (style.Any() || langAttribute != null || runStartMark != null || listItemRunAttribute != null)
            {
                style.AddIfMissing("margin", "0");
                style.AddIfMissing("padding", "0");
                var xe = new XElement(Xhtml.span, langAttribute, runStartMark, listItemRunAttribute, content, runEndMark);

                xe.AddAnnotation(style);
                content = xe;
            }
            return content;
        }

        private static Dictionary<string, string> DefineRunStyle(XElement run, IFontHandler fontHandler)
        {
            var style = new Dictionary<string, string>();

            var rPr = run.Elements(W.rPr).First();

            var styleName = (string)run.Attribute(PtOpenXml.StyleName);
            if (styleName != null)
            {
                style.Add("PtStyleName", styleName);
            }

            // W.bdr
            if (rPr.Element(W.bdr) != null && (string)rPr.Elements(W.bdr).Attributes(W.val).FirstOrDefault() != "none")
            {
                style.AddIfMissing("border", "solid windowtext 1.0pt");
                style.AddIfMissing("padding", "0");
            }

            // W.color
            var color = (string)rPr.Elements(W.color).Attributes(W.val).FirstOrDefault();
            if (color != null)
            {
                CreateColorProperty("color", color, style);
            }

            // W.highlight
            var highlight = (string)rPr.Elements(W.highlight).Attributes(W.val).FirstOrDefault();
            if (highlight != null)
            {
                CreateColorProperty("background", highlight, style);
            }

            // W.shd
            var shade = (string)rPr.Elements(W.shd).Attributes(W.fill).FirstOrDefault();
            if (shade != null)
            {
                CreateColorProperty("background", shade, style);
            }

            // Pt.FontName
            var font = fontHandler.TranslateRunStyleFont(run);

            if (font != null)
            {
                CreateFontCssProperty(font, style);
            }

            // W.sz
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);
            var sz = GetFontSize(languageType, rPr);
            if (sz != null)
            {
                style.AddIfMissing("font-size", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", sz / 2.0m));
            }

            // W.caps
            if (GetBoolProp(rPr, W.caps))
            {
                style.AddIfMissing("text-transform", "uppercase");
            }

            // W.smallCaps
            if (GetBoolProp(rPr, W.smallCaps))
            {
                style.AddIfMissing("font-variant", "small-caps");
            }

            // W.spacing
            var spacingInTwips = (decimal?)rPr.Elements(W.spacing).Attributes(W.val).FirstOrDefault();
            if (spacingInTwips != null)
            {
                style.AddIfMissing("letter-spacing",
                    spacingInTwips > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingInTwips / 20)
                        : "0");
            }

            // W.position
            var position = (decimal?)rPr.Elements(W.position).Attributes(W.val).FirstOrDefault();
            if (position != null)
            {
                style.AddIfMissing("position", "relative");
                style.AddIfMissing("top", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", -(position / 2)));
            }

            // W.vanish
            if (GetBoolProp(rPr, W.vanish) && !GetBoolProp(rPr, W.specVanish))
            {
                style.AddIfMissing("display", "none");
            }

            // W.u
            if (rPr.Element(W.u) != null && (string)rPr.Elements(W.u).Attributes(W.val).FirstOrDefault() != "none")
            {
                style.AddIfMissing("text-decoration", "underline");
            }

            // W.i
            style.AddIfMissing("font-style", GetBoolProp(rPr, W.i) ? "italic" : "normal");

            // W.b
            style.AddIfMissing("font-weight", GetBoolProp(rPr, W.b) ? "bold" : "normal");

            // W.strike
            if (GetBoolProp(rPr, W.strike) || GetBoolProp(rPr, W.dstrike))
            {
                style.AddIfMissing("text-decoration", "line-through");
            }

            return style;
        }

        private static decimal? GetFontSize(XElement e)
        {
            var languageType = (string)e.Attribute(PtOpenXml.LanguageType);
            if (e.Name == W.p)
            {
                return GetFontSize(languageType, e.Elements(W.pPr).Elements(W.rPr).FirstOrDefault());
            }
            if (e.Name == W.r)
            {
                return GetFontSize(languageType, e.Element(W.rPr));
            }
            return null;
        }

        private static decimal? GetFontSize(string languageType, XElement rPr)
        {
            if (rPr == null)
            {
                return null;
            }

            return languageType == "bidi" ? (decimal?)rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault() : (decimal?)rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
        }

        private static void DetermineRunMarks(XElement run, XElement rPr, Dictionary<string, string> style, out XEntity? runStartMark, out XEntity? runEndMark)
        {
            runStartMark = null;
            runEndMark = null;

            // Only do the following for text runs.
            if (run.Element(W.t) == null)
            {
                return;
            }

            // Can't add directional marks if the font-family is symbol - they are visible, and display as a ?
            var addDirectionalMarks = true;
            if (style.ContainsKey("font-family") && style["font-family"].ToLower() == "symbol")
            {
                addDirectionalMarks = false;
            }
            if (!addDirectionalMarks)
            {
                return;
            }

            var isRtl = rPr.Element(W.rtl) != null;
            if (isRtl)
            {
                runStartMark = new XEntity("#x200f"); // RLM
                runEndMark = new XEntity("#x200f"); // RLM
            }
            else
            {
                var paragraph = run.Ancestors(W.p).First();
                var paraIsBidi = paragraph
                    .Elements(W.pPr)
                    .Elements(W.bidi)
                    .Any(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);

                if (paraIsBidi)
                {
                    runStartMark = new XEntity("#x200e"); // LRM
                    runEndMark = new XEntity("#x200e"); // LRM
                }
            }
        }

        private static XAttribute? GetLangAttribute(XElement run)
        {
            const string defaultLanguage = "en-US"; // todo need to get defaultLanguage

            var rPr = run.Elements(W.rPr).First();
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);

            string? lang = null;
            if (languageType == "western")
            {
                lang = (string)rPr.Elements(W.lang).Attributes(W.val).FirstOrDefault();
            }
            else if (languageType == "bidi")
            {
                lang = (string)rPr.Elements(W.lang).Attributes(W.bidi).FirstOrDefault();
            }
            else if (languageType == "eastAsia")
            {
                lang = (string)rPr.Elements(W.lang).Attributes(W.eastAsia).FirstOrDefault();
            }

            if (lang == null)
            {
                lang = defaultLanguage;
            }

            return lang != defaultLanguage ? new XAttribute("lang", lang) : null;
        }

        private static void AdjustTableBorders(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done
            // for all content parts, not just the main document part.

            var xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
            {
                AdjustTableBorders(tbl);
            }

            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static void AdjustTableBorders(XElement tbl)
        {
            var ta = tbl
                .Elements(W.tr)
                .Select(r => r
                    .Elements(W.tc)
                    .SelectMany(c =>
                        Enumerable.Repeat(c,
                            (int?)c.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault() ?? 1))
                    .ToArray())
                .ToArray();

            for (var y = 0; y < ta.Length; y++)
            {
                for (var x = 0; x < ta[y].Length; x++)
                {
                    var thisCell = ta[y][x];
                    FixTopBorder(ta, thisCell, x, y);
                    FixLeftBorder(ta, thisCell, x, y);
                    FixBottomBorder(ta, thisCell, x, y);
                    FixRightBorder(ta, thisCell, x, y);
                }
            }
        }

        private static void FixTopBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y > 0)
            {
                var rowAbove = ta[y - 1];
                if (x < rowAbove.Length - 1)
                {
                    var cellAbove = ta[y - 1][x];
                    if (cellAbove != null &&
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                        cellAbove.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault(),
                            cellAbove.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault());
                    }
                }
            }
        }

        private static void FixLeftBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x > 0)
            {
                var cellLeft = ta[y][x - 1];
                if (cellLeft != null &&
                    thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                    cellLeft.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault(),
                        cellLeft.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault());
                }
            }
        }

        private static void FixBottomBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y < ta.Length - 1)
            {
                var rowBelow = ta[y + 1];
                if (x < rowBelow.Length - 1)
                {
                    var cellBelow = ta[y + 1][x];
                    if (cellBelow != null &&
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                        cellBelow.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault(),
                            cellBelow.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault());
                    }
                }
            }
        }

        private static void FixRightBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x < ta[y].Length - 1)
            {
                var cellRight = ta[y][x + 1];
                if (cellRight != null &&
                    thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                    cellRight.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault(),
                        cellRight.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault());
                }
            }
        }

        private static readonly Dictionary<string, int> BorderTypePriority = new Dictionary<string, int>()
        {
            { "single", 1 },
            { "thick", 2 },
            { "double", 3 },
            { "dotted", 4 },
        };

        private static readonly Dictionary<string, int> BorderNumber = new Dictionary<string, int>()
        {
            {"single", 1 },
            {"thick", 2 },
            {"double", 3 },
            {"dotted", 4 },
            {"dashed", 5 },
            {"dotDash", 5 },
            {"dotDotDash", 5 },
            {"triple", 6 },
            {"thinThickSmallGap", 6 },
            {"thickThinSmallGap", 6 },
            {"thinThickThinSmallGap", 6 },
            {"thinThickMediumGap", 6 },
            {"thickThinMediumGap", 6 },
            {"thinThickThinMediumGap", 6 },
            {"thinThickLargeGap", 6 },
            {"thickThinLargeGap", 6 },
            {"thinThickThinLargeGap", 6 },
            {"wave", 7 },
            {"doubleWave", 7 },
            {"dashSmallGap", 5 },
            {"dashDotStroked", 5 },
            {"threeDEmboss", 7 },
            {"threeDEngrave", 7 },
            {"outset", 7 },
            {"inset", 7 },
        };

        private static void ResolveCellBorder(XElement border1, XElement border2)
        {
            if (border1 == null || border2 == null)
            {
                return;
            }

            if ((string)border1.Attribute(W.val) == "nil" || (string)border2.Attribute(W.val) == "nil")
            {
                return;
            }

            if ((string)border1.Attribute(W.sz) == "nil" || (string)border2.Attribute(W.sz) == "nil")
            {
                return;
            }

            var border1Val = (string)border1.Attribute(W.val);
            var border1Weight = 1;
            if (BorderNumber.ContainsKey(border1Val))
            {
                border1Weight = BorderNumber[border1Val];
            }

            var border2Val = (string)border2.Attribute(W.val);
            var border2Weight = 1;
            if (BorderNumber.ContainsKey(border2Val))
            {
                border2Weight = BorderNumber[border2Val];
            }

            if (border1Weight != border2Weight)
            {
                if (border1Weight < border2Weight)
                {
                    BorderOverride(border2, border1);
                }
                else
                {
                    BorderOverride(border1, border2);
                }
            }

            if ((decimal)border1.Attribute(W.sz) > (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border1, border2);
                return;
            }

            if ((decimal)border1.Attribute(W.sz) < (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border2, border1);
                return;
            }

            var border1Type = (string)border1.Attribute(W.val);
            var border2Type = (string)border2.Attribute(W.val);
            if (BorderTypePriority.ContainsKey(border1Type) &&
                BorderTypePriority.ContainsKey(border2Type))
            {
                var border1Pri = BorderTypePriority[border1Type];
                var border2Pri = BorderTypePriority[border2Type];
                if (border1Pri < border2Pri)
                {
                    BorderOverride(border2, border1);
                    return;
                }
                if (border2Pri < border1Pri)
                {
                    BorderOverride(border1, border2);
                    return;
                }
            }

            var color1Str = (string)border1.Attribute(W.color);
            if (color1Str == "auto")
            {
                color1Str = "000000";
            }

            var color2Str = (string)border2.Attribute(W.color);
            if (color2Str == "auto")
            {
                color2Str = "000000";
            }

            if (color1Str != null && color2Str != null && color1Str != color2Str)
            {
                try
                {
                    var color1 = Convert.ToInt32(color1Str, 16);
                    var color2 = Convert.ToInt32(color2Str, 16);
                    if (color1 < color2)
                    {
                        BorderOverride(border1, border2);
                        return;
                    }
                    if (color2 < color1)
                    {
                        BorderOverride(border2, border1);
                    }
                }
                catch (ArgumentException)
                {
                    // Ignore
                }
                catch (FormatException)
                {
                    // Ignore
                }
                catch (OverflowException)
                {
                    // Ignore
                }
            }
        }

        private static void BorderOverride(XElement fromBorder, XElement toBorder)
        {
            toBorder.Attribute(W.val).Value = fromBorder.Attribute(W.val).Value;
            if (fromBorder.Attribute(W.color) != null)
            {
                toBorder.SetAttributeValue(W.color, fromBorder.Attribute(W.color).Value);
            }

            if (fromBorder.Attribute(W.sz) != null)
            {
                toBorder.SetAttributeValue(W.sz, fromBorder.Attribute(W.sz).Value);
            }

            if (fromBorder.Attribute(W.themeColor) != null)
            {
                toBorder.SetAttributeValue(W.themeColor, fromBorder.Attribute(W.themeColor).Value);
            }

            if (fromBorder.Attribute(W.themeTint) != null)
            {
                toBorder.SetAttributeValue(W.themeTint, fromBorder.Attribute(W.themeTint).Value);
            }
        }

        private static void CalculateSpanWidthForTabs(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done for all content parts, not just the main document part.

            // w:defaultTabStop in settings
            var sxd = wordDoc.MainDocumentPart?.DocumentSettingsPart?.GetXDocument();
            var defaultTabStopValue = (string)sxd?.Descendants(W.defaultTabStop).Attributes(W.val).FirstOrDefault();
            var defaultTabStop = defaultTabStopValue != null ? WordprocessingMLUtil.StringToTwips(defaultTabStopValue) : 720;

            var pxd = wordDoc.MainDocumentPart?.GetXDocument();
            var root = pxd?.Root;
            if (root == null)
            {
                return;
            }

            var newRoot = (XElement)CalculateSpanWidthTransform(root, defaultTabStop);
            root.ReplaceWith(newRoot);
            wordDoc.MainDocumentPart?.PutXDocument();
        }

        private static object CalculateSpanWidthTransform(XNode node, int defaultTabStop)
        {
            if (!(node is XElement element))
            {
                return node;
            }

            // if it is not a paragraph or if there are no tabs in the paragraph, then no need to continue processing.
            if (element.Name != W.p || !element.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.r).Elements(W.tab).Any())
            {
                // TODO: Revisit. Can we just return the node if it is a paragraph that does not have any tab?
                return new XElement(element.Name, element.Attributes(), element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop)));
            }

            var clonedPara = new XElement(element);

            var leftInTwips = 0;
            var firstInTwips = 0;

            var ind = clonedPara.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
            if (ind != null)
            {
                // todo need to handle start and end attributes

                var left = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.left));
                if (left != null)
                {
                    leftInTwips = (int)left;
                }

                var firstLine = 0;
                var firstLineAtt = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.firstLine));
                if (firstLineAtt != null)
                {
                    firstLine = (int)firstLineAtt;
                }

                var hangingAtt = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.hanging));
                if (hangingAtt != null)
                {
                    firstLine = -(int)hangingAtt;
                }

                firstInTwips = leftInTwips + firstLine;
            }

            // calculate the tab stops, in twips
            var tabs = clonedPara.Elements(W.pPr).Elements(W.tabs).FirstOrDefault();

            if (tabs == null)
            {
                if (leftInTwips == 0)
                {
                    tabs = new XElement(W.tabs,
                        Enumerable.Range(1, 100)
                            .Select(r => new XElement(W.tab,
                                new XAttribute(W.val, "left"),
                                new XAttribute(W.pos, r * defaultTabStop))));
                }
                else
                {
                    tabs = new XElement(W.tabs,
                        new XElement(W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, leftInTwips)));
                    tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
                }
            }
            else
            {
                if (leftInTwips != 0)
                {
                    tabs.Add(
                        new XElement(W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, leftInTwips)));
                }
                tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
            }

            var twipCounter = firstInTwips;
            var contentToMeasure = element.DescendantsTrimmed(z => z.Name == W.txbxContent || z.Name == W.pPr || z.Name == W.rPr).ToArray();
            var currentElementIdx = 0;
            while (true)
            {
                var currentElement = contentToMeasure[currentElementIdx];

                if (currentElement.Name == W.br)
                {
                    twipCounter = leftInTwips;

                    currentElement.Add(new XAttribute(PtOpenXml.TabWidth,
                        string.Format(NumberFormatInfo.InvariantInfo,
                            "{0:0.000}", firstInTwips / 1440m)));

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                    {
                        break;
                    }
                }

                if (currentElement.Name == W.tab)
                {
                    var runContainingTabToReplace = currentElement.Parent;
                    var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName") ??
                                      runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                    var testAmount = twipCounter;

                    var tabAfterText = tabs.Elements(W.tab).FirstOrDefault(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos)) > testAmount);

                    if (tabAfterText == null)
                    {
                        // something has gone wrong, so put 1/2 inch in
                        if (currentElement.Attribute(PtOpenXml.TabWidth) == null)
                        {
                            currentElement.Add(new XAttribute(PtOpenXml.TabWidth, 720m));
                        }

                        break;
                    }

                    var tabVal = (string)tabAfterText.Attribute(W.val);
                    if (tabVal == "right" || tabVal == "end")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements.TakeWhile(z => z.Name != W.tab && z.Name != W.br && z.Name != W.cr).ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun2 = new XElement(W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab));

                        var widthOfTextAfterTab = CalcWidthOfRunInTwips(dummyRun2);
                        var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfTextAfterTab - twipCounter;
                        if (delta2 < 0)
                        {
                            delta2 = 0;
                        }

                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", delta2 / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)), twipCounter + widthOfTextAfterTab);

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement == null)
                        {
                            break;
                        }

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                        {
                            break;
                        }

                        continue;
                    }
                    if (tabVal == "decimal")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z =>
                                z.Name != W.tab &&
                                z.Name != W.br &&
                                z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        if (textAfterTab.Contains("."))
                        {
                            var mantissa = textAfterTab.Split('.')[0];

                            var dummyRun4 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, mantissa));

                            var widthOfMantissa = CalcWidthOfRunInTwips(dummyRun4);
                            var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfMantissa - twipCounter;
                            if (delta2 < 0)
                            {
                                delta2 = 0;
                            }

                            currentElement.Add(
                                new XAttribute(PtOpenXml.TabWidth,
                                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", delta2 / 1440m)),
                                GetLeader(tabAfterText));

                            var decims = textAfterTab.Substring(textAfterTab.IndexOf('.'));
                            dummyRun4 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, decims));

                            var widthOfDecims = CalcWidthOfRunInTwips(dummyRun4);
                            twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) + widthOfDecims, twipCounter + widthOfMantissa + widthOfDecims);

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement == null)
                            {
                                break;
                            }

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                            {
                                break;
                            }

                            continue;
                        }
                        else
                        {
                            var dummyRun2 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, textAfterTab));

                            var widthOfTextAfterTab = CalcWidthOfRunInTwips(dummyRun2);
                            var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfTextAfterTab - twipCounter;
                            if (delta2 < 0)
                            {
                                delta2 = 0;
                            }

                            currentElement.Add(
                                new XAttribute(PtOpenXml.TabWidth,
                                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", delta2 / 1440m)),
                                GetLeader(tabAfterText));
                            twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)), twipCounter + widthOfTextAfterTab);

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement == null)
                            {
                                break;
                            }

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                            {
                                break;
                            }

                            continue;
                        }
                    }
                    if ((string)tabAfterText.Attribute(W.val) == "center")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z =>
                                z.Name != W.tab &&
                                z.Name != W.br &&
                                z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun4 = new XElement(W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab));

                        var widthOfText = CalcWidthOfRunInTwips(dummyRun4);
                        var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfText / 2 - twipCounter;
                        if (delta2 < 0)
                        {
                            delta2 = 0;
                        }

                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", delta2 / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) + widthOfText / 2, twipCounter + widthOfText);

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement == null)
                        {
                            break;
                        }

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                        {
                            break;
                        }

                        continue;
                    }
                    if (tabVal == "left" || tabVal == "start" || tabVal == "num")
                    {
                        var delta = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - twipCounter;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", delta / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos));

                        currentElementIdx++;
                        if (currentElementIdx >= contentToMeasure.Length)
                        {
                            break;
                        }

                        continue;
                    }
                }

                if (currentElement.Name == W.t)
                {
                    const int widthOfText = 0;
                    currentElement.Add(new XAttribute(PtOpenXml.TabWidth, string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", widthOfText / 1440m)));
                    twipCounter += widthOfText;

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                    {
                        break;
                    }

                    continue;
                }

                currentElementIdx++;
                if (currentElementIdx >= contentToMeasure.Length)
                {
                    break;
                }
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop)));
        }

        private static XAttribute? GetLeader(XElement tabAfterText)
        {
            var leader = (string)tabAfterText.Attribute(W.leader);
            if (leader == null)
            {
                return null;
            }

            return new XAttribute(PtOpenXml.Leader, leader);
        }

        private static XElement AddDefaultTabsAfterLastTab(XElement tabs, int defaultTabStop)
        {
            var lastTabElement = tabs
                .Elements(W.tab)
                .Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar")
                .OrderBy(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos)))
                .LastOrDefault();
            if (lastTabElement != null)
            {
                if (defaultTabStop == 0)
                {
                    defaultTabStop = 720;
                }

                var rangeStart = WordprocessingMLUtil.StringToTwips((string)lastTabElement.Attribute(W.pos)) / defaultTabStop + 1;
                var tempTabs = new XElement(W.tabs,
                    tabs.Elements().Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar"),
                    Enumerable.Range(rangeStart, 100)
                    .Select(r => new XElement(W.tab,
                        new XAttribute(W.val, "left"),
                        new XAttribute(W.pos, r * defaultTabStop))));
                tempTabs = new XElement(W.tabs,
                    tempTabs.Elements().OrderBy(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos))));
                return tempTabs;
            }
            else
            {
                tabs = new XElement(W.tabs,
                    Enumerable.Range(1, 100)
                    .Select(r => new XElement(W.tab,
                        new XAttribute(W.val, "left"),
                        new XAttribute(W.pos, r * defaultTabStop))));
            }
            return tabs;
        }

        private static readonly HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string>? _knownFamilies;

        private static HashSet<string> KnownFamilies
        {
            get
            {
                if (_knownFamilies == null)
                {
                    _knownFamilies = new HashSet<string>();
                    var families = SixLabors.Fonts.SystemFonts.Families;
                    foreach (var fam in families)
                    {
                        _knownFamilies.Add(fam.Name);
                    }
                }
                return _knownFamilies;
            }
        }

        private static int CalcWidthOfRunInTwips(XElement r)
        {
            var fontName = (string)r.Attribute(PtOpenXml.pt + "FontName") ??
                           (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
            {
                throw new OpenXmlPowerToolsException("Internal Error, should have FontName attribute");
            }

            if (UnknownFonts.Contains(fontName))
            {
                return 0;
            }

            var rPr = r.Element(W.rPr);
            if (rPr == null)
            {
                throw new OpenXmlPowerToolsException("Internal Error, should have run properties");
            }

            var sz = GetFontSize(r) ?? 22m;

            // unknown font families will throw ArgumentException, in which case just return 0
            if (!KnownFamilies.Contains(fontName))
            {
                return 0;
            }

            // in theory, all unknown fonts are found by the above test, but if not...
            SixLabors.Fonts.FontFamily ff;

            try
            {
                ff = SixLabors.Fonts.SystemFonts.Families.Single(font => font.Name == fontName);
            }
            catch (ArgumentException)
            {
                UnknownFonts.Add(fontName);

                return 0;
            }

            var fs = SixLabors.Fonts.FontStyle.Regular;
            if (GetBoolProp(rPr, W.b) || GetBoolProp(rPr, W.bCs))
            {
                fs |= SixLabors.Fonts.FontStyle.Bold;
            }

            if (GetBoolProp(rPr, W.i) || GetBoolProp(rPr, W.iCs))
            {
                fs |= SixLabors.Fonts.FontStyle.Italic;
            }

            // Appended blank as a quick fix to accommodate &nbsp; that will get appended to some layout-critical runs such as list item numbers. In some cases, this might not be required or even wrong, so this must be revisited.
            var runText = r.DescendantsTrimmed(W.txbxContent).Where(e => e.Name == W.t).Select(t => (string)t).StringConcatenate() + " ";

            var tabLength = r.DescendantsTrimmed(W.txbxContent).Where(e => e.Name == W.tab).Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth)).Sum();

            if (runText.Length == 0 && tabLength == 0)
            {
                return 0;
            }

            var multiplier = 1;
            if (runText.Length <= 2)
            {
                multiplier = 100;
            }
            else if (runText.Length <= 4)
            {
                multiplier = 50;
            }
            else if (runText.Length <= 8)
            {
                multiplier = 25;
            }
            else if (runText.Length <= 16)
            {
                multiplier = 12;
            }
            else if (runText.Length <= 32)
            {
                multiplier = 6;
            }

            if (multiplier != 1)
            {
                var sb = new StringBuilder();
                for (var i = 0; i < multiplier; i++)
                {
                    sb.Append(runText);
                }

                runText = sb.ToString();
            }

            var w = MetricsGetter.GetTextWidth(ff, fs, sz, runText);

            return (int)(w / 96m * 1440m / multiplier + tabLength * 1440m);
        }

        private static void InsertAppropriateNonbreakingSpaces(WordprocessingDocument wordDoc)
        {
            foreach (var part in wordDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                var root = pxd.Root;
                if (root == null)
                {
                    return;
                }

                var newRoot = (XElement)InsertAppropriateNonbreakingSpacesTransform(root);
                root.ReplaceWith(newRoot);
                part.PutXDocument();
            }
        }

        // Non-breaking spaces are not required if we use appropriate CSS, i.e., "white-space: pre-wrap;".
        // We only need to make sure that empty w:p elements are translated into non-empty h:p elements, because empty h:p elements would be ignored by browsers.
        // Further, in addition to not being required, non-breaking spaces would change the layout behavior of spans having consecutive spaces. Therefore, avoiding non-breaking spaces has the additional benefit of leading to a more faithful representation of the Word document in HTML.
        private static object InsertAppropriateNonbreakingSpacesTransform(XNode node)
        {
            if (node is XElement element)
            {
                // child content of run to look for
                // W.br
                // W.cr
                // W.dayLong
                // W.dayShort
                // W.drawing
                // W.monthLong
                // W.monthShort
                // W.noBreakHyphen
                // W.object
                // W.pgNum
                // W.pTab
                // W.separator
                // W.softHyphen
                // W.sym
                // W.t
                // W.tab
                // W.yearLong
                // W.yearShort
                if (element.Name == W.p)
                {
                    // Translate empty paragraphs to paragraphs having one run with
                    // a normal space. A non-breaking space, i.e., \x00A0, is not
                    // required if we use appropriate CSS.
                    var hasContent = element
                        .Elements()
                        .Where(e => e.Name != W.pPr)
                        .DescendantsAndSelf()
                        .Any(e =>
                            e.Name == W.dayLong ||
                            e.Name == W.dayShort ||
                            e.Name == W.drawing ||
                            e.Name == W.monthLong ||
                            e.Name == W.monthShort ||
                            e.Name == W.noBreakHyphen ||
                            e.Name == W._object ||
                            e.Name == W.pgNum ||
                            e.Name == W.ptab ||
                            e.Name == W.separator ||
                            e.Name == W.softHyphen ||
                            e.Name == W.sym ||
                            e.Name == W.t ||
                            e.Name == W.tab ||
                            e.Name == W.yearLong ||
                            e.Name == W.yearShort
                        );
                    if (!hasContent)
                    {
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n)),
                            new XElement(W.r,
                                element.Elements(W.pPr).Elements(W.rPr),
                                new XElement(W.t, " ")));
                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n)));
            }
            return node;
        }

        private sealed class SectionAnnotation
        {
            public XElement? SectionElement;
        }

        private static void AnnotateForSections(WordprocessingDocument wordDoc)
        {
            var xd = wordDoc?.MainDocumentPart?.GetXDocument();

            var document = xd?.Root;
            if (document == null)
            {
                return;
            }

            var body = document.Element(W.body);
            if (body == null)
            {
                return;
            }

            // move last sectPr into last paragraph
            var lastSectPr = body.Elements(W.sectPr).LastOrDefault();
            if (lastSectPr != null)
            {
                // if the last thing in the document is a table, Word will always insert a paragraph following that.
                var lastPara = body
                    .DescendantsTrimmed(W.txbxContent)
                    .LastOrDefault(p => p.Name == W.p);

                if (lastPara != null)
                {
                    var lastParaProps = lastPara.Element(W.pPr);
                    if (lastParaProps != null)
                    {
                        lastParaProps.Add(lastSectPr);
                    }
                    else
                    {
                        lastPara.Add(new XElement(W.pPr, lastSectPr));
                    }

                    lastSectPr.Remove();
                }
            }

            var reverseDescendants = xd?.Descendants().Reverse().ToList();
            var currentSection = InitializeSectionAnnotation(reverseDescendants);

            foreach (var d in reverseDescendants)
            {
                if (d.Name == W.sectPr)
                {
                    if (d.Attribute(XNamespace.Xmlns + "w") == null)
                    {
                        d.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));
                    }

                    currentSection = new SectionAnnotation()
                    {
                        SectionElement = d
                    };
                }
                else
                {
                    d.AddAnnotation(currentSection);
                }
            }
        }

        private static SectionAnnotation InitializeSectionAnnotation(IEnumerable<XElement> reverseDescendants)
        {
            var currentSection = new SectionAnnotation()
            {
                SectionElement = reverseDescendants.FirstOrDefault(e => e.Name == W.sectPr)
            };
            if (currentSection.SectionElement != null &&
                currentSection.SectionElement.Attribute(XNamespace.Xmlns + "w") == null)
            {
                currentSection.SectionElement.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));
            }

            // todo what should the default section props be?
            if (currentSection.SectionElement == null)
            {
                currentSection = new SectionAnnotation()
                {
                    SectionElement = new XElement(W.sectPr,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        new XElement(W.pgSz,
                            new XAttribute(W._w, 12240),
                            new XAttribute(W.h, 15840)),
                        new XElement(W.pgMar,
                            new XAttribute(W.top, 1440),
                            new XAttribute(W.right, 1440),
                            new XAttribute(W.bottom, 1440),
                            new XAttribute(W.left, 1440),
                            new XAttribute(W.header, 720),
                            new XAttribute(W.footer, 720),
                            new XAttribute(W.gutter, 0)),
                        new XElement(W.cols,
                            new XAttribute(W.space, 720)),
                        new XElement(W.docGrid,
                            new XAttribute(W.linePitch, 360)))
                };
            }

            return currentSection;
        }

        private static object CreateBorderDivs(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, IEnumerable<XElement> elements)
        {
            var htmlNestedListTypes = new List<string>();
            return elements.GroupAdjacent(e =>
                {
                    var pBdr = e.Elements(W.pPr).Elements(W.pBdr).FirstOrDefault();
                    if (pBdr != null)
                    {
                        var indStr = string.Empty;
                        var ind = e.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
                        if (ind != null)
                        {
                            indStr = ind.ToString(SaveOptions.DisableFormatting);
                        }

                        return pBdr.ToString(SaveOptions.DisableFormatting) + indStr;
                    }
                    return e.Name == W.tbl ? "table" : string.Empty;
                })
                .Select(g =>
                {
                    if (string.IsNullOrEmpty(g.Key))
                    {
                        return (object)GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, 0m, htmlNestedListTypes);
                    }
                    if (g.Key == "table")
                    {
                        return g.Select(gc => ConvertToHtmlTransform(wordDoc, settings, gc, false, 0));
                    }
                    var pPr = g.First().Elements(W.pPr).First();
                    var pBdr = pPr.Element(W.pBdr);
                    var style = new Dictionary<string, string>();
                    GenerateBorderStyle(pBdr, W.top, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.right, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.bottom, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.left, style, BorderType.Paragraph);

                    var currentMarginLeft = 0m;
                    var ind = pPr.Element(W.ind);
                    if (ind != null)
                    {
                        var leftInInches = (decimal?)ind.Attribute(W.left) / 1440m ?? 0;
                        var hangingInInches = -(decimal?)ind.Attribute(W.hanging) / 1440m ?? 0;
                        currentMarginLeft = leftInInches + hangingInInches;

                        style.AddIfMissing("margin-left",
                            currentMarginLeft > 0m
                                ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", currentMarginLeft)
                                : "0");
                    }

                    var div = new XElement(Xhtml.div,
                        GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, currentMarginLeft, htmlNestedListTypes));
                    div.AddAnnotation(style);
                    return div;
                })
            .ToList();
        }

        private static IEnumerable<object> GroupAndVerticallySpaceNumberedParagraphs(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements,
            decimal currentMarginLeft,
            IList<string?> htmlNestedListTypes)
        {
            var grouped = elements
                .GroupAdjacent(e =>
                {
                    var abstractNumId = (string)e.Attribute(PtOpenXml.pt + "AbstractNumId");
                    if (abstractNumId != null)
                    {
                        return "num:" + abstractNumId;
                    }

                    var contextualSpacing = e.Elements(W.pPr).Elements(W.contextualSpacing).FirstOrDefault();
                    if (contextualSpacing != null)
                    {
                        var styleName = (string)e.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                        if (styleName == null)
                        {
                            return "";
                        }

                        return "sty:" + styleName;
                    }
                    return "";
                })
                .ToList();
            var newContent = grouped
                .Select(g =>
                {
                    if (string.IsNullOrEmpty(g.Key))
                    {
                        return g.Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                    }

                    if (g.Key.StartsWith("num:", StringComparison.Ordinal))
                    {
                        return ListAwareConvertToHtmlTransform(wordDoc, settings, g, currentMarginLeft, htmlNestedListTypes);
                    }

                    var last = g.Count() - 1;
                    return g.Select((e, i) => ConvertToHtmlTransform(wordDoc, settings, e, i != last, currentMarginLeft));
                });
            return newContent;
        }

        private static object? ProcessAlternateContent(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            foreach (var choice in element.Elements(MC.Choice).Concat(element.Elements(MC.Fallback)))
            {
                var childElements = choice.Elements()
                    .Select(child => ConvertToHtmlTransform(wordDoc, settings, child, false, 0m))
                    .Where(htmlElement => htmlElement != null)
                    .ToList();
                if (childElements.Count > 0)
                {
                    return childElements;
                }
            }

            return null;
        }

        private class BorderMappingInfo
        {
            public string CssName = string.Empty;
            public decimal CssSize;
        }

        private static readonly Dictionary<string, BorderMappingInfo> BorderStyleMap = new Dictionary<string, BorderMappingInfo>()
        {
            { "single", new BorderMappingInfo { CssName = "solid", CssSize = 1.0m }},
            { "dotted", new BorderMappingInfo { CssName = "dotted", CssSize = 1.0m }},
            { "dashSmallGap", new BorderMappingInfo { CssName = "dashed", CssSize = 1.0m }},
            { "dashed", new BorderMappingInfo { CssName = "dashed", CssSize = 1.0m }},
            { "dotDash", new BorderMappingInfo { CssName = "dashed", CssSize = 1.0m }},
            { "dotDotDash", new BorderMappingInfo { CssName = "dashed", CssSize = 1.0m }},
            { "double", new BorderMappingInfo { CssName = "double", CssSize = 2.5m }},
            { "triple", new BorderMappingInfo { CssName = "double", CssSize = 2.5m }},
            { "thinThickSmallGap", new BorderMappingInfo { CssName = "double", CssSize = 4.5m }},
            { "thickThinSmallGap", new BorderMappingInfo { CssName = "double", CssSize = 4.5m }},
            { "thinThickThinSmallGap", new BorderMappingInfo { CssName = "double", CssSize = 6.0m }},
            { "thickThinMediumGap", new BorderMappingInfo { CssName = "double", CssSize = 6.0m }},
            { "thinThickMediumGap", new BorderMappingInfo { CssName = "double", CssSize = 6.0m }},
            { "thinThickThinMediumGap", new BorderMappingInfo { CssName = "double", CssSize = 9.0m }},
            { "thinThickLargeGap", new BorderMappingInfo { CssName = "double", CssSize = 5.25m }},
            { "thickThinLargeGap", new BorderMappingInfo { CssName = "double", CssSize = 5.25m }},
            { "thinThickThinLargeGap", new BorderMappingInfo { CssName = "double", CssSize = 9.0m }},
            { "wave", new BorderMappingInfo { CssName = "solid", CssSize = 3.0m }},
            { "doubleWave", new BorderMappingInfo { CssName = "double", CssSize = 5.25m }},
            { "dashDotStroked", new BorderMappingInfo { CssName = "solid", CssSize = 3.0m }},
            { "threeDEmboss", new BorderMappingInfo { CssName = "ridge", CssSize = 6.0m }},
            { "threeDEngrave", new BorderMappingInfo { CssName = "groove", CssSize = 6.0m }},
            { "outset", new BorderMappingInfo { CssName = "outset", CssSize = 4.5m }},
            { "inset", new BorderMappingInfo { CssName = "inset", CssSize = 4.5m }},
        };

        private static void GenerateBorderStyle(XElement pBdr, XName sideXName, Dictionary<string, string> style, BorderType borderType)
        {
            string whichSide;
            if (sideXName == W.top)
            {
                whichSide = "top";
            }
            else if (sideXName == W.right)
            {
                whichSide = "right";
            }
            else if (sideXName == W.bottom)
            {
                whichSide = "bottom";
            }
            else
            {
                whichSide = "left";
            }

            if (pBdr == null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right"))
                {
                    style.Add("padding-" + whichSide, "5.4pt");
                }

                return;
            }

            var side = pBdr.Element(sideXName);
            if (side == null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell && (whichSide == "left" || whichSide == "right"))
                {
                    style.Add("padding-" + whichSide, "5.4pt");
                }

                return;
            }
            var type = (string)side.Attribute(W.val);
            if (type == "nil" || type == "none")
            {
                style.Add("border-" + whichSide + "-style", "none");

                var space = (decimal?)side.Attribute(W.space) ?? 0;
                if (borderType == BorderType.Cell && (whichSide == "left" || whichSide == "right") && space < 5.4m)
                {
                    space = 5.4m;
                }

                style.Add("padding-" + whichSide,
                    space == 0 ? "0" : string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", space));
            }
            else
            {
                var sz = (int)side.Attribute(W.sz);
                var space = (decimal?)side.Attribute(W.space) ?? 0;
                var color = (string)side.Attribute(W.color);
                if (color == null || color == "auto")
                {
                    color = "windowtext";
                }
                else
                {
                    color = ConvertColor(color);
                }

                var borderWidthInPoints = Math.Max(1m, Math.Min(96m, Math.Max(2m, sz)) / 8m);

                var borderStyle = "solid";
                if (BorderStyleMap.ContainsKey(type))
                {
                    var borderInfo = BorderStyleMap[type];
                    borderStyle = borderInfo.CssName;
                    if (type == "double")
                    {
                        if (sz <= 8)
                        {
                            borderWidthInPoints = 2.5m;
                        }
                        else if (sz <= 18)
                        {
                            borderWidthInPoints = 6.75m;
                        }
                        else
                        {
                            borderWidthInPoints = sz / 3m;
                        }
                    }
                    else if (type == "triple")
                    {
                        if (sz <= 8)
                        {
                            borderWidthInPoints = 8m;
                        }
                        else if (sz <= 18)
                        {
                            borderWidthInPoints = 11.25m;
                        }
                        else
                        {
                            borderWidthInPoints = 11.25m;
                        }
                    }
                    else if (type.ToLower().Contains("dash"))
                    {
                        if (sz <= 4)
                        {
                            borderWidthInPoints = 1m;
                        }
                        else if (sz <= 12)
                        {
                            borderWidthInPoints = 1.5m;
                        }
                        else
                        {
                            borderWidthInPoints = 2m;
                        }
                    }
                    else if (type != "single")
                    {
                        borderWidthInPoints = borderInfo.CssSize;
                    }
                }
                if (type == "outset" || type == "inset")
                {
                    color = "";
                }

                var borderWidth = string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", borderWidthInPoints);

                style.Add("border-" + whichSide, borderStyle + " " + color + " " + borderWidth);
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right") && space < 5.4m)
                {
                    space = 5.4m;
                }

                style.Add("padding-" + whichSide, space == 0 ? "0" : string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", space));
            }
        }

        private static readonly Dictionary<string, Func<string, string, string>> ShadeMapper = new Dictionary<string, Func<string, string, string>>()
        {
            { "auto", (c, f) => c },
            { "clear", (c, f) => f },
            { "nil", (c, f) => f },
            { "solid", (c, f) => c },
            { "diagCross", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "diagStripe", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "horzCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "horzStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "pct10", (c, f) => ConvertColorFillPct(c, f, .1) },
            { "pct12", (c, f) => ConvertColorFillPct(c, f, .125) },
            { "pct15", (c, f) => ConvertColorFillPct(c, f, .15) },
            { "pct20", (c, f) => ConvertColorFillPct(c, f, .2) },
            { "pct25", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "pct30", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "pct35", (c, f) => ConvertColorFillPct(c, f, .35) },
            { "pct37", (c, f) => ConvertColorFillPct(c, f, .375) },
            { "pct40", (c, f) => ConvertColorFillPct(c, f, .4) },
            { "pct45", (c, f) => ConvertColorFillPct(c, f, .45) },
            { "pct50", (c, f) => ConvertColorFillPct(c, f, .50) },
            { "pct55", (c, f) => ConvertColorFillPct(c, f, .55) },
            { "pct60", (c, f) => ConvertColorFillPct(c, f, .60) },
            { "pct62", (c, f) => ConvertColorFillPct(c, f, .625) },
            { "pct65", (c, f) => ConvertColorFillPct(c, f, .65) },
            { "pct70", (c, f) => ConvertColorFillPct(c, f, .7) },
            { "pct75", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "pct80", (c, f) => ConvertColorFillPct(c, f, .8) },
            { "pct85", (c, f) => ConvertColorFillPct(c, f, .85) },
            { "pct87", (c, f) => ConvertColorFillPct(c, f, .875) },
            { "pct90", (c, f) => ConvertColorFillPct(c, f, .9) },
            { "pct95", (c, f) => ConvertColorFillPct(c, f, .95) },
            { "reverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinHorzCross", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "thinHorzStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinReverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinVertStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
        };

        private static readonly Dictionary<string, string> ShadeCache = new Dictionary<string, string>();

        // fill is the background, color is the foreground
        private static string ConvertColorFillPct(string color, string fill, double pct)
        {
            if (color == "auto")
            {
                color = "000000";
            }

            if (fill == "auto")
            {
                fill = "ffffff";
            }

            var key = color + fill + pct.ToString(CultureInfo.InvariantCulture);
            if (ShadeCache.ContainsKey(key))
            {
                return ShadeCache[key];
            }

            var fillRed = Convert.ToInt32(fill.Substring(0, 2), 16);
            var fillGreen = Convert.ToInt32(fill.Substring(2, 2), 16);
            var fillBlue = Convert.ToInt32(fill.Substring(4, 2), 16);
            var colorRed = Convert.ToInt32(color.Substring(0, 2), 16);
            var colorGreen = Convert.ToInt32(color.Substring(2, 2), 16);
            var colorBlue = Convert.ToInt32(color.Substring(4, 2), 16);
            var finalRed = (int)(fillRed - (fillRed - colorRed) * pct);
            var finalGreen = (int)(fillGreen - (fillGreen - colorGreen) * pct);
            var finalBlue = (int)(fillBlue - (fillBlue - colorBlue) * pct);
            var returnValue = string.Format("{0:x2}{1:x2}{2:x2}", finalRed, finalGreen, finalBlue);
            ShadeCache.Add(key, returnValue);
            return returnValue;
        }

        private static void CreateStyleFromShd(Dictionary<string, string> style, XElement shd)
        {
            if (shd == null)
            {
                return;
            }

            var shadeType = (string)shd.Attribute(W.val);
            var color = (string)shd.Attribute(W.color);
            var fill = (string)shd.Attribute(W.fill);
            if (ShadeMapper.ContainsKey(shadeType))
            {
                color = ShadeMapper[shadeType](color, fill);
            }
            if (color != null)
            {
                var cvtColor = ConvertColor(color);
                if (!string.IsNullOrEmpty(cvtColor))
                {
                    style.AddIfMissing("background", cvtColor);
                }
            }
        }

        private static readonly Dictionary<string, string> NamedColors = new Dictionary<string, string>()
        {
            {"black", "black"},
            {"blue", "blue" },
            {"cyan", "aqua" },
            {"green", "green" },
            {"magenta", "fuchsia" },
            {"red", "red" },
            {"yellow", "yellow" },
            {"white", "white" },
            {"darkBlue", "#00008B" },
            {"darkCyan", "#008B8B" },
            {"darkGreen", "#006400" },
            {"darkMagenta", "#800080" },
            {"darkRed", "#8B0000" },
            {"darkYellow", "#808000" },
            {"darkGray", "#A9A9A9" },
            {"lightGray", "#D3D3D3" },
            {"none", "" },
        };

        private static void CreateColorProperty(string propertyName, string color, Dictionary<string, string> style)
        {
            if (color == null)
            {
                return;
            }

            // "auto" color is black for "color" and white for "background" property.
            if (color == "auto")
            {
                color = propertyName == "color" ? "black" : "white";
            }

            if (NamedColors.ContainsKey(color))
            {
                var lc = NamedColors[color];
                if (string.IsNullOrEmpty(lc))
                {
                    return;
                }

                style.AddIfMissing(propertyName, lc);
                return;
            }
            style.AddIfMissing(propertyName, "#" + color);
        }

        private static string ConvertColor(string color)
        {
            // "auto" color is black for "color" and white for "background" property.
            // As this method is only called for "background" colors, "auto" is translated
            // to "white" and never "black".
            if (color == "auto")
            {
                color = "white";
            }

            if (NamedColors.ContainsKey(color))
            {
                var lc = NamedColors[color];
                if (string.IsNullOrEmpty(lc))
                {
                    return "black";
                }

                return lc;
            }
            return "#" + color;
        }

        private static readonly Dictionary<string, string> FontFallback = new Dictionary<string, string>()
        {
            { "Arial", @"'{0}', 'sans-serif'" },
            { "Arial Narrow", @"'{0}', 'sans-serif'" },
            { "Arial Rounded MT Bold", @"'{0}', 'sans-serif'" },
            { "Arial Unicode MS", @"'{0}', 'sans-serif'" },
            { "Baskerville Old Face", @"'{0}', 'serif'" },
            { "Berlin Sans FB", @"'{0}', 'sans-serif'" },
            { "Berlin Sans FB Demi", @"'{0}', 'sans-serif'" },
            { "Calibri Light", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT Condensed", @"'{0}', 'sans-serif'" },
            { "Lucida Sans", @"'{0}', 'sans-serif'" },
            { "Lucida Sans Unicode", @"'{0}', 'sans-serif'" },
            { "Segoe UI", @"'{0}', 'sans-serif'" },
            { "Segoe UI Light", @"'{0}', 'sans-serif'" },
            { "Segoe UI Semibold", @"'{0}', 'sans-serif'" },
            { "Tahoma", @"'{0}', 'sans-serif'" },
            { "Trebuchet MS", @"'{0}', 'sans-serif'" },
            { "Verdana", @"'{0}', 'sans-serif'" },
            { "Book Antiqua", @"'{0}', 'serif'" },
            { "Bookman Old Style", @"'{0}', 'serif'" },
            { "Californian FB", @"'{0}', 'serif'" },
            { "Cambria", @"'{0}', 'serif'" },
            { "Constantia", @"'{0}', 'serif'" },
            { "Garamond", @"'{0}', 'serif'" },
            { "Lucida Bright", @"'{0}', 'serif'" },
            { "Lucida Fax", @"'{0}', 'serif'" },
            { "Palatino Linotype", @"'{0}', 'serif'" },
            { "Times New Roman", @"'{0}', 'serif'" },
            { "Wide Latin", @"'{0}', 'serif'" },
            { "Courier New", @"'{0}'" },
            { "Lucida Console", @"'{0}'" },
        };

        private static void CreateFontCssProperty(string font, Dictionary<string, string> style)
        {
            if (FontFallback.ContainsKey(font))
            {
                style.AddIfMissing("font-family", string.Format(FontFallback[font], font));
                return;
            }
            style.AddIfMissing("font-family", font);
        }

        private static bool GetBoolProp(XElement runProps, XName xName)
        {
            var p = runProps.Element(xName);
            if (p == null)
            {
                return false;
            }

            var v = p.Attribute(W.val);
            if (v == null)
            {
                return true;
            }

            var s = v.Value.ToLower();
            if (s == "0" || s == "false")
            {
                return false;
            }

            if (s == "1" || s == "true")
            {
                return true;
            }

            return false;
        }

        private static object ConvertContentThatCanContainFields(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, IEnumerable<XElement> elements)
        {
            var grouped = elements.GroupAdjacent(e =>
                {
                    var stack = e.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                    return stack == null || !stack.Any() ? (int?)null : stack.Select(st => st.Id).Min();
                })
                .ToList();

            var txformed = grouped
                .Select(g =>
                {
                    var key = g.Key;
                    if (key == null)
                    {
                        return (object)g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));
                    }

                    var instrText = FieldRetriever.InstrText(g.First().Ancestors().Last(), (int)key)
                        .TrimStart('{').TrimEnd('}');

                    var parsed = FieldRetriever.ParseField(instrText);
                    if (parsed.FieldType != "HYPERLINK")
                    {
                        return g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));
                    }

                    var content = g.DescendantsAndSelf(W.r).Select(run => ConvertRun(wordDoc, settings, run));
                    var a = parsed.Arguments.Length > 0
                        ? new XElement(Xhtml.a, new XAttribute("href", parsed.Arguments[0]), content)
                        : new XElement(Xhtml.a, content);
                    var a2 = a;
                    if (!a2.Nodes().Any())
                    {
                        a2.Add(new XText(""));
                        return a2;
                    }
                    return a;
                })
                .ToList();

            return txformed;
        }

        // Don't process wmf files (with contentType == "image/x-wmf") because GDI consumes huge amounts of memory when dealing with wmf perhaps because it loads a DLL to do the rendering?
        // It actually works, but is not recommended.
        private static readonly List<string> ImageContentTypes = new List<string>
        {
            "image/png", "image/gif", "image/tiff", "image/jpeg"
        };

        internal static XElement? ProcessImage(WordprocessingDocument wordDoc, XElement element, IImageHandler imageHandler)
        {
            if (imageHandler == null)
            {
                return null;
            }
            if (element.Name == W.drawing)
            {
                return ProcessDrawing(wordDoc, element, imageHandler);
            }
            if (element.Name == W.pict || element.Name == W._object)
            {
                return ProcessPictureOrObject(wordDoc, element, imageHandler);
            }
            return null;
        }

        private static XElement? ProcessDrawing(WordprocessingDocument wordDoc, XElement element, IImageHandler imageHandler)
        {
            var containerElement = element.Elements()
                .FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            if (containerElement == null)
            {
                return null;
            }

            string? hyperlinkUri = null;
            var hyperlinkElement = element
                .Elements(WP.inline)
                .Elements(WP.docPr)
                .Elements(A.hlinkClick)
                .FirstOrDefault();
            if (hyperlinkElement != null)
            {
                var rId = (string)hyperlinkElement.Attribute(R.id);
                if (rId != null)
                {
                    var hyperlinkRel = wordDoc.MainDocumentPart?.HyperlinkRelationships.FirstOrDefault(hlr => hlr.Id == rId);
                    if (hyperlinkRel != null)
                    {
                        hyperlinkUri = hyperlinkRel.Uri.ToString();
                    }
                }
            }

            var extentCx = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cx).FirstOrDefault();
            var extentCy = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cy).FirstOrDefault();
            var altText = (string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.descr).FirstOrDefault() ?? (string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.name).FirstOrDefault() ?? "";
            var blipFill = containerElement.Elements(A.graphic).Elements(A.graphicData).Elements(Pic._pic).Elements(Pic.blipFill).FirstOrDefault();

            if (blipFill == null)
            {
                return null;
            }

            var imageRid = (string)blipFill.Elements(A.blip).Attributes(R.embed).FirstOrDefault();

            if (imageRid == null)
            {
                return null;
            }

            var pp3 = wordDoc.MainDocumentPart?.Parts.FirstOrDefault(pp => pp.RelationshipId == imageRid);

            if (pp3 == default)
            {
                return null;
            }

            var imagePart = pp3?.OpenXmlPart as ImagePart;

            if (imagePart == null)
            {
                return null;
            }

            // If the image markup points to a NULL image, then following will throw an ArgumentOutOfRangeException
            try
            {
                var mainDocumentPart = wordDoc.MainDocumentPart;
                imagePart = mainDocumentPart?.GetPartById(imageRid) as ImagePart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }

            var contentType = imagePart?.ContentType;
            if (contentType == null || !ImageContentTypes.Contains(contentType))
            {
                return null;
            }

            using var partStream = imagePart?.GetStream();
            if (extentCx != null && extentCy != null)
            {
                var imageInfo = new ImageInfo
                {
                    Image = partStream,
                    ImgStyleAttribute = new XAttribute("style",
                        string.Format(NumberFormatInfo.InvariantInfo,
                            "width: {0}in; height: {1}in",
                            (float)extentCx / ImageInfo.EmusPerInch,
                            (float)extentCy / ImageInfo.EmusPerInch)),
                    ContentType = contentType,
                    DrawingElement = element,
                    AltText = altText,
                };
                var imgElement2 = imageHandler.TransformImage(imageInfo);
                if (hyperlinkUri != null)
                {
                    return new XElement(XhtmlNoNamespace.a,
                        new XAttribute(XhtmlNoNamespace.href, hyperlinkUri),
                        imgElement2);
                }
                return imgElement2;
            }

            var imageInfo2 = new ImageInfo
            {
                Image = partStream,
                ContentType = contentType,
                DrawingElement = element,
                AltText = altText,
            };
            var imgElement = imageHandler.TransformImage(imageInfo2);

            if (hyperlinkUri != null)
            {
                return new XElement(XhtmlNoNamespace.a, new XAttribute(XhtmlNoNamespace.href, hyperlinkUri), imgElement);
            }
            return imgElement;
        }

        private static XElement? ProcessPictureOrObject(WordprocessingDocument wordDoc, XElement element, IImageHandler imageHandler)
        {
            var imageRid = (string?)element.Elements(VML.shape).Elements(VML.imagedata).Attributes(R.id).FirstOrDefault();
            if (imageRid == null)
            {
                // Check if this is horizontal line
                var rects = element.Elements(VML.rect).Take(2).ToList();
                if (rects.Count == 1)
                {
                    var rect = rects[0];
                    var horizontal = (string?)rect.Attribute(O.hr);
                    var horizontalStandard = (string?)rect.Attribute(O.hrstd);
                    if (ConvertTrueFalseValueToBool(horizontal)
                        && ConvertTrueFalseValueToBool(horizontalStandard))
                    {
                        return new XElement(Xhtml.hr);
                    }
                }

                return null;
            }

            try
            {
                var pp = wordDoc.MainDocumentPart?.Parts.FirstOrDefault(pp2 => pp2.RelationshipId == imageRid);
                if (!pp.HasValue)
                {
                    return null;
                }

                var imagePart = (ImagePart)pp.Value.OpenXmlPart;
                if (imagePart == null)
                {
                    return null;
                }

                var contentType = imagePart.ContentType;
                if (!ImageContentTypes.Contains(contentType))
                {
                    return null;
                }

                using var partStream = imagePart.GetStream();
                try
                {
                    var imageInfo = new ImageInfo()
                    {
                        Image = partStream,
                        ContentType = contentType,
                        DrawingElement = element
                    };

                    var style = (string)element.Elements(VML.shape).Attributes("style").FirstOrDefault();
                    if (style == null)
                    {
                        return imageHandler.TransformImage(imageInfo);
                    }

                    var tokens = style.Split(';');
                    var widthInPoints = WidthInPoints(tokens);
                    var heightInPoints = HeightInPoints(tokens);
                    if (widthInPoints != null && heightInPoints != null)
                    {
                        imageInfo.ImgStyleAttribute = new XAttribute("style",
                            string.Format(NumberFormatInfo.InvariantInfo,
                                "width: {0}pt; height: {1}pt", widthInPoints, heightInPoints));
                    }
                    return imageHandler.TransformImage(imageInfo);
                }
                catch (OutOfMemoryException)
                {
                    // the Bitmap class can throw OutOfMemoryException, which means the bitmap is messed up, so punt.
                    return null;
                }
                catch (ArgumentException)
                {
                    return null;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static float? HeightInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "height");
        }

        private static float? WidthInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "width");
        }

        private static float? SizeInPoints(IEnumerable<string> tokens, string name)
        {
            var sizeString = tokens
                .Select(t => new
                {
                    Name = t.Split(':').First(),
                    Value = t.Split(':').Skip(1).Take(1).FirstOrDefault()
                })
                .Where(p => p.Name == name)
                .Select(p => p.Value)
                .FirstOrDefault();

            if (sizeString != null && sizeString.Length > 2 && sizeString.Substring(sizeString.Length - 2) == "pt" && float.TryParse(sizeString.Substring(0, sizeString.Length - 2), out var size))
            {
                return size;
            }
            return null;
        }

        private static bool ConvertTrueFalseValueToBool(string? trueFalseValue)
        {
            var returnValue = false;
            if (trueFalseValue != null
                && (trueFalseValue.Equals("t", StringComparison.OrdinalIgnoreCase) || trueFalseValue.Equals("true", StringComparison.OrdinalIgnoreCase)))
            {
                returnValue = true;
            }

            return returnValue;
        }
    }
}