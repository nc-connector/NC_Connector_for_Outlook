/**
 * Copyright (c) 2026 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using Ganss.Xss;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Central HTML sanitizer for backend-provided template HTML.
     * Uses a Thunderbird-aligned allowlist/forbidlist policy and fails closed.
     */
    internal static class HtmlTemplateSanitizer
    {
        private static readonly HashSet<string> SanitizerDependencyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "HtmlSanitizer",
            "AngleSharp",
            "AngleSharp.Css",
            "System.Buffers",
            "System.Collections.Immutable",
            "System.Memory",
            "System.Runtime.CompilerServices.Unsafe",
            "System.Text.Encoding.CodePages"
        };

        private static readonly string[] ForbiddenTags = new[]
        {
            "script",
            "style",
            "iframe",
            "object",
            "embed",
            "link",
            "meta",
            "form",
            "input",
            "button",
            "textarea",
            "select",
            "option",
            "svg",
            "math"
        };

        private static readonly string[] AdditionalAttributes = new[]
        {
            "style",
            "target",
            "rel",
            "role",
            "width",
            "height",
            "colspan",
            "rowspan",
            "cellpadding",
            "cellspacing",
            "align",
            "valign"
        };

        private static readonly string[] AdditionalTags = new[]
        {
            "section",
            "article",
            "header",
            "footer"
        };

        // DOMPurify v3.3.1 (vendored in Thunderbird) html profile lists from vendor/purify.js:
        // - html$1 (tags)
        // - html (attributes)
        // TB runtime applies USE_PROFILES: { html: true }, ALLOW_DATA_ATTR: false,
        // then FORBID_TAGS + ADD_ATTR + ADD_TAGS.
        private static readonly string[] DomPurifyHtmlProfileTags = ParseTokenList(@"
a abbr acronym address area article aside audio b bdi bdo big blink blockquote body br button canvas caption center cite code col colgroup content data datalist dd decorator del details dfn dialog dir div dl dt element em fieldset figcaption figure font footer form h1 h2 h3 h4 h5 h6 head header hgroup hr html i img input ins kbd label legend li main map mark marquee menu menuitem meter nav nobr ol optgroup option output p picture pre progress q rp rt ruby s samp search section select shadow slot small source spacer span strike strong style sub summary sup table tbody td template textarea tfoot th thead time tr track tt u ul var video wbr #text
");

        private static readonly string[] DomPurifyHtmlProfileAttributes = ParseTokenList(@"
accept action align alt autocapitalize autocomplete autopictureinpicture autoplay background bgcolor border capture cellpadding cellspacing checked cite class clear color cols colspan controls controlslist coords crossorigin datetime decoding default dir disabled disablepictureinpicture disableremoteplayback download draggable enctype enterkeyhint exportparts face for headers height hidden high href hreflang id inert inputmode integrity ismap kind label lang list loading loop low max maxlength media method min minlength multiple muted name nonce noshade novalidate nowrap open optimum part pattern placeholder playsinline popover popovertarget popovertargetaction poster preload pubdate radiogroup readonly rel required rev reversed role rows rowspan spellcheck scope selected shape size sizes slot span srclang start src srcset step style summary tabindex title translate type usemap valign value width wrap xmlns xlink:href xml:id xlink:title xml:space xmlns:xlink
");

        private static readonly string[] DomPurifyAllowedSchemes = new[]
        {
            "ftp",
            "ftps",
            "http",
            "https",
            "mailto",
            "tel",
            "callto",
            "sms",
            "cid",
            "xmpp",
            "matrix"
        };

        private static readonly Lazy<HtmlSanitizer> Sanitizer = new Lazy<HtmlSanitizer>(CreateSanitizer, true);
        private static int _dependencyBootstrapDone;
        private static readonly Regex RgbaColorRegex = new Regex(
            @"rgba\(\s*(?<r>\d{1,3})\s*,\s*(?<g>\d{1,3})\s*,\s*(?<b>\d{1,3})\s*,\s*(?<a>0|0?\.\d+|1(?:\.0+)?)\s*\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex RgbColorRegex = new Regex(
            @"rgb\(\s*(?<r>\d{1,3})\s*,\s*(?<g>\d{1,3})\s*,\s*(?<b>\d{1,3})\s*\)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private sealed class HtmlStructureStats
        {
            internal int ElementCount;
            internal int AttributeCount;
            internal readonly Dictionary<string, int> TagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, int> AttributeCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        private struct NormalizationReport
        {
            internal int RgbaStyleConversions;
            internal int AnchorRelAdjustments;
        }

        static HtmlTemplateSanitizer()
        {
            EnsureSanitizerDependencies();
        }

        internal static string SanitizeShareTemplateHtml(string html)
        {
            return SanitizeTemplateHtml(html, "share");
        }

        internal static string SanitizeTalkTemplateHtml(string html)
        {
            return SanitizeTemplateHtml(html, "talk");
        }

        /**
         * Appointment-specific compatibility transform for Talk HTML before the
         * HTML->RTF bridge. Keeps this behavior explicit and isolated from other
         * template rendering paths.
         */
        internal static string PrepareTalkAppointmentHtmlForOutlookRtfBridge(string html)
        {
            return NormalizeForOutlookRtfBridge(html, "talk", true);
        }

        internal static string NormalizeForOutlookRtfBridge(string html, string templateType)
        {
            return NormalizeForOutlookRtfBridge(html, templateType, false);
        }

        private static string NormalizeForOutlookRtfBridge(string html, string templateType, bool appointmentCompatMode)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            try
            {
                var parser = new HtmlParser();
                var document = parser.ParseDocument(html);
                var body = document != null ? document.Body : null;
                if (body == null)
                {
                    return html;
                }

                bool changed = false;
                int fontColorWrappers = 0;
                int bgColorAttrs = 0;
                int alignAttrs = 0;
                int valignAttrs = 0;
                int strippedCssDeclarations = 0;
                int linkColorFontWrappers = 0;
                bool tableWrapped = false;
                int rgbaStyleConversions = 0;

                foreach (IElement element in body.QuerySelectorAll("[style]"))
                {
                    string style = element.GetAttribute("style");
                    if (string.IsNullOrWhiteSpace(style))
                    {
                        continue;
                    }

                    string color = NormalizeLegacyColorValue(ExtractCssProperty(style, "color"));
                    string backgroundColor = NormalizeLegacyColorValue(ExtractCssProperty(style, "background-color"));
                    string textAlign = NormalizeAlignValue(ExtractCssProperty(style, "text-align"));
                    string verticalAlign = NormalizeVerticalAlignValue(ExtractCssProperty(style, "vertical-align"));
                    int styleConversions;
                    string normalizedStyle = NormalizeOutlookStyleColors(style, out styleConversions);
                    rgbaStyleConversions += styleConversions;
                    if (appointmentCompatMode)
                    {
                        int strippedStyleDeclarations;
                        normalizedStyle = StripUnsupportedOutlookStyleDeclarations(normalizedStyle, out strippedStyleDeclarations);
                        strippedCssDeclarations += strippedStyleDeclarations;
                    }

                    if (!string.Equals(style, normalizedStyle, StringComparison.Ordinal))
                    {
                        if (string.IsNullOrWhiteSpace(normalizedStyle))
                        {
                            element.RemoveAttribute("style");
                        }
                        else
                        {
                            element.SetAttribute("style", normalizedStyle);
                        }
                        changed = true;
                    }

                    if (!string.IsNullOrWhiteSpace(color))
                    {
                        string tagName = (element.TagName ?? string.Empty).ToLowerInvariant();
                        if (string.Equals(tagName, "font", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.Equals(element.GetAttribute("color") ?? string.Empty, color, StringComparison.OrdinalIgnoreCase))
                            {
                                element.SetAttribute("color", color);
                                changed = true;
                                fontColorWrappers++;
                            }
                        }
                        else if (appointmentCompatMode && string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase))
                        {
                            if (TryWrapElementChildrenInFontColor(document, element, color))
                            {
                                changed = true;
                                fontColorWrappers++;
                                linkColorFontWrappers++;
                            }
                        }
                        else if (SupportsInlineFontColorFallback(tagName))
                        {
                            if (TryWrapElementChildrenInFontColor(document, element, color))
                            {
                                changed = true;
                                fontColorWrappers++;
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(backgroundColor) && SupportsBgColorAttribute((element.TagName ?? string.Empty).ToLowerInvariant()))
                    {
                        if (!string.Equals(element.GetAttribute("bgcolor") ?? string.Empty, backgroundColor, StringComparison.OrdinalIgnoreCase))
                        {
                            element.SetAttribute("bgcolor", backgroundColor);
                            changed = true;
                            bgColorAttrs++;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(textAlign) && SupportsAlignAttribute((element.TagName ?? string.Empty).ToLowerInvariant()))
                    {
                        if (!string.Equals(element.GetAttribute("align") ?? string.Empty, textAlign, StringComparison.OrdinalIgnoreCase))
                        {
                            element.SetAttribute("align", textAlign);
                            changed = true;
                            alignAttrs++;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(verticalAlign) && SupportsVAlignAttribute((element.TagName ?? string.Empty).ToLowerInvariant()))
                    {
                        if (!string.Equals(element.GetAttribute("valign") ?? string.Empty, verticalAlign, StringComparison.OrdinalIgnoreCase))
                        {
                            element.SetAttribute("valign", verticalAlign);
                            changed = true;
                            valignAttrs++;
                        }
                    }
                }

                if (appointmentCompatMode && body.QuerySelector("table") == null)
                {
                    var wrapperTable = document.CreateElement("table");
                    wrapperTable.SetAttribute("role", "presentation");
                    wrapperTable.SetAttribute("width", "100%");
                    wrapperTable.SetAttribute("cellpadding", "0");
                    wrapperTable.SetAttribute("cellspacing", "0");
                    wrapperTable.SetAttribute("border", "0");

                    var wrapperTbody = document.CreateElement("tbody");
                    var wrapperTr = document.CreateElement("tr");
                    var wrapperTd = document.CreateElement("td");
                    wrapperTd.SetAttribute("align", "left");
                    wrapperTd.SetAttribute("valign", "top");
                    alignAttrs++;
                    valignAttrs++;

                    while (body.FirstChild != null)
                    {
                        wrapperTd.AppendChild(body.FirstChild);
                    }

                    wrapperTr.AppendChild(wrapperTd);
                    wrapperTbody.AppendChild(wrapperTr);
                    wrapperTable.AppendChild(wrapperTbody);
                    body.AppendChild(wrapperTable);
                    changed = true;
                    tableWrapped = true;
                }

                string output = changed ? body.InnerHtml : html;
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Template outlook-rtf normalization ({0}): compatMode={1}, inputLen={2}, outputLen={3}, tableWrapped={4}, fontColorWrappers={5}, linkColorWrappers={6}, bgcolorAttrs={7}, alignAttrs={8}, valignAttrs={9}, strippedCssDecls={10}, rgbaConversions={11}",
                        templateType ?? "unknown",
                        appointmentCompatMode,
                        html.Length,
                        output.Length,
                        tableWrapped,
                        fontColorWrappers,
                        linkColorFontWrappers,
                        bgColorAttrs,
                        alignAttrs,
                        valignAttrs,
                        strippedCssDeclarations,
                        rgbaStyleConversions));

                return output;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Template outlook-rtf normalization failed (" + (templateType ?? "unknown") + ").",
                    ex);
                return html;
            }
        }

        private static void EnsureSanitizerDependencies()
        {
            if (Interlocked.Exchange(ref _dependencyBootstrapDone, 1) != 0)
            {
                return;
            }

            AppDomain.CurrentDomain.AssemblyResolve += ResolveSanitizerDependency;

            string assemblyDirectory = Path.GetDirectoryName(typeof(HtmlTemplateSanitizer).Assembly.Location) ?? string.Empty;
            TryLoadDependency(Path.Combine(assemblyDirectory, "AngleSharp.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "AngleSharp.Css.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "HtmlSanitizer.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "System.Buffers.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "System.Collections.Immutable.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "System.Memory.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "System.Runtime.CompilerServices.Unsafe.dll"));
            TryLoadDependency(Path.Combine(assemblyDirectory, "System.Text.Encoding.CodePages.dll"));
        }

        private static Assembly ResolveSanitizerDependency(object sender, ResolveEventArgs args)
        {
            if (string.IsNullOrWhiteSpace(args == null ? null : args.Name))
            {
                return null;
            }

            AssemblyName requestedAssembly;
            try
            {
                requestedAssembly = new AssemblyName(args.Name);
            }
            catch
            {
                return null;
            }

            string requestedName = requestedAssembly == null ? string.Empty : requestedAssembly.Name;
            if (!SanitizerDependencyNames.Contains(requestedName))
            {
                return null;
            }

            string assemblyDirectory = Path.GetDirectoryName(typeof(HtmlTemplateSanitizer).Assembly.Location) ?? string.Empty;
            string dependencyPath = Path.Combine(assemblyDirectory, requestedName + ".dll");
            if (!File.Exists(dependencyPath))
            {
                return null;
            }

            try
            {
                return Assembly.LoadFrom(dependencyPath);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve sanitizer dependency '" + requestedName + "'.", ex);
                return null;
            }
        }

        private static void TryLoadDependency(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return;
            }

            try
            {
                AssemblyName targetAssembly = AssemblyName.GetAssemblyName(path);
                bool alreadyLoaded = AppDomain.CurrentDomain
                    .GetAssemblies()
                    .Any(loaded => AssemblyName.ReferenceMatchesDefinition(loaded.GetName(), targetAssembly));
                if (alreadyLoaded)
                {
                    return;
                }

                Assembly.LoadFrom(path);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to preload sanitizer dependency '" + path + "'.", ex);
            }
        }

        private static HtmlSanitizer CreateSanitizer()
        {
            var sanitizer = new HtmlSanitizer();
            sanitizer.AllowDataAttributes = false;
            sanitizer.AllowedTags.Clear();
            sanitizer.AllowedAttributes.Clear();
            sanitizer.AllowedSchemes.Clear();

            foreach (string tag in DomPurifyHtmlProfileTags)
            {
                sanitizer.AllowedTags.Add(tag);
            }

            foreach (string attribute in DomPurifyHtmlProfileAttributes)
            {
                sanitizer.AllowedAttributes.Add(attribute);
            }

            foreach (string scheme in DomPurifyAllowedSchemes)
            {
                sanitizer.AllowedSchemes.Add(scheme);
            }

            foreach (string tag in ForbiddenTags)
            {
                sanitizer.AllowedTags.Remove(tag);
            }

            foreach (string tag in AdditionalTags)
            {
                sanitizer.AllowedTags.Add(tag);
            }

            foreach (string attribute in AdditionalAttributes)
            {
                sanitizer.AllowedAttributes.Add(attribute);
            }

            return sanitizer;
        }

        private static string[] ParseTokenList(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return Array.Empty<string>();
            }

            return Regex
                .Split(text, @"[\s,]+")
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static string SanitizeTemplateHtml(string html, string templateType)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            try
            {
                HtmlStructureStats inputStats = AnalyzeHtmlStructure(html);
                string sanitized = Sanitizer.Value.Sanitize(html);
                if (string.IsNullOrWhiteSpace(sanitized))
                {
                    LogSanitizationSummary(
                        templateType,
                        html.Length,
                        0,
                        0,
                        inputStats,
                        null,
                        new NormalizationReport(),
                        true);
                    return string.Empty;
                }

                NormalizationReport normalizationReport;
                string normalized = NormalizeSanitizedHtml(sanitized, out normalizationReport);
                HtmlStructureStats outputStats = AnalyzeHtmlStructure(normalized);
                LogSanitizationSummary(
                    templateType,
                    html.Length,
                    sanitized.Length,
                    normalized.Length,
                    inputStats,
                    outputStats,
                    normalizationReport,
                    false);
                return normalized;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Template sanitization failed (" + (templateType ?? "unknown") + ").",
                    ex);
                throw new InvalidOperationException(
                    "Template sanitization failed (" + (templateType ?? "unknown") + ").",
                    ex);
            }
        }

        private static void LogSanitizationSummary(
            string templateType,
            int inputLength,
            int sanitizedLength,
            int normalizedLength,
            HtmlStructureStats inputStats,
            HtmlStructureStats outputStats,
            NormalizationReport normalizationReport,
            bool emptied)
        {
            if (inputStats == null)
            {
                inputStats = new HtmlStructureStats();
            }

            if (outputStats == null)
            {
                outputStats = new HtmlStructureStats();
            }

            string removedTags = FormatRemovedEntries(inputStats.TagCounts, outputStats.TagCounts, 8);
            string removedAttributes = FormatRemovedEntries(inputStats.AttributeCounts, outputStats.AttributeCounts, 8);

            var builder = new StringBuilder();
            builder.Append("Template sanitization ");
            builder.Append(emptied ? "emptied" : "completed");
            builder.Append(" (");
            builder.Append(templateType ?? "unknown");
            builder.Append("): ");
            builder.Append("inputLen=").Append(inputLength);
            builder.Append(", sanitizedLen=").Append(sanitizedLength);
            builder.Append(", normalizedLen=").Append(normalizedLength);
            builder.Append(", inputElements=").Append(inputStats.ElementCount);
            builder.Append(", outputElements=").Append(outputStats.ElementCount);
            builder.Append(", inputAttrs=").Append(inputStats.AttributeCount);
            builder.Append(", outputAttrs=").Append(outputStats.AttributeCount);
            builder.Append(", removedTags=").Append(removedTags);
            builder.Append(", removedAttrs=").Append(removedAttributes);
            builder.Append(", rgbaConversions=").Append(normalizationReport.RgbaStyleConversions);
            builder.Append(", anchorRelAdjustments=").Append(normalizationReport.AnchorRelAdjustments);

            DiagnosticsLogger.Log(LogCategories.Core, builder.ToString());
        }

        private static string FormatRemovedEntries(
            Dictionary<string, int> inputCounts,
            Dictionary<string, int> outputCounts,
            int maxEntries)
        {
            if (inputCounts == null || inputCounts.Count == 0)
            {
                return "none";
            }

            var removed = new List<KeyValuePair<string, int>>();
            foreach (var input in inputCounts)
            {
                int outputValue;
                if (!outputCounts.TryGetValue(input.Key, out outputValue))
                {
                    outputValue = 0;
                }

                int delta = input.Value - outputValue;
                if (delta > 0)
                {
                    removed.Add(new KeyValuePair<string, int>(input.Key, delta));
                }
            }

            if (removed.Count == 0)
            {
                return "none";
            }

            return string.Join(
                ";",
                removed
                    .OrderByDescending(pair => pair.Value)
                    .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                    .Take(Math.Max(1, maxEntries))
                    .Select(pair => pair.Key + ":-" + pair.Value));
        }

        private static HtmlStructureStats AnalyzeHtmlStructure(string html)
        {
            var stats = new HtmlStructureStats();
            if (string.IsNullOrWhiteSpace(html))
            {
                return stats;
            }

            try
            {
                var parser = new HtmlParser();
                var document = parser.ParseDocument(html);
                var body = document != null ? document.Body : null;
                if (body == null)
                {
                    return stats;
                }

                foreach (IElement element in body.QuerySelectorAll("*"))
                {
                    stats.ElementCount++;

                    string tagName = (element.TagName ?? string.Empty).ToLowerInvariant();
                    if (!string.IsNullOrWhiteSpace(tagName))
                    {
                        int currentTagCount;
                        stats.TagCounts.TryGetValue(tagName, out currentTagCount);
                        stats.TagCounts[tagName] = currentTagCount + 1;
                    }

                    foreach (IAttr attribute in element.Attributes)
                    {
                        stats.AttributeCount++;
                        string attrName = (attribute.Name ?? string.Empty).ToLowerInvariant();
                        if (string.IsNullOrWhiteSpace(attrName))
                        {
                            continue;
                        }

                        int currentAttrCount;
                        stats.AttributeCounts.TryGetValue(attrName, out currentAttrCount);
                        stats.AttributeCounts[attrName] = currentAttrCount + 1;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to analyze HTML structure for sanitizer summary.", ex);
            }

            return stats;
        }

        private static string NormalizeSanitizedHtml(string html, out NormalizationReport report)
        {
            report = new NormalizationReport();
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            var parser = new HtmlParser();
            var document = parser.ParseDocument(html);
            var body = document != null ? document.Body : null;
            if (body == null)
            {
                return html;
            }

            bool changed = false;
            foreach (IElement element in body.QuerySelectorAll("[style]"))
            {
                string style = element.GetAttribute("style");
                if (string.IsNullOrWhiteSpace(style))
                {
                    continue;
                }

                int styleConversions;
                string normalizedStyle = NormalizeOutlookStyleColors(style, out styleConversions);
                if (!string.Equals(style, normalizedStyle, StringComparison.Ordinal))
                {
                    element.SetAttribute("style", normalizedStyle);
                    changed = true;
                }

                report.RgbaStyleConversions += styleConversions;
            }

            foreach (IElement anchor in body.QuerySelectorAll("a[target]"))
            {
                string target = anchor.GetAttribute("target");
                if (!string.Equals(target, "_blank", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var relTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                string rel = anchor.GetAttribute("rel") ?? string.Empty;
                string[] split = rel.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string token in split)
                {
                    relTokens.Add(token);
                }

                bool hadNoOpener = relTokens.Contains("noopener");
                bool hadNoReferrer = relTokens.Contains("noreferrer");
                if (!hadNoOpener || !hadNoReferrer)
                {
                    relTokens.Add("noopener");
                    relTokens.Add("noreferrer");
                    anchor.SetAttribute("rel", string.Join(" ", relTokens.ToArray()));
                    changed = true;
                    report.AnchorRelAdjustments++;
                }
            }

            return changed ? body.InnerHtml : html;
        }

        private static string NormalizeOutlookStyleColors(string style, out int rgbaConversions)
        {
            rgbaConversions = 0;
            if (string.IsNullOrWhiteSpace(style))
            {
                return string.Empty;
            }

            int conversionCount = 0;
            string normalized = RgbaColorRegex.Replace(
                style,
                delegate (Match match)
                {
                    int red = ParseCssChannel(match.Groups["r"].Value);
                    int green = ParseCssChannel(match.Groups["g"].Value);
                    int blue = ParseCssChannel(match.Groups["b"].Value);
                    double alpha = ParseCssAlpha(match.Groups["a"].Value);

                    if (alpha <= 0.001d)
                    {
                        conversionCount++;
                        return "transparent";
                    }

                    // Outlook Word renderer is unreliable with rgba(); keep opaque rgb().
                    conversionCount++;
                    return string.Format(
                        CultureInfo.InvariantCulture,
                        "rgb({0}, {1}, {2})",
                        red,
                        green,
                        blue);
                });

            rgbaConversions = conversionCount;
            return normalized;
        }

        private static string StripUnsupportedOutlookStyleDeclarations(string style, out int strippedCount)
        {
            strippedCount = 0;
            if (string.IsNullOrWhiteSpace(style))
            {
                return string.Empty;
            }

            string[] declarations = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            var kept = new List<string>(declarations.Length);
            foreach (string rawDeclaration in declarations)
            {
                string declaration = (rawDeclaration ?? string.Empty).Trim();
                if (declaration.Length == 0)
                {
                    continue;
                }

                int separatorIndex = declaration.IndexOf(':');
                if (separatorIndex <= 0 || separatorIndex >= declaration.Length - 1)
                {
                    kept.Add(declaration);
                    continue;
                }

                string property = declaration.Substring(0, separatorIndex).Trim().ToLowerInvariant();
                string value = declaration.Substring(separatorIndex + 1).Trim();
                if (ShouldStripOutlookStyleDeclaration(property, value))
                {
                    strippedCount++;
                    continue;
                }

                kept.Add(property + ": " + value);
            }

            return kept.Count == 0 ? string.Empty : string.Join("; ", kept) + ";";
        }

        private static bool ShouldStripOutlookStyleDeclaration(string property, string value)
        {
            if (string.IsNullOrWhiteSpace(property))
            {
                return false;
            }

            if (property.StartsWith("flex", StringComparison.OrdinalIgnoreCase)
                || property.StartsWith("grid", StringComparison.OrdinalIgnoreCase)
                || property.Contains("radius")
                || string.Equals(property, "object-fit", StringComparison.OrdinalIgnoreCase)
                || string.Equals(property, "overflow", StringComparison.OrdinalIgnoreCase)
                || property.StartsWith("overflow-", StringComparison.OrdinalIgnoreCase)
                || string.Equals(property, "user-select", StringComparison.OrdinalIgnoreCase)
                || string.Equals(property, "-webkit-user-select", StringComparison.OrdinalIgnoreCase)
                || string.Equals(property, "-moz-user-select", StringComparison.OrdinalIgnoreCase)
                || string.Equals(property, "-ms-user-select", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(property, "display", StringComparison.OrdinalIgnoreCase))
            {
                string normalizedValue = (value ?? string.Empty).Trim().ToLowerInvariant();
                if (normalizedValue.Contains("flex") || normalizedValue.Contains("grid"))
                {
                    return true;
                }
            }

            return false;
        }

        private static string ExtractCssProperty(string style, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            Match match = Regex.Match(
                style,
                @"(?:^|;)\s*" + Regex.Escape(propertyName) + @"\s*:\s*(?<value>[^;]+)",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return string.Empty;
            }

            string value = match.Groups["value"].Value ?? string.Empty;
            value = Regex.Replace(value, @"\s*!important\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return value;
        }

        private static bool SupportsInlineFontColorFallback(string tagName)
        {
            return string.Equals(tagName, "span", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "strong", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "b", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "em", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "i", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "p", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "div", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "td", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "th", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryWrapElementChildrenInFontColor(IDocument document, IElement targetElement, string color)
        {
            if (document == null || targetElement == null || string.IsNullOrWhiteSpace(color))
            {
                return false;
            }

            var currentSingleChild = targetElement.Children.Length == 1 ? targetElement.Children[0] : null;
            if (currentSingleChild != null
                && string.Equals(currentSingleChild.TagName, "FONT", StringComparison.OrdinalIgnoreCase)
                && string.Equals(currentSingleChild.GetAttribute("color") ?? string.Empty, color, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(targetElement.TagName, "A", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(targetElement.TextContent))
            {
                return false;
            }

            var fontElement = document.CreateElement("font");
            fontElement.SetAttribute("color", color);
            while (targetElement.FirstChild != null)
            {
                fontElement.AppendChild(targetElement.FirstChild);
            }

            targetElement.AppendChild(fontElement);
            return true;
        }

        private static bool SupportsBgColorAttribute(string tagName)
        {
            return string.Equals(tagName, "table", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "tbody", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "thead", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "tfoot", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "tr", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "td", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "th", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "body", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "div", StringComparison.OrdinalIgnoreCase);
        }

        private static bool SupportsAlignAttribute(string tagName)
        {
            return string.Equals(tagName, "p", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "div", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "table", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "tr", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "td", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "th", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "img", StringComparison.OrdinalIgnoreCase);
        }

        private static bool SupportsVAlignAttribute(string tagName)
        {
            return string.Equals(tagName, "tr", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "td", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "th", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "img", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeAlignValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value.Trim().ToLowerInvariant();
            if (normalized == "left" || normalized == "right" || normalized == "center" || normalized == "justify")
            {
                return normalized;
            }

            return string.Empty;
        }

        private static string NormalizeVerticalAlignValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value.Trim().ToLowerInvariant();
            if (normalized == "top" || normalized == "middle" || normalized == "bottom" || normalized == "baseline")
            {
                return normalized;
            }

            return string.Empty;
        }

        private static string NormalizeLegacyColorValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value.Trim();
            if (normalized.Length == 0)
            {
                return string.Empty;
            }

            if (normalized.StartsWith("#", StringComparison.Ordinal))
            {
                if (normalized.Length == 4)
                {
                    char r = normalized[1];
                    char g = normalized[2];
                    char b = normalized[3];
                    return "#" + r + r + g + g + b + b;
                }

                if (normalized.Length == 7)
                {
                    return normalized;
                }
            }

            Match rgbMatch = RgbColorRegex.Match(normalized);
            if (rgbMatch.Success)
            {
                int red = ParseCssChannel(rgbMatch.Groups["r"].Value);
                int green = ParseCssChannel(rgbMatch.Groups["g"].Value);
                int blue = ParseCssChannel(rgbMatch.Groups["b"].Value);
                return string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", red, green, blue);
            }

            // Allow named colors as fallback for legacy HTML attributes.
            if (Regex.IsMatch(normalized, @"^[a-zA-Z][a-zA-Z0-9_-]*$"))
            {
                return normalized;
            }

            return string.Empty;
        }

        private static int ParseCssChannel(string value)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
            {
                return 0;
            }

            if (parsed < 0)
            {
                return 0;
            }

            if (parsed > 255)
            {
                return 255;
            }

            return parsed;
        }

        private static double ParseCssAlpha(string value)
        {
            double parsed;
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                return 1d;
            }

            if (parsed < 0d)
            {
                return 0d;
            }

            if (parsed > 1d)
            {
                return 1d;
            }

            return parsed;
        }
    }
}
