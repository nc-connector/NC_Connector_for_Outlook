// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class HtmlToPlainTextConverter
    {
        private static readonly Regex ExcessBlankLinesRegex = new Regex(@"\n{3,}", RegexOptions.Compiled);
        private static readonly Regex InlineWhitespaceRegex = new Regex(@"[ \t\f\v]+", RegexOptions.Compiled);

        internal static string Convert(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            try
            {
                var parser = new HtmlParser();
                var document = parser.ParseDocument(html);
                IElement body = document != null ? document.Body : null;
                if (body == null)
                {
                    return NormalizePlainText(html);
                }

                var builder = new StringBuilder();
                var context = new RenderContext();
                AppendChildren(body, builder, context);
                return NormalizePlainText(builder.ToString());
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to convert HTML template to plain text.", ex);
                return string.Empty;
            }
        }

        private static void AppendChildren(INode node, StringBuilder builder, RenderContext context)
        {
            if (node == null)
            {
                return;
            }

            foreach (INode child in node.ChildNodes)
            {
                AppendNode(child, builder, context);
            }
        }

        private static void AppendNode(INode node, StringBuilder builder, RenderContext context)
        {
            if (node == null)
            {
                return;
            }

            var textNode = node as IText;
            if (textNode != null)
            {
                AppendText(builder, textNode.Data);
                return;
            }

            var element = node as IElement;
            if (element == null)
            {
                return;
            }

            string tag = (element.TagName ?? string.Empty).ToLowerInvariant();
            if (tag == "script" || tag == "style" || tag == "template")
            {
                return;
            }

            if (tag == "br")
            {
                AppendLineBreak(builder);
                return;
            }

            if (tag == "img")
            {
                AppendText(builder, element.GetAttribute("alt"));
                return;
            }

            if (tag == "a")
            {
                AppendAnchor(element, builder, context);
                return;
            }

            if (tag == "ul" || tag == "ol")
            {
                AppendList(element, builder, context, tag == "ol");
                return;
            }

            if (tag == "li")
            {
                AppendListItem(element, builder, context);
                return;
            }

            if (tag == "table")
            {
                AppendBlockStart(builder);
                AppendChildren(element, builder, context);
                AppendBlockEnd(builder);
                return;
            }

            if (tag == "tr")
            {
                AppendBlockStart(builder);
                AppendChildren(element, builder, context);
                AppendLineBreak(builder);
                return;
            }

            if (tag == "td" || tag == "th")
            {
                AppendBlockStart(builder);
                AppendChildren(element, builder, context);
                AppendLineBreak(builder);
                return;
            }

            if (IsBlockElement(tag))
            {
                AppendBlockStart(builder);
                AppendChildren(element, builder, context);
                AppendBlockEnd(builder);
                return;
            }

            AppendChildren(element, builder, context);
        }

        private static void AppendAnchor(IElement element, StringBuilder builder, RenderContext context)
        {
            var labelBuilder = new StringBuilder();
            AppendChildren(element, labelBuilder, context);
            string label = NormalizeInlineText(labelBuilder.ToString());
            string href = (element.GetAttribute("href") ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(label))
            {
                AppendText(builder, href);
                return;
            }

            if (string.IsNullOrWhiteSpace(href)
                || string.Equals(label, href, StringComparison.OrdinalIgnoreCase))
            {
                AppendText(builder, label);
                return;
            }

            AppendText(builder, label + " (" + href + ")");
        }

        private static void AppendList(IElement element, StringBuilder builder, RenderContext context, bool ordered)
        {
            AppendBlockStart(builder);
            context.ListStack.Push(new ListState { Ordered = ordered, Index = 1 });
            AppendChildren(element, builder, context);
            context.ListStack.Pop();
            AppendBlockEnd(builder);
        }

        private static void AppendListItem(IElement element, StringBuilder builder, RenderContext context)
        {
            AppendLineStart(builder);
            ListState state = context.ListStack.Count > 0 ? context.ListStack.Peek() : null;
            if (state != null && state.Ordered)
            {
                AppendText(builder, state.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ". ");
                state.Index++;
            }
            else
            {
                AppendText(builder, "- ");
            }

            AppendChildren(element, builder, context);
            AppendLineBreak(builder);
        }

        private static bool IsBlockElement(string tag)
        {
            switch (tag)
            {
                case "address":
                case "article":
                case "aside":
                case "blockquote":
                case "caption":
                case "center":
                case "dd":
                case "details":
                case "div":
                case "dl":
                case "dt":
                case "figcaption":
                case "figure":
                case "footer":
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                case "header":
                case "hr":
                case "main":
                case "nav":
                case "p":
                case "pre":
                case "section":
                    return true;
                default:
                    return false;
            }
        }

        private static void AppendText(StringBuilder builder, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            string text = value
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\u00A0', ' ');
            text = InlineWhitespaceRegex.Replace(text, " ");
            if (text.Length == 0)
            {
                return;
            }

            if (builder.Length > 0
                && !char.IsWhiteSpace(builder[builder.Length - 1])
                && !char.IsWhiteSpace(text[0]))
            {
                builder.Append(' ');
            }
            builder.Append(text);
        }

        private static void AppendBlockStart(StringBuilder builder)
        {
            if (builder.Length == 0)
            {
                return;
            }
            AppendLineBreak(builder);
        }

        private static void AppendBlockEnd(StringBuilder builder)
        {
            AppendLineBreak(builder);
            AppendLineBreak(builder);
        }

        private static void AppendLineStart(StringBuilder builder)
        {
            if (builder.Length == 0)
            {
                return;
            }
            if (builder[builder.Length - 1] != '\n')
            {
                AppendLineBreak(builder);
            }
        }

        private static void AppendLineBreak(StringBuilder builder)
        {
            if (builder.Length == 0 || builder[builder.Length - 1] == '\n')
            {
                return;
            }
            builder.Append('\n');
        }

        private static string NormalizePlainText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\u00A0', ' ');
            string[] lines = normalized.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = NormalizeInlineText(lines[i]).TrimEnd();
            }

            normalized = string.Join("\n", lines).Trim('\n', ' ', '\t');
            normalized = ExcessBlankLinesRegex.Replace(normalized, "\n\n");
            return normalized.Replace("\n", "\r\n").Trim();
        }

        private static string NormalizeInlineText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            string normalized = value
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\u00A0', ' ');
            return InlineWhitespaceRegex.Replace(normalized, " ").Trim();
        }

        private sealed class RenderContext
        {
            internal Stack<ListState> ListStack = new Stack<ListState>();
        }

        private sealed class ListState
        {
            internal bool Ordered;
            internal int Index;
        }
    }

    internal static class EmailSignatureContentBuilder
    {
        internal static string BuildManagedHtml(string sanitizedHtml)
        {
            return "<div data-nc-connector-signature=\"true\">"
                   + (sanitizedHtml ?? string.Empty)
                   + "</div>";
        }

        internal static string BuildPlainText(string sanitizedHtml)
        {
            string plainText = HtmlToPlainTextConverter.Convert(sanitizedHtml);
            if (string.IsNullOrWhiteSpace(plainText))
            {
                return string.Empty;
            }

            string normalized = plainText
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Trim('\r', '\n');
            return "-- \r\n" + normalized.Replace("\n", "\r\n") + "\r\n";
        }
    }
}
