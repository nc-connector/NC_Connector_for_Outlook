/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Controllers
{
    /**
     * Encapsulates rendering/removal of NC Talk description blocks in plain-text and HTML bodies.
     */
    internal static class TalkDescriptionTemplateController
    {
        private const string BodySectionHeader = "Nextcloud Talk";
        private const string TalkHelpUrlMarker = "join_a_call_or_chat_as_guest.html";
        private const string HtmlTalkBlockStartMarker = "<!-- NC4OL_TALK_BLOCK_START -->";
        private const string HtmlTalkBlockEndMarker = "<!-- NC4OL_TALK_BLOCK_END -->";

        internal static string UpdateBodyWithTalkBlock(string existingBody, string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            string body = RemoveExistingTalkBlock(existingBody ?? string.Empty);
            string block = BuildTalkBodyBlock(roomUrl, password, languageOverride, invitationTemplate);
            if (string.IsNullOrWhiteSpace(block))
            {
                return body.TrimEnd('\r', '\n');
            }
            if (string.IsNullOrWhiteSpace(body))
            {
                return block;
            }
            return body.TrimEnd('\r', '\n') + "\r\n\r\n" + block;
        }

        internal static string UpdateHtmlBodyWithTalkBlock(string existingHtmlBody, string existingBody, string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            string html = PrepareHtmlBody(existingHtmlBody, existingBody);
            html = RemoveExistingTalkBlockHtml(html);

            string block = BuildTalkBodyBlockHtml(roomUrl, password, languageOverride, invitationTemplate);
            if (string.IsNullOrWhiteSpace(block))
            {
                return string.IsNullOrWhiteSpace(html) ? string.Empty : html.Trim();
            }
            if (string.IsNullOrWhiteSpace(html))
            {
                return block;
            }
            return html.TrimEnd() + Environment.NewLine + block;
        }

        internal static string BuildInitialRoomDescription(string password, string languageOverride, string invitationTemplate)
        {
            string normalizedLanguage = NormalizeTalkDescriptionLanguage(languageOverride);
            if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(invitationTemplate))
            {
                return RenderTalkInvitationTemplate(invitationTemplate, string.Empty, password);
            }
            if (string.IsNullOrWhiteSpace(password))
            {
                return string.Empty;
            }
            if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase))
            {
                normalizedLanguage = "default";
            }
            string passwordLineFormat = Strings.GetInLanguage(normalizedLanguage, "ui_description_password_line", "Password: {0}");
            return string.Format(CultureInfo.InvariantCulture, passwordLineFormat, password.Trim());
        }

        /**
         * Outlook appointment bodies are plain text. Convert backend HTML/text
         * templates into a stable plain-text block before inserting them.
         */
        internal static string ConvertHtmlTemplateToPlainText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            string text = value.Replace("\r\n", "\n").Replace('\r', '\n');
            text = Regex.Replace(text, @"<\s*br\s*/?\s*>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<\s*/\s*(p|div|li|tr|table|h[1-6])\s*>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<\s*(p|div|li|tr|table|h[1-6])[^>]*>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<[^>]+>", string.Empty, RegexOptions.IgnoreCase);
            text = HttpUtility.HtmlDecode(text) ?? string.Empty;
            text = text.Replace('\u00A0', ' ');

            var lines = new List<string>();
            bool lastBlank = false;
            foreach (string rawLine in text.Split('\n'))
            {
                string line = (rawLine ?? string.Empty).Trim();
                if (line.Length == 0)
                {
                    if (!lastBlank && lines.Count > 0)
                    {
                        lines.Add(string.Empty);
                    }
                    lastBlank = true;
                    continue;
                }

                lines.Add(line);
                lastBlank = false;
            }
            return string.Join("\r\n", lines).Trim();
        }

        private static string RemoveExistingTalkBlock(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return body;
            }
            var lines = new List<string>();
            using (var reader = new StringReader(body))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }
            bool removed = false;
            int index = 0;

            while (index < lines.Count)
            {
                if (!IsTalkBlockHeaderLine(lines[index]))
                {
                    index++;
                    continue;
                }
                int blockEnd = FindTalkBlockEnd(lines, index);
                if (blockEnd < 0)
                {
                    index++;
                    continue;
                }
                int blockStart = index;
                while (blockStart > 0 && string.IsNullOrWhiteSpace(lines[blockStart - 1]))
                {
                    blockStart--;
                }
                int removeEnd = blockEnd;
                while (removeEnd + 1 < lines.Count && string.IsNullOrWhiteSpace(lines[removeEnd + 1]))
                {
                    removeEnd++;
                }

                lines.RemoveRange(blockStart, removeEnd - blockStart + 1);
                removed = true;
                index = blockStart;
            }
            if (!removed)
            {
                return body;
            }
            return string.Join("\r\n", lines).Trim('\r', '\n');
        }

        private static string RemoveExistingTalkBlockHtml(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }
            string updated = html;
            int startIndex = updated.IndexOf(HtmlTalkBlockStartMarker, StringComparison.OrdinalIgnoreCase);
            while (startIndex >= 0)
            {
                int endIndex = updated.IndexOf(HtmlTalkBlockEndMarker, startIndex, StringComparison.OrdinalIgnoreCase);
                if (endIndex < 0)
                {
                    break;
                }

                endIndex += HtmlTalkBlockEndMarker.Length;
                updated = updated.Remove(startIndex, endIndex - startIndex);
                startIndex = updated.IndexOf(HtmlTalkBlockStartMarker, StringComparison.OrdinalIgnoreCase);
            }
            return updated.Trim();
        }

        private static string PrepareHtmlBody(string existingHtmlBody, string existingBody)
        {
            string html = string.IsNullOrWhiteSpace(existingHtmlBody) ? string.Empty : existingHtmlBody;
            string body = string.IsNullOrWhiteSpace(existingBody) ? string.Empty : existingBody;
            if (!string.IsNullOrWhiteSpace(html) && html.IndexOf(HtmlTalkBlockStartMarker, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return html;
            }
            string cleanedBody = RemoveExistingTalkBlock(body);
            bool bodyHadLegacyBlock = !string.Equals(cleanedBody, body, StringComparison.Ordinal);
            if (bodyHadLegacyBlock || string.IsNullOrWhiteSpace(html))
            {
                return ConvertPlainTextToHtml(cleanedBody);
            }
            return html;
        }

        private static string BuildTalkBodyBlock(string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            string normalizedLanguage = NormalizeTalkDescriptionLanguage(languageOverride);
            if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(invitationTemplate))
            {
                return RenderTalkInvitationTemplate(invitationTemplate, roomUrl, password);
            }
            if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase))
            {
                normalizedLanguage = "default";
            }
            string joinLabel = Strings.GetInLanguage(normalizedLanguage, "ui_description_join_label", "Join the meeting now:");
            string passwordLineFormat = Strings.GetInLanguage(normalizedLanguage, "ui_description_password_line", "Password: {0}");
            string helpLabel = Strings.GetInLanguage(normalizedLanguage, "ui_description_help_label", "Need help?");
            string helpUrl = Strings.GetInLanguage(
                normalizedLanguage,
                "ui_description_help_url",
                "https://docs.nextcloud.com/server/latest/user_manual/en/talk/join_a_call_or_chat_as_guest.html");

            var lines = new List<string>
            {
                BodySectionHeader,
                string.Empty,
                joinLabel,
                roomUrl ?? string.Empty,
                string.Empty
            };

            if (!string.IsNullOrWhiteSpace(password))
            {
                lines.Add(string.Format(CultureInfo.InvariantCulture, passwordLineFormat, password.Trim()));
                lines.Add(string.Empty);
            }

            lines.Add(helpLabel);
            lines.Add(string.Empty);
            lines.Add(helpUrl);
            return string.Join("\r\n", lines).TrimEnd('\r', '\n');
        }

        private static string BuildTalkBodyBlockHtml(string roomUrl, string password, string languageOverride, string invitationTemplate)
        {
            string normalizedLanguage = NormalizeTalkDescriptionLanguage(languageOverride);
            string innerHtml;

            if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(invitationTemplate))
            {
                innerHtml = RenderTalkInvitationTemplateHtml(invitationTemplate, roomUrl, password);
            }
            else
            {
                if (string.Equals(normalizedLanguage, "custom", StringComparison.OrdinalIgnoreCase))
                {
                    normalizedLanguage = "default";
                }
                string joinLabel = Strings.GetInLanguage(normalizedLanguage, "ui_description_join_label", "Join the meeting now:");
                string passwordLineFormat = Strings.GetInLanguage(normalizedLanguage, "ui_description_password_line", "Password: {0}");
                string helpLabel = Strings.GetInLanguage(normalizedLanguage, "ui_description_help_label", "Need help?");
                string helpUrl = Strings.GetInLanguage(
                    normalizedLanguage,
                    "ui_description_help_url",
                    "https://docs.nextcloud.com/server/latest/user_manual/en/talk/join_a_call_or_chat_as_guest.html");

                var html = new StringBuilder();
                html.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"border-collapse:collapse;\">");
                html.Append("<tbody>");
                html.Append("<tr><td valign=\"top\" align=\"left\"><strong>")
                    .Append(HttpUtility.HtmlEncode(BodySectionHeader))
                    .Append("</strong></td></tr>");
                html.Append("<tr><td valign=\"top\" align=\"left\">")
                    .Append(HttpUtility.HtmlEncode(joinLabel))
                    .Append("<br><a href=\"")
                    .Append(HttpUtility.HtmlAttributeEncode(roomUrl ?? string.Empty))
                    .Append("\">")
                    .Append(HttpUtility.HtmlEncode(roomUrl ?? string.Empty))
                    .Append("</a></td></tr>");

                if (!string.IsNullOrWhiteSpace(password))
                {
                    html.Append("<tr><td valign=\"top\" align=\"left\">")
                        .Append(HttpUtility.HtmlEncode(string.Format(CultureInfo.InvariantCulture, passwordLineFormat, password.Trim())))
                        .Append("</td></tr>");
                }

                html.Append("<tr><td valign=\"top\" align=\"left\">")
                    .Append(HttpUtility.HtmlEncode(helpLabel))
                    .Append("<br><a href=\"")
                    .Append(HttpUtility.HtmlAttributeEncode(helpUrl))
                    .Append("\">")
                    .Append(HttpUtility.HtmlEncode(helpUrl))
                    .Append("</a></td></tr>");
                html.Append("</tbody>");
                html.Append("</table>");
                innerHtml = html.ToString();
            }
            if (string.IsNullOrWhiteSpace(innerHtml))
            {
                return string.Empty;
            }
            return HtmlTalkBlockStartMarker
                + "<table data-nc4ol-talk-block=\"true\" role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:16px;border-collapse:collapse;\">"
                + "<tbody><tr><td valign=\"top\" align=\"left\">"
                + innerHtml.Trim()
                + "</td></tr></tbody></table>"
                + HtmlTalkBlockEndMarker;
        }

        private static bool IsTalkBlockHeaderLine(string line)
        {
            string normalized = NormalizeTalkBlockLine(line);
            return string.Equals(normalized, BodySectionHeader, StringComparison.OrdinalIgnoreCase);
        }

        private static int FindTalkBlockEnd(List<string> lines, int headerIndex)
        {
            int maxIndex = Math.Min(lines.Count - 1, headerIndex + 40);
            for (int i = headerIndex; i <= maxIndex; i++)
            {
                string rawLine = lines[i] ?? string.Empty;
                string normalized = NormalizeTalkBlockLine(rawLine);
                if (rawLine.IndexOf(TalkHelpUrlMarker, StringComparison.OrdinalIgnoreCase) >= 0
                    || normalized.IndexOf(TalkHelpUrlMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return i;
                }
            }
            return -1;
        }

        private static string RenderTalkInvitationTemplate(string template, string meetingUrl, string password)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                return string.Empty;
            }
            string rendered = template
                .Replace("{MEETING_URL}", HttpUtility.HtmlEncode(meetingUrl ?? string.Empty))
                .Replace("{PASSWORD}", HttpUtility.HtmlEncode(password ?? string.Empty));
            string sanitized = HtmlTemplateSanitizer.SanitizeTalkTemplateHtml(rendered);
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                throw new InvalidOperationException("Talk invitation template sanitized to empty output.");
            }
            return ConvertHtmlTemplateToPlainText(sanitized);
        }

        private static string RenderTalkInvitationTemplateHtml(string template, string meetingUrl, string password)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                return string.Empty;
            }
            string rendered = template
                .Replace("{MEETING_URL}", HttpUtility.HtmlEncode(meetingUrl ?? string.Empty))
                .Replace("{PASSWORD}", HttpUtility.HtmlEncode(password ?? string.Empty));
            string sanitized = HtmlTemplateSanitizer.SanitizeTalkTemplateHtml(rendered);
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                throw new InvalidOperationException("Talk invitation HTML template sanitized to empty output.");
            }
            return sanitized.Trim();
        }

        private static string NormalizeTalkBlockLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return string.Empty;
            }
            string normalized = Regex.Replace(line, "<[^>]+>", string.Empty);
            normalized = HttpUtility.HtmlDecode(normalized) ?? string.Empty;
            normalized = normalized.Replace('\u00A0', ' ');
            return normalized.Trim();
        }

        private static string ConvertPlainTextToHtml(string value)
        {
            string normalized = (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }
            string[] paragraphs = Regex.Split(normalized, @"\n{2,}");
            var html = new StringBuilder();
            for (int i = 0; i < paragraphs.Length; i++)
            {
                string paragraph = (paragraphs[i] ?? string.Empty).Trim();
                if (paragraph.Length == 0)
                {
                    continue;
                }
                if (html.Length > 0)
                {
                    html.Append(Environment.NewLine);
                }

                html.Append("<p>")
                    .Append(HttpUtility.HtmlEncode(paragraph).Replace("\n", "<br>"))
                    .Append("</p>");
            }
            return html.ToString();
        }

        private static string NormalizeTalkDescriptionLanguage(string languageOverride)
        {
            if (string.IsNullOrWhiteSpace(languageOverride))
            {
                return "default";
            }
            string trimmed = languageOverride.Trim();
            if (string.Equals(trimmed, "custom", StringComparison.OrdinalIgnoreCase))
            {
                return "custom";
            }
            return Strings.NormalizeLanguageOverride(trimmed);
        }
    }
}
