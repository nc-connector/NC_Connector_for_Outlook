/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Builds the HTML block inserted into mail compose windows (download link, password, expiration, permissions).
     * The layout is intentionally kept static to remain stable when embedded in Outlook.
     */
    internal static class FileLinkHtmlBuilder
    {
        private static readonly Lazy<string> HeaderBase64 = new Lazy<string>(LoadHeaderBase64);

        /**
         * Creates the HTML block including branding and share information.
         */
        internal static string Build(FileLinkResult result, FileLinkRequest request, string languageOverride)
        {
            if (result == null)
            {
                throw new ArgumentNullException("result");
            }

            string intro = Strings.GetInLanguage(
                languageOverride,
                "sharing_html_intro",
                "I would like to share files securely and protect your privacy. Click the link below to download your files.");

            string footerFormat = Strings.GetInLanguage(
                languageOverride,
                "sharing_html_footer",
                "{0} is a solution for secure email and file exchange.");

            string downloadLabel = Strings.GetInLanguage(languageOverride, "sharing_html_download_label", "Download link");
            string passwordLabel = Strings.GetInLanguage(languageOverride, "sharing_html_password_label", "Password");
            string expireLabel = Strings.GetInLanguage(languageOverride, "sharing_html_expire_label", "Expiration date");
            string permissionsLabel = Strings.GetInLanguage(languageOverride, "sharing_html_permissions_label", "Your permissions");

            string permissionRead = Strings.GetInLanguage(languageOverride, "sharing_permission_read", "Read");
            string permissionCreate = Strings.GetInLanguage(languageOverride, "sharing_permission_create", "Upload");
            string permissionWrite = Strings.GetInLanguage(languageOverride, "sharing_permission_write", "Modify");
            string permissionDelete = Strings.GetInLanguage(languageOverride, "sharing_permission_delete", "Delete");

            string brandBlue = BrandingAssets.BrandBlueHex;

            var builder = new StringBuilder();
            builder.AppendLine("<div style=\"font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:11pt;margin:16px 0;\">");
            builder.AppendLine("<table role=\"presentation\" width=\"640\" style=\"border-collapse:separate;border-spacing:0;width:640px;margin:0;background-color:transparent;border:1px solid #d7d7db;border-radius:8px;overflow:hidden;\">");
            builder.AppendLine("<tr>");
            builder.AppendLine("<td style=\"padding:0;\">");
            builder.AppendLine("<table role=\"presentation\" width=\"640\" style=\"border-collapse:collapse;width:640px;margin:0;background-color:transparent;\">");
            builder.AppendLine("<tr>");
            builder.AppendFormat(CultureInfo.InvariantCulture, "<td height=\"32\" bgcolor=\"{0}\" style=\"padding:0;background-color:{0};text-align:center;height:32px;line-height:0;font-size:0;mso-line-height-rule:exactly;\">", brandBlue);
            builder.AppendLine();
            builder.AppendLine("<a href=\"https://github.com/nc-connector/NC_Connector_for_Outlook\" style=\"display:block;text-decoration:none;line-height:0;font-size:0;\" target=\"_blank\" rel=\"noopener\">");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<img alt=\"NC Connector\" height=\"32\" style=\"display:block;width:auto;height:32px;max-width:164px;object-fit:contain;border:0;margin:0 auto;\" src=\"data:image/png;base64,{0}\" />",
                HeaderBase64.Value);
            builder.AppendLine("</a>");
            builder.AppendLine("</td>");
            builder.AppendLine("</tr>");
            builder.AppendLine("</table>");
            builder.AppendLine("<div style=\"padding:18px 18px 12px 18px;\">");
            if (request != null && request.NoteEnabled && !string.IsNullOrWhiteSpace(request.Note))
            {
                builder.AppendFormat(
                    CultureInfo.InvariantCulture,
                    "<p style=\"margin:0 0 14px 0;line-height:1.4;\">{0}</p>",
                    HttpUtility.HtmlEncode(request.Note));
            }
            builder.AppendLine("<p style=\"margin:0 0 14px 0;line-height:1.4;\">" + HttpUtility.HtmlEncode(intro) + "<br /></p>");
            builder.AppendLine("<table style=\"width:100%;border-collapse:collapse;margin-bottom:10px;\">");

            AppendRow(builder, downloadLabel, string.Format(
                CultureInfo.InvariantCulture,
                "<a href=\"{0}\" style=\"color:{1};text-decoration:none;\">{0}</a>",
                HttpUtility.HtmlEncode(result.ShareUrl),
                brandBlue));

            if (!string.IsNullOrEmpty(result.Password))
            {
                var passwordBuilder = new StringBuilder();
                passwordBuilder.Append("<span style=\"display:inline-block;font-family:'Consolas','Courier New',monospace;padding:2px 6px;border:1px solid #c7c7c7;border-radius:3px;-ms-user-select:all;user-select:all;\" ondblclick=\"try{window.getSelection().selectAllChildren(this);}catch(e){}\" onclick=\"try{window.getSelection().selectAllChildren(this);}catch(e){}\">");
                passwordBuilder.Append(HttpUtility.HtmlEncode(result.Password));
                passwordBuilder.Append("</span>");
                string passwordHtml = passwordBuilder.ToString();
                AppendRow(builder, passwordLabel, passwordHtml);
            }

            if (result.ExpireDate.HasValue)
            {
                AppendRow(builder, expireLabel, HttpUtility.HtmlEncode(result.ExpireDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
            }

            AppendRow(builder, permissionsLabel, BuildPermissions(result.Permissions, permissionRead, permissionCreate, permissionWrite, permissionDelete));

            builder.AppendLine("</table>");
            builder.AppendLine("</div>");
            builder.AppendLine("<div style=\"padding:10px 18px 16px 18px;font-size:9pt;font-style:italic;\">");
            string nextcloudLink = string.Format(CultureInfo.InvariantCulture, "<a href=\"https://nextcloud.com/\" style=\"color:{0};text-decoration:none;\">Nextcloud</a>", brandBlue);
            builder.AppendLine(string.Format(CultureInfo.InvariantCulture, footerFormat, nextcloudLink));
            builder.AppendLine("</div>");
            builder.AppendLine("</td>");
            builder.AppendLine("</tr>");
            builder.AppendLine("</table>");
            builder.AppendLine("</div>");
            return builder.ToString();
        }

        /**
         * Adds a table row with label and content.
         */
        private static void AppendRow(StringBuilder builder, string label, string valueHtml)
        {
            builder.AppendLine("<tr>");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<th style=\"text-align:left;width:12ch;vertical-align:top;padding:6px 10px 6px 0;\">{0}</th>",
                HttpUtility.HtmlEncode(label));
            builder.Append("<td style=\"padding:6px 0;max-width:50ch;word-break:break-word;\">");
            builder.Append(valueHtml ?? string.Empty);
            builder.Append("</td>");
            builder.AppendLine("</tr>");
        }

        /**
         * Renders the permissions badges (checkmarks / red crosses).
         */
        private static string BuildPermissions(FileLinkPermissionFlags permissions, string readLabel, string createLabel, string writeLabel, string deleteLabel)
        {
            var builder = new StringBuilder();
            builder.Append("<table style=\"border-collapse:collapse;\">");
            builder.Append("<tr>");
            AppendPermissionCell(builder, readLabel, (permissions & FileLinkPermissionFlags.Read) == FileLinkPermissionFlags.Read);
            AppendPermissionCell(builder, createLabel, (permissions & FileLinkPermissionFlags.Create) == FileLinkPermissionFlags.Create);
            AppendPermissionCell(builder, writeLabel, (permissions & FileLinkPermissionFlags.Write) == FileLinkPermissionFlags.Write);
            AppendPermissionCell(builder, deleteLabel, (permissions & FileLinkPermissionFlags.Delete) == FileLinkPermissionFlags.Delete);
            builder.Append("</tr>");
            builder.Append("</table>");
            return builder.ToString();
        }

        /**
         * Builds the cell for a single permission.
         */
        private static void AppendPermissionCell(StringBuilder builder, string label, bool enabled)
        {
            builder.Append("<td style=\"padding:0 18px 6px 0;\">");
            builder.Append("<span style=\"display:inline-flex;align-items:center;\">");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<span style=\"display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border:1px solid {0};color:{0};font-size:13px;font-weight:700;\">{1}</span>",
                enabled ? BrandingAssets.BrandBlueHex : "#c62828",
                enabled ? "&#10003;" : "&#10007;");
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<span style=\"padding-left:6px;font-weight:600;\">{0}</span>",
                HttpUtility.HtmlEncode(label));
            builder.Append("</span>");
            builder.Append("</td>");
        }

        /**
         * Loads the embedded header banner as a Base64 string.
         */
        private static string LoadHeaderBase64()
        {
            const string resource = "NcTalkOutlookAddIn.Resources.header-solid-blue-164x48.png";
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (stream == null)
                {
                    return string.Empty;
                }

                using (var reader = new BinaryReader(stream))
                {
                    byte[] data = reader.ReadBytes((int)stream.Length);
                    return Convert.ToBase64String(data);
                }
            }
        }
    }
}
