/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using NcTalkOutlookAddIn.Settings;

namespace NcTalkOutlookAddIn.Utilities
{
    internal enum HttpFailureKind
    {
        Generic = 0,
        TlsHandshake = 1,
        CertificateValidation = 2,
        DnsResolution = 3,
        ProxyOrConnect = 4,
        Timeout = 5
    }

    internal sealed class HttpFailureInfo
    {
        internal HttpFailureInfo(HttpFailureKind kind, string summary, string guidance, string technicalDetail)
        {
            Kind = kind;
            Summary = summary ?? string.Empty;
            Guidance = guidance ?? string.Empty;
            TechnicalDetail = technicalDetail ?? string.Empty;
        }

        internal HttpFailureKind Kind { get; private set; }
        internal string Summary { get; private set; }
        internal string Guidance { get; private set; }
        internal string TechnicalDetail { get; private set; }

        internal string BuildUserMessage()
        {
            if (string.IsNullOrWhiteSpace(Guidance))
            {
                return Summary;
            }

            return Summary + Environment.NewLine + Environment.NewLine + Guidance;
        }
    }

    internal static class HttpFailureDiagnostics
    {
        internal static HttpFailureInfo Analyze(WebException ex)
        {
            if (ex == null)
            {
                return BuildInfo(HttpFailureKind.Generic, string.Empty);
            }

            string detail = FlattenMessages(ex);
            HttpFailureKind kind = DetectKind(ex, detail);
            return BuildInfo(kind, detail);
        }

        internal static string BuildLogSummary(WebException ex, HttpFailureInfo info)
        {
            if (info == null)
            {
                return "kind=unknown";
            }

            string status = ex != null ? ex.Status.ToString() : "Unknown";
            string securityProtocol = ServicePointManager.SecurityProtocol.ToString();
            string technical = info.TechnicalDetail ?? string.Empty;
            if (technical.Length > 400)
            {
                technical = technical.Substring(0, 400);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "kind={0}, webStatus={1}, securityProtocol={2}, detail={3}",
                info.Kind,
                status,
                securityProtocol,
                technical);
        }

        private static HttpFailureInfo BuildInfo(HttpFailureKind kind, string technicalDetail)
        {
            switch (kind)
            {
                case HttpFailureKind.CertificateValidation:
                    return new HttpFailureInfo(
                        kind,
                        Strings.ConnectionFailureCertificateSummary,
                        Strings.ConnectionFailureCertificateGuidance,
                        technicalDetail);
                case HttpFailureKind.DnsResolution:
                    return new HttpFailureInfo(
                        kind,
                        Strings.ConnectionFailureDnsSummary,
                        Strings.ConnectionFailureDnsGuidance,
                        technicalDetail);
                case HttpFailureKind.ProxyOrConnect:
                    return new HttpFailureInfo(
                        kind,
                        Strings.ConnectionFailureProxySummary,
                        Strings.ConnectionFailureProxyGuidance,
                        technicalDetail);
                case HttpFailureKind.Timeout:
                    return new HttpFailureInfo(
                        kind,
                        Strings.ConnectionFailureTimeoutSummary,
                        Strings.ConnectionFailureTimeoutGuidance,
                        technicalDetail);
                case HttpFailureKind.TlsHandshake:
                    return new HttpFailureInfo(
                        kind,
                        Strings.ConnectionFailureTlsSummary,
                        Strings.ConnectionFailureTlsGuidance,
                        technicalDetail);
                default:
                    return new HttpFailureInfo(
                        HttpFailureKind.Generic,
                        Strings.ConnectionFailureGenericSummary,
                        Strings.ConnectionFailureGenericGuidance,
                        technicalDetail);
            }
        }

        private static HttpFailureKind DetectKind(WebException ex, string technicalDetail)
        {
            string detail = (technicalDetail ?? string.Empty).ToLowerInvariant();

            if (ex.Status == WebExceptionStatus.TrustFailure ||
                ContainsAny(detail,
                    "remote certificate",
                    "certificate chain",
                    "certificate is invalid",
                    "trust relationship",
                    "name mismatch",
                    "untrustedroot"))
            {
                return HttpFailureKind.CertificateValidation;
            }

            if (ex.Status == WebExceptionStatus.NameResolutionFailure ||
                ex.Status == WebExceptionStatus.ProxyNameResolutionFailure ||
                ContainsAny(detail,
                    "no such host is known",
                    "name or service not known",
                    "name resolution",
                    "host not found"))
            {
                return HttpFailureKind.DnsResolution;
            }

            if (ex.Status == WebExceptionStatus.Timeout ||
                ContainsAny(detail, "timed out", "timeout"))
            {
                return HttpFailureKind.Timeout;
            }

            if (ex.Status == WebExceptionStatus.SecureChannelFailure ||
                ContainsAny(detail,
                    "secure channel",
                    "ssl/tls",
                    "tls handshake",
                    "authentication failed because"))
            {
                return HttpFailureKind.TlsHandshake;
            }

            if (ex.Status == WebExceptionStatus.ConnectFailure ||
                ex.Status == WebExceptionStatus.SendFailure ||
                ex.Status == WebExceptionStatus.ReceiveFailure ||
                ex.Status == WebExceptionStatus.RequestCanceled ||
                ContainsAny(detail, "proxy", "connection refused", "actively refused", "network path was not found"))
            {
                return HttpFailureKind.ProxyOrConnect;
            }

            SocketException socket = FindSocketException(ex);
            if (socket != null)
            {
                switch (socket.SocketErrorCode)
                {
                    case SocketError.HostNotFound:
                    case SocketError.TryAgain:
                    case SocketError.NoData:
                        return HttpFailureKind.DnsResolution;
                    case SocketError.TimedOut:
                        return HttpFailureKind.Timeout;
                    case SocketError.ConnectionRefused:
                    case SocketError.HostUnreachable:
                    case SocketError.NetworkUnreachable:
                        return HttpFailureKind.ProxyOrConnect;
                }
            }

            return HttpFailureKind.Generic;
        }

        private static SocketException FindSocketException(Exception ex)
        {
            while (ex != null)
            {
                SocketException socket = ex as SocketException;
                if (socket != null)
                {
                    return socket;
                }

                ex = ex.InnerException;
            }

            return null;
        }

        private static string FlattenMessages(Exception ex)
        {
            var parts = new List<string>();
            while (ex != null)
            {
                if (!string.IsNullOrWhiteSpace(ex.Message))
                {
                    parts.Add(ex.GetType().Name + ": " + ex.Message.Trim());
                }

                ex = ex.InnerException;
            }

            return string.Join(" | ", parts.ToArray());
        }

        private static bool ContainsAny(string source, params string[] fragments)
        {
            if (string.IsNullOrEmpty(source) || fragments == null)
            {
                return false;
            }

            foreach (string fragment in fragments)
            {
                if (!string.IsNullOrEmpty(fragment) &&
                    source.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }

    internal static class TransportSecurityConfigurator
    {
        private const SecurityProtocolType Tls13Protocol = (SecurityProtocolType)12288;

        internal static SecurityProtocolType ApplyFromSettings(AddinSettings settings, string source)
        {
            bool useSystemDefault = settings != null && settings.TransportTlsUseSystemDefault;
            bool enableTls12 = settings == null || settings.TransportTlsEnable12;
            bool enableTls13 = settings != null && settings.TransportTlsEnable13;
            return Apply(useSystemDefault, enableTls12, enableTls13, source);
        }

        internal static SecurityProtocolType Apply(bool useSystemDefault, bool enableTls12, bool enableTls13, string source)
        {
            SecurityProtocolType protocol = BuildProtocol(useSystemDefault, enableTls12, enableTls13);
            try
            {
                ServicePointManager.SecurityProtocol = protocol;
            }
            catch (Exception ex)
            {
                if ((protocol & Tls13Protocol) == Tls13Protocol)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "TLS 1.3 protocol flag was selected but rejected by this runtime. No auto-fallback is applied.",
                        ex);
                    throw new InvalidOperationException(
                        "TLS 1.3 was selected but is not supported by this runtime. NC Connector does not auto-fallback to TLS 1.2.",
                        ex);
                }

                throw;
            }
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Transport security applied (source="
                + (source ?? string.Empty)
                + ", useSystemDefault="
                + useSystemDefault.ToString(CultureInfo.InvariantCulture)
                + ", enableTls12="
                + enableTls12.ToString(CultureInfo.InvariantCulture)
                + ", enableTls13="
                + enableTls13.ToString(CultureInfo.InvariantCulture)
                + ", securityProtocol="
                + protocol
                + ").");
            return protocol;
        }

        internal static SecurityProtocolType BuildProtocol(bool useSystemDefault, bool enableTls12, bool enableTls13)
        {
            if (useSystemDefault)
            {
                return SecurityProtocolType.SystemDefault;
            }

            SecurityProtocolType protocol = 0;
            if (enableTls12)
            {
                protocol |= SecurityProtocolType.Tls12;
            }
            if (enableTls13)
            {
                protocol |= Tls13Protocol;
            }

            if (protocol == 0)
            {
                protocol = SecurityProtocolType.Tls12;
            }

            return protocol;
        }
    }
}
