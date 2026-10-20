param([string]$ProjectRoot = '.', [switch]$SkipBuild, [string]$PreviewDirectory = '')
$ErrorActionPreference = 'Stop'
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
if (-not $SkipBuild) {
    $refs = & (Join-Path $ProjectRoot 'tools/ci/Resolve-OfficeExtensibilityReference.ps1') -OutputDirectory (Join-Path $ProjectRoot '.ci/refs')
    & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" (Join-Path $ProjectRoot 'NcTalkOutlookAddIn.sln') /nologo /verbosity:minimal /t:Build /p:Configuration=Release "/p:ReferencePath=$refs"
    if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
}
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('nc4ol-ribbon-tests-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
try {
    $test = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Serialization;
using System.Windows.Forms;
using System.Xml;

internal static class RibbonAccessTests
{
    private static Assembly assembly;
    private static int assertions;
    private static string previewDirectory;
    private const BindingFlags All = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    private static Type T(string name) { return assembly.GetType("NcTalkOutlookAddIn." + name, true); }
    private static object New(string name) { return Activator.CreateInstance(T(name), true); }
    private static object Get(object item, string name)
    {
        PropertyInfo property = item.GetType().GetProperty(name, All);
        return property != null ? property.GetValue(item, null) : item.GetType().GetField(name, All).GetValue(item);
    }
    private static void Set(object item, string name, object value)
    {
        PropertyInfo property = item.GetType().GetProperty(name, All);
        if (property != null) { property.SetValue(item, value, null); }
        else { item.GetType().GetField(name, All).SetValue(item, value); }
    }
    private static object Call(object item, string name, params object[] args)
    {
        Type type = item as Type ?? item.GetType();
        return type.GetMethod(name, All).Invoke(item is Type ? null : item, args);
    }
    private static string Text(string name) { return (string)T("Utilities.Strings").GetProperty(name, All).GetValue(null, null); }
    private static void Check(bool condition, string message) { assertions++; if (!condition) { throw new Exception(message); } }
    private static object Settings(bool managed, bool complete)
    {
        object settings = New("Settings.AddinSettings");
        Set(settings, "IsEnterpriseRollout", managed);
        if (complete)
        {
            Set(settings, "ServerUrl", "https://cloud.example.test");
            Set(settings, "Username", "test-user");
            Set(settings, "AppPassword", "test-only");
        }
        return settings;
    }
    private static object Status(string mode, bool valid, bool assigned, string seat, string access = "ACTIVE")
    {
        object status = FormatterServices.GetUninitializedObject(T("Models.BackendPolicyStatus"));
        Set(status, "EndpointAvailable", true); Set(status, "FetchSucceeded", true);
        Set(status, "Mode", mode); Set(status, "IsValid", valid); Set(status, "SeatAssigned", assigned);
        Set(status, "SeatState", seat); Set(status, "AccessStatus", access); Set(status, "LicenseStatus", access);
        Set(status, "Reason", ""); Set(status, "LicenseActivationState", "");
        return status;
    }
    private static object Evaluate(object settings, object status, bool requiresSeat, string pause = "")
    {
        return Call(T("Controllers.RibbonActionController"), "EvaluateAccess", settings, status, requiresSeat, pause);
    }
    private static bool CanStart(object state) { return (bool)Get(state, "CanStart"); }
    private static string Notice(object state) { return (string)Get(state, "Notice"); }
    private static string ActionNotice(object settings, object status, bool requiresSeat)
    {
        return (string)Call(T("Utilities.PolicyUiHelper"), "GetActionUnavailableNotice", settings, status, requiresSeat);
    }
    private static void TestAccessMatrix()
    {
        foreach (bool managed in new[] { false, true })
        foreach (bool requiresSeat in new[] { false, true })
        foreach (bool valid in new[] { false, true })
        foreach (bool assigned in new[] { false, true })
        foreach (string seat in new[] { "none", "active", "suspended_overlimit", "unknown" })
        foreach (bool administrator in new[] { false, true })
        {
            object settings = Settings(managed, true);
            object community = Status("community", valid, assigned, seat);
            object pro = Status("pro", valid, assigned, seat);
            Set(community, "CanManageLicense", administrator); Set(pro, "CanManageLicense", administrator);
            object state = Evaluate(settings, community, requiresSeat);
            object paired = Evaluate(settings, pro, requiresSeat);
            bool expected = !(managed || requiresSeat) || (valid && assigned && seat == "active");
            Check(CanStart(state), "Seat refusal keeps its explanation reachable without granting execution");
            Check(CanStart(state) == CanStart(paired) && Notice(state) == Notice(paired), "Community and Pro produce identical access and notices");
            Check(Notice(state) == ActionNotice(settings, community, requiresSeat), "Ribbon and execution gate use the same cause text");
            Check((Notice(state).Length == 0) == expected, "Execution permission still requires personal entitlement");
            if ((managed || requiresSeat) && !assigned)
            {
                Check(Notice(state) == Text(managed ? "EnterpriseRolloutSeatRequired" : "SharingPasswordSeparateNoSeatTooltip"), "No-seat notice explains the actual managed or feature requirement");
            }
            if (managed)
            {
                object resource = Evaluate(settings, community, true);
                Check(CanStart(state) == CanStart(resource) && Notice(state) == Notice(resource), "Managed Share, Talk and resources share the same gate");
            }
        }
        foreach (string reason in new[] { "EXPIRED", "INACTIVE", "INVALID", "OFFLINE_EXPIRED", "ACTIVATION_REQUIRED", "UNKNOWN" })
        {
            object settings = Settings(true, true);
            object status = Status("community", false, true, "active", reason);
            object state = Evaluate(settings, status, false);
            Check(CanStart(state) && Notice(state) == (string)Call(T("Utilities.PolicyUiHelper"), "GetSeatFeatureUnavailableNotice", status), "License refusal remains invocable and retains its exact shared explanation: " + reason);
            Check(Notice(state) != Text("SharingPasswordSeparateNoSeatTooltip"), "A license refusal is not relabelled missing Seat");
        }
        object grace = Status("community", true, true, "active", "GRACE");
        Check(CanStart(Evaluate(Settings(true, true), grace, false)), "Backend-confirmed grace preserves an active Seat");
        object suspended = Status("pro", true, true, "suspended_overlimit");
        Check(Notice(Evaluate(Settings(true, true), suspended, false)).StartsWith(Text("PolicyWarningSeatSuspended"), StringComparison.Ordinal), "Personal capacity pause retains its dedicated notice");
        foreach (bool managed in new[] { false, true })
        foreach (bool requiresSeat in new[] { false, true })
        {
            object settings = Settings(managed, true);
            object missing = Status("local", false, false, "none");
            Set(missing, "EndpointAvailable", false); Set(missing, "FetchSucceeded", false); Set(missing, "Reason", "backend_missing");
            object state = Evaluate(settings, missing, requiresSeat);
            Check(CanStart(state), "Confirmed missing backend keeps the requirement explanation reachable");
            Check(Notice(state) == (managed ? Text("EnterpriseRolloutBackendRequired") : requiresSeat ? Text("SharingPasswordSeparateBackendRequiredTooltip") : ""), "Missing backend has a precise requirement notice");
            foreach (object pending in new object[] { null, Status("local", false, false, "none") })
            {
                if (pending != null) { Set(pending, "FetchSucceeded", false); Set(pending, "Reason", "backend_unavailable"); }
                state = Evaluate(settings, pending, requiresSeat);
                Check(CanStart(state), "Unknown or failed check keeps explicit verification reachable");
                Check(Notice(state) == (managed ? Text("EnterpriseRolloutStatusUnavailable") : requiresSeat ? Text("PolicyWarningBackendUnavailable") : ""), "Unknown state does not claim the backend is absent");
                if (managed || requiresSeat) { Check(ActionNotice(settings, pending, requiresSeat).Length != 0, "Reachable entry is not execution permission for unknown status"); }
            }
            object denial = Status("pro", false, false, "none");
            state = Evaluate(Settings(managed, false), denial, requiresSeat);
            Check(CanStart(state) && Notice(state) == Text("ErrorMissingCredentials"), "Initial authentication stays reachable despite a previous refusal");
            state = Evaluate(settings, denial, requiresSeat, "auth_required");
            Check(CanStart(state) && Notice(state) == Text("ConnectionAuthRequired"), "Rejected credentials offer the same recovery entry for every action");
            state = Evaluate(Settings(managed, false), denial, requiresSeat, "rate_limited");
            Check(!CanStart(state) && Notice(state) == Text("ConnectionRateLimited"), "Rate limit prevents missing-credential login bypass");
            state = Evaluate(settings, grace, requiresSeat, "rate_limited");
            Check(!CanStart(state) && Notice(state) == Text("ConnectionRateLimited"), "Rate limit overrides confirmed entitlement");
        }
        object tlsSettings = Settings(true, false);
        object policy = FormatterServices.GetUninitializedObject(T("Settings.ManagedSetupPolicy"));
        Set(policy, "HasTransportTlsEnable12Value", true);
        Set(tlsSettings, "_managedSetupPolicy", policy);
        foreach (bool requiresSeat in new[] { false, true })
        {
            object state = Evaluate(tlsSettings, grace, requiresSeat, "auth_required");
            Check(!CanStart(state) && Notice(state) == Text("ManagedTlsPolicyInvalid"), "Invalid managed TLS blocks before authentication entry");
            Check(ActionNotice(tlsSettings, grace, requiresSeat) == Notice(state), "TLS tooltip and execution failure share a message");
        }
        Check(!CanStart(Evaluate(null, null, false)), "Uninitialized settings do not start an action");
    }

    private static void TestActionNoticeDialog()
    {
        object offline = Status("pro", false, true, "active", "OFFLINE_EXPIRED");
        Set(offline, "LicenseConnectionError", true);
        Set(offline, "LicenseLastSyncAtIso", "2026-10-01T10:00:00Z");
        Set(offline, "LicenseOfflineUntilIso", "2026-10-15T10:00:00Z");
        foreach (string locale in new[] { "de", "en", "fr", "cs", "es", "hu", "it", "ja", "nl", "pl", "pt_BR", "pt_PT", "ru", "zh_CN", "zh_TW" })
        {
            Call(T("Utilities.Strings"), "SetPreferredUiLanguage", locale);
            foreach (string notice in new[] { Text("EnterpriseRolloutSeatRequired"), Text("EnterpriseRolloutBackendRequired"), Text("EnterpriseRolloutStatusUnavailable"), ActionNotice(Settings(true, true), offline, false) })
            {
                using (Form dialog = (Form)Call(T("Utilities.PolicyUiHelper"), "CreateActionUnavailableDialog", notice))
                {
                    dialog.StartPosition = FormStartPosition.Manual;
                    dialog.Location = new Point(-20000, -20000);
                    dialog.Opacity = 0;
                    dialog.Show();
                    Application.DoEvents();
                    dialog.PerformLayout();
                    var content = (TableLayoutPanel)dialog.Controls[0];
                    content.PerformLayout();
                    var label = (Label)content.GetControlFromPosition(0, 0);
                    var ok = (Button)content.GetControlFromPosition(0, 1);
                    Color expected = SystemInformation.HighContrast ? SystemColors.WindowText
                        : (Color)Get(Call(T("Utilities.UiThemeManager"), "DetectPalette"), "ErrorText");
                    Check(notice.Length > 0 && label.Text == notice && !label.UseMnemonic, locale + " action dialog preserves the complete existing localized notice");
                    Check(label.ForeColor == expected, locale + " blocked text uses theme red or accessible high-contrast colors");
                    Check(label.AutoSize && label.MaximumSize.Height == 0 && label.MaximumSize.Width > 0, locale + " action notice wraps without a fixed height");
                    Check(ok.Text == Text("DialogOk") && ok.DialogResult == DialogResult.OK && dialog.AcceptButton == ok && dialog.CancelButton == ok, locale + " action dialog retains one localized OK action");
                    Check(label.Bottom <= ok.Top && content.ClientRectangle.Contains(label.Bounds) && content.ClientRectangle.Contains(ok.Bounds), locale + " actual displayed dialog fits the complete notice and OK button");
                    if (locale == "de" && !string.IsNullOrEmpty(previewDirectory))
                    {
                        string name = notice == Text("EnterpriseRolloutSeatRequired") ? "seat" : notice == Text("EnterpriseRolloutBackendRequired") ? "backend" : notice == Text("EnterpriseRolloutStatusUnavailable") ? "status" : "offline";
                        using (var bitmap = new Bitmap(dialog.Width, dialog.Height))
                        {
                            dialog.DrawToBitmap(bitmap, new Rectangle(Point.Empty, dialog.Size));
                            bitmap.Save(System.IO.Path.Combine(previewDirectory, "action-notice-" + name + ".png"));
                        }
                    }
                    dialog.Scale(new SizeF(1.5f, 1.5f));
                    dialog.PerformLayout();
                    content.PerformLayout();
                    Check(label.Bottom <= ok.Top && content.ClientRectangle.Contains(label.Bounds) && content.ClientRectangle.Contains(ok.Bounds), locale + " scaled dialog retains separated complete text and button bounds");
                    dialog.Hide();
                }
            }
        }
        Call(T("Utilities.Strings"), "SetPreferredUiLanguage", "de");
        Check(Text("EnterpriseRolloutSeatRequired") == "Diese Installation wird zentral verwaltet und erfordert einen zugewiesenen, gültigen NC Connector Seat.", "German managed no-seat explanation retains the previous wording");
    }

    private static void TestResourceContextBlocks()
    {
        object addin = New("NextcloudTalkAddIn");
        object resources = Get(addin, "_calendarResourceController");
        Set(resources, "_opening", true);
        Check(!(bool)Call(resources, "IsContextAvailable", new object[] { null }), "Busy resource entry stays disabled even when a Seat explanation is reachable");
        Set(resources, "_opening", false);
        Check(!(bool)Call(resources, "IsContextAvailable", new object[] { null }), "Missing appointment context stays disabled");
    }
    private sealed class RibbonProxy : RealProxy
    {
        internal int Invalidations;
        internal bool Closed;
        internal RibbonProxy(Type type) : base(type) { }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            if (call.MethodName == "Invalidate")
            {
                Invalidations++;
                if (Closed) { return new ReturnMessage(new System.Runtime.InteropServices.COMException("Test closed ribbon"), call); }
            }
            if (call.MethodName == "Equals") { return new ReturnMessage(Object.ReferenceEquals(GetTransparentProxy(), call.Args[0]), null, 0, call.LogicalCallContext, call); }
            if (call.MethodName == "GetHashCode") { return new ReturnMessage(GetHashCode(), null, 0, call.LogicalCallContext, call); }
            return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
        }
    }
    private static void TestRibbonLifetime()
    {
        Type controllerType = T("Controllers.RibbonActionController");
        object controller = Activator.CreateInstance(controllerType, All, null, new object[] { null, null }, null);
        Type ribbonType = controllerType.GetMethod("Register", All).GetParameters()[0].ParameterType;
        var first = new RibbonProxy(ribbonType); var second = new RibbonProxy(ribbonType); var closed = new RibbonProxy(ribbonType) { Closed = true };
        FieldInfo subscriptions = T("Services.NextcloudConnectionState").GetField("StateChanged", All);
        int before = subscriptions.GetValue(null) == null ? 0 : ((Delegate)subscriptions.GetValue(null)).GetInvocationList().Length;
        try
        {
            Call(controller, "Register", first.GetTransparentProxy());
            Check(Get(controller, "_refreshTimer") == null, "Pre-start ribbon registration does not create a running timer");
            Call(controller, "Start"); Call(controller, "Start");
            Check(((Delegate)subscriptions.GetValue(null)).GetInvocationList().Length == before + 1, "Repeated Start registers one connection-state listener");
            Check(Get(controller, "_refreshTimer") != null, "Start also schedules refresh for an already registered ribbon");
            Call(controller, "Register", first.GetTransparentProxy()); Call(controller, "Register", second.GetTransparentProxy());
            Call(controller, "Register", first.GetTransparentProxy()); Call(controller, "Register", closed.GetTransparentProxy());
            Check(((IList)Get(controller, "_ribbons")).Count == 3, "Multiple windows are retained without duplicate registration");
            var timer = (Timer)Get(controller, "_refreshTimer");
            Check(timer != null && timer.Enabled && timer.Interval == 15000, "Disabled callbacks have an active UI-only refresh timer");
            Call(controller, "Invalidate");
            Check(first.Invalidations == 1 && second.Invalidations == 1 && closed.Invalidations == 1, "One closed ribbon does not prevent refreshing the other windows");
            Check(((IList)Get(controller, "_ribbons")).Count == 2, "Only the failed ribbon reference is removed");
            Call(controller, "Invalidate");
            Check(first.Invalidations == 2 && second.Invalidations == 2 && closed.Invalidations == 1, "Subsequent refreshes skip the dead window");
            Call(controller, "Stop"); Call(controller, "Stop");
            int after = subscriptions.GetValue(null) == null ? 0 : ((Delegate)subscriptions.GetValue(null)).GetInvocationList().Length;
            Check(after == before, "Stop removes the connection-state subscription");
            Check(Get(controller, "_refreshTimer") == null && ((IList)Get(controller, "_ribbons")).Count == 0, "Stop releases timer and all ribbon references");
            Call(controller, "Invalidate");
            Check(first.Invalidations == 2 && second.Invalidations == 2, "Stopped controller does not touch former windows");
        }
        finally { Call(controller, "Stop"); }
    }
    private static void TestRibbonXml()
    {
        object addin = New("NextcloudTalkAddIn");
        var buttons = new Dictionary<string, XmlElement>(StringComparer.Ordinal);
        foreach (string surface in new[] { "Microsoft.Outlook.Appointment", "Microsoft.Outlook.Mail.Compose", "Microsoft.Outlook.Explorer" })
        {
            var doc = new XmlDocument(); doc.LoadXml((string)Call(addin, "GetCustomUI", surface));
            foreach (XmlElement button in doc.SelectNodes("//*[local-name()='button']")) { buttons[button.GetAttribute("id")] = button; }
        }
        foreach (string id in new[] { "NcTalkCreateButton", "NcTalkFileLinkButton", "NcTalkInlineFileLinkButton", "NcResourceAppointmentButton", "NcResourceSchedulingButton" })
        {
            Check(buttons.ContainsKey(id), "Expected action button exists: " + id);
            Check(buttons[id].GetAttribute("getEnabled") == "OnGetActionEnabled" && buttons[id].GetAttribute("getSupertip") == "OnGetActionTooltip", "Action button uses both shared access callbacks: " + id);
        }
    }
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            assembly = Assembly.LoadFrom(args[0]);
            previewDirectory = args.Length > 1 ? args[1] : string.Empty;
            Application.EnableVisualStyles();
            TestAccessMatrix(); TestActionNoticeDialog(); TestResourceContextBlocks(); TestRibbonLifetime(); TestRibbonXml();
            Console.WriteLine("Ribbon access OK: " + assertions + " assertions."); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
'@
    $source = Join-Path $scratch 'RibbonAccessTests.cs'
    [IO.File]::WriteAllText($source, $test, (New-Object Text.UTF8Encoding($false)))
    $exe = Join-Path $scratch 'RibbonAccessTests.exe'
    & "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll $source
    if ($LASTEXITCODE -ne 0) { throw 'Ribbon access tests failed to compile.' }
    if ($PreviewDirectory) { $PreviewDirectory = (New-Item -ItemType Directory -Force -Path $PreviewDirectory).FullName }
    & $exe (Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/bin/Release/NcTalkOutlookAddIn.dll') $PreviewDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Ribbon access tests failed.' }
} finally {
    $resolvedScratch = (Resolve-Path -LiteralPath $scratch).Path
    if (-not $resolvedScratch.StartsWith((Join-Path ([IO.Path]::GetTempPath()) 'nc4ol-ribbon-tests-'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected scratch path.' }
    Remove-Item -LiteralPath $scratch -Recurse -Force
}
