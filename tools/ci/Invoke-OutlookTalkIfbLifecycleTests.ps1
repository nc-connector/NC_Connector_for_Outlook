Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"

function Assert-SourceContract {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Pattern
    )

    $text = Get-Content -LiteralPath $Path -Raw
    if ($text -notmatch $Pattern) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-SourceAbsent {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Pattern
    )

    $text = Get-Content -LiteralPath $Path -Raw
    if ($text -match $Pattern) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-PathAbsent {
    param(
        [string]$Name,
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        throw "Source check failed: $Name"
    }
    Write-Host "[OK] $Name"
}

function Get-SourceSlice {
    param(
        [string]$Path,
        [string]$Start,
        [string]$End
    )

    $source = Get-Content -LiteralPath $Path -Raw
    $startIndex = $source.IndexOf($Start, [StringComparison]::Ordinal)
    $endIndex = $source.IndexOf(
        $End,
        $startIndex + $Start.Length,
        [StringComparison]::Ordinal)
    if ($startIndex -lt 0 -or $endIndex -le $startIndex) {
        throw "Could not isolate source slice starting at '$Start'."
    }
    return $source.Substring($startIndex, $endIndex - $startIndex)
}

function Assert-LiteralOrder {
    param(
        [string]$Name,
        [string]$Source,
        [string[]]$Fragments
    )

    $position = 0
    foreach ($fragment in $Fragments) {
        $position = $Source.IndexOf(
            $fragment,
            $position,
            [StringComparison]::Ordinal)
        if ($position -lt 0) {
            throw "Source check failed: $Name"
        }
        $position += $fragment.Length
    }
    Write-Host "[OK] $Name"
}

$calendarLifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkCalendarLifecycle.cs"
$calendarMonitor = Join-Path $SourceRoot "Services\TalkCalendarLifecycleMonitor.cs"
$appointmentController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.cs"
$appointmentUserProperties = Join-Path $SourceRoot "Controllers\TalkAppointmentController.UserProperties.cs"
$appointmentSyncController = Join-Path $SourceRoot "Controllers\TalkAppointmentController.Sync.cs"
$talkRibbonController = Join-Path $SourceRoot "Controllers\TalkRibbonController.cs"
$addin = Join-Path $SourceRoot "NextcloudTalkAddIn.cs"
$appointmentSubscription = Join-Path $SourceRoot "NextcloudTalkAddIn.AppointmentSubscription.cs"
$calendarSelection = Join-Path $SourceRoot "NextcloudTalkAddIn.CalendarSelection.cs"
$hooks = Join-Path $SourceRoot "NextcloudTalkAddIn.Hooks.cs"
$subscriptionEnsure = Join-Path $SourceRoot "NextcloudTalkAddIn.SubscriptionEnsure.cs"
$appointmentSync = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkAppointmentSync.cs"
$talkLifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.TalkRoomLifecycle.cs"
$lifecycle = Join-Path $SourceRoot "NextcloudTalkAddIn.Lifecycle.cs"
$talkCoordinator = Join-Path $SourceRoot "Services\TalkRoomLifecycleCoordinator.cs"
$ifbManager = Join-Path $SourceRoot "Services\FreeBusyManager.cs"
$ifbServer = Join-Path $SourceRoot "Services\FreeBusyServer.cs"
$ifbCache = Join-Path $SourceRoot "Services\IfbAddressBookCache.cs"
$ifbOwnership = Join-Path $SourceRoot "Services\IfbRegistryOwnershipManager.cs"
$protectedStateStore = Join-Path $SourceRoot "Services\ProtectedJsonStateStore.cs"
$talkStore = Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"
$ifbStore = Join-Path $SourceRoot "Services\IfbRegistryStateStore.cs"

Assert-PathAbsent `
    "Global Talk calendar lifecycle partial is removed" `
    $calendarLifecycle
Assert-PathAbsent `
    "Global Talk calendar monitor is removed" `
    $calendarMonitor
foreach ($path in @(
    $appointmentSubscription,
    $calendarSelection,
    $hooks,
    $subscriptionEnsure,
    $talkLifecycle,
    $lifecycle,
    $talkCoordinator,
    $talkStore
)) {
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not enumerate Outlook stores in $path" `
        $path `
        '(?i)(?:session|application\.Session)\.Stores'
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not scan calendar item collections in $path" `
        $path `
        'TryScanAll|ReconcileSuccessfulScan'
    Assert-SourceAbsent `
        "Targeted Talk lifecycle does not subscribe to folder item events in $path" `
        $path `
        'ItemsEvents_Item(?:Add|Change|Remove)'
}
Assert-SourceContract `
    "Explorer lifecycle subscribes and unsubscribes calendar selection changes" `
    $hooks `
    'SelectionChange\s*\+=\s*selectionChangeHandler[\s\S]*SelectionChange\s*-=\s*selectionChangeHandler'
Assert-SourceContract `
    "Existing Explorer selection is processed immediately after the hook is installed" `
    $hooks `
    '_hookedExplorers\[explorerKey\]\s*=\s*explorer;[\s\S]{0,300}OnExplorerSelectionChanged\(explorerKey\)'
Assert-SourceContract `
    "Selected appointments reuse the targeted appointment subscription path" `
    $calendarSelection `
    'selection\s*=\s*explorer\.Selection[\s\S]*selection\.Count[\s\S]*EnsureSubscriptionForSelectedAppointment'
Assert-SourceContract `
    "Duplicate deferred selection callbacks release their unowned appointment reference" `
    $subscriptionEnsure `
    'if\s*\(\s*!_deferredAppointmentEnsureState\.TryQueuePendingKey\(ensureKey\)\s*\)\s*\{\s*return false;'
Assert-SourceAbsent `
    "Calendar selection binding does not enumerate stores, folders, or calendar Items collections" `
    $calendarSelection `
    '(?i)(?:session|application\.Session)\.Stores|\.Folders\b|\.Items\b'
Assert-SourceContract `
    "Outlook startup initializes only the pending Talk deletion queue" `
    $lifecycle `
    'InitializeTalkRoomLifecycle\s*\('
Assert-SourceAbsent `
    "Outlook startup does not start calendar reconciliation" `
    $lifecycle `
    'StartTalkRoomLifecycleRecovery|EnsureTalkCalendarLifecycleMonitor'
Assert-SourceContract `
    "Appointment BeforeDelete rejects cancelled deletion and recurrence instances" `
    $appointmentSubscription `
    'OnBeforeDelete[\s\S]*if\s*\(\s*cancel\s*\)[\s\S]*IsRoomDeletionAllowedForRecurrence\s*\(\s*\)[\s\S]*QueueSavedTalkRoomDeletion\s*\(\s*\)'
Assert-SourceContract `
    "Recurring occurrences and exceptions retain the shared Talk room" `
    $appointmentSubscription `
    'OlRecurrenceState\.olApptNotRecurring[\s\S]*OlRecurrenceState\.olApptMaster'
Assert-SourceContract `
    "Saved appointment deletion skips rooms delegated to another user" `
    $appointmentSubscription `
    'QueueSavedTalkRoomDeletion[\s\S]*IsDelegatedToOtherUser[\s\S]*return;[\s\S]*_owner\.QueueSavedTalkRoomDeletion'
Assert-SourceContract `
    "Saved appointment deletion creates a policy-required durable job" `
    $talkLifecycle `
    'QueueSavedTalkRoomDeletion[\s\S]*QueueTalkRoomDeletion\([\s\S]*true\s*\)'
Assert-SourceContract `
    "Unsaved appointment cleanup creates an unconditional durable job" `
    $talkLifecycle `
    'QueueUnsavedTalkRoomDeletion[\s\S]*QueueTalkRoomDeletion\([\s\S]*false\s*\)'

$talkDialogFlow = Get-SourceSlice `
    $talkRibbonController `
    "private bool RunTalkDialogOnUiThread(" `
    "private void CleanupCreatedRoom("
Assert-LiteralOrder `
    "Talk replacement captures and attaches before retiring the existing room" `
    $talkDialogFlow `
    @(
        "TryCaptureAppointmentState(appointment",
        "result = service.CreateRoom(request);",
        "_owner.ApplyRoomToAppointment(appointment, request, result);",
        "_owner.TryDeleteRoom(existingToken, existingIsEvent)"
    )
Assert-LiteralOrder `
    "Failed Talk attachment restores the appointment before cleaning the created room" `
    $talkDialogFlow `
    @(
        "if (!roomAttached)",
        "RestoreAppointmentState(appointment, appointmentState)",
        "CleanupCreatedRoom(result)"
    )
Assert-SourceContract `
    "Failed old-room cleanup is queued without rolling back the replacement" `
    $talkRibbonController `
    'if\s*\(\s*!_owner\.TryDeleteRoom\(existingToken,\s*existingIsEvent\)\s*\)[\s\S]{0,240}_owner\.QueueUnsavedTalkRoomDeletion\(existingToken,\s*existingIsEvent\)'

$createdRoomCleanup = Get-SourceSlice `
    $talkRibbonController `
    "private void CleanupCreatedRoom(" `
    "private static void ShowAppointmentAttachError("
Assert-LiteralOrder `
    "Failed created-room cleanup is retried through the durable queue" `
    $createdRoomCleanup `
    @(
        "_owner.TryDeleteRoom(roomToken, result.CreatedAsEventConversation, false)",
        "_owner.QueueUnsavedTalkRoomDeletion("
    )
Assert-SourceContract `
    "Talk dialog and appointment mutation remain on the captured Outlook STA" `
    $talkRibbonController `
    'RunOnOutlookUiThreadAsync\([\s\S]{0,160}RunTalkDialogOnUiThread\('

$applyRoom = Get-SourceSlice `
    $appointmentController `
    "internal bool ApplyRoomToAppointment(" `
    "internal bool PersistCoreIcalProperties("
Assert-LiteralOrder `
    "Talk subscription and synchronization start only after local metadata succeeds" `
    $applyRoom `
    @(
        "if (!metadataStored)",
        "_owner.RegisterSubscription(appointment, result);",
        "_owner.QueueTalkAppointmentSync(descriptionSnapshot);"
    )
Assert-SourceContract `
    "Talk appointment apply result is returned to the ribbon flow" `
    $addin `
    'internal\s+bool\s+ApplyRoomToAppointment[\s\S]{0,220}return\s+_talkAppointmentController\.ApplyRoomToAppointment'

$registration = Get-SourceSlice `
    $addin `
    "internal void RegisterSubscription(Outlook.AppointmentItem appointment, string roomToken, string roomUrl" `
    "private void UnregisterSubscription("
Assert-SourceContract `
    "Same-appointment subscription reuse also requires the same room token" `
    $addin `
    'existingByEntry\.IsFor\(appointment\)\s*&&\s*existingByEntry\.MatchesToken\(normalizedRoomToken\)[\s\S]{0,120}return;'
Assert-LiteralOrder `
    "Changed-token subscription is constructed before the old appointment binding is disposed" `
    $registration `
    @(
        "var subscription = new AppointmentSubscription(",
        "staleByEntry.Dispose();"
    )
Assert-SourceContract `
    "Partial Talk appointment event registration is rolled back" `
    $appointmentSubscription `
    'AttachEventHandlers[\s\S]{0,900}catch[\s\S]{0,180}DetachEventHandlers\([\s\S]{0,180}throw;'
Assert-SourceContract `
    "Talk appointment disposal uses the complete event-detach path" `
    $appointmentSubscription `
    'public\s+void\s+Dispose\(\)[\s\S]{0,320}DetachEventHandlers\(true,\s*true,\s*true\)'

$appointmentSources = (Get-Content -LiteralPath $appointmentController -Raw) `
    + (Get-Content -LiteralPath $appointmentUserProperties -Raw)
if ([regex]::Matches($appointmentSources, 'appointment\.UserProperties').Count -ne 1) {
    throw "Source check failed: Talk appointment user properties have more than one COM ownership path."
}
Write-Host "[OK] Talk appointment user properties have one COM ownership path"

$userPropertyMethods = @(
    @{ Start = "internal static bool SetUserProperty("; End = "internal static string GetUserPropertyText(" },
    @{ Start = "internal static string GetUserPropertyText("; End = "internal static bool HasUserProperty(" },
    @{ Start = "internal static bool HasUserProperty("; End = "internal static bool GetUserPropertyBool(" },
    @{ Start = "internal static bool GetUserPropertyBool("; End = "internal static bool RemoveUserProperty(" },
    @{ Start = "internal static bool RemoveUserProperty("; End = "private static TResult UseUserProperty<TResult>(" }
)
foreach ($method in $userPropertyMethods) {
    $methodSource = Get-SourceSlice `
        $appointmentUserProperties `
        $method.Start `
        $method.End
    if (-not $methodSource.Contains("UseUserProperty(")) {
        throw "Source check failed: $($method.Start) bypasses the shared COM ownership helper."
    }
    if ($methodSource.Contains("appointment.UserProperties")) {
        throw "Source check failed: $($method.Start) owns Outlook user properties directly."
    }
}
Write-Host "[OK] All Talk appointment user-property operations use the shared COM ownership helper"

$userPropertyOwnership = Get-SourceSlice `
    $appointmentUserProperties `
    "private static TResult UseUserProperty<TResult>(" `
    "private static bool UserPropertyValueEquals("
Assert-LiteralOrder `
    "Talk user-property COM objects are released child before collection" `
    $userPropertyOwnership `
    @(
        "finally",
        "ComInteropScope.TryRelease(property,",
        "ComInteropScope.TryRelease(properties,"
    )
if ($userPropertyOwnership.Contains("TryRelease(appointment")) {
    throw "Source check failed: Borrowed AppointmentItem is released by the user-property helper."
}
Write-Host "[OK] User-property helper does not release the borrowed AppointmentItem"

foreach ($propertyName in @(
    "IcalToken",
    "IcalUrl",
    "IcalLobby",
    "IcalStart",
    "IcalEvent",
    "IcalObjectId",
    "IcalAddUsers",
    "IcalAddGuests",
    "IcalDelegate",
    "IcalDelegateName",
    "IcalDelegated",
    "IcalDelegateReady"
)) {
    Assert-SourceContract `
        "Talk replacement snapshot includes $propertyName" `
        $appointmentController `
        ("TalkPropertyNames[\s\S]{0,900}NextcloudTalkAddIn\." + $propertyName)
}
Assert-SourceContract `
    "Talk replacement snapshot preserves subject, location, and cloned RTF body" `
    $appointmentController `
    'TryCaptureAppointmentState[\s\S]*Subject\s*=\s*appointment\.Subject[\s\S]*Location\s*=\s*appointment\.Location[\s\S]*RtfBody\s*=\s*\(byte\[\]\)rtfBody\.Clone\(\)'
Assert-SourceAbsent `
    "Remote promotion does not leave the room before Outlook persistence" `
    $appointmentSyncController `
    '\.LeaveRoom\s*\('
Assert-SourceContract `
    "Outlook saves the handoff before the background leave call" `
    $appointmentSyncController `
    'appointment\.Save\(\)'
Assert-SourceContract `
    "LeaveRoom runs only from the post-persistence completion path" `
    $appointmentSync `
    'handoffPersisted[\s\S]*\.LeaveRoom\s*\('
Assert-SourceContract `
    "IFB server does not log raw authenticated request URLs" `
    $ifbServer `
    'TryParseAuthorizedEmailPath'
Assert-SourceContract `
    "Outlook IFB URL keeps the attendee domain placeholder" `
    $ifbManager `
    '/freebusy/%NAME%@%SERVER%\.vfb'
Assert-SourceAbsent `
    "Outlook IFB URL does not guess a domain from connector settings" `
    $ifbManager `
    'GuessDefaultDomain'
Assert-SourceAbsent `
    "IFB address lookup has no local-part fallback" `
    $ifbCache `
    'TryResolveEmail|_localPartToEmail'
Assert-SourceContract `
    "IFB server looks up the complete attendee address" `
    $ifbServer `
    'TryGetUid\(\s*_configuration,\s*_cacheHours,\s*email'
Assert-SourceContract `
    "IFB server limits concurrent requests directly" `
    $ifbServer `
    'SemaphoreSlim\s+_requestSlots'
Assert-SourceContract `
    "IFB cache scope includes profile, base URL, and configured username" `
    $ifbCache `
    'BuildScopeFingerprint\(\s*_profileScope,\s*serverBaseUrl,\s*username'
if ((Get-Content -LiteralPath $ifbServer -Raw) -match 'RawUrl') {
    throw "Source check failed: IFB server references RawUrl."
}
Write-Host "[OK] IFB server never references RawUrl"
foreach ($path in @($appointmentController, $appointmentSubscription)) {
    $text = Get-Content -LiteralPath $path -Raw
    if ($text -match 'CreateTalkService\s*\(') {
        throw "Source check failed: appointment callback path creates a synchronous Talk service in $path."
    }
}
Write-Host "[OK] Appointment callback paths contain no synchronous Talk service creation"
Assert-SourceContract `
    "IFB policy keys are checked before user-owned registry writes" `
    $ifbOwnership `
    'ThrowIfPolicyConflicts'
Assert-SourceContract `
    "Protected state journals use durable backup replacement" `
    $protectedStateStore `
    'DurableFileReplace\.CommitPreparedFile'
Assert-SourceContract `
    "Talk lifecycle journal uses the protected JSON state store" `
    $talkStore `
    'ProtectedJsonStateStore<TalkRoomLifecycleState>'
Assert-SourceContract `
    "IFB ownership journal uses the protected JSON state store" `
    $ifbStore `
    'ProtectedJsonStateStore<IfbRegistryState>'

$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    "nc4ol-talk-ifb-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "TalkIfbLifecycleTests.cs"
    $testCode = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static readonly List<string> Messages = new List<string>();
        internal static void Log(string category, string message) { Messages.Add(message); }
        internal static void LogException(
            string category,
            string message,
            Exception ex) { Messages.Add(message + " " + ex.Message); }
        internal static IDisposable BeginOperation(string category, string operation)
        {
            return new Scope();
        }
        private sealed class Scope : IDisposable
        {
            public void Dispose() { }
        }
    }

    internal static class Strings
    {
        internal const string ErrorMissingCredentials = "Missing credentials";
        internal const string TalkSystemAddressbookInvalidResponse = "Invalid address book";
        internal const string TalkSystemAddressbookFetchFailed = "Address book fetch failed";
    }
}

namespace NcTalkOutlookAddIn.Settings
{
    internal class AddinSettings
    {
        public string ServerUrl { get; set; }
        public string Username { get; set; }
        public string AppPassword { get; set; }
        public string IfbPreviousFreeBusyPath { get; set; }
    }
}

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class NcHttpRequestOptions
    {
        internal string Method, Url, Accept;
        internal int TimeoutMs;
        internal bool IncludeAuthHeader, IncludeOcsApiHeader, ParseJson;
    }

    internal sealed class NcHttpResponse
    {
        internal bool HasHttpResponse;
        internal HttpStatusCode StatusCode;
        internal string ContentType, ResponseText;
        internal Exception TransportException;
    }

    internal sealed class NcHttpClient
    {
        internal static NcHttpResponse NextResponse;
        internal static NcHttpRequestOptions LastOptions;
        internal static int SendCount;
        internal static string LastLogin;
        internal NcHttpClient(TalkServiceConfiguration configuration)
        {
            LastLogin = configuration.Username;
        }
        internal NcHttpResponse Send(NcHttpRequestOptions options)
        {
            LastOptions = options;
            SendCount++;
            return NextResponse;
        }
    }

    internal static class NextcloudUserIdentityService
    {
        internal static string CanonicalUserId = "alice";

        internal static string ResolveCurrentUserId(
            TalkServiceConfiguration configuration,
            bool forceRefresh = false)
        {
            return CanonicalUserId;
        }
    }

    internal sealed class TalkService
    {
        internal readonly List<string> Users = new List<string>();
        internal readonly List<string> Guests = new List<string>();
        internal TalkService(
            TalkServiceConfiguration configuration) { }

        internal void DeleteRoom(
            string roomToken,
            bool isEventConversation) { }
        internal void AddUserParticipant(string roomToken, string uid) { Users.Add(uid); }
        internal void AddGuestParticipant(string roomToken, string email) { Guests.Add(email); }
    }
}

internal static class NextcloudTalkAddIn
{
    internal static void LogTalkMessage(string message)
    {
        DiagnosticsLogger.Log("TALK", message);
    }
}

internal static class TalkIfbLifecycleTests
{
    private static int failures;

    private static void Check(
        string name,
        bool condition,
        string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }
        failures++;
        Console.Error.WriteLine(
            "[FAIL] "
            + name
            + (detail.Length == 0 ? "" : ": " + detail));
    }

    public static int Main()
    {
        TestLifecycleAccountMatching();
        TestPolicyRequiredDeletion();
        TestUnconditionalDeletion();
        TestDeletionRetryAcrossRestart();
        TestPendingStoreMigration();
        TestSyncCoalescing();
        TestProtectedStateStoreCompatibility();
        TestLegacyIfbRegistryMigration();
        TestDurableReplacement();
        TestSystemAddressbookResponses();
        TestSystemAddressbookCacheRecovery();
        TestSystemAddressbookScope();
        TestSystemAddressbookParticipants();

        if (failures > 0)
        {
            return 1;
        }
        Console.WriteLine("All Talk/IFB lifecycle tests passed.");
        return 0;
    }

    private const string AddressbookExport =
        "BEGIN:VCARD\r\nVERSION:3.0\r\nUID:alice\r\nEMAIL:alice@example.test\r\nEND:VCARD\r\n"
        + "BEGIN:VCARD\r\nVERSION:3.0\r\nUID:bob\r\nEMAIL;TYPE=WORK:bob@example.test\r\n"
        + "item1.EMAIL;TYPE=\"HOME\":bob.other@example.test\r\nEND:VCARD\r\n"
        + "BEGIN:VCARD\r\nVERSION:3.0\r\nUID:moderator-only\r\nFN:Moderator\r\nEND:VCARD\r\n";

    private static NcHttpResponse AddressbookResponse(int status, string content, string contentType)
    {
        return new NcHttpResponse
        {
            HasHttpResponse = true,
            StatusCode = (HttpStatusCode)status,
            ResponseText = content,
            ContentType = contentType
        };
    }

    private static void TestSystemAddressbookResponses()
    {
        string root = NewTestRoot("addressbook-response");
        try
        {
            var cache = new IfbAddressBookCache(root, "responses");
            var configuration = NewConfiguration();
            string[] mediaTypes = { "text/vcard", "text/x-vcard; charset=utf-8", "TEXT/DIRECTORY", "text/html", "application/json", null };
            string[] invalid =
            {
                "<html>private-response-marker</html>",
                "{\"error\":\"private-response-marker\"}",
                "BEGIN:VCARD\r\nUID:alice\r\n",
                "BEGIN:VCARD\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID-FAKE:alice\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nUID:bob\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\\nbob\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nBEGIN:VCARD\r\nUID:bob\r\nEND:VCARD\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nBEGIN:VEVENT\r\nEND:VEVENT\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nEND:VEVENT\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nEND:VCARD:extra\r\n",
                "BEGIN:VCARD:extra\r\nUID:alice\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID:alice\r\nproperty-without-colon\r\nEND:VCARD\r\n",
                "BEGIN:VCARD\r\nUID;VALUE=\"unterminated:alice\r\nEND:VCARD\r\n",
                AddressbookExport + "END:VCARD\r\n",
                AddressbookExport + "BEGIN:VCARD\r\nUID:unfinished\r\n",
                AddressbookExport + "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n",
                AddressbookExport + "<html>private-response-marker</html>",
                "{\"error\":true}\r\n" + AddressbookExport
            };
            int cases = 0;
            foreach (int status in new[] { 200, 404 })
            {
                foreach (string mediaType in mediaTypes)
                {
                    NcHttpClient.NextResponse = AddressbookResponse(status, AddressbookExport, mediaType);
                    var result = cache.GetSystemAddressbookStatus(configuration, 24, true);
                    Check("Valid export HTTP " + status + " / " + (mediaType ?? "missing"), result.Available && result.Count == 3);
                    cases++;
                    foreach (string content in invalid)
                    {
                        NcHttpClient.NextResponse = AddressbookResponse(status, content, mediaType);
                        result = cache.GetSystemAddressbookStatus(configuration, 24, true);
                        Check("Invalid export case " + cases, !result.Available && result.Count == 0
                            && result.Error == Strings.TalkSystemAddressbookInvalidResponse);
                        cases++;
                    }
                    foreach (string empty in new[] { "", " \r\n\t\n", null })
                    {
                        NcHttpClient.NextResponse = AddressbookResponse(status, empty, mediaType);
                        result = cache.GetSystemAddressbookStatus(configuration, 24, true);
                        bool expected = status == 200 && mediaType != null
                            && mediaType != "text/html" && mediaType != "application/json";
                        Check("Empty export case " + cases, result.Available == expected && result.Count == 0);
                        cases++;
                    }
                }
            }
            foreach (int status in new[] { 301, 400, 401, 403, 500, 503 })
            {
                NcHttpClient.NextResponse = AddressbookResponse(status, AddressbookExport, "text/vcard");
                var result = cache.GetSystemAddressbookStatus(configuration, 24, true);
                Check("Non-404 error rejects valid export HTTP " + status,
                    !result.Available && result.Error == Strings.TalkSystemAddressbookFetchFailed);
            }
            NcHttpClient.NextResponse = AddressbookResponse(404,
                "begin:vcard\r\nVERSION:3.0\r\nUID;VALUE=text:canonical\\,id\r\n"
                + "item1.EMAIL;TYPE=\"work:primary\";PREF=1:folded@exa\r\n mple.test\r\n"
                + "NOTE:folded content\r\n\tEND:VCARD is text\r\nend:vcard\r\n", "text/html");
            var users = cache.GetUsers(configuration, 24, true);
            string uid;
            Check("Folded and grouped properties retain canonical escaped UID",
                users.Count == 1 && users[0].UserId == "canonical,id"
                && cache.TryGetUid(configuration, 24, "folded@example.test", out uid) && uid == "canonical,id");
            Check("Response contents never enter address book diagnostics",
                !string.Join("\n", DiagnosticsLogger.Messages.ToArray()).Contains("private-response-marker"));
            Check("Response matrix executes full cache fetch path", cases == 288);
        }
        finally { Directory.Delete(root, true); }
    }

    private static object CacheField(IfbAddressBookCache cache, string name)
    {
        return typeof(IfbAddressBookCache).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(cache);
    }

    private static void TestSystemAddressbookCacheRecovery()
    {
        string root = NewTestRoot("addressbook-cache");
        try
        {
            var configuration = NewConfiguration();
            var cache = new IfbAddressBookCache(root, "cache");
            NcHttpClient.NextResponse = AddressbookResponse(404, AddressbookExport, null);
            var users = cache.GetUsers(configuration, 24, true);
            string cachePath = Directory.GetFiles(root, "ifb-addressbook-cache-*.json")[0];
            string saved = File.ReadAllText(cachePath);
            object savedTime = CacheField(cache, "_generatedUtc");
            object savedEmails = CacheField(cache, "_emailToUid");
            int sends = NcHttpClient.SendCount;
            var reloaded = new IfbAddressBookCache(root, "cache").GetUsers(configuration, 24, false);
            Check("UID-only moderator survives disk cache reload", users.Count == 3 && reloaded.Count == 3
                && reloaded[2].UserId == "moderator-only" && reloaded[2].Email == ""
                && NcHttpClient.SendCount == sends);
            string uid;
            Check("Every email for a contact resolves to its UID",
                cache.TryGetUid(configuration, 24, "BOB.OTHER@example.test", out uid) && uid == "bob");
            foreach (NcHttpResponse response in new[]
            {
                AddressbookResponse(404, "", "text/vcard"),
                AddressbookResponse(200, "private-response-marker", "text/vcard"),
                AddressbookResponse(401, AddressbookExport, "text/vcard"),
                AddressbookResponse(403, AddressbookExport, "text/vcard"),
                AddressbookResponse(503, AddressbookExport, "text/vcard"),
                new NcHttpResponse { TransportException = new IOException("private-response-marker") },
                null
            })
            {
                NcHttpClient.NextResponse = response;
                var failed = cache.GetSystemAddressbookStatus(configuration, 24, true);
                Check("Failed refresh preserves contacts, timestamp and disk cache",
                    !failed.Available && failed.Count == 0
                    && object.ReferenceEquals(savedEmails, CacheField(cache, "_emailToUid"))
                    && savedTime.Equals(CacheField(cache, "_generatedUtc"))
                    && File.ReadAllText(cachePath) == saved);
                sends = NcHttpClient.SendCount;
                failed = cache.GetSystemAddressbookStatus(configuration, 24, false);
                Check("Failed refresh bypasses still-fresh memory and disk on retry",
                    !failed.Available && NcHttpClient.SendCount == sends + 1);
                bool rejected = false;
                try { cache.TryGetUid(configuration, 24, "bob@example.test", out uid); }
                catch (InvalidOperationException) { rejected = true; }
                catch (InvalidDataException) { rejected = true; }
                Check("Identity lookup throws instead of reporting a failed fetch as a miss", rejected);
                sends = NcHttpClient.SendCount;
                failed = new IfbAddressBookCache(root, "cache").GetSystemAddressbookStatus(configuration, 24, false);
                Check("Another consumer cannot reuse disk after the current scope failed",
                    !failed.Available && NcHttpClient.SendCount == sends + 1);
            }
            NcHttpClient.NextResponse = AddressbookResponse(404, AddressbookExport.Replace("bob", "recovered"), "application/json");
            Check("A valid 404 retry restores availability",
                cache.GetSystemAddressbookStatus(configuration, 24, false).Available
                && cache.TryGetUid(configuration, 24, "recovered@example.test", out uid) && uid == "recovered");
            sends = NcHttpClient.SendCount;
            cache.GetUsers(configuration, 24, false);
            Check("Successful recovery re-enables cache hits", NcHttpClient.SendCount == sends);
            NcHttpClient.NextResponse = AddressbookResponse(200, "", "text/directory");
            var empty = cache.GetSystemAddressbookStatus(configuration, 24, true);
            Check("Confirmed empty export replaces old contacts and persists empty success",
                empty.Available && empty.Count == 0
                && new IfbAddressBookCache(root, "cache").GetUsers(configuration, 24, false).Count == 0);
            Check("Transport response contents never reach error logs",
                !string.Join("\n", DiagnosticsLogger.Messages.ToArray()).Contains("private-response-marker"));
        }
        finally { Directory.Delete(root, true); }
    }

    private static void TestSystemAddressbookScope()
    {
        string root = NewTestRoot("addressbook-scope");
        try
        {
            var configuration = NewConfiguration();
            var cache = new IfbAddressBookCache(root, "scope");
            NcHttpClient.NextResponse = AddressbookResponse(200, AddressbookExport, "text/vcard");
            cache.GetUsers(configuration, 24, true);
            NextcloudUserIdentityService.CanonicalUserId = "canonical/user";
            var other = new TalkServiceConfiguration(configuration.GetNormalizedBaseUrl(), "second-login@example.test", "test-password");
            NcHttpClient.NextResponse = AddressbookResponse(404,
                "BEGIN:VCARD\r\nUID:other-user\r\nEND:VCARD\r\n", "text/plain");
            var users = cache.GetUsers(other, 24, false);
            Check("Account switch resolves only the new account export", users.Count == 1 && users[0].UserId == "other-user");
            Check("DAV path uses canonical UID while authentication uses login",
                NcHttpClient.LastOptions.Url.Contains("/canonical%2Fuser/z-server-generated--system?export")
                && NcHttpClient.LastLogin == "second-login@example.test"
                && NcHttpClient.LastOptions.IncludeAuthHeader && !NcHttpClient.LastOptions.IncludeOcsApiHeader
                && !NcHttpClient.LastOptions.ParseJson && NcHttpClient.LastOptions.Method == "GET");
            int sends = NcHttpClient.SendCount;
            Check("Switching back loads only the matching previous account",
                cache.GetUsers(configuration, 24, false).Count == 3 && NcHttpClient.SendCount == sends);
            NcHttpClient.NextResponse = AddressbookResponse(503, AddressbookExport, "text/vcard");
            foreach (var changed in new[]
            {
                new TalkServiceConfiguration("https://other.example.test", configuration.Username, "test-password"),
                new TalkServiceConfiguration(configuration.GetNormalizedBaseUrl() + "/subpath", configuration.Username, "test-password")
            })
            {
                Check("Server and subpath change cannot reuse another scope",
                    !cache.GetSystemAddressbookStatus(changed, 24, false).Available);
            }
            Check("Profile change cannot reuse another scope",
                !new IfbAddressBookCache(root, "other-profile").GetSystemAddressbookStatus(configuration, 24, false).Available);
            Check("Failed other scope does not poison the matching cached account",
                cache.GetSystemAddressbookStatus(configuration, 24, false).Available);
        }
        finally
        {
            NextcloudUserIdentityService.CanonicalUserId = "alice";
            Directory.Delete(root, true);
        }
    }

    private static void TestSystemAddressbookParticipants()
    {
        string root = NewTestRoot("addressbook-participants");
        try
        {
            var configuration = NewConfiguration();
            var snapshot = new TalkAppointmentSyncSnapshot
            {
                Configuration = configuration, DataDirectory = root, ProfileScope = "participants", CacheHours = 24,
                AddUsers = true, AddGuests = true, RoomToken = "room",
                AttendeeEmails = new List<string> { "alice@example.test", "bob@example.test", "external@example.test" }
            };
            NcHttpClient.NextResponse = AddressbookResponse(404, AddressbookExport, "text/html");
            var service = new TalkService(configuration);
            ExecuteParticipantSync(service, snapshot);
            Check("Actual participant sync accepts 404, omits self, and distinguishes internal from guest",
                service.Users.Count == 1 && service.Users[0] == "bob"
                && service.Guests.Count == 1 && service.Guests[0] == "external@example.test");
            var cache = new IfbAddressBookCache(root, "participants");
            foreach (NcHttpResponse response in new[]
            {
                AddressbookResponse(200, AddressbookExport + "END:VCARD\r\n", "text/vcard"),
                AddressbookResponse(404, "", "text/vcard"),
                AddressbookResponse(401, AddressbookExport, "text/vcard"),
                new NcHttpResponse { TransportException = new IOException("offline") }
            })
            {
                NcHttpClient.NextResponse = response;
                cache.GetSystemAddressbookStatus(configuration, 24, true);
                service = new TalkService(configuration);
                ExecuteParticipantSync(service, snapshot);
                Check("Actual participant sync stops on failure even with previous disk contacts",
                    service.Users.Count == 0 && service.Guests.Count == 0);
            }
            NcHttpClient.NextResponse = AddressbookResponse(404, AddressbookExport, null);
            service = new TalkService(configuration);
            ExecuteParticipantSync(service, snapshot);
            Check("Actual participant sync resumes only after successful retry",
                service.Users.Count == 1 && service.Users[0] == "bob" && service.Guests.Count == 1);
        }
        finally { Directory.Delete(root, true); }
    }

    // EXECUTE_PARTICIPANT_SYNC_SOURCE

    private static void TestLifecycleAccountMatching()
    {
        var record = new TalkRoomLifecycleRecord
        {
            ServerBaseUrl =
                "https://cloud.example.org/nextcloud",
            AccountLogin = "old-login",
            AccountId = "alice",
            RoomToken = "room",
            PendingDeletion = true
        };
        Check(
            "Canonical account ID accepts a changed login alias",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "new-login",
                "alice"));
        Check(
            "Canonical account ID rejects another account",
            !TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "old-login",
                "bob"));

        record.AccountId = string.Empty;
        record.AccountLogin = "me@example.org";
        Check(
            "Unbound deletion job requires its exact login",
            TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "me@example.org",
                "alice")
            && !TalkRoomLifecycleCoordinator.MatchesAccount(
                record,
                record.ServerBaseUrl,
                "other@example.org",
                "alice"));
    }

    private static void TestPolicyRequiredDeletion()
    {
        string root = NewTestRoot("policy");
        const string profile = "policy-profile";
        var policyStarted = new ManualResetEventSlim(false);
        var releasePolicy = new ManualResetEventSlim(false);
        int deleteCount = 0;
        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () =>
            {
                policyStarted.Set();
                releasePolicy.Wait(TimeSpan.FromSeconds(5));
                return false;
            },
            configuration => "alice",
            (configuration, token, isEventConversation) =>
                Interlocked.Increment(ref deleteCount));
        try
        {
            Check(
                "Policy-required deletion accepts a complete job",
                coordinator.QueueDeletion(
                    "policy-room",
                    true,
                    NewConfiguration(),
                    true));
            Check(
                "Policy-required deletion reaches the policy gate",
                policyStarted.Wait(TimeSpan.FromSeconds(5)));
            TalkRoomLifecycleState pending =
                LoadLifecycleState(root, profile);
            Check(
                "Policy-required deletion is persisted before evaluation",
                pending.Records.Count == 1
                && pending.Records[0].PolicyRequired
                && pending.Records[0].PendingDeletion
                && pending.Records[0].AccountId == "alice");

            releasePolicy.Set();
            Check(
                "Disabled saved-event deletion removes the pending job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        coordinator,
                        "policy-room") == null));
            Check(
                "Policy-off removal is saved to the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
            Check(
                "Disabled saved-event deletion does not call Talk",
                Volatile.Read(ref deleteCount) == 0);
        }
        finally
        {
            releasePolicy.Set();
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestUnconditionalDeletion()
    {
        string root = NewTestRoot("unconditional");
        const string profile = "unconditional-profile";
        var deleted = new ManualResetEventSlim(false);
        var releaseDelete = new ManualResetEventSlim(false);
        int policyCount = 0;
        string deletedToken = string.Empty;
        bool deletedEventRoom = false;
        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () =>
            {
                Interlocked.Increment(ref policyCount);
                return false;
            },
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                deletedToken = token;
                deletedEventRoom = isEventConversation;
                deleted.Set();
                releaseDelete.Wait(TimeSpan.FromSeconds(5));
            });
        try
        {
            Check(
                "Unsaved cleanup accepts a complete job",
                coordinator.QueueDeletion(
                    "unsaved-room",
                    true,
                    NewConfiguration(),
                    false));
            Check(
                "Unsaved cleanup executes without the saved-event policy",
                deleted.Wait(TimeSpan.FromSeconds(5)));
            Check(
                "Unsaved cleanup preserves token and room type",
                deletedToken == "unsaved-room"
                && deletedEventRoom);
            Check(
                "Unsaved cleanup does not evaluate saved-event policy",
                Volatile.Read(ref policyCount) == 0);
            TalkRoomLifecycleRecord pending =
                ReadCoordinatorRecord(
                    coordinator,
                    "unsaved-room");
            Check(
                "Unsaved cleanup is durable before the Talk call completes",
                pending != null
                && pending.PendingDeletion
                && !pending.PolicyRequired);
            releaseDelete.Set();
            Check(
                "Successful unsaved cleanup clears the pending job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        coordinator,
                        "unsaved-room") == null));
            Check(
                "Successful unsaved cleanup updates the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
        }
        finally
        {
            releaseDelete.Set();
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestDeletionRetryAcrossRestart()
    {
        string root = NewTestRoot("retry");
        const string profile = "retry-profile";
        var failed = new ManualResetEventSlim(false);
        var releaseFailure = new ManualResetEventSlim(false);
        var firstCoordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                failed.Set();
                releaseFailure.Wait(TimeSpan.FromSeconds(5));
                throw new InvalidOperationException("retry");
            });
        try
        {
            Check(
                "Retry test accepts a complete deletion job",
                firstCoordinator.QueueDeletion(
                    "retry-room",
                    false,
                    NewConfiguration(),
                    true));
            Check(
                "Failed deletion reaches the retry worker",
                failed.Wait(TimeSpan.FromSeconds(5)));
            TalkRoomLifecycleRecord pending =
                ReadCoordinatorRecord(
                    firstCoordinator,
                    "retry-room");
            Check(
                "Deletion job remains pending while Talk is in flight",
                pending != null
                && pending.PendingDeletion
                && pending.AttemptCount == 0);
            releaseFailure.Set();
            Check(
                "Failed deletion remains persisted with retry metadata",
                WaitUntil(
                    () =>
                    {
                        TalkRoomLifecycleRecord record =
                            ReadCoordinatorRecord(
                                firstCoordinator,
                                "retry-room");
                        return record != null
                            && record.AttemptCount == 1
                            && record.NextAttemptUtc
                                > DateTime.UtcNow;
                    }));
            TalkRoomLifecycleState savedRetry =
                LoadLifecycleState(root, profile);
            Check(
                "Retry metadata is written to the durable queue",
                savedRetry.Records.Count == 1
                && savedRetry.Records[0].AttemptCount == 1
                && savedRetry.Records[0].AccountId == "alice");
        }
        finally
        {
            releaseFailure.Set();
            firstCoordinator.Dispose();
        }

        TalkRoomLifecycleState retryState =
            LoadLifecycleState(root, profile);
        retryState.Records[0].NextAttemptUtc = DateTime.MinValue;
        new TalkRoomLifecycleStore(root, profile).Save(retryState);

        var deleted = new ManualResetEventSlim(false);
        var releaseRestartDelete =
            new ManualResetEventSlim(false);
        var restartedCoordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            () => NewSettings("renamed-login"),
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) =>
            {
                if (token == "retry-room" && !isEventConversation)
                {
                    deleted.Set();
                    releaseRestartDelete.Wait(
                        TimeSpan.FromSeconds(5));
                }
            });
        try
        {
            restartedCoordinator.StartPendingProcessing();
            Check(
                "Restart matches the canonical account after a login alias change",
                deleted.Wait(TimeSpan.FromSeconds(5)));
            releaseRestartDelete.Set();
            Check(
                "Restarted deletion clears its durable job",
                WaitUntil(
                    () => ReadCoordinatorRecord(
                        restartedCoordinator,
                        "retry-room") == null));
            Check(
                "Restarted deletion updates the durable queue",
                LoadLifecycleState(
                    root,
                    profile).Records.Count == 0);
        }
        finally
        {
            releaseRestartDelete.Set();
            restartedCoordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static void TestPendingStoreMigration()
    {
        string root = NewTestRoot("migration");
        const string profile = "migration-profile";
        var state = new TalkRoomLifecycleState();
        state.Records.Add(
            NewLifecycleRecord("obsolete-room", false));
        state.Records.Add(
            NewLifecycleRecord("pending-room", true));
        new TalkRoomLifecycleStore(root, profile).Save(state);

        var coordinator = new TalkRoomLifecycleCoordinator(
            root,
            profile,
            NewSettings,
            () => true,
            configuration => "alice",
            (configuration, token, isEventConversation) => { });
        try
        {
            TalkRoomLifecycleState migrated =
                LoadLifecycleState(root, profile);
            Check(
                "Startup drops obsolete calendar tracking records",
                migrated.Records.Count == 1
                && migrated.Records[0].RoomToken == "pending-room"
                && migrated.Records[0].PendingDeletion);
        }
        finally
        {
            coordinator.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static TalkRoomLifecycleRecord NewLifecycleRecord(
        string token,
        bool pendingDeletion)
    {
        return new TalkRoomLifecycleRecord
        {
            RoomToken = token,
            ServerBaseUrl =
                "https://cloud.example.org/nextcloud",
            AccountLogin = "me@example.org",
            PendingDeletion = pendingDeletion,
            PolicyRequired = true,
            NextAttemptUtc = DateTime.UtcNow.AddHours(1)
        };
    }

    private static string NewTestRoot(string name)
    {
        string root = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-" + name + "-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }

    private static NcTalkOutlookAddIn.Settings.AddinSettings
        NewSettings()
    {
        return NewSettings("me@example.org");
    }

    private static NcTalkOutlookAddIn.Settings.AddinSettings
        NewSettings(string username)
    {
        return new NcTalkOutlookAddIn.Settings.AddinSettings
        {
            ServerUrl =
                "https://cloud.example.org/nextcloud",
            Username = username,
            AppPassword = "secret"
        };
    }

    private static TalkServiceConfiguration NewConfiguration()
    {
        NcTalkOutlookAddIn.Settings.AddinSettings settings =
            NewSettings();
        return new TalkServiceConfiguration(
            settings.ServerUrl,
            settings.Username,
            settings.AppPassword);
    }

    private static TalkRoomLifecycleState LoadLifecycleState(
        string root,
        string profile)
    {
        return new TalkRoomLifecycleStore(root, profile).Load();
    }

    private static TalkRoomLifecycleRecord ReadCoordinatorRecord(
        TalkRoomLifecycleCoordinator coordinator,
        string roomToken)
    {
        FieldInfo syncRootField =
            typeof(TalkRoomLifecycleCoordinator).GetField(
                "_syncRoot",
                BindingFlags.NonPublic
                | BindingFlags.Instance);
        FieldInfo stateField =
            typeof(TalkRoomLifecycleCoordinator).GetField(
                "_state",
                BindingFlags.NonPublic
                | BindingFlags.Instance);
        object syncRoot =
            syncRootField.GetValue(coordinator);
        lock (syncRoot)
        {
            TalkRoomLifecycleState state =
                (TalkRoomLifecycleState)stateField.GetValue(
                    coordinator);
            for (int i = 0; i < state.Records.Count; i++)
            {
                TalkRoomLifecycleRecord record =
                    state.Records[i];
                if (record != null
                    && string.Equals(
                        record.RoomToken,
                        roomToken,
                        StringComparison.Ordinal))
                {
                    return new TalkRoomLifecycleRecord
                    {
                        Id = record.Id,
                        RoomToken = record.RoomToken,
                        IsEventConversation =
                            record.IsEventConversation,
                        ServerBaseUrl =
                            record.ServerBaseUrl,
                        AccountLogin = record.AccountLogin,
                        AccountId = record.AccountId,
                        PendingDeletion =
                            record.PendingDeletion,
                        PolicyRequired =
                            record.PolicyRequired,
                        AttemptCount = record.AttemptCount,
                        NextAttemptUtc =
                            record.NextAttemptUtc
                    };
                }
            }
        }
        return null;
    }

    private static bool WaitUntil(Func<bool> condition)
    {
        DateTime deadline =
            DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }
            Thread.Sleep(25);
        }
        return condition();
    }

    private static void TestSyncCoalescing()
    {
        var firstStarted = new ManualResetEventSlim(false);
        var releaseFirst = new ManualResetEventSlim(false);
        var firstCompletionStarted =
            new ManualResetEventSlim(false);
        var releaseFirstCompletion =
            new ManualResetEventSlim(false);
        var completed = new ManualResetEventSlim(false);
        var executed = new List<string>();
        int completionCount = 0;
        var coordinator = new TalkAppointmentSyncCoordinator(
            snapshot =>
            {
                lock (executed)
                {
                    executed.Add(snapshot.RoomName);
                }
                if (snapshot.RoomName == "first")
                {
                    firstStarted.Set();
                    releaseFirst.Wait(TimeSpan.FromSeconds(5));
                }
                return new TalkAppointmentSyncResult(
                    snapshot.RoomToken);
            },
            result =>
            {
                int count =
                    Interlocked.Increment(ref completionCount);
                if (count == 1)
                {
                    firstCompletionStarted.Set();
                    return Task.Run(
                        () => releaseFirstCompletion.Wait(
                            TimeSpan.FromSeconds(5)));
                }
                if (count == 2)
                {
                    completed.Set();
                }
                return Task.FromResult(0);
            });

        coordinator.Queue(NewSnapshot("first"));
        Check(
            "Talk sync worker starts",
            firstStarted.Wait(TimeSpan.FromSeconds(5)));
        coordinator.Queue(NewSnapshot("superseded"));
        coordinator.Queue(NewSnapshot("latest"));
        releaseFirst.Set();
        Check(
            "Talk sync waits for Outlook result persistence",
            firstCompletionStarted.Wait(
                TimeSpan.FromSeconds(5)));
        lock (executed)
        {
            Check(
                "No later remote sync overtakes result persistence",
                executed.Count == 1);
        }
        releaseFirstCompletion.Set();
        Check(
            "Talk sync worker completes coalesced work",
            completed.Wait(TimeSpan.FromSeconds(5)));
        lock (executed)
        {
            Check(
                "Talk sync keeps first in-flight and latest pending snapshot",
                executed.Count == 2
                && executed[0] == "first"
                && executed[1] == "latest",
                string.Join(",", executed.ToArray()));
        }
        coordinator.Dispose();
    }

    private static TalkAppointmentSyncSnapshot NewSnapshot(
        string roomName)
    {
        return new TalkAppointmentSyncSnapshot
        {
            RoomToken = "same-room",
            RoomName = roomName
        };
    }

    private static void TestProtectedStateStoreCompatibility()
    {
        string root = NewTestRoot("protected-state");
        const string profile = " compatibility-profile ";
        string talkPath = Path.Combine(
            root,
            "talk-room-lifecycle-339954de9e466b4233c67981.dat");
        string ifbPath = Path.Combine(
            root,
            "ifb-registry-state-339954de9e466b4233c67981.dat");
        try
        {
            var talkState = new TalkRoomLifecycleState();
            talkState.Records.Add(
                NewLifecycleRecord("compatible-room", true));
            WriteCompatibleProtectedState(
                talkPath,
                talkState,
                "NC4OL::TalkRoomLifecycle::v1");

            var ifbState = new IfbRegistryState();
            ifbState.Ownership.Add(
                new IfbRegistryOwnership
                {
                    RegistryPath = @"Software\NC4OL\Test",
                    ValueName = "FreeBusySupport",
                    WrittenValue = "http://127.0.0.1/test"
                });
            WriteCompatibleProtectedState(
                ifbPath,
                ifbState,
                "NC4OL::IFB::RegistryState::v1");

            TalkRoomLifecycleState loadedTalk =
                new TalkRoomLifecycleStore(root, profile).Load();
            Check(
                "Talk store reads the existing protected state format",
                loadedTalk.Records.Count == 1
                && loadedTalk.Records[0].RoomToken
                    == "compatible-room");
            IfbRegistryState loadedIfb =
                new IfbRegistryStateStore(root, profile).Load();
            Check(
                "IFB store reads the existing protected state format",
                loadedIfb.Ownership.Count == 1
                && loadedIfb.Ownership[0].ValueName
                    == "FreeBusySupport");

            loadedTalk.Records[0].RoomToken = "saved-room";
            new TalkRoomLifecycleStore(root, profile).Save(
                loadedTalk);
            TalkRoomLifecycleState savedTalk =
                ReadCompatibleProtectedState<TalkRoomLifecycleState>(
                    talkPath,
                    "NC4OL::TalkRoomLifecycle::v1");
            Check(
                "Talk store keeps its file name, entropy, and JSON format",
                savedTalk.Records.Count == 1
                && savedTalk.Records[0].RoomToken == "saved-room"
                && !HasUtf8Bom(talkPath));

            loadedIfb.Ownership[0].WrittenValue =
                "http://127.0.0.1/saved";
            new IfbRegistryStateStore(root, profile).Save(loadedIfb);
            IfbRegistryState savedIfb =
                ReadCompatibleProtectedState<IfbRegistryState>(
                    ifbPath,
                    "NC4OL::IFB::RegistryState::v1");
            Check(
                "IFB store keeps its file name, entropy, and JSON format",
                savedIfb.Ownership.Count == 1
                && savedIfb.Ownership[0].WrittenValue.EndsWith(
                    "/saved",
                    StringComparison.Ordinal)
                && !HasUtf8Bom(ifbPath));

            string talkBackupPath = talkPath + ".bak";
            var backupTalkState = new TalkRoomLifecycleState();
            backupTalkState.Records.Add(
                NewLifecycleRecord("backup-room", true));
            WriteCompatibleProtectedState(
                talkBackupPath,
                backupTalkState,
                "NC4OL::TalkRoomLifecycle::v1");
            TalkRoomLifecycleState primaryTalk =
                new TalkRoomLifecycleStore(root, profile).Load();
            Check(
                "Talk store prefers a valid primary over its backup",
                primaryTalk.Records.Count == 1
                && primaryTalk.Records[0].RoomToken == "saved-room");
            File.WriteAllText(talkPath, "unreadable");
            TalkRoomLifecycleState recoveredTalk =
                new TalkRoomLifecycleStore(root, profile).Load();
            Check(
                "Talk store restores a valid backup after primary failure",
                recoveredTalk.Records.Count == 1
                && recoveredTalk.Records[0].RoomToken == "backup-room"
                && File.ReadAllText(talkPath)
                    == File.ReadAllText(talkBackupPath));

            var invalidTalk = new TalkRoomLifecycleState
            {
                Records = null
            };
            WriteCompatibleProtectedState(
                talkPath,
                invalidTalk,
                "NC4OL::TalkRoomLifecycle::v1");
            WriteCompatibleProtectedState(
                talkBackupPath,
                invalidTalk,
                "NC4OL::TalkRoomLifecycle::v1");
            var blockedTalkStore =
                new TalkRoomLifecycleStore(root, profile);
            Check(
                "Talk store rejects structurally invalid state",
                blockedTalkStore.Load().Records.Count == 0);
            bool talkWriteBlocked = false;
            try
            {
                blockedTalkStore.Save(loadedTalk);
            }
            catch (InvalidOperationException ex)
            {
                talkWriteBlocked = ex.Message
                    == "Talk room deletion queue is unreadable; existing data was preserved.";
            }
            Check(
                "Talk store blocks writes after failed recovery",
                talkWriteBlocked);

            string ifbBackupPath = ifbPath + ".bak";
            File.WriteAllText(ifbPath, "unreadable");
            File.WriteAllText(ifbBackupPath, "unreadable");
            var blockedIfbStore =
                new IfbRegistryStateStore(root, profile);
            Check(
                "IFB store returns an empty state after failed recovery",
                blockedIfbStore.Load().Ownership.Count == 0);
            bool ifbWriteBlocked = false;
            try
            {
                blockedIfbStore.Save(loadedIfb);
            }
            catch (InvalidOperationException ex)
            {
                ifbWriteBlocked = ex.Message
                    == "IFB registry state is unreadable; existing data was preserved.";
            }
            Check(
                "IFB store blocks writes after failed recovery",
                ifbWriteBlocked);
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }

    private static void WriteCompatibleProtectedState<TState>(
        string path,
        TState state,
        string entropy)
    {
        var serializer = new JavaScriptSerializer();
        byte[] clearBytes = Encoding.UTF8.GetBytes(
            serializer.Serialize(state));
        byte[] protectedBytes = ProtectedData.Protect(
            clearBytes,
            Encoding.UTF8.GetBytes(entropy),
            DataProtectionScope.CurrentUser);
        File.WriteAllText(
            path,
            Convert.ToBase64String(protectedBytes),
            new UTF8Encoding(false));
    }

    private static TState ReadCompatibleProtectedState<TState>(
        string path,
        string entropy)
    {
        byte[] protectedBytes = Convert.FromBase64String(
            File.ReadAllText(path, Encoding.UTF8));
        byte[] clearBytes = ProtectedData.Unprotect(
            protectedBytes,
            Encoding.UTF8.GetBytes(entropy),
            DataProtectionScope.CurrentUser);
        return new JavaScriptSerializer().Deserialize<TState>(
            Encoding.UTF8.GetString(clearBytes));
    }

    private static bool HasUtf8Bom(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        return bytes.Length >= 3
            && bytes[0] == 0xef
            && bytes[1] == 0xbb
            && bytes[2] == 0xbf;
    }

    private static void TestLegacyIfbRegistryMigration()
    {
        string root = NewTestRoot("ifb-upgrade");
        string outlookVersion =
            "NCConnectorTest" + Guid.NewGuid().ToString("N");
        string registryRoot =
            @"Software\Microsoft\Office\" + outlookVersion;
        string calendarPath =
            registryRoot + @"\Outlook\Options\Calendar";
        string internetPath =
            calendarPath + @"\Internet Free/Busy";
        const string calendarValue = "FreeBusySearchPath";
        const string internetValue = "Read URL";
        const string legacyUrl =
            "http://127.0.0.1:44777/nc-ifb/freebusy/%NAME%@example.org.vfb";
        const string desiredUrl =
            "http://127.0.0.1:44777/nc-ifb/new-token/freebusy/%NAME%@%SERVER%.vfb";
        const string modernUrl =
            "http://127.0.0.1:44777/nc-ifb/other-token/freebusy/%NAME%@%SERVER%.vfb";
        const string externalUrl = "https://freebusy.example.org/path.vfb";

        try
        {
            WriteRegistryValue(calendarPath, calendarValue, legacyUrl);
            WriteRegistryValue(internetPath, internetValue, legacyUrl);
            var legacySettings =
                new NcTalkOutlookAddIn.Settings.AddinSettings
                {
                    IfbPreviousFreeBusyPath = string.Empty
                };
            var legacyManager = new IfbRegistryOwnershipManager(
                root,
                "legacy-upgrade");
            legacyManager.Apply(
                outlookVersion,
                desiredUrl,
                legacySettings);
            Check(
                "Legacy IFB values migrate without a stored predecessor",
                ReadRegistryValue(calendarPath, calendarValue) == desiredUrl
                && ReadRegistryValue(internetPath, internetValue) == desiredUrl);
            legacyManager.Restore();
            Check(
                "Migrated IFB values without a predecessor are removed on restore",
                ReadRegistryValue(calendarPath, calendarValue) == null
                && ReadRegistryValue(internetPath, internetValue) == null);

            WriteRegistryValue(calendarPath, calendarValue, modernUrl);
            WriteRegistryValue(internetPath, internetValue, modernUrl);
            bool rejectedModernValue = false;
            try
            {
                new IfbRegistryOwnershipManager(root, "modern-value").Apply(
                    outlookVersion,
                    desiredUrl,
                    new NcTalkOutlookAddIn.Settings.AddinSettings());
            }
            catch (InvalidOperationException)
            {
                rejectedModernValue = true;
            }
            Check(
                "Unowned tokenized IFB values remain protected",
                rejectedModernValue
                && ReadRegistryValue(calendarPath, calendarValue) == modernUrl
                && ReadRegistryValue(internetPath, internetValue) == modernUrl);

            WriteRegistryValue(calendarPath, calendarValue, legacyUrl);
            WriteRegistryValue(internetPath, internetValue, legacyUrl);
            var externalSettings =
                new NcTalkOutlookAddIn.Settings.AddinSettings
                {
                    IfbPreviousFreeBusyPath = externalUrl
                };
            var externalManager = new IfbRegistryOwnershipManager(
                root,
                "external-predecessor");
            externalManager.Apply(
                outlookVersion,
                desiredUrl,
                externalSettings);
            externalManager.Restore();
            Check(
                "Legacy IFB migration restores its saved external predecessor",
                ReadRegistryValue(calendarPath, calendarValue) == externalUrl
                && ReadRegistryValue(internetPath, internetValue) == externalUrl);
        }
        finally
        {
            Registry.CurrentUser.DeleteSubKeyTree(registryRoot, false);
            Directory.Delete(root, true);
        }
    }

    private static void WriteRegistryValue(
        string path,
        string valueName,
        string value)
    {
        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(path))
        {
            key.SetValue(valueName, value, RegistryValueKind.String);
        }
    }

    private static string ReadRegistryValue(
        string path,
        string valueName)
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, false))
        {
            return key == null
                ? null
                : key.GetValue(
                    valueName,
                    null,
                    RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
        }
    }

    private static void TestDurableReplacement()
    {
        string root = Path.Combine(
            Path.GetTempPath(),
            "nc4ol-durable-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string primary = Path.Combine(root, "state.dat");
            string backup = primary + ".bak";
            string prepared = primary + ".tmp";
            File.WriteAllText(primary, "known-good");
            File.WriteAllText(backup, "older");
            File.WriteAllText(prepared, "new-state");
            DurableFileReplace.CommitPreparedFile(
                prepared,
                primary,
                backup);
            Check(
                "Durable replacement writes the new primary",
                File.ReadAllText(primary) == "new-state");
            Check(
                "Durable replacement preserves the prior primary as backup",
                File.ReadAllText(backup) == "known-good");

            MethodInfo fallback =
                typeof(DurableFileReplace).GetMethod(
                    "CommitWithPreservedCopyFallback",
                    BindingFlags.NonPublic
                    | BindingFlags.Static);
            File.WriteAllText(primary, "fallback-old");
            File.WriteAllText(backup, "fallback-older");
            File.WriteAllText(prepared, "fallback-new");
            string preserved = primary + ".preserved";
            File.WriteAllText(preserved, "fallback-old");
            fallback.Invoke(
                null,
                new object[] { prepared, primary, backup, preserved });
            Check(
                "Copy fallback writes the new primary",
                File.ReadAllText(primary) == "fallback-new");
            Check(
                "Copy fallback preserves the prior primary",
                File.ReadAllText(backup) == "fallback-old");
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }
}
'@
    $participantSyncSource = Get-SourceSlice `
        $appointmentSyncController `
        '        private static void ExecuteParticipantSync(' `
        '        private static void ExecuteDelegationSync('
    $testCode.Replace('// EXECUTE_PARTICIPANT_SYNC_SOURCE', $participantSyncSource) |
        Set-Content -LiteralPath $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path -LiteralPath $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $SourceRoot "Services\DurableFileReplace.cs"),
        (Join-Path $SourceRoot "Services\ProtectedJsonStateStore.cs"),
        (Join-Path $SourceRoot "Services\TalkServiceConfiguration.cs"),
        (Join-Path $SourceRoot "Services\IfbAddressBookCache.cs"),
        (Join-Path $SourceRoot "Services\TalkAppointmentSyncCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleCoordinator.cs"),
        (Join-Path $SourceRoot "Services\TalkRoomLifecycleStore.cs"),
        (Join-Path $SourceRoot "Services\IfbRegistryStateStore.cs"),
        (Join-Path $SourceRoot "Services\IfbRegistryOwnershipManager.cs"),
        (Join-Path $SourceRoot "Models\TalkAppointmentSyncSnapshot.cs"),
        (Join-Path $SourceRoot "Models\NextcloudUser.cs"),
        (Join-Path $SourceRoot "Models\TalkRoomLifecycleRecord.cs"),
        (Join-Path $SourceRoot "Utilities\AppDataPaths.cs"),
        (Join-Path $SourceRoot "Utilities\LogCategories.cs"),
        (Join-Path $SourceRoot "Utilities\NextcloudUriValidator.cs")
    )
    $exe = Join-Path $TempRoot "TalkIfbLifecycleTests.exe"
    & $csc `
        /nologo `
        /target:exe `
        "/out:$exe" `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Security.dll `
        /reference:System.Web.Extensions.dll `
        @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    if (Test-Path -LiteralPath $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
