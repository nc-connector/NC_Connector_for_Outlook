Param(
    [string]$ProjectRoot = ".",
    [string]$BuiltAddInPath = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Read-Source([string]$RelativePath) {
    return [IO.File]::ReadAllText(
        (Join-Path $ProjectRoot $RelativePath))
}

function Assert-True(
    [string]$Name,
    [bool]$Condition
) {
    if (-not $Condition) {
        throw "[FAIL] $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-Contains(
    [string]$Name,
    [string]$Source,
    [string]$Expected
) {
    Assert-True $Name $Source.Contains($Expected)
}

function Assert-NotContains(
    [string]$Name,
    [string]$Source,
    [string]$Unexpected
) {
    Assert-True $Name (-not $Source.Contains($Unexpected))
}

function Assert-Precedes(
    [string]$Name,
    [string]$Source,
    [string]$First,
    [string]$Second
) {
    $firstIndex = $Source.IndexOf($First, [StringComparison]::Ordinal)
    $secondIndex = $Source.IndexOf($Second, [StringComparison]::Ordinal)
    Assert-True $Name (
        $firstIndex -ge 0 `
            -and $secondIndex -gt $firstIndex)
}

function Get-MethodSlice(
    [string]$Source,
    [string]$Signature
) {
    $start = $Source.IndexOf($Signature, [StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "Could not locate method '$Signature'."
    }
    $lineStart = $Source.LastIndexOf([char]10, $start) + 1
    $indent = $Source.Substring($lineStart, $start - $lineStart)
    $closingLine = [string][char]10 + $indent + "}"
    $end = $Source.IndexOf($closingLine, $start, [StringComparison]::Ordinal)
    if ($end -lt 0) {
        throw "Could not isolate method '$Signature'."
    }
    return $Source.Substring($start, $end + $closingLine.Length - $start)
}

$composePath = "src\NcTalkOutlookAddIn\Controllers\SeparatePasswordDeliveryController.cs"
$trackerPath = "src\NcTalkOutlookAddIn\Controllers\ComposeShareCleanupTracker.cs"
$sendPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.Send.cs"
$shareCleanupPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs"
$addinPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.cs"
$subscriptionPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.cs"
$attachmentFlowPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs"
$subscriptionRegistryPath = "src\NcTalkOutlookAddIn\Controllers\MailComposeSubscriptionRegistryController.cs"
$hooksPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.Hooks.cs"
$fileLinkPath = "src\NcTalkOutlookAddIn\Controllers\FileLinkLaunchController.cs"
$fileLinkLaunchOptionsPath = "src\NcTalkOutlookAddIn\Models\FileLinkWizardLaunchOptions.cs"
$fileLinkWizardPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.cs"
$fileLinkWizardFilesPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.Files.cs"
$fileLinkWizardDragDropPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.DragDrop.cs"
$projectPath = "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"

$compose = Read-Source $composePath
$tracker = Read-Source $trackerPath
$send = Read-Source $sendPath
$shareCleanup = Read-Source $shareCleanupPath
$addin = Read-Source $addinPath
$subscription = Read-Source $subscriptionPath
$attachmentFlow = Read-Source $attachmentFlowPath
$attachmentPolicy = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentPolicy.cs"
$attachmentMaterialization = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentMaterialization.cs"
$attachmentQueue = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentQueue.cs"
$subscriptionRegistry = Read-Source $subscriptionRegistryPath
$hooks = Read-Source $hooksPath
$fileLink = Read-Source $fileLinkPath
$fileLinkLaunchOptions = Read-Source $fileLinkLaunchOptionsPath
$fileLinkWizard = Read-Source $fileLinkWizardPath
$fileLinkWizardFiles = Read-Source $fileLinkWizardFilesPath
$fileLinkWizardDragDrop = Read-Source $fileLinkWizardDragDropPath
$project = Read-Source $projectPath
$directPasswordDispatch = Get-MethodSlice `
    $compose `
    "internal void DispatchSeparatePasswordMailQueue("
$passwordMailPopulation = Get-MethodSlice `
    $compose `
    "private List<string> PopulatePasswordMail("
$manualPasswordFallback = Get-MethodSlice `
    $compose `
    "private bool TryOpenSeparatePasswordFallback("
$fileLinkWizardUi = Get-MethodSlice `
    $fileLink `
    "private bool RunFileLinkWizardOnUiThread("

$attachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnAttachmentAdd("
$beforeAttachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnBeforeAttachmentAdd("
$snapshotAttachments = Get-MethodSlice `
    $attachmentMaterialization `
    "private List<AttachmentSnapshot> SnapshotAttachments()"
$automationFileAttachment = Get-MethodSlice `
    $attachmentMaterialization `
    "private bool IsAutomationFileAttachment("
$collectAttachments = Get-MethodSlice `
    $attachmentMaterialization `
    "private void CollectAttachmentSelectionsForShare("
$startAttachmentShareFlow = Get-MethodSlice `
    $attachmentQueue `
    "private async Task StartComposeAttachmentShareFlowAsync("
$prepareAttachmentSelections = Get-MethodSlice `
    $attachmentQueue `
    "private bool PrepareComposeAttachmentSelections("
$removeAttachmentsByIndices = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveAttachmentsByIndices("
$lastAddedBatch = Get-MethodSlice `
    $attachmentMaterialization `
    "private AttachmentBatchInfo BuildLastAddedBatchInfo("
$readAttachmentSettings = Get-MethodSlice `
    $attachmentPolicy `
    "private AttachmentAutomationSettings ReadAttachmentAutomationSettings()"
$attachmentSendGate = Get-MethodSlice `
    $attachmentPolicy `
    "private bool TryValidateAttachmentPolicyBeforeSend("
$attachmentSettingsFreshness = Get-MethodSlice `
    $attachmentPolicy `
    "private bool HasFreshAttachmentAutomationSettingsSnapshot()"
$requiredAttachmentNotice = Get-MethodSlice `
    $attachmentPolicy `
    "private static void ShowRequiredAttachmentRoutingNotice()"
$removeAttachmentOriginals = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveAttachmentOriginals("
$removeLastAttachmentBatch = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveLastAddedAttachmentBatch("
$queueSelectionScan = Get-MethodSlice `
    $fileLinkWizardFiles `
    "private async Task AddSelectionsAsync("
$initialFileSelection = Get-MethodSlice `
    $fileLinkWizardDragDrop `
    "private bool TryAddInitialFileSelection("
$initialQueueSelections = Get-MethodSlice `
    $fileLinkWizardFiles `
    "private void AddInitialSelections("
$reserveQueueSelection = Get-MethodSlice `
    $fileLinkWizardDragDrop `
    "private bool TryReserveSelection("
$composeConstructor = Get-MethodSlice $subscription "internal MailComposeSubscription("
$attachmentEvaluation = Get-MethodSlice $attachmentFlow "private async Task EvaluateAttachmentAutomationAsync()"

Assert-Precedes `
    "Initial attachment evaluation follows compose event registration" `
    $composeConstructor `
    "_events.AttachmentAdd += OnAttachmentAdd;" `
    "ScheduleAttachmentEvaluation();"
Assert-Precedes `
    "Initial attachment evaluation reuses the pending policy refresh" `
    $composeConstructor `
    "BeginAttachmentAutomationSettingsRefresh();" `
    "ScheduleAttachmentEvaluation();"
Assert-Contains `
    "Attachment evaluation returns to the Outlook UI before inspecting attachments" `
    $attachmentEvaluation `
    "EvaluateAttachmentAutomationOnUiAsync(settings)"

Assert-Precedes `
    "Body resources are ignored before post-add batching" `
    $attachmentAdd `
    "IsAutomationFileAttachment(attachment, true)" `
    "_pendingAddedBatch.Add("
Assert-Precedes `
    "Body resources are allowed before automation preflight" `
    $beforeAttachmentAdd `
    "IsAutomationFileAttachment(attachment, true)" `
    "ReadAttachmentAutomationSettings()"
Assert-Precedes `
    "Body resources are excluded from threshold snapshots" `
    $snapshotAttachments `
    "IsAutomationFileAttachment(attachment)" `
    "snapshots.Add("
Assert-Contains `
    "Attachment classification reads PR_ATTACHMENT_HIDDEN" `
    $automationFileAttachment `
    "0x7FFE000B"
Assert-Contains `
    "Attachment classification releases its PropertyAccessor" `
    $automationFileAttachment `
    "ComInteropScope.TryRelease("
Assert-Precedes `
    "Body resources are excluded before FileLink materialization" `
    $collectAttachments `
    "IsAutomationFileAttachment(attachment)" `
    "TryResolveAttachmentLocalPath("
Assert-Precedes `
    "Post-add attachments are finalized only after the sharing result" `
    $startAttachmentShareFlow `
    "accepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);" `
    "await FinalizeAttachmentShareFlowAsync(originals, launchOptions, accepted);"
Assert-NotContains `
    "Queue adoption does not remove original Outlook attachments" `
    $startAttachmentShareFlow `
    'RemoveAttachmentsByIndices('
Assert-Contains `
    "Attachment launch options expose the queue-adoption boundary" `
    $fileLinkLaunchOptions `
    "internal Action OnInitialQueueAdopted { get; set; }"
Assert-NotContains `
    "Attachment positions are not captured before server prefetch" `
    $startAttachmentShareFlow `
    "CollectAttachmentSelectionsForShare("
Assert-Contains `
    "Post-add capture is deferred until the UI handoff" `
    $startAttachmentShareFlow `
    "launchOptions.PrepareInitialSelections = () =>"
Assert-Contains `
    "The UI handoff captures the current attachment collection" `
    $prepareAttachmentSelections `
    "CollectAttachmentSelectionsForShare(selections, originals, tempFiles);"
Assert-NotContains `
    "Attachment capture cannot yield before queue adoption" `
    $prepareAttachmentSelections `
    "await "
Assert-Precedes `
    "Current attachments are captured on the STA before wizard construction" `
    $fileLinkWizardUi `
    "!launchOptions.PrepareInitialSelections()" `
    "new FileLinkWizardForm("
Assert-NotContains `
    "Wizard construction and attachment adoption do not yield" `
    $fileLinkWizardUi `
    "await "
Assert-Contains `
    "The wizard reports its accepted queue size" `
    $fileLinkWizard `
    "internal int QueuedSelectionCount"
Assert-Precedes `
    "The complete attachment queue is checked before adoption" `
    $fileLinkWizardUi `
    "wizard.QueuedSelectionCount" `
    "launchOptions.OnInitialQueueAdopted();"
Assert-Precedes `
    "Attachment ownership transfers before the wizard opens" `
    $fileLinkWizardUi `
    "launchOptions.OnInitialQueueAdopted();" `
    "wizard.ShowDialog()"
Assert-Precedes `
    "Synchronous attachment events refresh an expired settings snapshot" `
    $readAttachmentSettings `
    "HasFreshAttachmentAutomationSettingsSnapshot()" `
    "BeginAttachmentAutomationSettingsRefresh();"
Assert-Contains `
    "The send gate checks the actual availability separately from known rules" `
    $attachmentSendGate `
    "TryGetCurrentBackendPolicyCheck(configuration, out check)"
Assert-NotContains `
    "Snapshot expiry alone does not block the attachment send gate" `
    $attachmentSendGate `
    "HasFreshAttachmentAutomationSettingsSnapshot()"
Assert-NotContains `
    "Pending attachment policy does not create a blanket Send notice" `
    $attachmentSendGate `
    "Strings.AttachmentPolicyPending"
Assert-Contains `
    "Required attachment routing has its own localized notice" `
    $requiredAttachmentNotice `
    "Strings.AttachmentRoutingRequired"
foreach ($misleadingNotice in @(
    "Strings.FileLinkWizardAttachmentModeReasonAlways",
    "Strings.FileLinkWizardUploadFailed"
)) {
    Assert-NotContains `
        ("Attachment send notices do not reuse wizard errors: " + $misleadingNotice) `
        ($attachmentSendGate + $requiredAttachmentNotice) `
        $misleadingNotice
}
Assert-NotContains `
    "Attachment events no longer call the misleading forced-processing error" `
    $attachmentFlow `
    "ShowForcedAttachmentProcessingError"
Assert-Contains `
    "Settings changes invalidate open compose attachment caches" `
    $addin `
    ".RefreshAttachmentAutomationSettings();"
Assert-Contains `
    "The compose registry refreshes every open subscription" `
    $subscriptionRegistry `
    "current[i].RefreshAttachmentAutomationSettings();"
Assert-Contains `
    "Superseded attachment-policy requests cannot replace current settings" `
    $attachmentPolicy `
    "== _attachmentAutomationSettingsRefreshGeneration"
Assert-Contains `
    "The threshold prompt uses the last attachment size" `
    $lastAddedBatch `
    "latestBatchEntry.SizeBytes)"
Assert-NotContains `
    "The threshold prompt does not label a batch total as the last file size" `
    $lastAddedBatch `
    "total +="
Assert-Contains `
    "Local queue snapshots are built off the wizard thread" `
    $queueSelectionScan `
    "FileLinkQueueNode snapshot = await Task.Run("
Assert-Contains `
    "Local queue snapshots use the cancellable scan token" `
    $queueSelectionScan `
    "selection,`r`n                                token)"
Assert-NotContains `
    "Interactive queue scans do not use an uncancellable token" `
    $queueSelectionScan `
    "CancellationToken.None"
Assert-Precedes `
    "Queue snapshots return to the wizard before UI state changes" `
    $queueSelectionScan `
    "FileLinkQueueNode snapshot = await Task.Run(" `
    "AddPreparedSelection(selection, snapshot);"
Assert-NotContains `
    "Queue scans retain the captured WinForms context" `
    $queueSelectionScan `
    "ConfigureAwait(false)"
Assert-Contains `
    "The wizard busy state includes local queue scans" `
    $fileLinkWizard `
    "|| _queueScanInProgress;"
Assert-Contains `
    "Drag and drop awaits the background queue scan" `
    $fileLinkWizardDragDrop `
    "await AddSelectionsAsync(selections);"
foreach ($selectionAdmission in @($queueSelectionScan, $initialQueueSelections)) {
    Assert-Contains `
        "Queue admission uses the model's source-aware identity comparer" `
        $selectionAdmission `
        "FileLinkSelection.IdentityComparer"
    Assert-Contains `
        "Queue admission reserves typed selections" `
        $selectionAdmission `
        "new HashSet<FileLinkSelection>("
}
Assert-Contains `
    "Queue reservation uses the source-aware selection set" `
    $reserveQueueSelection `
    "return existingSelections.Add(selection);"
Assert-Contains `
    "Attachment mode still bypasses source deduplication" `
    $reserveQueueSelection `
    "if (!_attachmentMode && existingSelections != null)"
Assert-Precedes `
    "Only individual attachment files use synchronous initial capture" `
    $initialFileSelection `
    "!= FileLinkSelectionType.File" `
    "CancellationToken.None"
Assert-Precedes `
    "Successful sharing preserves body resources" `
    $removeAttachmentOriginals `
    "IsAutomationFileAttachment(original.Attachment)" `
    "original.Attachment.Delete();"
Assert-Contains `
    "Successful sharing only removes files present in the completed upload plan" `
    $removeAttachmentOriginals `
    "shared.Contains(new FileLinkSelection(FileLinkSelectionType.File, original.LocalPath))"
Assert-Precedes `
    "Completed local paths are published only after body insertion" `
    $fileLinkWizardUi `
    "if (!inserted)" `
    "wizard.GetSharedLocalPaths()"
Assert-NotContains `
    "Unsafe filename-based removal is no longer present" `
    $attachmentMaterialization `
    "RemoveSuppressedBeforeAddAttachmentByName"
Assert-Contains `
    "Threshold pre-add sharing waits until the host can add the native file" `
    $beforeAttachmentAdd `
    'QueueBeforeAddAttachmentShareFlow("threshold_preadd"'
Assert-Contains `
    "Last-batch removal starts from the visible attachment snapshot" `
    $removeLastAttachmentBatch `
    "List<AttachmentSnapshot> attachments = SnapshotAttachments();"
Assert-Contains `
    "Last-batch removal uses the filtered attachment indices" `
    $removeLastAttachmentBatch `
    'RemoveAttachmentsByIndices(removeIndices, "remove_last_batch");'
Assert-NotContains `
    "Last-batch removal does not delete the physical collection tail" `
    $removeLastAttachmentBatch `
    "attachments.Remove(attachments.Count)"

Assert-Precedes `
    "Primary Send captures recipients before direct password dispatch" `
    $send `
    "CapturePasswordDispatchRecipients();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-Precedes `
    "Primary Send captures the sender before direct password dispatch" `
    $send `
    "CapturePasswordDispatchSender();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-Precedes `
    "The password queue is consumed before direct dispatch" `
    $send `
    "_passwordDispatchQueue.Clear();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-NotContains `
    "Primary Send no longer waits for Sent-folder confirmation" `
    $send `
    "TryArmPendingPasswordDrafts"
Assert-NotContains `
    "Primary Send does not cancel because password auto-send failed" `
    $send `
    "SharingPasswordMailPrepareFailed"

Assert-Precedes `
    "Final password dispatch is prepared before Outlook mail creation" `
    $directPasswordDispatch `
    "PrepareSeparatePasswordDispatch(" `
    "_owner.OutlookApplication.CreateItem("
Assert-Precedes `
    "Final sender identity is applied before the password body" `
    $passwordMailPopulation `
    "ApplyAndVerifySeparatePasswordSender(" `
    "ApplySeparatePasswordBody("
Assert-Precedes `
    "Password body is complete before backend signature insertion" `
    $passwordMailPopulation `
    "ApplySeparatePasswordBody(" `
    "ApplySeparatePasswordBackendSignature("
Assert-Precedes `
    "Recipients are resolved after final body and signature assembly" `
    $passwordMailPopulation `
    "ApplySeparatePasswordBackendSignature(" `
    "ApplySeparatePasswordRecipientsForSend("
Assert-Precedes `
    "Password mail is fully populated before direct Send" `
    $directPasswordDispatch `
    "PopulatePasswordMail(" `
    "((Outlook._MailItem)passwordMail).Send();"
Assert-NotContains `
    "Direct password delivery does not persist an intermediate draft" `
    $directPasswordDispatch `
    ".Save();"
Assert-Contains `
    "Automatic Send failure offers a fully prepared manual fallback" `
    $directPasswordDispatch `
    "TryOpenSeparatePasswordFallback("
Assert-Contains `
    "Ambiguous submission suppresses a duplicate manual send" `
    $directPasswordDispatch `
    "ReadSubmittedOrAmbiguous(passwordMail)"
Assert-Contains `
    "A definite Send and fallback failure is reported to the user" `
    $directPasswordDispatch `
    "ShowPasswordMailFailure(ex.Message);"
Assert-Contains `
    "Secrets fallback warning also covers a prepared manual fallback" `
    $directPasswordDispatch `
    "(sent > 0 || manual > 0)"
Assert-Contains `
    "Manual Send fallback displays the prepared mail" `
    $manualPasswordFallback `
    "fallback.Display(false);"
Assert-Contains `
    "Build fallback uses normalized To recipients without resolution" `
    $manualPasswordFallback `
    "fallback.To = toRecipients;"
Assert-Contains `
    "Build fallback uses normalized Cc recipients without resolution" `
    $manualPasswordFallback `
    "fallback.CC = ccRecipients;"
Assert-Contains `
    "Build fallback uses normalized Bcc recipients without resolution" `
    $manualPasswordFallback `
    "fallback.BCC = bccRecipients;"
Assert-NotContains `
    "Build fallback does not repeat automatic recipient resolution" `
    $manualPasswordFallback `
    "ApplySeparatePasswordRecipientsForSend("
Assert-Precedes `
    "Build fallback reconciles the managed signature after display" `
    $manualPasswordFallback `
    "fallback.Display(false);" `
    "ApplySeparatePasswordBackendSignatureToDisplayedFallback("
Assert-Contains `
    "Unexpected password dispatch failures do not escape the primary Send callback" `
    $send `
    "The primary send continues."
Assert-True `
    "Pending Sent-folder controller is removed" `
    (-not (Test-Path -LiteralPath (
        Join-Path $ProjectRoot `
            "src\NcTalkOutlookAddIn\Controllers\PendingPasswordDraftController.cs")))
Assert-NotContains `
    "Project no longer compiles the pending Sent-folder controller" `
    $project `
    "PendingPasswordDraftController.cs"

$insertedIndex = $fileLink.IndexOf(
    "bool inserted =",
    [StringComparison]::Ordinal)
$insertFailureIndex = $fileLink.IndexOf(
    "if (!inserted)",
    $insertedIndex,
    [StringComparison]::Ordinal)
$armIndex = $fileLink.IndexOf(
    "composeSubscription.ArmShareCleanup(",
    $insertFailureIndex,
    [StringComparison]::Ordinal)
$passwordRegistrationIndex = $fileLink.IndexOf(
    "if (registerSeparatePassword)",
    $armIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "Successful FileLink insertion arms compose cleanup before follow-up handling" `
    ($insertedIndex -ge 0 `
        -and $insertFailureIndex -gt $insertedIndex `
        -and $armIndex -gt $insertFailureIndex `
        -and $passwordRegistrationIndex -gt $armIndex)

Assert-Contains `
    "Compose cleanup arms the focused tracker" `
    $shareCleanup `
    "_shareCleanupTracker.Arm(record)"
Assert-Contains `
    "Compose cleanup tracker exposes ReleaseAll" `
    $tracker `
    "internal int ReleaseAll()"
Assert-Contains `
    "Compose cleanup tracker exposes Drain" `
    $tracker `
    "internal List<ComposeShareCleanupRecord> Drain()"

$afterWriteIndex = $shareCleanup.IndexOf(
    "private void OnAfterWrite()",
    [StringComparison]::Ordinal)
$releaseIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.ReleaseAll();",
    $afterWriteIndex,
    [StringComparison]::Ordinal)
$unloadIndex = $shareCleanup.IndexOf(
    "private void OnUnload()",
    [StringComparison]::Ordinal)
$inspectorCloseIndex = $shareCleanup.IndexOf(
    "private void OnInspectorClosed()",
    [StringComparison]::Ordinal)
$completionIndex = $shareCleanup.IndexOf(
    "private void CompleteComposeShareCleanup(",
    [StringComparison]::Ordinal)
$drainIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.Drain();",
    $completionIndex,
    [StringComparison]::Ordinal)
$disposeIndex = $shareCleanup.IndexOf(
    "Dispose(detachItemEvents);",
    $drainIndex,
    [StringComparison]::Ordinal)
$cleanupQueueIndex = $shareCleanup.IndexOf(
    "_owner.QueueCreatedShareCleanup(",
    $disposeIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "AfterWrite releases shares that Outlook persisted" `
    ($afterWriteIndex -ge 0 `
        -and $releaseIndex -gt $afterWriteIndex `
        -and $releaseIndex -lt $unloadIndex)
Assert-True `
    "Inspector close and inline unload share one captured-state finalizer" `
    ($unloadIndex -ge 0 `
        -and $inspectorCloseIndex -gt $unloadIndex `
        -and $completionIndex -gt $inspectorCloseIndex `
        -and $drainIndex -gt $completionIndex `
        -and $disposeIndex -gt $drainIndex `
        -and $cleanupQueueIndex -gt $disposeIndex)
Assert-NotContains `
    "Share cleanup terminal handlers do not inspect the MailItem" `
    $shareCleanup `
    "_mail"

Assert-Contains `
    "Compose subscription hooks AfterWrite" `
    $subscription `
    "_events.AfterWrite += OnAfterWrite;"
Assert-Contains `
    "Compose subscription hooks Unload" `
    $subscription `
    "_events.Unload += OnUnload;"
Assert-Contains `
    "Compose subscription unhooks AfterWrite" `
    $subscription `
    "_events.AfterWrite -= OnAfterWrite;"
Assert-Contains `
    "Compose subscription unhooks Unload" `
    $subscription `
    "_events.Unload -= OnUnload;"
Assert-Contains `
    "Compose subscription hooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close += OnInspectorClosed;"
Assert-Contains `
    "Compose subscription unhooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close -= OnInspectorClosed;"
Assert-Contains `
    "A concrete replacement Inspector rebinds the lifecycle sink" `
    $subscription `
    "ComInteropScope.AreSameObject("
Assert-Contains `
    "Fallback binding does not replace an existing Inspector sink" `
    $subscription `
    "(_inspectorEvents != null && inspector == null)"
Assert-Contains `
    "NewInspector passes the concrete Inspector to the compose subscription" `
    ($hooks.Replace("`r`n", "`n")) `
    "null,`n                        inspector);"

$composeCleanupSources = $subscription + "`n" + $send + "`n" + $shareCleanup
foreach ($obsoleteClosePath in @(
    "_events.Close += OnClose;",
    "ScheduleSurfaceCloseVerification",
    "OnCleanupGraceTimerTick",
    "IsMailComposeSurfaceOpen",
    "_cleanupGraceTimer"
)) {
    Assert-NotContains `
        ("Compose cleanup does not use close polling: " + $obsoleteClosePath) `
        $composeCleanupSources `
        $obsoleteClosePath
}
Assert-NotContains `
    "Compose cleanup does not poll Inspector state" `
    $shareCleanup `
    "IsMailComposeSurfaceOpen"

Assert-NotContains `
    "FileLink launch no longer blocks on current-user lookup" `
    $fileLink `
    "currentUserIdTask"
Assert-Contains `
    "Lifecycle origin remains attached to compose share state" `
    $fileLink `
    "ComposeLifecycleOrigin.Create("

$deleteMethod = Read-Source "src\NcTalkOutlookAddIn\Services\ComposeShareCleanupService.cs"
Assert-Contains `
    "Compose cleanup method is present" `
    $deleteMethod `
    "internal bool QueueCleanup("
Assert-NotContains `
    "Password delivery does not own remote cleanup" `
    $compose `
    "TryDeleteComposeShareFolder"
Assert-NotContains `
    "Remote cleanup has no Outlook COM dependency" `
    $deleteMethod `
    "Microsoft.Office.Interop"
Assert-Contains `
    "Compose cleanup persists the captured canonical account" `
    $deleteMethod `
    "entry.Origin.AccountId"
Assert-Contains `
    "Compose cleanup uses the current verified account credentials" `
    $deleteMethod `
    "new FileLinkService(identity.Configuration)"
Assert-NotContains `
    "Compose cleanup never falls back to current settings" `
    $deleteMethod `
    "_owner.CurrentSettings"

$removed = @(
    "Models\ComposeLifecycleRecord.cs",
    "Controllers\ComposeLifecycleCoordinator.cs",
    "Controllers\ComposeLifecycleCoordinator.FolderLookup.cs",
    "Controllers\ComposeLifecycleCoordinator.Marker.cs",
    "Controllers\ComposeLifecycleCoordinator.Processing.cs",
    "Controllers\ComposeLifecycleCoordinator.Recovery.cs",
    "Controllers\ComposeLifecycleCoordinator.SentItems.cs",
    "Controllers\SeparatePasswordDispatchController.cs",
    "Controllers\SeparatePasswordDispatchController.Preparation.cs",
    "Controllers\SeparatePasswordDispatchController.Recipients.cs",
    "Controllers\SeparatePasswordDispatchController.Sender.cs",
    "Controllers\SeparatePasswordDispatchController.Signature.cs",
    "Services\ComposeLifecycleJournal.cs"
)
foreach ($relative in $removed) {
    Assert-True `
        ("Legacy source removed: " + $relative) `
        (-not (Test-Path (
            Join-Path `
                $ProjectRoot `
                ("src\NcTalkOutlookAddIn\" + $relative))))
    Assert-NotContains `
        ("Legacy project include removed: " + $relative) `
        $project `
        $relative
}
Assert-Contains `
    "Project includes the compact cleanup record" `
    $project `
    'Models\ComposeShareCleanupRecord.cs'
Assert-Contains `
    "Project includes the compose cleanup tracker" `
    $project `
    'Controllers\ComposeShareCleanupTracker.cs'
Assert-Contains `
    "Project includes the compose cleanup subscription partial" `
    $project `
    'NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs'

$attachmentHandoffHarness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Models
{
    internal enum FileLinkSelectionType { File }
    internal sealed class FileLinkSelection
    {
        internal string LocalPath;
        internal FileLinkSelection(FileLinkSelectionType type, string path) { LocalPath = path; }
        internal static readonly IEqualityComparer<FileLinkSelection> IdentityComparer = new SelectionComparer();
        private sealed class SelectionComparer : IEqualityComparer<FileLinkSelection>
        {
            public bool Equals(FileLinkSelection left, FileLinkSelection right)
            { return string.Equals(left.LocalPath, right.LocalPath, StringComparison.OrdinalIgnoreCase); }
            public int GetHashCode(FileLinkSelection item)
            { return StringComparer.OrdinalIgnoreCase.GetHashCode(item.LocalPath); }
        }
    }
    __LAUNCH_OPTIONS__
}
namespace Outlook
{
    internal sealed class Attachment
    {
        internal string Name, LocalPath;
        internal long SizeBytes;
        internal bool Hidden, Deleted;
        internal int Releases;
        internal Attachments Owner;
        internal void Delete() { Deleted = true; Owner.Items.Remove(this); }
    }
    internal sealed class Attachments
    {
        internal readonly List<Attachment> Items = new List<Attachment>();
        internal int Count { get { return Items.Count; } }
        internal Attachment this[int index] { get { return Items[index - 1]; } }
        internal Attachment Add(string name)
        {
            var item = new Attachment { Name = name, LocalPath = "temp/" + name, SizeBytes = 1, Owner = this };
            Items.Add(item);
            return item;
        }
    }
}
public static class AttachmentHandoffRegression
{
    private static class Strings
    {
        internal const string FileLinkWizardAttachmentModeReasonAlways = "always";
        internal const string FileLinkWizardAttachmentModeReasonThreshold = "threshold: {0} > {1}; last: {2} ({3})";
        internal const string AttachmentPromptLastUnknown = "unknown";
    }
    private static class SizeFormatting
    {
        __FORMAT_MEGABYTES__
    }
    private sealed class Wizard
    {
        private readonly FileLinkWizardLaunchOptions _launchOptions;
        internal Wizard(FileLinkWizardLaunchOptions options) { _launchOptions = options; }
        internal string InfoText() { return BuildAttachmentModeInfoText(); }
        __ATTACHMENT_MODE_INFO__
    }
    private static class File { internal static bool Exists(string path) { return !string.IsNullOrEmpty(path); } }
    private static class OutlookAttachmentAutomationGuardService { internal sealed class GuardState { } }
    private static class DiagnosticsLogger { internal static void LogException(string category, string message, Exception ex) { } }
    private static class LogCategories { internal const string FileLink = "FILELINK"; }
    private static class ComInteropScope
    {
        internal static bool AreSameObject(object left, object right, string category, string leftName, string rightName)
        { return object.ReferenceEquals(left, right); }
        internal static void TryRelease(object value, string category, string message)
        {
            var attachment = value as Outlook.Attachment;
            if (attachment != null) { attachment.Releases++; }
        }
    }
    private sealed class AttachmentBatchInfo { internal string Name; internal long SizeBytes; }
    private sealed class AddinSettings { internal string ServerUrl, Username, AppPassword; }
    private sealed class TalkServiceConfiguration
    { internal TalkServiceConfiguration(string url, string username, string password) { } }
    private sealed class Mail { internal readonly Outlook.Attachments Attachments = new Outlook.Attachments(); }
    private sealed class Owner
    {
        internal readonly TaskCompletionSource<bool> Prefetch = new TaskCompletionSource<bool>();
        internal readonly List<string> Queue = new List<string>();
        internal AddinSettings _currentSettings = new AddinSettings();
        internal bool RejectQueue, Accept, ThrowAfterCapture;
        internal int Invalidations, UiDispatches;
        internal FileLinkWizardLaunchOptions LastLaunchOptions;
        internal string LastInfoText;
        internal Action<Mail, FileLinkWizardLaunchOptions> WizardAction;
        internal bool TryGetAttachmentAutomationGuardState(string stage, string key, out OutlookAttachmentAutomationGuardService.GuardState state)
        { state = null; return false; }
        internal void InvalidateCurrentBackendPolicyCheck(TalkServiceConfiguration configuration) { Invalidations++; }
        internal Task<T> RunOnOutlookUiThreadAsync<T>(Func<T> action)
        { UiDispatches++; return Task.FromResult(action()); }
        internal async Task<bool> RunFileLinkWizardForMailAsync(Mail mail, FileLinkWizardLaunchOptions options)
        {
            await Prefetch.Task;
            if (options.PrepareInitialSelections != null && !options.PrepareInitialSelections()) { return false; }
            if (RejectQueue) { return false; }
            foreach (FileLinkSelection item in options.InitialSelections) { Queue.Add(item.LocalPath); }
            if (options.OnInitialQueueAdopted != null) { options.OnInitialQueueAdopted(); }
            LastLaunchOptions = options;
            LastInfoText = new Wizard(options).InfoText();
            if (ThrowAfterCapture) { throw new InvalidOperationException("upload failed"); }
            if (WizardAction != null) { WizardAction(mail, options); }
            return Accept;
        }
    }
    private sealed class Subscription
    {
        __ORIGINAL_CLASS__
        __BATCH_CLASS__
        __QUEUE_CLASS__
        internal readonly Mail _mail = new Mail();
        internal readonly Owner _owner = new Owner();
        internal bool _disposed;
        internal bool _sendAccepted;
        private bool CanApplyComposeChanges { get { return !_disposed && !_sendAccepted; } }
        private readonly string _composeKey = "test";
        private bool _attachmentSuppressed, _beforeAddShareFlowRunning;
        private readonly List<BeforeAddShareEntry> _pendingBeforeAddShareEntries = new List<BeforeAddShareEntry>();
        private readonly List<string> _pendingAddedBatch = new List<string>();
        internal int Captures, CleanupCalls, SignatureSchedules;
        internal Task Start() { return StartComposeAttachmentShareFlowAsync("threshold", 12, 1, null); }
        internal Task Start(string trigger, long totalBytes, int thresholdMb)
        { return StartComposeAttachmentShareFlowAsync(trigger, totalBytes, thresholdMb, null); }
        internal Task StartQueued(Outlook.Attachment item)
        {
            QueueNotice(item, "always_preadd", 1);
            return StartQueued();
        }
        internal void QueueNotice(Outlook.Attachment item, string trigger, int thresholdMb)
        {
            _pendingBeforeAddShareEntries.Add(new BeforeAddShareEntry
            {
                Candidate = new AttachmentBatchEntry { OriginalAttachment = item, Name = item.Name, SizeBytes = item.SizeBytes },
                LocalPath = item.LocalPath, ThresholdMb = thresholdMb, Trigger = trigger
            });
        }
        internal Task StartQueued() { return RunQueuedBeforeAddAttachmentShareFlowAsync(); }
        internal bool FlowRunning { get { return _beforeAddShareFlowRunning || _attachmentSuppressed; } }
        private void LogFileLink(string message) { }
        private void EndAttachmentSuppression(string reason) { _attachmentSuppressed = false; }
        private void CleanupTemporaryFiles(List<string> files) { CleanupCalls++; }
        private void RestartBeforeAddShareTimerIfNeeded() { }
        private void BeginAttachmentAutomationSettingsRefresh() { }
        private void ScheduleEmailSignatureApplication(string reason) { SignatureSchedules++; }
        private bool IsAutomationFileAttachment(Outlook.Attachment item, bool deferBodyReferences = false) { return item != null && !item.Hidden; }
        private static string ReadAttachmentName(Outlook.Attachment item) { return item.Name; }
        private bool TryResolveAttachmentLocalPath(Outlook.Attachment item, string name, List<string> files, out string path)
        { path = item.LocalPath; return !string.IsNullOrEmpty(path); }
        __START_FLOW__
        __PREPARE_SELECTIONS__
        __QUEUED_FLOW__
        __FINALIZE_FLOW__
        __COLLECT__
        __CAPTURE_BEFORE_ADD__
        __REMOVE_SHARED__
        __RELEASE_ORIGINALS__
    }
    private static void Check(bool condition, string message)
    { if (!condition) { throw new InvalidOperationException(message); } }
    private static void CheckQueuedNotice(string[] triggers, long[] sizes, string expectedTrigger, bool thresholdNotice)
    {
        var compose = new Subscription();
        long totalBytes = 0;
        for (int index = 0; index < triggers.Length; index++)
        {
            var attachment = compose._mail.Attachments.Add("notice-" + index + ".pdf");
            attachment.SizeBytes = sizes[index];
            totalBytes += sizes[index];
            compose.QueueNotice(attachment, triggers[index], 20);
        }
        Task flow = compose.StartQueued();
        compose._owner.Prefetch.SetResult(true);
        flow.GetAwaiter().GetResult();
        FileLinkWizardLaunchOptions options = compose._owner.LastLaunchOptions;
        Check(options != null && options.AttachmentTrigger == expectedTrigger,
            "The before-add batch lost its actual routing reason: " + string.Join(",", triggers));
        Check(options.AttachmentTotalBytes == totalBytes && options.AttachmentThresholdMb == 20,
            "The before-add batch changed its size or threshold context.");
        string expectedNotice = thresholdNotice
            ? "threshold: " + SizeFormatting.FormatMegabytes(totalBytes) + " > 20 MB; last: notice-"
                + (sizes.Length - 1) + ".pdf (" + SizeFormatting.FormatMegabytes(sizes[sizes.Length - 1]) + ")"
            : Strings.FileLinkWizardAttachmentModeReasonAlways;
        Check(compose._owner.LastInfoText == expectedNotice,
            "The queued routing reason produced the wrong wizard notice: " + string.Join(",", triggers));
    }
    private static void CheckAttachmentModeNotices()
    {
        CultureInfo previousCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            const long megabyte = 1024L * 1024L;
            long smallBytes = 21L * megabyte / 10L;
            var direct = new Subscription();
            direct._mail.Attachments.Add("direct.pdf").SizeBytes = smallBytes;
            Task flow = direct.Start("always", smallBytes, 20);
            direct._owner.Prefetch.SetResult(true);
            flow.GetAwaiter().GetResult();
            Check(direct._owner.LastLaunchOptions.AttachmentTrigger == "always"
                && direct._owner.LastInfoText == Strings.FileLinkWizardAttachmentModeReasonAlways,
                "Direct always-sharing claimed that 2.1 MB exceeds 20 MB.");
            Check(new Wizard(null).InfoText() == Strings.FileLinkWizardAttachmentModeReasonAlways,
                "A missing launch context invented a threshold reason.");
            foreach (string trigger in new[] { "always", "always_preadd", "threshold_preadd", "preadd_batch", "unknown", "", null })
            {
                var options = new FileLinkWizardLaunchOptions
                { AttachmentTrigger = trigger, AttachmentTotalBytes = 21L * megabyte, AttachmentThresholdMb = 20 };
                Check(new Wizard(options).InfoText() == Strings.FileLinkWizardAttachmentModeReasonAlways,
                    "A non-threshold trigger invented a threshold reason: " + trigger);
            }
            foreach (long totalBytes in new[] { 0L, smallBytes, 20L * megabyte - 1L, 20L * megabyte })
            {
                var options = new FileLinkWizardLaunchOptions
                { AttachmentTrigger = "threshold", AttachmentTotalBytes = totalBytes, AttachmentThresholdMb = 20 };
                Check(new Wizard(options).InfoText() == Strings.FileLinkWizardAttachmentModeReasonAlways,
                    "A threshold notice claimed an unexceeded limit: " + totalBytes);
            }
            var exceeded = new FileLinkWizardLaunchOptions
            {
                AttachmentTrigger = "ThReShOlD", AttachmentTotalBytes = 21L * megabyte, AttachmentThresholdMb = 20,
                AttachmentLastName = " last.pdf ", AttachmentLastSizeBytes = smallBytes
            };
            Check(new Wizard(exceeded).InfoText() == "threshold: 21.0 MB > 20 MB; last: last.pdf (2.1 MB)",
                "An actual threshold crossing lost its total, limit, name or last-file size.");
            foreach (long totalBytes in new[] { megabyte - 1L, megabyte })
            {
                var options = new FileLinkWizardLaunchOptions
                { AttachmentTrigger = "threshold", AttachmentTotalBytes = totalBytes, AttachmentThresholdMb = 0 };
                Check(new Wizard(options).InfoText() == Strings.FileLinkWizardAttachmentModeReasonAlways,
                    "The minimum effective threshold was not applied before choosing the notice.");
            }
            CheckQueuedNotice(new[] { "always_preadd" }, new[] { smallBytes }, "always", false);
            CheckQueuedNotice(new[] { "threshold_preadd" }, new[] { 21L * megabyte }, "threshold", true);
            CheckQueuedNotice(new[] { "threshold_preadd", "THRESHOLD_PREADD" }, new[] { 11L * megabyte, 10L * megabyte }, "threshold", true);
            CheckQueuedNotice(new[] { "always_preadd", "threshold_preadd" }, new[] { smallBytes, 21L * megabyte }, "always", false);
            CheckQueuedNotice(new[] { "threshold_preadd", "always_preadd" }, new[] { 21L * megabyte, smallBytes }, "always", false);
            CheckQueuedNotice(new[] { "unknown", "threshold_preadd" }, new[] { smallBytes, 21L * megabyte }, "always", false);
            CheckQueuedNotice(new[] { "threshold_preadd" }, new[] { smallBytes }, "threshold", false);
            CheckQueuedNotice(new[] { "threshold_preadd" }, new[] { 20L * megabyte }, "threshold", false);
        }
        finally { CultureInfo.CurrentCulture = previousCulture; }
    }
    public static void Run()
    {
        CheckAttachmentModeNotices();
        foreach (string scenario in new[] { "cancel", "cancel-after-error", "cancel-sent", "success", "subset", "queue-rejected", "prefetch-failed", "upload-failed", "closed", "removed", "sent" })
        {
            var compose = new Subscription();
            Outlook.Attachment first = compose._mail.Attachments.Add("A.pdf");
            Outlook.Attachment second = compose._mail.Attachments.Add("B.pdf");
            compose._owner.Accept = scenario == "success" || scenario == "subset" || scenario == "sent";
            compose._owner.RejectQueue = scenario == "queue-rejected";
            compose._owner.ThrowAfterCapture = scenario == "upload-failed";
            Outlook.Attachment later = null;
            compose._owner.WizardAction = (mail, options) =>
            {
                if (scenario == "sent" || scenario == "cancel-sent") { compose._sendAccepted = true; }
                options.CancelledByUser = scenario == "cancel" || scenario == "cancel-after-error" || scenario == "cancel-sent";
                options.UnexpectedFailureObserved = scenario == "cancel-after-error";
                later = mail.Attachments.Add("A.pdf"); // Same name must never transfer ownership.
                foreach (FileLinkSelection item in options.InitialSelections)
                {
                    if (scenario != "subset" || item.LocalPath == second.LocalPath) { options.SharedLocalPaths.Add(item.LocalPath); }
                }
            };
            Task flow = compose.Start();
            Check(compose._owner.Queue.Count == 0, "Capture occurred before prefetch: " + scenario);
            if (scenario == "closed") { compose._disposed = true; }
            if (scenario == "removed") { first.Delete(); second.Delete(); }
            if (scenario == "prefetch-failed") { compose._owner.Prefetch.SetException(new InvalidOperationException("prefetch")); }
            else { compose._owner.Prefetch.SetResult(true); }
            try { flow.GetAwaiter().GetResult(); }
            catch (InvalidOperationException)
            { Check(scenario == "prefetch-failed" || scenario == "upload-failed", "Unexpected flow error: " + scenario); }
            if (scenario == "success") { Check(first.Deleted && second.Deleted && !later.Deleted, "Successful sharing did not remove only exact originals."); }
            else if (scenario == "cancel" || scenario == "cancel-after-error")
            { Check(first.Deleted && second.Deleted && !later.Deleted, "Cancellation did not discard only the adopted original attachments: " + scenario); }
            else if (scenario == "subset") { Check(!first.Deleted && second.Deleted && !later.Deleted, "Removing a file from the wizard removed its native original."); }
            else if (scenario != "removed") { Check(!first.Deleted && !second.Deleted, "Unsuccessful sharing deleted native attachments: " + scenario); }
            Check(!compose.FlowRunning && compose.CleanupCalls == 1, "Flow state or temporary-file cleanup leaked: " + scenario);
            Check(compose._owner.UiDispatches == 1, "Original cleanup did not marshal to the Outlook UI: " + scenario);
            Check(compose._owner.Invalidations == (scenario == "upload-failed" || scenario == "prefetch-failed" || scenario == "cancel-after-error" ? 1 : 0),
                "User cancellation was confused with service failure: " + scenario);
        }

        foreach (string outcome in new[] { "success", "cancel", "failure" })
        {
            var compose = new Subscription();
            var original = compose._mail.Attachments.Add("same.pdf");
            var unrelated = compose._mail.Attachments.Add("same.pdf");
            var hidden = compose._mail.Attachments.Add("logo.png");
            hidden.Hidden = true;
            compose._owner.Accept = outcome == "success";
            compose._owner.WizardAction = (mail, options) =>
            {
                options.CancelledByUser = outcome == "cancel";
                if (outcome == "success") { options.SharedLocalPaths.Add(original.LocalPath); }
            };
            Task flow = compose.StartQueued(original);
            compose._owner.Prefetch.SetResult(true);
            flow.GetAwaiter().GetResult();
            Check(original.Deleted == (outcome != "failure") && !unrelated.Deleted && !hidden.Deleted,
                "Before-add sharing used filenames instead of the event attachment identity.");
            Check(!compose.FlowRunning, "Queued attachment flow did not release state.");
            Check(original.Releases == 1 && unrelated.Releases == 0, "Owned COM references leaked or borrowed references were released.");
        }

        var changing = new Subscription();
        var old = changing._mail.Attachments.Add("old.pdf");
        Task changingFlow = changing.Start();
        old.Delete();
        var current = changing._mail.Attachments.Add("current.pdf");
        changing._owner.Prefetch.SetResult(true);
        changingFlow.GetAwaiter().GetResult();
        Check(changing._owner.Queue.Count == 1 && changing._owner.Queue[0] == current.LocalPath && !current.Deleted,
            "Deferred capture used a stale collection position.");

        var resources = new Subscription();
        var logo = resources._mail.Attachments.Add("logo.png");
        logo.Hidden = true;
        var visible = resources._mail.Attachments.Add("doc.pdf");
        resources._owner.Accept = true;
        resources._owner.WizardAction = (mail, options) =>
        { foreach (FileLinkSelection item in options.InitialSelections) { options.SharedLocalPaths.Add(item.LocalPath); } };
        Task resourceFlow = resources.Start();
        resources._owner.Prefetch.SetResult(true);
        resourceFlow.GetAwaiter().GetResult();
        Check(!logo.Deleted && visible.Deleted && resources._owner.Queue.Count == 1, "Hidden body resources entered sharing.");
    }
}
'@
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__LAUNCH_OPTIONS__", (Get-MethodSlice $fileLinkLaunchOptions "internal sealed class FileLinkWizardLaunchOptions"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__ORIGINAL_CLASS__", (Get-MethodSlice $attachmentMaterialization "private sealed class AttachmentShareOriginal"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__BATCH_CLASS__", (Get-MethodSlice $attachmentMaterialization "private sealed class AttachmentBatchEntry"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__QUEUE_CLASS__", (Get-MethodSlice $attachmentQueue "private sealed class BeforeAddShareEntry"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__START_FLOW__", $startAttachmentShareFlow)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__PREPARE_SELECTIONS__", $prepareAttachmentSelections)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__QUEUED_FLOW__", (Get-MethodSlice $attachmentQueue "private async Task RunQueuedBeforeAddAttachmentShareFlowAsync()"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__FINALIZE_FLOW__", (Get-MethodSlice $attachmentQueue "private async Task FinalizeAttachmentShareFlowAsync("))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__COLLECT__", $collectAttachments)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__CAPTURE_BEFORE_ADD__", (Get-MethodSlice $attachmentMaterialization "private void CaptureBeforeAddAttachmentOriginal("))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__REMOVE_SHARED__", $removeAttachmentOriginals)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__RELEASE_ORIGINALS__", (Get-MethodSlice $attachmentMaterialization "private static void ReleaseAttachmentShareOriginals("))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__ATTACHMENT_MODE_INFO__", (Get-MethodSlice $fileLinkWizard "private string BuildAttachmentModeInfoText()"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__FORMAT_MEGABYTES__", (Get-MethodSlice (Read-Source "src\NcTalkOutlookAddIn\Utilities\SizeFormatting.cs") "internal static string FormatMegabytes("))
Add-Type -TypeDefinition $attachmentHandoffHarness -Language CSharp -IgnoreWarnings -WarningAction SilentlyContinue
[AttachmentHandoffRegression]::Run()
Write-Host "[OK] Exact attachment ownership, successful subset sharing, cancellation, changed collection, hidden resources, queue rejection and failed prefetch/upload"
Write-Host "[OK] Attachment wizard notices: direct/pre-add always, preserved threshold reason, exact boundary, unknown triggers and mixed batches"

$wizardClosingHarness = @'
using System;
using System.Threading;

public static class AttachmentWizardClosingRegression
{
    private enum CloseReason { UserClosing, WindowsShutDown, FormOwnerClosing }
    private sealed class FormClosingEventArgs
    {
        internal CloseReason CloseReason;
        internal bool Cancel;
    }
    private class BaseForm
    {
        protected virtual void OnFormClosing(FormClosingEventArgs e) { }
    }
    private sealed class Wizard : BaseForm
    {
        internal bool _shareFinalized, _closeAfterCancellation, _closeRequested;
        internal CancellationTokenSource _cancellationSource;
        internal bool CancelledByUser { get; private set; }
        internal bool Close(CloseReason reason)
        {
            var args = new FormClosingEventArgs { CloseReason = reason };
            OnFormClosing(args);
            return !args.Cancel;
        }
        __CLOSING__
    }
    public static void Run()
    {
        var cancelled = new Wizard();
        if (!cancelled.Close(CloseReason.UserClosing) || !cancelled.CancelledByUser)
        { throw new Exception("Cancel/window X did not record the user's cancellation."); }
        foreach (CloseReason reason in new[] { CloseReason.WindowsShutDown, CloseReason.FormOwnerClosing })
        {
            var shutdown = new Wizard();
            if (!shutdown.Close(reason) || shutdown.CancelledByUser)
            { throw new Exception("Host shutdown was mistaken for an attachment cancellation."); }
        }
        var finished = new Wizard { _shareFinalized = true };
        if (!finished.Close(CloseReason.UserClosing) || finished.CancelledByUser)
        { throw new Exception("Successful sharing was mistaken for cancellation."); }
        using (var source = new CancellationTokenSource())
        {
            var uploading = new Wizard { _cancellationSource = source };
            if (uploading.Close(CloseReason.UserClosing) || !uploading.CancelledByUser
                || !source.IsCancellationRequested || !uploading._closeAfterCancellation)
            { throw new Exception("Cancelling an active upload lost the user's intent."); }
            uploading._cancellationSource = null;
            if (!uploading.Close(CloseReason.UserClosing) || !uploading.CancelledByUser)
            { throw new Exception("Deferred close lost the attachment cancellation."); }
        }
        using (var source = new CancellationTokenSource())
        {
            var shutdown = new Wizard { _cancellationSource = source };
            if (shutdown.Close(CloseReason.FormOwnerClosing) || shutdown.CancelledByUser)
            { throw new Exception("Host close during upload discarded attachments."); }
            shutdown._cancellationSource = null;
            if (!shutdown.Close(CloseReason.UserClosing) || shutdown.CancelledByUser)
            { throw new Exception("Deferred host close was reclassified as user cancellation."); }
        }
    }
}
'@
$wizardClosingHarness = $wizardClosingHarness.Replace("__CLOSING__", (Get-MethodSlice (Read-Source "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.Upload.cs") "protected override void OnFormClosing("))
Add-Type -TypeDefinition $wizardClosingHarness -Language CSharp -IgnoreWarnings -WarningAction SilentlyContinue
[AttachmentWizardClosingRegression]::Run()
Assert-Contains "Wizard cancellation is propagated separately from failure" $fileLinkWizardUi "&& wizard.CancelledByUser;"
Assert-Contains "Successful wizard results do not discard unshared originals" $fileLinkWizardUi "launchOptions.CancelledByUser = wizardResult != DialogResult.OK"
Write-Host "[OK] Attachment cancellation distinguishes user close, active upload, successful completion and host shutdown"

# Exercise delayed attachment evaluation and removal with the production methods.
$attachmentSendRaceHarness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;
using Outlook = AttachmentSendRaceBoundary.Outlook;
namespace AttachmentSendRaceBoundary
{
    internal static class Outlook
    {
        internal sealed class Attachment
        {
            internal string Name;
            internal long Size;
            internal bool Hidden, Deleted;
        }
        internal sealed class Attachments
        {
            internal readonly List<Attachment> Items = new List<Attachment>();
            internal Action CountAction, RemoveAction;
            internal int Count { get { if (CountAction != null) { CountAction(); } return Items.Count; } }
            internal Attachment this[int index] { get { return Items[index - 1]; } }
            internal void Remove(int index)
            {
                Items[index - 1].Deleted = true;
                Items.RemoveAt(index - 1);
                if (RemoveAction != null) { RemoveAction(); }
            }
        }
    }
    public static class AttachmentSendRaceRegression
    {
        private static class OutlookAttachmentAutomationGuardService { internal sealed class GuardState { } }
        private static class DiagnosticsLogger { internal static void LogException(string category, string message, Exception ex) { } }
        private static class LogCategories { internal const string FileLink = "FILELINK"; }
        private static class ComInteropScope { internal static void TryRelease(object value, string category, string message) { } }
        private static class SizeFormatting { internal static string FormatMegabytes(long bytes) { return bytes.ToString(); } }
        private static class File { internal static bool Exists(string path) { return !string.IsNullOrEmpty(path); } }
        private static class Strings
        {
            internal const string AttachmentPromptReason = "{0} {1} {2} {3}", AttachmentPromptLastUnknown = "unknown";
        }
        private enum ComposeAttachmentPromptDecision { Share, RemoveLast }
        private static class ComposeAttachmentPromptForm
        {
            internal static int Calls;
            internal static Action PromptAction;
            internal static ComposeAttachmentPromptDecision Decision;
            internal static ComposeAttachmentPromptDecision ShowPrompt(object owner, string reason)
            {
                Calls++;
                if (PromptAction != null) { PromptAction(); }
                return Decision;
            }
        }
        private sealed class AddinSettings { }
        private sealed class Timer
        {
            internal int Interval, Starts;
            internal bool Enabled;
            internal void Start() { Starts++; Enabled = true; }
            internal void Stop() { Enabled = false; }
        }
        private sealed class Mail
        {
            internal readonly Outlook.Attachments Items = new Outlook.Attachments();
            internal int Reads;
            internal Outlook.Attachments Attachments { get { Reads++; return Items; } }
        }
        private sealed class MailInterop
        { internal object TryCreateMailInspectorDialogOwner(Mail mail) { return null; } }
        private sealed class Owner
        {
            internal readonly MailInterop _mailInteropController = new MailInterop();
            internal Action<string> GuardAction;
            internal int UiCalls;
            internal Task<T> RunOnOutlookUiThreadAsync<T>(Func<T> action)
            { UiCalls++; return Task.FromResult(action()); }
            internal bool TryGetAttachmentAutomationGuardState(string stage, string key, out OutlookAttachmentAutomationGuardService.GuardState state)
            {
                state = null;
                if (GuardAction != null) { GuardAction(stage); }
                return false;
            }
        }
        private sealed class Subscription
        {
            __SETTINGS_CLASS__
            __SNAPSHOT_CLASS__
            __BATCH_INFO_CLASS__
            __BATCH_ENTRY_CLASS__
            __SURFACE_STATE__
            internal readonly Mail _mail = new Mail();
            internal readonly Owner _owner = new Owner();
            internal readonly Timer _attachmentEvalTimer = new Timer();
            private readonly Timer _emailSignatureTimer = new Timer();
            internal bool _disposed, _sendAccepted, _attachmentSuppressed;
            private bool CanApplyComposeChanges { get { return !_disposed && !_sendAccepted; } }
            private bool _beforeAddShareFlowRunning, _attachmentPromptOpen;
            private bool _attachmentEvaluationPending, _attachmentEvaluationRunning;
            private ComposeSurfaceState _composeSurfaceState = ComposeSurfaceState.Inspector;
            private string _activeInspectorIdentityKey = "inspector", _inlineExplorerIdentityKey = "";
            private const int EmailSignatureApplyDebounceMs = 900, EmailSignatureInlineApplyDebounceMs = 250;
            private readonly string _composeKey = "race";
            private readonly List<string> _pendingBeforeAddShareEntries = new List<string>();
            private readonly List<AttachmentBatchEntry> _pendingAddedBatch = new List<AttachmentBatchEntry>();
            private readonly TaskCompletionSource<AttachmentAutomationSettings> _settings = new TaskCompletionSource<AttachmentAutomationSettings>();
            private readonly TaskCompletionSource<bool> _sharing = new TaskCompletionSource<bool>();
            internal Action AvailabilityAction;
            internal int Shares, BeforeAddCaptures, PolicyReads;
            internal bool DelayShare;
            internal bool PromptOpen { get { return _attachmentPromptOpen; } }
            internal bool EvaluationPending { get { return _attachmentEvaluationPending; } }
            internal bool EvaluationRunning { get { return _attachmentEvaluationRunning; } }
            internal void RequestEvaluation() { ScheduleAttachmentEvaluation(); }
            internal Task Evaluate() { return EvaluateAttachmentAutomationAsync(); }
            internal Task Tick() { _attachmentEvalTimer.Stop(); return EvaluateAttachmentAutomationAsync(); }
            internal void FinishSharing() { _sharing.SetResult(true); }
            internal void BeforeAddBusy(bool running, bool queued)
            {
                _beforeAddShareFlowRunning = running;
                _pendingBeforeAddShareEntries.Clear();
                if (queued) { _pendingBeforeAddShareEntries.Add("queued"); }
            }
            internal void ReleaseSuppression() { EndAttachmentSuppression("test"); }
            internal void FinishSettings(bool always)
            {
                _settings.SetResult(new AttachmentAutomationSettings
                { AlwaysConnector = always, OfferAboveEnabled = !always, ThresholdMb = 1, ThresholdBytes = 1 });
            }
            internal Outlook.Attachment Add(string name, bool hidden)
            {
                var item = new Outlook.Attachment { Name = name, Size = 2, Hidden = hidden };
                _mail.Items.Items.Add(item);
                return item;
            }
            internal void RemoveLast() { RemoveLastAddedAttachmentBatch(new AttachmentBatchInfo { Count = 1 }); }
            internal void RemoveIndices() { RemoveAttachmentsByIndices(new List<int> { 1, 2 }, "race"); }
            internal void BeforeAdd(Outlook.Attachment item, ref bool cancel) { OnBeforeAttachmentAdd(item, ref cancel); }
            private AttachmentAutomationSettings ReadAttachmentAutomationSettings()
            { return new AttachmentAutomationSettings { OfferAboveEnabled = true, ThresholdMb = 1, ThresholdBytes = 1 }; }
            private Task<AttachmentAutomationSettings> ReadAttachmentAutomationSettingsAsync() { PolicyReads++; return _settings.Task; }
            private bool TryBuildBeforeAddAttachmentCandidate(Outlook.Attachment item, out AttachmentBatchEntry candidate, out string path, out bool temporary)
            {
                BeforeAddCaptures++;
                candidate = new AttachmentBatchEntry { Name = item.Name, SizeBytes = item.Size };
                path = "source";
                temporary = false;
                return true;
            }
            private void QueueBeforeAddAttachmentShareFlow(string trigger, AttachmentBatchEntry candidate, string path, int thresholdMb, bool cleanup)
            { Shares++; }
            private void CleanupTemporaryFiles(List<string> files) { }
            private bool PauseUnavailableAttachmentAutomation(AttachmentAutomationSettings settings, long bytes)
            { if (AvailabilityAction != null) { AvailabilityAction(); } return false; }
            private Task StartComposeAttachmentShareFlowAsync(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            { Shares++; return DelayShare ? _sharing.Task : Task.FromResult(true); }
            private bool ResumeDeferredEmailSignatureApplication(string trigger) { return false; }
            private void ScheduleEmailSignatureApplication(string trigger) { }
            private void LogEmailSignature(string message) { }
            private void LogFileLink(string message) { }
            private bool IsAutomationFileAttachment(Outlook.Attachment item, bool deferBodyReferences = false) { return item != null && !item.Hidden; }
            private static string ReadAttachmentName(Outlook.Attachment item) { return item.Name; }
            private static long ReadAttachmentSizeBytes(Outlook.Attachment item) { return item.Size; }
            __EVALUATE__
            __EVALUATE_UI__
            __SCHEDULE__
            __MARK_INLINE__
            __MARK_INSPECTOR__
            __MARK_CLOSED__
            __END_SUPPRESSION__
            __BEFORE_ADD__
            __SNAPSHOT__
            __SUM__
            __LAST_BATCH__
            __REMOVE_LAST__
            __REMOVE_INDICES__
        }
        private static void Check(bool condition, string message)
        { if (!condition) { throw new InvalidOperationException(message); } }
        private static void ResetPrompt(ComposeAttachmentPromptDecision decision)
        { ComposeAttachmentPromptForm.Calls = 0; ComposeAttachmentPromptForm.PromptAction = null; ComposeAttachmentPromptForm.Decision = decision; }
        private static void CheckInitialAttachmentEvaluation()
        {
            foreach (string surface in new[] { "forward_inspector", "forward_inline", "reopened_draft" })
            foreach (bool always in new[] { false, true })
            {
                ResetPrompt(ComposeAttachmentPromptDecision.Share);
                var compose = new Subscription();
                var inherited = compose.Add("inherited.pdf", false);
                var logo = compose.Add("quoted-logo.png", true);
                if (surface == "forward_inline") { compose.MarkInlineResponse("explorer"); }
                compose.RequestEvaluation();
                Check(compose.EvaluationPending && compose._attachmentEvalTimer.Enabled,
                    "Opening an existing attachment did not schedule evaluation: " + surface);
                Task evaluation = compose.Tick();
                Check(compose.EvaluationRunning && !evaluation.IsCompleted && compose._mail.Reads == 0,
                    "Initial attachment evaluation did not await policy: " + surface);
                compose.FinishSettings(always);
                evaluation.GetAwaiter().GetResult();
                Check(compose.Shares == 1 && ComposeAttachmentPromptForm.Calls == (always ? 0 : 1)
                    && !inherited.Deleted && !logo.Deleted && !compose.EvaluationPending && !compose.EvaluationRunning,
                    "Pre-existing files were not evaluated exactly once without an add event: " + surface + "/" + always);
                Check(compose._owner.UiCalls >= 2, "Initial policy completion did not marshal evaluation and completion to Outlook UI.");
                compose.MarkInlineResponse("explorer");
                compose.MarkInlineResponseClosed("explorer");
                compose.MarkInspector("popout");
                Check(!compose._attachmentEvalTimer.Enabled && !compose.EvaluationPending && compose.Shares == 1,
                    "A handled attachment set reopened its wizard after a surface transition.");
            }

            foreach (bool always in new[] { false, true })
            {
                ResetPrompt(ComposeAttachmentPromptDecision.Share);
                var images = new Subscription();
                var logo = images.Add("quoted-logo.png", true);
                images.RequestEvaluation();
                images.FinishSettings(always);
                images.Tick().GetAwaiter().GetResult();
                Check(images.Shares == 0 && ComposeAttachmentPromptForm.Calls == 0 && !logo.Deleted && !images.EvaluationPending,
                    "An image-only initial compose opened attachment UI.");
            }

            ResetPrompt(ComposeAttachmentPromptDecision.Share);
            var coalesced = new Subscription();
            coalesced.Add("inherited.pdf", false);
            coalesced.MarkInlineResponse("explorer");
            coalesced.RequestEvaluation();
            Task waiting = coalesced.Tick();
            coalesced.RequestEvaluation();
            coalesced.MarkInlineResponseClosed("explorer");
            coalesced.MarkInspector("popout");
            coalesced.Tick().GetAwaiter().GetResult();
            Check(coalesced.PolicyReads == 1 && coalesced.EvaluationRunning && !coalesced._attachmentEvalTimer.Enabled,
                "A surface transition started a second policy evaluation while the first was pending.");
            coalesced.FinishSettings(true);
            waiting.GetAwaiter().GetResult();
            Check(coalesced.Shares == 1 && !coalesced.EvaluationPending && !coalesced._attachmentEvalTimer.Enabled,
                "Pending policy and pop-out produced duplicate sharing.");

            var detached = new Subscription();
            detached.Add("inherited.pdf", false);
            detached.MarkInlineResponse("explorer");
            detached.RequestEvaluation();
            Task detachedWait = detached.Tick();
            detached.MarkInlineResponseClosed("explorer");
            detached.FinishSettings(true);
            detachedWait.GetAwaiter().GetResult();
            Check(detached.EvaluationPending && !detached.EvaluationRunning && !detached._attachmentEvalTimer.Enabled
                && detached._mail.Reads == 0 && detached.Shares == 0,
                "Late initial policy completion touched a detached compose or lost its pending evaluation.");
            detached.MarkInspector("popout");
            Check(detached._attachmentEvalTimer.Enabled, "Pop-out did not resume the pending initial evaluation.");
            detached.Tick().GetAwaiter().GetResult();
            Check(detached.Shares == 1 && !detached.EvaluationPending, "Pop-out did not evaluate the existing attachment once.");

            foreach (string boundary in new[] { "suppression", "preadd_running", "preadd_queued" })
            {
                var busy = new Subscription();
                busy.Add("inherited.pdf", false);
                busy.RequestEvaluation();
                Task busyWait = busy.Tick();
                busy._attachmentSuppressed = boundary == "suppression";
                busy.BeforeAddBusy(boundary == "preadd_running", boundary == "preadd_queued");
                busy.FinishSettings(true);
                busyWait.GetAwaiter().GetResult();
                Check(busy.EvaluationPending && !busy.EvaluationRunning && !busy._attachmentEvalTimer.Enabled
                    && busy._mail.Reads == 0 && busy.Shares == 0,
                    "Initial evaluation ignored a busy attachment boundary: " + boundary);
                busy.BeforeAddBusy(false, false);
                busy.ReleaseSuppression();
                Check(busy._attachmentEvalTimer.Enabled, "Pending evaluation did not resume after " + boundary);
                busy.Tick().GetAwaiter().GetResult();
                Check(busy.Shares == 1 && !busy.EvaluationPending, "Resumed initial evaluation duplicated or lost files: " + boundary);
            }

            var sharing = new Subscription { DelayShare = true };
            sharing.Add("inherited.pdf", false);
            sharing.MarkInlineResponse("explorer");
            sharing.RequestEvaluation();
            sharing.FinishSettings(true);
            Task openWizard = sharing.Tick();
            Check(sharing.Shares == 1 && sharing.EvaluationRunning && !sharing.EvaluationPending && !openWizard.IsCompleted,
                "The sharing boundary was not held open.");
            sharing.MarkInlineResponseClosed("explorer");
            sharing.MarkInspector("popout");
            sharing.Tick().GetAwaiter().GetResult();
            Check(sharing.Shares == 1 && sharing.PolicyReads == 1 && !sharing._attachmentEvalTimer.Enabled,
                "Pop-out opened a second wizard while the first was running.");
            sharing.FinishSharing();
            openWizard.GetAwaiter().GetResult();
            Check(!sharing.EvaluationRunning && !sharing.EvaluationPending && !sharing._attachmentEvalTimer.Enabled,
                "Completed sharing rescheduled an already handled initial attachment set.");

            foreach (string terminal in new[] { "sent", "closed" })
            {
                var ended = new Subscription();
                ended.Add("inherited.pdf", false);
                ended.RequestEvaluation();
                Task endedWait = ended.Tick();
                ended._sendAccepted = terminal == "sent";
                ended._disposed = terminal == "closed";
                ended.FinishSettings(true);
                endedWait.GetAwaiter().GetResult();
                ended.MarkInspector("later");
                Check(ended.Shares == 0 && ended._mail.Reads == 0 && !ended.EvaluationRunning && !ended._attachmentEvalTimer.Enabled,
                    "Late initial evaluation touched or restarted a terminal compose: " + terminal);
            }
        }
        public static void Run()
        {
            CheckInitialAttachmentEvaluation();
            foreach (bool always in new[] { false, true })
            foreach (string transition in new[] { "sent", "closed", "suppressed" })
            {
                ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
                var compose = new Subscription();
                var original = compose.Add("document.pdf", false);
                Task evaluation = compose.Evaluate();
                Check(!evaluation.IsCompleted, "The delayed policy boundary was not exercised.");
                if (transition == "sent") { compose._sendAccepted = true; }
                if (transition == "closed") { compose._disposed = true; }
                if (transition == "suppressed") { compose._attachmentSuppressed = true; }
                compose.FinishSettings(always);
                evaluation.GetAwaiter().GetResult();
                Check(!original.Deleted && compose._mail.Reads == 0 && compose.Shares == 0 && ComposeAttachmentPromptForm.Calls == 0,
                    "Late policy completion touched an unavailable compose item: " + transition + "/" + always);
            }
            foreach (bool always in new[] { false, true })
            {
                ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
                var compose = new Subscription();
                var original = compose.Add("document.pdf", false);
                compose.AvailabilityAction = () => { compose._sendAccepted = true; };
                compose.FinishSettings(always);
                compose.Evaluate().GetAwaiter().GetResult();
                Check(!original.Deleted && compose.Shares == 0 && ComposeAttachmentPromptForm.Calls == 0,
                    "Accepted Send was ignored at the sharing/prompt boundary.");
            }
            foreach (ComposeAttachmentPromptDecision decision in new[] { ComposeAttachmentPromptDecision.RemoveLast, ComposeAttachmentPromptDecision.Share })
            {
                ResetPrompt(decision);
                var compose = new Subscription();
                var original = compose.Add("document.pdf", false);
                ComposeAttachmentPromptForm.PromptAction = () => { compose._sendAccepted = true; };
                compose.FinishSettings(false);
                compose.Evaluate().GetAwaiter().GetResult();
                Check(!original.Deleted && compose.Shares == 0 && !compose.PromptOpen,
                    "A prompt result changed an already accepted Send.");
            }
            ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
            var atRemoval = new Subscription();
            var kept = atRemoval.Add("document.pdf", false);
            atRemoval._owner.GuardAction = stage => { if (stage == "prompt_action") { atRemoval._sendAccepted = true; } };
            atRemoval.FinishSettings(false);
            atRemoval.Evaluate().GetAwaiter().GetResult();
            Check(!kept.Deleted, "RemoveLast did not recheck compose availability at its own boundary.");

            var direct = new Subscription();
            var directFirst = direct.Add("first.pdf", false);
            var directSecond = direct.Add("second.pdf", false);
            direct._sendAccepted = true;
            direct.RemoveLast();
            direct.RemoveIndices();
            Check(!directFirst.Deleted && !directSecond.Deleted && direct._mail.Reads == 0,
                "Direct removal touched an already accepted Send.");

            var duringCount = new Subscription();
            var countFirst = duringCount.Add("first.pdf", false);
            var countSecond = duringCount.Add("second.pdf", false);
            duringCount._mail.Items.CountAction = () => { duringCount._sendAccepted = true; };
            duringCount.RemoveIndices();
            Check(!countFirst.Deleted && !countSecond.Deleted, "Removal did not recheck after an Outlook collection call.");

            var duringRemoval = new Subscription();
            var remaining = duringRemoval.Add("first.pdf", false);
            var removed = duringRemoval.Add("second.pdf", false);
            duringRemoval._mail.Items.RemoveAction = () => { duringRemoval._sendAccepted = true; };
            duringRemoval.RemoveIndices();
            Check(removed.Deleted && !remaining.Deleted, "Further attachments were removed after Send became accepted.");

            ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
            var active = new Subscription();
            var logo = active.Add("logo.png", true);
            var visible = active.Add("document.pdf", false);
            active.FinishSettings(false);
            active.Evaluate().GetAwaiter().GetResult();
            Check(visible.Deleted && !logo.Deleted && ComposeAttachmentPromptForm.Calls == 1 && !active.PromptOpen,
                "The active optional prompt no longer removes the selected visible attachment.");

            ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
            var beforeAccepted = new Subscription();
            var candidate = beforeAccepted.Add("document.pdf", false);
            beforeAccepted._sendAccepted = true;
            bool cancelled = false;
            beforeAccepted.BeforeAdd(candidate, ref cancelled);
            Check(!cancelled && beforeAccepted.BeforeAddCaptures == 0 && ComposeAttachmentPromptForm.Calls == 0,
                "BeforeAttachmentAdd processed an already accepted Send.");

            foreach (ComposeAttachmentPromptDecision decision in new[] { ComposeAttachmentPromptDecision.RemoveLast, ComposeAttachmentPromptDecision.Share })
            {
                ResetPrompt(decision);
                var beforePrompt = new Subscription();
                var incoming = beforePrompt.Add("document.pdf", false);
                ComposeAttachmentPromptForm.PromptAction = () => { beforePrompt._sendAccepted = true; };
                bool cancel = false;
                beforePrompt.BeforeAdd(incoming, ref cancel);
                Check(!cancel && beforePrompt.Shares == 0 && !beforePrompt.PromptOpen,
                    "A before-add prompt result changed an already accepted Send.");
            }
            ResetPrompt(ComposeAttachmentPromptDecision.RemoveLast);
            var beforeBoundary = new Subscription();
            var beforeCandidate = beforeBoundary.Add("document.pdf", false);
            beforeBoundary.AvailabilityAction = () => { beforeBoundary._sendAccepted = true; };
            bool beforeCancel = false;
            beforeBoundary.BeforeAdd(beforeCandidate, ref beforeCancel);
            Check(!beforeCancel && beforeBoundary.Shares == 0 && ComposeAttachmentPromptForm.Calls == 0,
                "BeforeAttachmentAdd ignored accepted Send before opening its prompt.");
        }
    }
}
'@
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SETTINGS_CLASS__", (Get-MethodSlice $attachmentPolicy "private sealed class AttachmentAutomationSettings"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SNAPSHOT_CLASS__", (Get-MethodSlice $attachmentMaterialization "private sealed class AttachmentSnapshot"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__BATCH_INFO_CLASS__", (Get-MethodSlice $attachmentMaterialization "private sealed class AttachmentBatchInfo"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__BATCH_ENTRY_CLASS__", (Get-MethodSlice $attachmentMaterialization "private sealed class AttachmentBatchEntry"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SURFACE_STATE__", (Get-MethodSlice $subscription "private enum ComposeSurfaceState"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__EVALUATE__", $attachmentEvaluation)
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__EVALUATE_UI__", (Get-MethodSlice $attachmentFlow "private async Task EvaluateAttachmentAutomationOnUiAsync("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SCHEDULE__", (Get-MethodSlice $attachmentFlow "private void ScheduleAttachmentEvaluation()"))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__MARK_INLINE__", (Get-MethodSlice $subscription "internal void MarkInlineResponse("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__MARK_INSPECTOR__", (Get-MethodSlice $subscription "internal void MarkInspector("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__MARK_CLOSED__", (Get-MethodSlice $subscription "internal void MarkInlineResponseClosed("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__END_SUPPRESSION__", (Get-MethodSlice $attachmentQueue "private void EndAttachmentSuppression("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__BEFORE_ADD__", $beforeAttachmentAdd)
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SNAPSHOT__", $snapshotAttachments)
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__SUM__", (Get-MethodSlice $attachmentMaterialization "private static long SumAttachmentBytes("))
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__LAST_BATCH__", $lastAddedBatch)
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__REMOVE_LAST__", $removeLastAttachmentBatch)
$attachmentSendRaceHarness = $attachmentSendRaceHarness.Replace("__REMOVE_INDICES__", $removeAttachmentsByIndices)
Add-Type -TypeDefinition $attachmentSendRaceHarness -Language CSharp -IgnoreWarnings -WarningAction SilentlyContinue
[AttachmentSendRaceBoundary.AttachmentSendRaceRegression]::Run()
Write-Host "[OK] Initial forward/draft attachments: inherited files, image-only messages, pending policy, pop-out, detached surfaces, busy flows and single wizard"
Write-Host "[OK] Delayed attachment policy, prompt completion and direct removal leave accepted Send unchanged"

# Execute the production attachment gate with inert backend and UI boundaries.
$attachmentSendGateHarness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

public static class AttachmentSendGateRegression
{
    private enum MessageBoxButtons { OK }
    private enum MessageBoxIcon { Information, Warning }
    private static class Strings
    {
        internal const string DialogTitle = "dialog-title";
        internal const string AttachmentRoutingRequired = "routing-required";
    }
    private static class MessageBox
    {
        internal static readonly List<string> Notices = new List<string>();
        internal static void Show(string text, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        { Notices.Add(text); }
    }
    private sealed class TalkServiceConfiguration
    {
        internal readonly string Key;
        internal TalkServiceConfiguration(string url, string username, string password) { Key = url + "|" + username + "|" + password; }
        internal bool IsComplete() { return true; }
    }
    private sealed class BackendPolicyStatus
    {
        internal bool FetchSucceeded, Always, ThresholdMandatory, RolloutBlocked;
        internal string Reason = "";
        internal int ThresholdMb = 10;
        internal bool IsEndpointMissing { get { return Reason == "backend_missing"; } }
        internal bool IsServiceUnavailable
        { get { return !FetchSucceeded && (Reason == "nextcloud_unavailable" || Reason == "backend_unavailable" || IsEndpointMissing || Reason == "rate_limited"); } }
        internal bool IsDomainActive(string domain) { return FetchSucceeded && (Always || ThresholdMandatory); }
        internal bool IsLocked(string domain, string key) { return ThresholdMandatory; }
        internal bool HasPolicyKey(string domain, string key) { return ThresholdMandatory; }
    }
    private sealed class AddinSettings
    {
        internal bool IsEnterpriseRollout, SendPolicyFailClosed;
        internal string ServerUrl = "https://cloud.example.test", Username = "user", AppPassword = "test-only";
        internal bool SharingAttachmentsAlwaysConnector, SharingAttachmentsOfferAboveEnabled;
        internal int SharingAttachmentsOfferAboveMb = 10;
        internal AddinSettings ResolvePolicyDefaults(BackendPolicyStatus status)
        {
            return new AddinSettings
            {
                SharingAttachmentsAlwaysConnector = SharingAttachmentsAlwaysConnector || (status != null && status.Always),
                SharingAttachmentsOfferAboveEnabled = SharingAttachmentsOfferAboveEnabled || (status != null && status.ThresholdMandatory),
                SharingAttachmentsOfferAboveMb = status != null ? status.ThresholdMb : SharingAttachmentsOfferAboveMb
            };
        }
    }
    private static class PolicyUiHelper
    {
        internal static string GetEnterpriseRolloutNotice(AddinSettings settings, BackendPolicyStatus status)
        { return status.RolloutBlocked ? "rollout-blocked" : string.Empty; }
    }
    private static class OutlookAttachmentAutomationGuardService
    {
        internal static int NormalizeThresholdMb(int value) { return Math.Max(1, value); }
    }
    private sealed class Owner
    {
        internal AddinSettings _currentSettings = new AddinSettings();
        internal BackendPolicyStatus Confirmed, Check;
        internal bool CheckCurrent = true;
        internal string ConfirmedKey;
        internal bool TryGetCachedEmailSignaturePolicyStatus(TalkServiceConfiguration configuration, out BackendPolicyStatus status)
        { status = configuration.Key == ConfirmedKey ? Confirmed : null; return status != null; }
        internal bool TryGetCurrentBackendPolicyCheck(TalkServiceConfiguration configuration, out BackendPolicyStatus status)
        { status = configuration.Key == ConfirmedKey ? Check : null; return CheckCurrent && status != null; }
        internal Task<BackendPolicyStatus> GetEmailSignaturePolicyStatusAsync(TalkServiceConfiguration configuration, string trigger)
        { return Task.FromResult(Confirmed ?? Check); }
    }
    private sealed class Subscription
    {
        __SETTINGS_CLASS__
        private AttachmentAutomationSettings _attachmentAutomationSettingsSnapshot;
        private DateTime _attachmentAutomationSettingsSnapshotUtc;
        private Task<AttachmentAutomationSettings> _attachmentAutomationSettingsRefreshTask;
        private int _attachmentAutomationSettingsRefreshGeneration;
        private static readonly TimeSpan AttachmentAutomationSettingsCacheLifetime = TimeSpan.FromMinutes(5);
        private readonly string _composeKey = "test";
        internal readonly Owner _owner = new Owner();
        internal bool _disposed, LocalAlways;
        internal int AttachmentCount = 1, CountCalls, RefreshCalls, Warnings, Blocks;
        internal long TotalBytes = 11L * 1024L * 1024L;
        internal readonly List<string> Logs = new List<string>();
        internal bool Validate(ref bool cancel)
        { MessageBox.Notices.Clear(); Warnings = Blocks = 0; return TryValidateAttachmentPolicyBeforeSend(ref cancel); }
        internal bool Pause() { return PauseUnavailableAttachmentAutomation(ReadAttachmentAutomationSettings(), TotalBytes); }
        internal void Know(BackendPolicyStatus status, BackendPolicyStatus check, bool closed)
        {
            _owner.Confirmed = status;
            _owner.Check = check;
            _owner._currentSettings.SendPolicyFailClosed = closed;
            AddinSettings settings = _owner._currentSettings;
            _owner.ConfirmedKey = new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword).Key;
        }
        private int CountPolicyRelevantAttachments(out long totalBytes)
        { CountCalls++; totalBytes = TotalBytes; return AttachmentCount; }
        private void BeginAttachmentAutomationSettingsRefresh() { RefreshCalls++; }
        private AttachmentAutomationSettings ReadLocalAttachmentAutomationSettings()
        {
            _owner._currentSettings.SharingAttachmentsAlwaysConnector = LocalAlways;
            return ApplyAttachmentAutomationPolicy(new AttachmentAutomationSettings { LocalSettings = _owner._currentSettings }, null);
        }
        private void ShowSendPolicyWarning() { }
        private void RecordSendPolicyWarning(bool signature, bool attachments, BackendPolicyStatus check)
        { if (!attachments || signature || !check.IsServiceUnavailable) { throw new InvalidOperationException("Wrong attachment warning"); } Warnings++; }
        private bool BlockSendPolicyFailure(ref bool cancel, BackendPolicyStatus check, bool unknown)
        { cancel = true; Blocks++; return false; }
        private void LogFileLink(string message) { Logs.Add(message); }
        __FRESHNESS__
        __READ_SETTINGS__
        __READ_SETTINGS_ASYNC__
        __REFRESH_SETTINGS__
        __INVALIDATE_SETTINGS__
        __SEND_GATE__
        __REQUIRED_NOTICE__
        __APPLY_POLICY__
        __BUILD_SETTINGS__
        __PAUSE_AUTOMATION__
    }
    private static void Check(bool condition, string message)
    { if (!condition) { throw new InvalidOperationException(message); } }
    public static void Run()
    {
        foreach (bool closed in new[] { false, true })
        {
            foreach (string checkState in new[] { "online", "nextcloud_unavailable", "backend_unavailable", "backend_missing", "rate_limited", "authentication_rejected", "invalid_payload", "check_failed", "pending" })
            {
                foreach (string rule in new[] { "none", "always", "threshold-required", "threshold-optional", "refused" })
                {
                    var compose = new Subscription();
                    var known = new BackendPolicyStatus
                    {
                        FetchSucceeded = true,
                        Always = rule == "always",
                        ThresholdMandatory = rule == "threshold-required",
                        RolloutBlocked = rule == "refused"
                    };
                    var check = new BackendPolicyStatus { FetchSucceeded = checkState == "online", Reason = checkState };
                    compose.Know(known, check, closed);
                    compose._owner.CheckCurrent = checkState != "pending";
                    bool cancel = false;
                    bool result = compose.Validate(ref cancel);
                    bool required = rule == "always" || rule == "threshold-required";
                    bool allow = !required
                        || (check.IsServiceUnavailable && !closed);
                    Check(result == allow && cancel == !allow, "Wrong Send result: " + rule + "/" + checkState + "/" + closed);
                    Check(compose.Warnings == (required && check.IsServiceUnavailable && !closed ? 1 : 0),
                        "Wrong outage warning: " + rule + "/" + checkState + "/" + closed);
                    Check(MessageBox.Notices.Count == (required && (checkState == "online" || (checkState == "pending" && !closed)) ? 1 : 0),
                        "Unrelated policy generated routing UI: " + rule + "/" + checkState + "/" + closed);
                    if (MessageBox.Notices.Count > 0)
                    { Check(MessageBox.Notices[0] == Strings.AttachmentRoutingRequired, "Misleading wizard failure reused at Send."); }
                }
            }

            foreach (long bytes in new[] { 0L, 10L * 1024L * 1024L, 10L * 1024L * 1024L + 1L, long.MaxValue })
            {
                var compose = new Subscription();
                compose.Know(new BackendPolicyStatus { FetchSucceeded = true, ThresholdMandatory = true },
                    new BackendPolicyStatus { FetchSucceeded = true }, closed);
                compose.TotalBytes = bytes;
                bool cancel = false;
                bool allowed = bytes <= 10L * 1024L * 1024L;
                Check(compose.Validate(ref cancel) == allowed && cancel == !allowed, "Wrong mandatory threshold boundary.");
            }

            var cold = new Subscription();
            cold._owner._currentSettings.SendPolicyFailClosed = closed;
            cold._owner.Check = new BackendPolicyStatus { Reason = "nextcloud_unavailable" };
            bool coldCancel = false;
            Check(cold.Validate(ref coldCancel) && !coldCancel && cold.Warnings == 0 && MessageBox.Notices.Count == 0,
                "The attachment path fabricated a rule from an unknown initial state; unknown closed mode belongs to the shared Send gate.");

            var account = new Subscription();
            account.Know(new BackendPolicyStatus { FetchSucceeded = true, Always = true },
                new BackendPolicyStatus { FetchSucceeded = true }, closed);
            account._owner._currentSettings.Username = "different";
            bool accountCancel = false;
            Check(account.Validate(ref accountCancel) && !accountCancel && account.Warnings == 0,
                "A different account inherited an attachment policy.");

            var empty = new Subscription();
            empty.Know(new BackendPolicyStatus { FetchSucceeded = true, Always = true },
                new BackendPolicyStatus { Reason = "nextcloud_unavailable" }, closed);
            empty.AttachmentCount = 0;
            bool emptyCancel = false;
            Check(empty.Validate(ref emptyCancel) && !emptyCancel && empty.RefreshCalls == 0 && empty.Warnings == 0,
                "No relevant attachment still checked or enforced a rule.");

            foreach (BackendPolicyStatus previous in new[] { null, new BackendPolicyStatus { FetchSucceeded = true } })
            {
                var localOnly = new Subscription { LocalAlways = true };
                localOnly.Know(previous, new BackendPolicyStatus { Reason = "backend_missing" }, closed);
                Check(!localOnly.Pause() && localOnly.Warnings == 0,
                    "A missing optional backend paused local attachment sharing.");
                bool localCancel = false;
                Check(!localOnly.Validate(ref localCancel) && localCancel && localOnly.Warnings == 0
                    && MessageBox.Notices.Count == 1,
                    "A missing optional backend bypassed local AlwaysConnector with native attachments.");
                localOnly.AttachmentCount = 0;
                localCancel = false;
                Check(localOnly.Validate(ref localCancel) && !localCancel && localOnly.Warnings == 0,
                    "Discarded attachment originals still blocked the attachment Send gate.");
            }

            foreach (bool managed in new[] { false, true })
            {
                var requiredBackend = new Subscription { LocalAlways = true };
                requiredBackend._owner._currentSettings.IsEnterpriseRollout = managed;
                requiredBackend.Know(managed ? null : new BackendPolicyStatus { FetchSucceeded = true, Always = true },
                    new BackendPolicyStatus { Reason = "backend_missing" }, closed);
                Check(requiredBackend.Pause(), "A required or previously used backend was treated as optional.");
                bool missingCancel = false;
                Check(requiredBackend.Validate(ref missingCancel) == !closed && missingCancel == closed,
                    "Missing required backend ignored the selected outage mode.");
            }

            foreach (string failure in new[] { "nextcloud_unavailable", "backend_unavailable", "authentication_rejected", "invalid_payload", "rate_limited" })
            {
                var failedLocal = new Subscription { LocalAlways = true };
                failedLocal.Know(null, new BackendPolicyStatus { Reason = failure }, closed);
                Check(failedLocal.Pause(), "A real failure was mistaken for a missing optional backend: " + failure);
            }
        }

        var freshRefusal = new Subscription();
        freshRefusal.LocalAlways = true;
        freshRefusal.Know(new BackendPolicyStatus { FetchSucceeded = true, RolloutBlocked = true },
            new BackendPolicyStatus { FetchSucceeded = true }, true);
        bool refusalCancel = false;
        Check(freshRefusal.Validate(ref refusalCancel) && !refusalCancel && freshRefusal.Warnings == 0,
            "Confirmed missing Seat blocked ordinary attachments.");

        var disposed = new Subscription { _disposed = true };
        bool disposedCancel = false;
        Check(disposed.Validate(ref disposedCancel) && disposed.CountCalls == 0, "Disposed subscription inspected attachments.");
        var host = new Subscription();
        bool hostCancel = true;
        Check(!host.Validate(ref hostCancel) && hostCancel && host.CountCalls == 0, "Already-cancelled Send was changed.");
    }
}
'@
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__SETTINGS_CLASS__", (Get-MethodSlice $attachmentPolicy "private sealed class AttachmentAutomationSettings"))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__FRESHNESS__", $attachmentSettingsFreshness)
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__READ_SETTINGS__", $readAttachmentSettings)
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__READ_SETTINGS_ASYNC__", (Get-MethodSlice $attachmentPolicy "private async Task<AttachmentAutomationSettings> ReadAttachmentAutomationSettingsAsync()"))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__REFRESH_SETTINGS__", (Get-MethodSlice $attachmentPolicy "private async Task<AttachmentAutomationSettings> RefreshAttachmentAutomationSettingsAsync("))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__INVALIDATE_SETTINGS__", (Get-MethodSlice $attachmentPolicy "internal void RefreshAttachmentAutomationSettings()"))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__SEND_GATE__", $attachmentSendGate)
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__REQUIRED_NOTICE__", $requiredAttachmentNotice)
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__APPLY_POLICY__", (Get-MethodSlice $attachmentPolicy "private static AttachmentAutomationSettings ApplyAttachmentAutomationPolicy("))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__BUILD_SETTINGS__", (Get-MethodSlice $attachmentPolicy "private static AttachmentAutomationSettings BuildAttachmentAutomationSettings("))
$attachmentSendGateHarness = $attachmentSendGateHarness.Replace("__PAUSE_AUTOMATION__", (Get-MethodSlice $attachmentPolicy "private bool PauseUnavailableAttachmentAutomation("))
Add-Type -TypeDefinition $attachmentSendGateHarness -Language CSharp -IgnoreWarnings -WarningAction SilentlyContinue
[AttachmentSendGateRegression]::Run()
Write-Host "[OK] Attachment Send matrix: fail modes, locked threshold, availability, rate limiting, authentication, unknown state, account change and Seat refusal"
Write-Host "[OK] Missing optional backend permits local automation; managed/remembered backend rules and actual outages retain their enforcement"

$classificationRoot = Join-Path ([IO.Path]::GetTempPath()) ("nc4ol-attachment-classification-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $classificationRoot | Out-Null
try {
    $classificationHarness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Css.Parser;
using AngleSharp.Css.Values;
using NcTalkOutlookAddIn.Utilities;
using Outlook = AttachmentClassificationOutlook;
namespace NcTalkOutlookAddIn.Utilities
{
    internal static class DiagnosticsLogger
    {
        internal static bool IsEnabled { get { return false; } }
        internal static void Log(string category, string message) { }
        internal static void LogException(string category, string message, Exception ex) { }
    }
    internal static class LogCategories { internal const string Core = "CORE"; }
}
namespace AttachmentClassificationOutlook
{
    public enum OlBodyFormat { olFormatPlain = 1, olFormatHTML = 2, olFormatRichText = 3 }
    public sealed class PropertyAccessor
    {
        public readonly Dictionary<string, object> Values = new Dictionary<string, object>();
        public int Failure;
        public object GetProperty(string name)
        {
            if (Failure != 0) { throw new COMException("property failure", Failure); }
            object value;
            if (Values.TryGetValue(name.Substring(name.LastIndexOf('/') + 1), out value)) { return value; }
            throw new COMException("missing optional property", unchecked((int)0x8004010F));
        }
    }
    public sealed class Attachment
    {
        public readonly PropertyAccessor PropertyAccessor = new PropertyAccessor();
        public string Name = "file.jpg";
        public int Size;
    }
    public sealed class Attachments
    {
        public readonly List<Attachment> Items = new List<Attachment>();
        public int Count { get { return Items.Count; } }
        public Attachment this[int index] { get { return Items[index - 1]; } }
        public void Remove(int index) { Items.RemoveAt(index - 1); }
    }
    public sealed class Mail
    {
        public OlBodyFormat BodyFormat = OlBodyFormat.olFormatHTML;
        public string HTMLBody = "";
        public readonly Attachments Attachments = new Attachments();
    }
}
public static class AttachmentClassificationRegression
{
    private static class LogCategories { internal const string FileLink = "FILELINK"; }
    private static class DiagnosticsLogger
    { internal static void LogException(string category, string message, Exception ex) { } }
    private static class ComInteropScope
    { internal static void TryRelease(object item, string category, string message) { } }
    private sealed class Subscription
    {
        internal readonly Outlook.Mail _mail = new Outlook.Mail();
        private readonly string _composeKey = "classification";
        private bool CanApplyComposeChanges { get { return true; } }
        private bool _attachmentSuppressed;
        private readonly List<string> _pendingAddedBatch = new List<string>();
        private void EndAttachmentSuppression(string reason) { _attachmentSuppressed = false; }
        private void LogFileLink(string message) { }
        private static string ReadAttachmentName(Outlook.Attachment item) { return item.Name; }
        private static long ReadAttachmentSizeBytes(Outlook.Attachment item) { return item.Size; }
        internal bool IsFile(Outlook.Attachment item, bool defer) { return IsAutomationFileAttachment(item, defer); }
        internal int Count(out long bytes) { return CountPolicyRelevantAttachments(out bytes); }
        internal int SnapshotCount() { return SnapshotAttachments().Count; }
        internal void RemoveAll() { RemoveAttachmentsByIndices(new List<int> { 1, 2, 3 }, "test"); }
        __METHODS__
    }
    private static void Check(bool condition, string message)
    { if (!condition) { throw new InvalidOperationException(message); } }
    private static Outlook.Attachment File(string id, int size)
    {
        var item = new Outlook.Attachment { Size = size };
        item.PropertyAccessor.Values["0x37140003"] = 0;
        if (id != null) { item.PropertyAccessor.Values["0x3712001F"] = id; }
        return item;
    }
    public static void Main()
    {
        var compose = new Subscription();
        var logo = File("logo@example.test", 0);
        string[] embedded = {
            "<blockquote><img src='cid:logo@example.test'><img src='cid:logo@example.test'></blockquote>",
            "<IMG SRC=CiD:logo%40example.test>",
            "<v:imagedata src='cid:logo@example.test'></v:imagedata>",
            "<body background='cid:logo@example.test'>body</body>",
            "<table><tr><td background='cid:logo@example.test'>body</td></tr></table>",
            "<div style=\"background-image:url('cid:logo@example.test')\">body</div>",
            "<div style=\"background:url(cid:logo@example.test) no-repeat\">body</div>",
            "<div style=\"background-image:url(cid:other),url(cid:logo@example.test)\">body</div>",
            "<!--[if mso]><v:imagedata src='cid:logo@example.test'></v:imagedata><![endif]-->",
            "<!--[if gte mso 9]><v:fill src='cid:logo@example.test'></v:fill><![endif]-->"
        };
        foreach (string html in embedded)
        {
            compose._mail.HTMLBody = html;
            Check(!compose.IsFile(logo, false), "Body resource counted as a file: " + html);
        }
        string[] notEmbedded = {
            "logo@example.test cid:logo@example.test", "<a href='cid:logo@example.test'>file</a>",
            "<!-- <img src='cid:logo@example.test'> -->", "<img alt='cid:logo@example.test'>",
            "<img data-src='cid:logo@example.test'>", "<img src='cid:logo@example.test.extra'>",
            "<img src='https://example.test/logo@example.test'>",
            "&lt;img src='cid:logo@example.test'&gt;",
            "<div style=\"content:'url(cid:logo@example.test)'\">body</div>",
            "<!--[if !mso]><img src='cid:logo@example.test'><![endif]-->"
        };
        foreach (string html in notEmbedded)
        {
            compose._mail.HTMLBody = html;
            Check(compose.IsFile(logo, false), "Non-resource text exempted a real file: " + html);
        }
        compose._mail.HTMLBody = "<img src='cid:logo@example.test'>";
        logo.PropertyAccessor.Values["0x7FFE000B"] = false;
        Check(!compose.IsFile(logo, false), "Explicit non-hidden value masked a body reference.");
        logo.PropertyAccessor.Values["0x3712001F"] = "<logo@example.test>";
        Check(!compose.IsFile(logo, false), "Content-ID brackets were not normalized.");
        logo.PropertyAccessor.Values.Remove("0x3712001F");
        logo.PropertyAccessor.Values["0x3712001E"] = "logo@example.test";
        Check(!compose.IsFile(logo, false), "ANSI Content-ID was ignored.");
        var located = File(null, 100);
        located.PropertyAccessor.Values["0x3713001F"] = "images/logo.png";
        compose._mail.HTMLBody = "<img src='images/logo.png'>";
        Check(!compose.IsFile(located, false), "Content-Location body image counted as file.");
        compose._mail.HTMLBody = "<img src='cid:a&amp;b@example.test'>";
        Check(!compose.IsFile(File("a&b@example.test", 10), false), "HTML entity in resource reference was ignored.");
        compose._mail.HTMLBody = "<img src='cid:logo@example.test'>";
        foreach (Outlook.OlBodyFormat format in new[] { Outlook.OlBodyFormat.olFormatPlain, Outlook.OlBodyFormat.olFormatRichText })
        {
            compose._mail.BodyFormat = format;
            Check(compose.IsFile(logo, false), "Non-HTML body text exempted a file.");
        }
        compose._mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
        var hidden = File(null, 100);
        hidden.PropertyAccessor.Values["0x7FFE000B"] = true;
        Check(!compose.IsFile(hidden, false), "Hidden resource counted as file.");
        hidden.PropertyAccessor.Values["0x7FFE000B"] = false;
        hidden.PropertyAccessor.Values["0x37140003"] = 7;
        Check(!compose.IsFile(hidden, false), "Rendered-in-body bit was ignored with other flags.");
        Check(compose.IsFile(File(null, 0), false), "Zero-byte file was exempted.");
        compose._mail.HTMLBody = "";
        Check(!compose.IsFile(logo, true) && compose.IsFile(logo, false), "Early CID candidate was not deferred and rechecked.");
        compose._mail.HTMLBody = "<img src='cid:logo@example.test'>";
        compose._mail.Attachments.Items.Add(logo);
        compose._mail.Attachments.Items.Add(File(null, 25));
        compose._mail.Attachments.Items.Add(File("unreferenced@example.test", 75));
        long bytes;
        Check(compose.Count(out bytes) == 2 && bytes == 100 && compose.SnapshotCount() == 2,
            "Real production count/snapshot did not separate inline resources and ordinary files.");
        compose.RemoveAll();
        Check(compose._mail.Attachments.Count == 1 && object.ReferenceEquals(compose._mail.Attachments[1], logo),
            "Removal discarded an inline resource.");
        logo.PropertyAccessor.Failure = unchecked((int)0x80004005);
        Check(compose.Count(out bytes) == int.MaxValue && bytes == long.MaxValue,
            "Unexpected COM failure silently bypassed file policy.");
        Console.WriteLine("[OK] Production attachment classification: unsaved replies, CID/location, HTML/CSS/VML, true file negatives, deferred events, counts, snapshots and removal");
    }
}
'@
    $classificationMethods = @(
        'private sealed class AttachmentSnapshot',
        'private int CountPolicyRelevantAttachments(',
        'private bool IsAutomationFileAttachment(',
        'private static object ReadOptionalAttachmentProperty(',
        'private static string NormalizeAttachmentContentId(',
        'private static bool IsAttachmentReferencedInBody(',
        '[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]',
        'private static bool MatchesAttachmentBodyResource(',
        'private List<AttachmentSnapshot> SnapshotAttachments()',
        'private void RemoveAttachmentsByIndices('
    ) | ForEach-Object { Get-MethodSlice $attachmentMaterialization $_ }
    $classificationHarness = $classificationHarness.Replace('__METHODS__', ($classificationMethods -join "`r`n"))
    $sourcePath = Join-Path $classificationRoot 'AttachmentClassification.cs'
    [IO.File]::WriteAllText($sourcePath, $classificationHarness, [Text.UTF8Encoding]::new($false))
    $addInDirectory = Join-Path $classificationRoot 'addin'
    New-Item -ItemType Directory -Path $addInDirectory | Out-Null
    $library = Join-Path $addInDirectory 'AttachmentClassification.dll'
    $vendor = Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/vendor/htmlsanitizer'
    $references = @('/reference:System.dll', '/reference:System.Core.dll', '/reference:System.Numerics.dll', '/reference:System.Web.dll')
    Get-ChildItem $vendor -Filter '*.dll' | ForEach-Object {
        $references += "/reference:$($_.FullName)"
        Copy-Item -LiteralPath $_.FullName -Destination $addInDirectory
    }
    $sanitizerSource = Join-Path $ProjectRoot 'src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs'
    $csc = "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    & $csc /nologo /nowarn:1702 /optimize+ /target:library "/out:$library" @references $sourcePath $sanitizerSource
    if ($LASTEXITCODE -ne 0) { throw 'Attachment classification harness failed to compile.' }
    # Outlook hosts the DLL without adopting its .config; keep the host and dependencies separate.
    $hostSource = @'
using System;
using System.Reflection;
internal static class AttachmentHost
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = new System.Text.UTF8Encoding(false);
        try
        {
            Assembly assembly = Assembly.LoadFrom(args[0]);
            Type regression = assembly.GetType("AttachmentClassificationRegression", false);
            if (regression != null)
            {
                regression.GetMethod("Main").Invoke(null, null);
            }
            else
            {
                Type compose = assembly.GetType("NcTalkOutlookAddIn.NextcloudTalkAddIn+MailComposeSubscription", true);
                MethodInfo match = compose.GetMethod("IsAttachmentReferencedInBody", BindingFlags.Static | BindingFlags.NonPublic);
                string[] html = { "<img src='cid:logo'>", "<div style='background-image:url(cid:logo)'>text</div>",
                    "<!--[if mso]><v:imagedata src='cid:logo'></v:imagedata><![endif]-->" };
                foreach (string body in html)
                {
                    if (!(bool)match.Invoke(null, new object[] { body, "logo", "" }) ||
                        (bool)match.Invoke(null, new object[] { body, "real-file", "" }))
                    { throw new InvalidOperationException("Packaged attachment classification failed."); }
                }
            }
            Console.WriteLine("[OK] Cold attachment parsing in an external .NET Framework host without binding redirects");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
'@
    $hostPath = Join-Path $classificationRoot 'AttachmentHost.cs'
    [IO.File]::WriteAllText($hostPath, $hostSource, [Text.UTF8Encoding]::new($false))
    $exe = Join-Path $classificationRoot 'AttachmentHost.exe'
    & $csc /nologo /optimize+ /target:exe "/out:$exe" $hostPath
    if ($LASTEXITCODE -ne 0) { throw 'Attachment host failed to compile.' }
    & $exe $library
    if ($LASTEXITCODE -ne 0) { throw 'Attachment classification regression failed.' }
    if (-not [string]::IsNullOrWhiteSpace($BuiltAddInPath)) {
        & $exe (Resolve-Path -LiteralPath $BuiltAddInPath).Path
        if ($LASTEXITCODE -ne 0) { throw 'Built add-in cold attachment classification failed.' }
    }
}
finally {
    $resolvedClassificationRoot = (Resolve-Path -LiteralPath $classificationRoot).Path
    if ([IO.Path]::GetDirectoryName($resolvedClassificationRoot) -eq [IO.Path]::GetTempPath().TrimEnd('\') -and
        [IO.Path]::GetFileName($resolvedClassificationRoot).StartsWith('nc4ol-attachment-classification-')) {
        Remove-Item -LiteralPath $resolvedClassificationRoot -Recurse -Force
    }
}
Write-Host "All Outlook compose lifecycle regression checks passed."
