Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"
$SignaturePath = Join-Path $SourceRoot "NextcloudTalkAddIn.MailComposeSubscription.Signature.cs"
$SendPath = Join-Path $SourceRoot "NextcloudTalkAddIn.MailComposeSubscription.SendCleanup.cs"
$InteropPath = Join-Path $SourceRoot "Controllers\MailInteropController.cs"
$ComposeShareLifecyclePath = Join-Path $SourceRoot "Controllers\ComposeShareLifecycleController.cs"
$PolicyPath = Join-Path $SourceRoot "NextcloudTalkAddIn.PolicyTemplates.cs"
$SignatureContentPath = Join-Path $SourceRoot "Utilities\HtmlToPlainTextConverter.cs"
$SignaturePlacementPath = Join-Path $SourceRoot "Utilities\EmailSignatureSlotPlacementPolicy.cs"

foreach ($requiredPath in @($SignaturePath, $SendPath, $InteropPath, $ComposeShareLifecyclePath, $PolicyPath, $SignatureContentPath, $SignaturePlacementPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required email-signature source file not found: $requiredPath"
    }
}

$SignatureSource = Get-Content -Raw -LiteralPath $SignaturePath
$SendSource = Get-Content -Raw -LiteralPath $SendPath
$InteropSource = Get-Content -Raw -LiteralPath $InteropPath
$ComposeShareLifecycleSource = Get-Content -Raw -LiteralPath $ComposeShareLifecyclePath
$PolicySource = Get-Content -Raw -LiteralPath $PolicyPath
$SignatureContentSource = Get-Content -Raw -LiteralPath $SignatureContentPath
$SignaturePlacementSource = Get-Content -Raw -LiteralPath $SignaturePlacementPath
$Failures = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    Param([string]$Message)
    $Failures.Add($Message)
}

function Require-Pattern {
    Param(
        [string]$Text,
        [string]$Pattern,
        [string]$Message
    )
    if (-not [regex]::IsMatch($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
        Add-Failure $Message
    }
}

function Forbid-Pattern {
    Param(
        [string]$Text,
        [string]$Pattern,
        [string]$Message
    )
    if ([regex]::IsMatch($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
        Add-Failure $Message
    }
}

function Require-Order {
    Param(
        [string]$Text,
        [string]$First,
        [string]$Second,
        [string]$Message
    )
    $firstIndex = $Text.IndexOf($First, [StringComparison]::Ordinal)
    $secondIndex = $Text.IndexOf($Second, [StringComparison]::Ordinal)
    if ($firstIndex -lt 0 -or $secondIndex -lt 0 -or $firstIndex -ge $secondIndex) {
        Add-Failure $Message
    }
}

function Require-PatternBeforeLiteral {
    Param(
        [string]$Text,
        [string]$FirstPattern,
        [string]$Second,
        [string]$Message
    )
    $firstMatch = [regex]::Match($Text, $FirstPattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $secondIndex = $Text.IndexOf($Second, [StringComparison]::Ordinal)
    if (-not $firstMatch.Success -or $secondIndex -lt 0 -or $firstMatch.Index -ge $secondIndex) {
        Add-Failure $Message
    }
}

function Get-CSharpMethodBlock {
    Param(
        [string]$Text,
        [string]$MethodName
    )

    $escapedName = [regex]::Escape($MethodName)
    $definition = [regex]::Match(
        $Text,
        "(?m)^[ \t]*(?:private|internal|public)[ \t]+(?:static[ \t]+)?(?:async[ \t]+)?[A-Za-z0-9_<>,\.\[\]\?]+[ \t]+$escapedName[ \t]*\(")
    if (-not $definition.Success) {
        return $null
    }

    $braceStart = $Text.IndexOf("{", $definition.Index + $definition.Length, [StringComparison]::Ordinal)
    if ($braceStart -lt 0) {
        return $null
    }

    $depth = 0
    $state = "code"
    for ($index = $braceStart; $index -lt $Text.Length; $index++) {
        $character = $Text[$index]
        $nextCharacter = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }

        if ($state -eq "line_comment") {
            if ($character -eq "`r" -or $character -eq "`n") {
                $state = "code"
            }
            continue
        }
        if ($state -eq "block_comment") {
            if ($character -eq '*' -and $nextCharacter -eq '/') {
                $state = "code"
                $index++
            }
            continue
        }
        if ($state -eq "string") {
            if ($character -eq '\') {
                $index++
            } elseif ($character -eq '"') {
                $state = "code"
            }
            continue
        }
        if ($state -eq "verbatim_string") {
            if ($character -eq '"' -and $nextCharacter -eq '"') {
                $index++
            } elseif ($character -eq '"') {
                $state = "code"
            }
            continue
        }
        if ($state -eq "character") {
            if ($character -eq '\') {
                $index++
            } elseif ($character -eq "'") {
                $state = "code"
            }
            continue
        }

        if ($character -eq '/' -and $nextCharacter -eq '/') {
            $state = "line_comment"
            $index++
        } elseif ($character -eq '/' -and $nextCharacter -eq '*') {
            $state = "block_comment"
            $index++
        } elseif ($character -eq '"') {
            $state = if ($index -gt 0 -and $Text[$index - 1] -eq '@') { "verbatim_string" } else { "string" }
        } elseif ($character -eq "'") {
            $state = "character"
        } elseif ($character -eq '{') {
            $depth++
        } elseif ($character -eq '}') {
            $depth--
            if ($depth -eq 0) {
                return $Text.Substring($definition.Index, $index - $definition.Index + 1)
            }
        }
    }
    return $null
}

# Compose subscriptions only orchestrate policy and the shared WordEditor reconciler.
foreach ($forbidden in @(
    @{ Pattern = '\bHTMLBody\b'; Message = 'Signature subscription accesses MailItem.HTMLBody directly.' },
    @{ Pattern = '\bRTFBody\b'; Message = 'Signature subscription accesses MailItem.RTFBody directly.' },
    @{ Pattern = '\.Body\s*=(?!=)'; Message = 'Signature subscription writes MailItem.Body directly.' },
    @{ Pattern = '\.BodyFormat\s*=(?!=)'; Message = 'Signature subscription changes MailItem.BodyFormat.' },
    @{ Pattern = '\bSelection\b'; Message = 'Signature subscription accesses the current Word selection directly.' },
    @{ Pattern = '\bTypeParagraph\b'; Message = 'Signature subscription inserts paragraphs through the current Word selection.' },
    @{ Pattern = '\bEmailSignaturePlainTextController\b'; Message = 'Signature subscription bypasses the shared WordEditor reconciler for plain text.' },
    @{ Pattern = '\bManagedSignatureBlockRegex\b'; Message = 'Signature subscription still owns HTML marker replacement.' },
    @{ Pattern = '\bReplaceComposeSignatureSlot\b'; Message = 'Signature subscription still owns raw compose-slot replacement.' },
    @{ Pattern = '\bEnsureHtmlBodyForEmailSignature\b'; Message = 'Signature subscription still converts RTF/plain compose bodies to HTML.' },
    @{ Pattern = '\bTryWriteEmailSignatureHtmlBody\b'; Message = 'Signature subscription still owns direct HTML-body writes.' },
    @{ Pattern = '\bTryWriteMailHtmlBodyPreservingSelection\b'; Message = 'Signature subscription calls the legacy HTMLBody writer.' },
    @{ Pattern = '\bTryReplaceInspectorSignatureSlot\b'; Message = 'Signature subscription calls the legacy Inspector slot path.' },
    @{ Pattern = '\bTryReplaceActiveInlineResponseSignatureSlot\b'; Message = 'Signature subscription calls the legacy inline slot path.' },
    @{ Pattern = '\bFetchBackendPolicyStatus\b'; Message = 'Signature subscription performs a synchronous backend policy fetch.' },
    @{ Pattern = '\.Result\s*(?:[;,\)\]\}]|\?\?)|\.Wait\s*\(|GetAwaiter\(\)\.GetResult\('; Message = 'Signature subscription blocks while loading policy state.' }
)) {
    Forbid-Pattern $SignatureSource $forbidden.Pattern $forbidden.Message
}

foreach ($required in @(
    @{ Pattern = '\bGetEmailSignaturePolicyStatusAsync\s*\('; Message = 'Signature subscription does not use the asynchronous signature-policy cache.' },
    @{ Pattern = '\bRunOnOutlookUiThreadAsync\s*\('; Message = 'Signature reconciliation is not marshalled back to the Outlook UI STA.' },
    @{ Pattern = '\.\s*ApplyManagedEmailSignature\s*\('; Message = 'Signature subscription does not apply through MailInteropController.' },
    @{ Pattern = '\.\s*ClearManagedEmailSignature\s*\('; Message = 'Signature subscription does not clear its managed Word bookmark through MailInteropController.' },
    @{ Pattern = '\.\s*ClearInitialEmailSignatureSlot\s*\('; Message = 'Signature subscription does not clear the exact initial Outlook signature slot through MailInteropController.' },
    @{ Pattern = 'EmailSignatureContentBuilder\.BuildPlainText\s*\('; Message = 'Signature subscription bypasses the shared plain-text content builder.' },
    @{ Pattern = 'EmailSignatureContentBuilder\.BuildManagedHtml\s*\('; Message = 'Signature subscription bypasses the shared managed HTML wrapper.' }
)) {
    Require-Pattern $SignatureSource $required.Pattern $required.Message
}

$signatureApplyPolicy = Get-CSharpMethodBlock $SignatureSource 'ApplyEmailSignaturePolicy'
if ($null -eq $signatureApplyPolicy) {
    Add-Failure 'ApplyEmailSignaturePolicy could not be parsed.'
} else {
    Require-Pattern $signatureApplyPolicy '\.ApplyManagedEmailSignature\(\s*_mail\s*,\s*_isInlineResponse\s*,[\s\S]*?_composeKey\s*,[\s\S]*?_inlineExplorerIdentityKey\s*\)' 'Managed signature application does not forward the tracked inline Explorer identity.'
}

$signatureClearManaged = Get-CSharpMethodBlock $SignatureSource 'ClearManagedEmailSignature'
if ($null -eq $signatureClearManaged) {
    Add-Failure 'ClearManagedEmailSignature could not be parsed.'
} else {
    Require-Pattern $signatureClearManaged '\.ClearManagedEmailSignature\(\s*_mail\s*,\s*_isInlineResponse\s*,\s*_composeKey\s*,[\s\S]*?_inlineExplorerIdentityKey\s*\)' 'Managed signature cleanup does not forward the tracked inline Explorer identity.'
}

$signatureClearInitial = Get-CSharpMethodBlock $SignatureSource 'ClearInitialEmailSignatureSlot'
if ($null -eq $signatureClearInitial) {
    Add-Failure 'ClearInitialEmailSignatureSlot could not be parsed.'
} else {
    Require-Pattern $signatureClearInitial '\.ClearInitialEmailSignatureSlot\(\s*_mail\s*,\s*_isInlineResponse\s*,\s*_composeKey\s*,[\s\S]*?_inlineExplorerIdentityKey\s*\)' 'Initial signature-slot cleanup does not forward the tracked inline Explorer identity.'
}

$signatureTimerBlock = Get-CSharpMethodBlock $SignatureSource 'OnEmailSignatureTimerTick'
if ($null -eq $signatureTimerBlock) {
    Add-Failure 'OnEmailSignatureTimerTick could not be parsed.'
} else {
    Require-Pattern $signatureTimerBlock '\bGetEmailSignaturePolicyStatusAsync\s*\(' 'Signature timer does not load policy through the asynchronous cache.'
    Require-Pattern $signatureTimerBlock '\bRunOnOutlookUiThreadAsync\s*\(' 'Signature timer does not return Outlook/Word COM work to the UI STA.'
    Require-Pattern $signatureTimerBlock '\.ConfigureAwait\(false\)' 'Signature timer retains a worker-thread continuation context.'
    Require-Pattern $signatureTimerBlock 'generation\s*==\s*_emailSignatureRequestGeneration' 'Signature timer does not reject stale policy responses.'
    Require-Pattern $signatureTimerBlock 'catch\s*\(\s*Exception\s+\w+\s*\)' 'Async signature timer has no exception boundary.'
}

# The shared interop path owns all supported body formats and exact Word slots.
foreach ($required in @(
    @{ Pattern = 'ManagedEmailSignatureBookmarkName\s*=\s*"NcConnectorSignature"'; Message = 'MailInteropController does not declare the managed signature bookmark.' },
    @{ Pattern = 'OutlookAutoSignatureBookmarkName\s*=\s*"_MailAutoSig"'; Message = 'MailInteropController does not declare Outlook''s auto-signature bookmark.' },
    @{ Pattern = 'OutlookOriginalMessageBookmarkName\s*=\s*"_MailOriginal"'; Message = 'MailInteropController does not declare Outlook''s original-message bookmark.' },
    @{ Pattern = 'OutlookOriginalMessageProtectedGap\s*=\s*2\s*;'; Message = 'MailInteropController does not preserve Outlook''s two-character gap before _MailOriginal.' },
    @{ Pattern = '\bReconcileEmailSignatureWordSlot\s*\('; Message = 'MailInteropController does not expose one shared WordEditor reconciliation path.' },
    @{ Pattern = '\bTryResolveSafeEmailSignatureInsertionPoint\s*\('; Message = 'MailInteropController lacks a safe insertion-point resolver.' },
    @{ Pattern = '\bTryAddMissingEmailSignatureLeadingParagraphs\s*\('; Message = 'MailInteropController does not normalize the leading signature gap.' },
    @{ Pattern = '\bCaptureEmailSignatureSelectionBookmark\s*\('; Message = 'MailInteropController does not preserve the cursor with a temporary bookmark.' },
    @{ Pattern = '\bRestoreEmailSignatureSelection\s*\('; Message = 'MailInteropController does not restore the cursor from its temporary bookmark.' }
)) {
    Require-Pattern $InteropSource $required.Pattern $required.Message
}

$interopStart = $InteropSource.IndexOf('internal EmailSignatureReconcileResult ApplyManagedEmailSignature', [StringComparison]::Ordinal)
$interopEnd = $InteropSource.IndexOf('private static void TryShowHiddenBookmarks', [StringComparison]::Ordinal)
if ($interopStart -lt 0 -or $interopEnd -le $interopStart) {
    Add-Failure 'MailInteropController signature-reconciler section could not be isolated.'
} else {
    $interopSignatureSection = $InteropSource.Substring($interopStart, $interopEnd - $interopStart)
    Forbid-Pattern $interopSignatureSection '\.(?:HTMLBody|RTFBody)\s*=(?!=)' 'WordEditor signature reconciliation writes an Outlook body property directly.'
    Forbid-Pattern $interopSignatureSection '\.Body\s*=(?!=)' 'WordEditor signature reconciliation writes MailItem.Body directly.'
    Forbid-Pattern $interopSignatureSection '\.BodyFormat\s*=(?!=)' 'WordEditor signature reconciliation changes the compose body format.'
    Forbid-Pattern $interopSignatureSection '\bTryInsertAtSelection\b|\bTypeParagraph\b' 'WordEditor signature reconciliation falls back to the current selection.'
    Forbid-Pattern $interopSignatureSection 'insertionPosition\s*=\s*originalSelection' 'WordEditor signature reconciliation uses the current cursor as an insertion target.'
    Require-Order $interopSignatureSection 'ManagedEmailSignatureBookmarkName' 'OutlookAutoSignatureBookmarkName' 'Managed signature bookmark must be resolved before Outlook''s native signature slot.'
    Require-Order $interopSignatureSection 'OutlookAutoSignatureBookmarkName' 'OutlookOriginalMessageBookmarkName' 'Outlook''s native signature slot must be resolved before the reply quote boundary.'
    Require-Pattern $interopSignatureSection 'source\s*=\s*"document_end"' 'New-message fallback is not anchored at the document end.'
    Require-Pattern $interopSignatureSection 'source\s*=\s*"mail_original"' 'Reply/forward fallback is not anchored at _MailOriginal.'
    Require-Pattern $interopSignatureSection 'originalStart\s*-\s*OutlookOriginalMessageProtectedGap' 'Reply/forward fallback is not kept above Outlook''s _MailOriginal boundary.'
    Require-Pattern $interopSignatureSection 'source\s*=\s*"quote_boundary_unavailable"' 'Reply/forward reconciliation does not fail closed when no quote boundary exists.'
    Require-Pattern $interopSignatureSection 'ManagedEmailSignatureBookmarkName[\s\S]*TryDeleteEmailSignatureSlot' 'Managed insertion is not paired with exact old-slot deletion.'
}

$signatureReconciler = Get-CSharpMethodBlock $InteropSource 'ReconcileEmailSignatureWordSlot'
if ($null -eq $signatureReconciler) {
    Add-Failure 'ReconcileEmailSignatureWordSlot could not be parsed.'
} else {
    Require-Pattern $signatureReconciler 'bool\s+initialSlotManaged\s*=\s*hasInitialSlot\s*&&\s*string\.Equals\(\s*slotSource\s*,\s*"managed"' 'Signature reconciliation does not preserve the initial managed-slot identity before diagnostic source mutations.'
    Require-Order $signatureReconciler 'bool initialSlotManaged' 'EmailSignatureSlotPlacementPolicy.Resolve' 'Initial managed-slot identity must be captured before placement decisions mutate the diagnostic source.'
    Require-Pattern $signatureReconciler 'TryResolveSafeEmailSignatureInsertionPoint\(\s*context\s*,\s*bookmarks\s*,\s*isReplyOrForward\s*,\s*true\s*,\s*slotStart\s*,\s*slotEnd' 'Existing signature validation does not exclude the complete current slot from quote-border detection.'
    Require-Pattern $signatureReconciler 'if\s*\(\s*isReplyOrForward\s*&&\s*!hasSafeInsertionPoint\s*\)\s*\{[\s\S]{0,300}?return\s+result\s*;' 'Reply/forward reconciliation does not fail closed when an existing slot has no safe quote boundary.'
    Require-Pattern $signatureReconciler '\bEmailSignatureSlotPlacementPolicy\.Resolve\s*\(' 'Existing managed/_MailAutoSig slots bypass the tested placement policy.'
    Require-Pattern $signatureReconciler 'EmailSignatureSlotPlacementPolicy\.Resolve\(\s*isReplyOrForward\s*,\s*slotStart\s*,\s*slotEnd\s*,\s*safeInsertionPosition\s*,\s*quoteBoundaryPosition\s*,\s*hasMeaningfulTextBetween\s*\)' 'Signature placement conflates the protected insertion target with the actual quote boundary.'
    Require-Pattern $signatureReconciler 'if\s*\(\s*placementDecision\s*==\s*EmailSignatureSlotPlacementDecision\.MoveToSafeInsertionPoint\s*\)\s*\{[\s\S]{0,900}?insertionPosition\s*=\s*safeInsertionPosition\s*;[\s\S]{0,400}?safeFallback\s*=\s*true\s*;' 'A move decision does not actually select the safe insertion point.'
    Require-Pattern $signatureReconciler 'if\s*\(\s*placementDecision\s*==\s*EmailSignatureSlotPlacementDecision\.UnsafeQuoteBoundaryOverlap\s*\)\s*\{[\s\S]{0,1500}?return\s+result\s*;' 'A slot overlapping the quote boundary does not fail closed.'
    Require-Pattern $signatureReconciler 'if\s*\(\s*hasInitialSlot\s*&&\s*safeFallback\s*&&\s*insertionPosition\s*<\s*slotStart\s*\)' 'Managed and _MailAutoSig slots moved upward are not both tracked before mutation.'
    Require-Pattern $signatureReconciler 'if\s*\(\s*!TryAddEmailSignatureBookmark\(\s*context\.Document\s*,\s*bookmarks\s*,\s*previousSlotBookmarkName[\s\S]{0,500}?result\.Source\s*=\s*"initial_slot_track_failed"\s*;[\s\S]{0,150}?return\s+result\s*;' 'Failure to create the previous-slot bookmark does not abort before mutation.'
    Require-Order $signatureReconciler 'previousSlotBookmarkName = "NcSigPrevious"' 'TryAddMissingEmailSignatureLeadingParagraphs' 'The previous slot is not tracked before leading paragraphs can shift it.'
    Require-Pattern $signatureReconciler 'TryAddEmailSignatureBookmark\(\s*context\.Document\s*,\s*bookmarks\s*,\s*stagedBookmarkName\s*,\s*stagedMutationStart\s*,\s*insertedEnd\s*\)' 'The staged rollback bookmark does not include leading paragraphs.'
    Require-Pattern $signatureReconciler '!TryGetBookmarkRange\(\s*bookmarks\s*,\s*previousSlotBookmarkName\s*,\s*out\s+currentSlotStart\s*,\s*out\s+currentSlotEnd\s*\)[\s\S]{0,500}?TryRollbackStagedEmailSignatureMutation[\s\S]{0,500}?result\.Source\s*=\s*"initial_slot_track_lost"\s*;[\s\S]{0,150}?return\s+result\s*;' 'Loss of the previous-slot bookmark does not roll back and abort before replacement.'
    Require-Pattern $signatureReconciler 'TryDeleteEmailSignatureSlot\(\s*context\.Document\s*,\s*currentSlotStart\s*,\s*currentSlotEnd\s*,\s*slotIsTable' 'The old signature slot is deleted using stale absolute positions.'
    Require-Pattern $signatureReconciler 'finally\s*\{[\s\S]*?TryDeleteEmailSignatureBookmark\(\s*bookmarks\s*,\s*previousSlotBookmarkName\s*\)' 'The previous-slot tracking bookmark is not cleaned up in the reconciliation finally block.'
    Require-Pattern $signatureReconciler 'bool\s+replacingManaged\s*=\s*initialSlotManaged\s*;' 'Managed-slot replacement ownership is derived from a mutable diagnostic source instead of the preserved initial slot identity.'
    Forbid-Pattern $signatureReconciler 'replacingManaged\s*=\s*[^;]*(?:resolvedSlotSource|safeSource|slotSource)' 'Managed-slot replacement ownership depends on a logged or mutated source string.'
}

$safeInsertionResolver = Get-CSharpMethodBlock $InteropSource 'TryResolveSafeEmailSignatureInsertionPoint'
if ($null -eq $safeInsertionResolver) {
    Add-Failure 'TryResolveSafeEmailSignatureInsertionPoint could not be parsed.'
} else {
    Require-Pattern $safeInsertionResolver 'TryGetBookmarkRange\(\s*bookmarks\s*,\s*OutlookOriginalMessageBookmarkName[\s\S]*?quoteBoundaryPosition\s*=\s*originalStart\s*;[\s\S]*?originalStart\s*-\s*OutlookOriginalMessageProtectedGap[\s\S]*?source\s*=\s*"mail_original"\s*;' 'The _MailOriginal target does not keep the actual quote boundary separate from its protected insertion point.'
    Require-Pattern $safeInsertionResolver 'if\s*\(\s*TryFindInlineQuoteSeparatorStart\(\s*context\.Document\s*,\s*documentStart\s*,\s*hasExcludedRange\s*,\s*excludedStart\s*,\s*excludedEnd\s*,\s*out\s+separatorStart\s*\)\s*\)\s*\{[\s\S]*?quoteBoundaryPosition\s*=\s*separatorStart\s*;[\s\S]*?source\s*=\s*"quote_separator"\s*;[\s\S]*?return\s+true\s*;' 'The paragraph-border quote fallback is not active or does not preserve its actual quote boundary.'
}

Require-Pattern $SignaturePlacementSource 'slotStart\s*<=\s*quoteBoundaryPosition\s*&&\s*slotEnd\s*>\s*quoteBoundaryPosition' 'Signature overlap is not evaluated against the actual quote boundary.'
Require-Pattern $SignaturePlacementSource 'slotStart\s*>\s*quoteBoundaryPosition' 'A signature entirely below quoted content is not detected against the actual quote boundary.'
Require-Pattern $SignaturePlacementSource 'safeInsertionPosition\s*>\s*slotEnd\s*&&\s*hasMeaningfulTextBetween' 'Authored-content placement no longer uses the safe insertion target.'

$quoteSeparatorFinder = Get-CSharpMethodBlock $InteropSource 'TryFindInlineQuoteSeparatorStart'
if ($null -eq $quoteSeparatorFinder) {
    Add-Failure 'TryFindInlineQuoteSeparatorStart could not be parsed.'
} else {
    Require-Pattern $quoteSeparatorFinder 'hasExcludedRange\s*&&\s*(?:paragraphStart\s*<\s*excludedEnd\s*&&\s*paragraphEnd\s*>\s*excludedStart|paragraphEnd\s*>\s*excludedStart\s*&&\s*paragraphStart\s*<\s*excludedEnd)[\s\S]{0,200}?continue\s*;' 'The quote-separator fallback can mistake a border inside the current signature slot for the quote boundary.'
}

$meaningfulTextProbe = Get-CSharpMethodBlock $InteropSource 'TryHasMeaningfulEmailSignatureText'
if ($null -eq $meaningfulTextProbe) {
    Add-Failure 'TryHasMeaningfulEmailSignatureText could not be parsed.'
} else {
    Require-Pattern $meaningfulTextProbe 'InvokeMethod\(\s*wordEditor\s*,\s*"Range"\s*,\s*new\s+object\[\]\s*\{\s*start\s*,\s*end\s*\}\s*\)' 'Authored-content detection does not inspect exactly the range between the existing slot and safe target.'
    Require-Pattern $meaningfulTextProbe 'hasMeaningfulText\s*=\s*text\.Trim\([\s\S]*?\)\.Length\s*>\s*0\s*;' 'Authored-content detection does not ignore whitespace-only text.'
}

$signatureWordEditorOpen = Get-CSharpMethodBlock $InteropSource 'TryOpenEmailSignatureWordEditor'
if ($null -eq $signatureWordEditorOpen) {
    Add-Failure 'TryOpenEmailSignatureWordEditor could not be parsed.'
} else {
    Require-Pattern $signatureWordEditorOpen 'OutlookWordEditorContext\.TryOpenInline\([\s\S]*?composeKey\s*,\s*inlineExplorerIdentityKey\s*,\s*out\s+context\s*\)' 'Signature reconciliation does not pass the tracked Explorer identity to the targeted TryOpenInline overload.'
    Forbid-Pattern $signatureWordEditorOpen 'TryOpenInline\([\s\S]*?composeKey\s*,\s*out\s+context\s*\)' 'Signature reconciliation retains the untargeted TryOpenInline overload.'
    Require-Pattern $signatureWordEditorOpen 'if\s*\(\s*isInlineResponse\s*\|\|\s*activeInline\s*\)\s*\{[\s\S]*?TryOpenInline\([\s\S]*?inlineExplorerIdentityKey[\s\S]*?return\s+false\s*;\s*\}\s*if\s*\(\s*OutlookWordEditorContext\.TryOpenInspector' 'Tracked inline lookup can fall through to an unrelated Inspector Word editor.'
}

$htmlShareInsert = Get-CSharpMethodBlock $InteropSource 'InsertHtmlIntoMail'
if ($null -eq $htmlShareInsert) {
    Add-Failure 'InsertHtmlIntoMail could not be parsed.'
} else {
    Require-Order $htmlShareInsert 'TryInsertHtmlIntoInspectorWordEditor' 'TryInsertHtmlIntoMailBody' 'Inspector HTML sharing rewrites HTMLBody before attempting bookmark-preserving WordEditor insertion.'
}

# Separate password follow-up mail uses only a successful cached policy snapshot
# and applies a sender-matching managed signature before automatic send.
$passwordSignatureSnapshot = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'BuildSeparatePasswordSignatureSnapshot'
if ($null -eq $passwordSignatureSnapshot) {
    Add-Failure 'BuildSeparatePasswordSignatureSnapshot could not be parsed.'
} else {
    Require-Pattern $passwordSignatureSnapshot '\bTryGetCachedEmailSignaturePolicyStatus\(\s*configuration\s*,\s*out\s+policyStatus\s*\)' 'Separate password signature does not use the cached backend-policy snapshot.'
    Require-Pattern $passwordSignatureSnapshot 'policyStatus\s*==\s*null\s*\|\|\s*!policyStatus\.FetchSucceeded' 'Separate password signature accepts an unsuccessful backend-policy snapshot.'
    Require-Order $passwordSignatureSnapshot 'policyStatus.FetchSucceeded' 'snapshot.Active = true' 'Separate password signature is activated before the cached policy snapshot is verified.'
    Require-Pattern $passwordSignatureSnapshot 'EmailSignatureContentBuilder\.BuildPlainText\(\s*sanitized\s*\)' 'Separate password signature bypasses the shared plain-text content builder.'
    Forbid-Pattern $passwordSignatureSnapshot '\bFetchBackendPolicyStatus\s*\(|\bGetEmailSignaturePolicyStatusAsync\s*\(' 'Separate password signature performs a backend policy fetch instead of using the cache.'
    Forbid-Pattern $passwordSignatureSnapshot '\.Result\s*(?:[;,\)\]\}]|\?\?)|\.Wait\s*\(|GetAwaiter\(\)\.GetResult\(' 'Separate password signature blocks on a backend policy fetch.'
}

$passwordDispatch = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'DispatchSeparatePasswordMailQueue'
if ($null -eq $passwordDispatch) {
    Add-Failure 'DispatchSeparatePasswordMailQueue could not be parsed.'
} else {
    Require-Order $passwordDispatch 'ApplyAndVerifySeparatePasswordSender' 'ApplySeparatePasswordBackendSignature' 'Separate password automatic dispatch applies the backend signature before verifying the effective sender.'
    Require-Order $passwordDispatch 'ApplySeparatePasswordBackendSignature' ').Send();' 'Separate password automatic dispatch sends before applying the backend signature.'
}

$passwordSignatureApply = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'ApplySeparatePasswordBackendSignature'
if ($null -eq $passwordSignatureApply) {
    Add-Failure 'ApplySeparatePasswordBackendSignature could not be parsed.'
} else {
    Require-Pattern $passwordSignatureApply 'normalizedEffectiveSender\s*=\s*NormalizeSmtpAddress\(\s*effectiveSenderEmail\s*\)[\s\S]*?string\.Equals\(\s*normalizedEffectiveSender\s*,\s*signatureSnapshot\.UserEmail\s*,\s*StringComparison\.OrdinalIgnoreCase\s*\)' 'Separate password signature is not gated by the effective-sender/policy identity match.'
    Require-Order $passwordSignatureApply 'string.Equals' 'mail.Body = CombinePlainTextSegments' 'Separate password plain signature is written before the effective sender is matched to policy.'
    Require-Order $passwordSignatureApply 'string.Equals' 'mail.HTMLBody = AppendHtmlSignature' 'Separate password HTML signature is written before the effective sender is matched to policy.'
}

$sharedPlainSignature = Get-CSharpMethodBlock $SignatureContentSource 'BuildPlainText'
if ($null -eq $sharedPlainSignature) {
    Add-Failure 'EmailSignatureContentBuilder.BuildPlainText could not be parsed.'
} else {
    Require-Pattern $sharedPlainSignature 'return\s+"-- \\r\\n"\s*\+\s*normalized\.Replace\("\\n"\s*,\s*"\\r\\n"\)\s*\+\s*"\\r\\n"\s*;' 'Shared plain-text signature does not use the standard "-- " signature block.'
}

$managedHtmlSignature = Get-CSharpMethodBlock $SignatureContentSource 'BuildManagedHtml'
if ($null -eq $managedHtmlSignature) {
    Add-Failure 'EmailSignatureContentBuilder.BuildManagedHtml could not be parsed.'
} else {
    Require-Pattern $managedHtmlSignature '<div data-nc-connector-signature=\\"true\\">' 'Shared HTML signature lacks the managed signature wrapper.'
}

$passwordHtmlAppend = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'AppendHtmlSignature'
if ($null -eq $passwordHtmlAppend) {
    Add-Failure 'AppendHtmlSignature could not be parsed.'
} else {
    Require-Pattern $passwordHtmlAppend 'EmailSignatureContentBuilder\.BuildManagedHtml\(\s*sanitizedSignature\s*\)' 'Separate password HTML signature bypasses the shared managed wrapper.'
}

$passwordFallback = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'TryOpenSeparatePasswordFallback'
if ($null -eq $passwordFallback) {
    Add-Failure 'TryOpenSeparatePasswordFallback could not be parsed.'
} else {
    Require-Order $passwordFallback 'ApplySeparatePasswordBody' 'fallback.Display(false)' 'Separate password manual fallback must populate its body before display.'
    Require-Order $passwordFallback 'fallback.Display(false)' 'ApplySeparatePasswordBackendSignatureToDisplayedFallback' 'Separate password manual fallback must display before managed signature reconciliation.'
    $fallbackDisplayIndex = $passwordFallback.IndexOf('fallback.Display(false)', [StringComparison]::Ordinal)
    if ($fallbackDisplayIndex -lt 0) {
        Add-Failure 'Separate password manual fallback has no display boundary.'
    } else {
        $passwordFallbackBeforeDisplay = $passwordFallback.Substring(0, $fallbackDisplayIndex)
        Forbid-Pattern $passwordFallbackBeforeDisplay '\bApplySeparatePasswordBackendSignature(?:ToDisplayedFallback)?\s*\(|\bApplyManagedEmailSignature\s*\(|\bAppendHtmlSignature\s*\(' 'Separate password manual fallback appends a signature before Outlook displays and initializes the draft.'
    }
}

$passwordDisplayedFallbackSignature = Get-CSharpMethodBlock $ComposeShareLifecycleSource 'ApplySeparatePasswordBackendSignatureToDisplayedFallback'
if ($null -eq $passwordDisplayedFallbackSignature) {
    Add-Failure 'ApplySeparatePasswordBackendSignatureToDisplayedFallback could not be parsed.'
} else {
    Require-Pattern $passwordDisplayedFallbackSignature 'ResolveSeparatePasswordEffectiveSenderEmail\(\s*mail\s*,\s*composeKey\s*\)[\s\S]*?string\.Equals\(\s*effectiveSenderEmail\s*,\s*signatureSnapshot\.UserEmail\s*,\s*StringComparison\.OrdinalIgnoreCase\s*\)' 'Displayed password fallback is not gated by the effective-sender/policy identity match.'
    Require-Pattern $passwordDisplayedFallbackSignature 'dispatch\.IsPlainText\s*\?\s*signatureSnapshot\.PlainText\s*:\s*EmailSignatureContentBuilder\.BuildManagedHtml\(\s*signatureSnapshot\.Html\s*\)' 'Displayed password fallback does not provide plain content or the shared managed HTML wrapper to reconciliation.'
    Require-Order $passwordDisplayedFallbackSignature 'ResolveSeparatePasswordEffectiveSenderEmail' '_passwordMailInteropController.ApplyManagedEmailSignature' 'Displayed password fallback reconciles its signature before resolving the effective sender.'
}

# Compose policy reads use an asynchronous, keyed, last-known-good cache.
foreach ($required in @(
    @{ Pattern = 'EmailSignaturePolicyCacheLifetime\s*=\s*TimeSpan\.FromMinutes\('; Message = 'Signature policy cache has no bounded lifetime.' },
    @{ Pattern = 'Task<BackendPolicyStatus>\s+_emailSignaturePolicyFetchTask'; Message = 'Signature policy cache does not coalesce an in-flight fetch.' },
    @{ Pattern = '\bGetEmailSignaturePolicyStatusAsync\s*\('; Message = 'Asynchronous signature policy accessor is missing.' },
    @{ Pattern = '\bTryGetCachedEmailSignaturePolicyStatus\s*\('; Message = 'Send gate has no synchronous cached-policy accessor.' },
    @{ Pattern = 'string\.Equals\(_emailSignaturePolicyCacheKey,\s*cacheKey'; Message = 'Signature policy cache is not scoped to the active backend credentials.' },
    @{ Pattern = 'DateTime\.UtcNow\s*-\s*_emailSignaturePolicyCacheFetchedAtUtc\s*<=\s*EmailSignaturePolicyCacheLifetime'; Message = 'Signature policy cache lifetime is not applied to normal compose reads.' },
    @{ Pattern = 'return\s+_emailSignaturePolicyFetchTask\s*;'; Message = 'Concurrent signature policy reads do not join the in-flight request.' },
    @{ Pattern = '\bTask\.Run\s*\('; Message = 'Backend signature policy fetch is not moved off the Outlook UI thread.' },
    @{ Pattern = '\.ConfigureAwait\(false\)'; Message = 'Backend signature policy fetch captures the Outlook UI context.' },
    @{ Pattern = '\bfetched\.FetchSucceeded\b'; Message = 'Failed policy responses can overwrite the last successful cache value.' },
    @{ Pattern = 'using last successful snapshot'; Message = 'Policy cache has no logged last-known-good fallback.' }
)) {
    Require-Pattern $PolicySource $required.Pattern $required.Message
}

# Sending is gated before any success/cleanup state is armed and never waits on the network.
$onSendBlock = Get-CSharpMethodBlock $SendSource 'OnSend'
if ($null -eq $onSendBlock) {
    Add-Failure 'Mail compose OnSend handler could not be parsed.'
} else {
    Require-PatternBeforeLiteral $onSendBlock 'TryFinalizeEmailSignatureBeforeSend\s*\(\s*ref\s+cancel\s*\)' '_sendPending = true' 'Signature send gate must run before send-success tracking is armed.'
}

$signatureSendGate = Get-CSharpMethodBlock $SignatureSource 'TryFinalizeEmailSignatureBeforeSend'
if ($null -eq $signatureSendGate) {
    Add-Failure 'TryFinalizeEmailSignatureBeforeSend(ref bool cancel) is missing.'
} else {
    Require-Pattern $signatureSendGate '\bTryGetCachedEmailSignaturePolicyStatus\s*\(' 'Signature send gate does not use the cached policy snapshot.'
    Require-Pattern $signatureSendGate '\bBlockEmailSignatureSend\s*\(\s*ref\s+cancel\s*,' 'Signature send gate has no fail-closed cancellation path.'
    Forbid-Pattern $signatureSendGate '\bGetEmailSignaturePolicyStatusAsync\s*\(|\bFetchBackendPolicyStatus\s*\(' 'Signature send gate performs network policy work.'
    Forbid-Pattern $signatureSendGate '\.Result\s*(?:[;,\)\]\}]|\?\?)|\.Wait\s*\(|GetAwaiter\(\)\.GetResult\(' 'Signature send gate blocks on an asynchronous operation.'
}

$signatureSendBlocker = Get-CSharpMethodBlock $SignatureSource 'BlockEmailSignatureSend'
if ($null -eq $signatureSendBlocker) {
    Add-Failure 'BlockEmailSignatureSend(ref bool cancel, ...) is missing.'
} else {
    Require-Pattern $signatureSendBlocker '\bcancel\s*=\s*true\s*;' 'Signature send blocker does not cancel Outlook send.'
    Require-Pattern $signatureSendBlocker '\breturn\s+false\s*;' 'Signature send blocker does not report gate failure.'
}

if ($Failures.Count -gt 0) {
    foreach ($failure in $Failures) {
        Write-Host ("ERROR: " + $failure) -ForegroundColor Red
    }
    throw "Email signature architecture check failed with $($Failures.Count) issue(s)."
}

Write-Host "Email signature architecture OK: cached policy, send gate, and exact WordEditor slots are wired."
