<a id="betriebsanleitung--nc-connector-für-outlook"></a>
<a id="operations-guide--nc-connector-for-outlook"></a>
<a id="zweck-und-zuständigkeiten"></a>
<a id="scope-and-responsibilities"></a>

# NC Connector for Outlook – Administration

Installation, managed configuration, and troubleshooting for version 3.4.3.

## Contents

- [Requirements](#requirements)
- [Installation and sign-in](#installation-and-sign-in)
- [Enterprise Rollout](#enterprise-rollout)
- [Registry reference](#registry-reference)
- [Backend defaults and signatures](#backend-defaults-and-signatures)
- [Operating notes](#operating-notes)
- [Configure Internet Free/Busy](#configure-internet-freebusy)
- [Update, backup, and uninstall](#update-backup-and-uninstall)
- [Troubleshooting](#troubleshooting)
- [Logs and support](#logs-and-support)

<a id="voraussetzungen-und-rollout-planung"></a>
<a id="requirements-and-rollout-planning"></a>
<a id="client-voraussetzungen"></a>
<a id="client-requirements"></a>
<a id="nextcloud-voraussetzungen"></a>
<a id="nextcloud-requirements"></a>
<a id="nextcloud-server-vorbereiten"></a>
<a id="nextcloud-server-preparation"></a>

## Requirements

| Area | Requirement |
| --- | --- |
| Workstation | 64-bit Windows 10 or 11; .NET Framework 4.7.2 |
| Outlook | Outlook classic 2019 or newer, 32-bit or 64-bit. The new Outlook is not supported. |
| Installation | Administrator rights; save open work and close Outlook in every Windows session when possible before setup. |
| Nextcloud | Version 32 or newer, reachable through HTTPS |
| Sharing | Nextcloud Files Sharing with public link shares permitted, enough storage, and write access to the target folder |
| Meetings | Nextcloud Talk |
| User search and availability | An exposed Nextcloud system address book |
| Managed settings and signatures, separate password mail, Enterprise Rollout | NC Connector Backend and a valid assigned NC Connector Seat for each user |
| One-time password links | The Nextcloud Secrets app in addition |

Without central management, sharing, Talk, and Internet Free/Busy can use local settings without the backend.

<a id="netzwerk-voraussetzungen"></a>
<a id="network-requirements"></a>

Use the public Nextcloud URL as the server address, for example `https://cloud.example.com` or `https://cloud.example.com/nextcloud`. Do not append `/index.php`, credentials, or query parameters. The workstation must trust the certificate, and the proxy and firewall must allow OCS and WebDAV traffic.

For update notices, the add-in queries `https://nc-connector.de/wp-json/ncc/v1/update-check`. MSI downloads come from GitHub. IFB operates only on the local computer and needs no inbound firewall rule for other computers.

<a id="bereitstellung-und-anwendungslebenszyklus"></a>
<a id="deployment-and-application-lifecycle"></a>
<a id="installation"></a>
<a id="install"></a>
<a id="inbetriebnahme"></a>
<a id="initial-setup"></a>

## Installation and sign-in

1. Save open work and close Outlook in every Windows session when possible.
2. Install the MSI with administrator rights.
3. Start Outlook and open **NC Connector → Settings**.
4. Enter the Nextcloud URL and sign in. Login Flow opens a browser; manual sign-in requires a username and app password.
5. Test the connection and save the settings.

If Outlook is still open, Windows Installer can offer to close it cleanly. An unattended installation attempts this automatically. If Outlook remains active, for example in another Windows session, setup stops before replacing the installed version. Close Outlook manually and run the installation again. Setup does not force termination or promise to restart Outlook automatically.

For an unattended installation with a log:

```powershell
msiexec.exe /i "NCConnectorForOutlook-3.4.3.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-install.log"
```

<a id="basisprüfungen"></a>
<a id="base-service-checks"></a>
<a id="regelmäßige-betriebsprüfungen"></a>
<a id="routine-operating-checks"></a>

NC Connector should then appear under **File → Options → Add-ins → COM Add-ins**. Before broad deployment, create a share or Talk link on one workstation with the intended configuration. After later server, proxy, or policy changes, test the connection and the affected function again.

<a id="verwaltete-konfiguration-und-benutzerdaten"></a>
<a id="managed-configuration-and-user-data"></a>

## Enterprise Rollout

Registry settings can distribute the server address and settings, lock selected areas, and hide the Settings entry.

**As soon as any value from the registry reference is present, every user needs the backend and a valid assigned NC Connector Seat.** This also applies to disabled, empty, or invalid values. When upgrading to 3.4.3, existing registry URL settings also activate this mode. Prepare the backend and seat assignments before deployment.

<a id="checkliste-vor-dem-rollout"></a>
<a id="pre-deployment-checklist"></a>
<a id="rollout-und-vorbelegung"></a>
<a id="rollout-and-pre-seeding"></a>

### Set up the rollout

1. Assign seats to the intended users in the Nextcloud backend and configure the required policies.
2. Distribute the required registry values through software deployment or Group Policy.
3. Install or update the MSI and restart Outlook.
4. Test initial sign-in with an intended user account.

For a mostly prepared sign-in, set `NextcloudUrl`, `NextcloudUrlLocked=1`, and `AuthMode=LoginFlow`. To hide the main tab and Settings, also set `ShowMainRibbonTab=0`. The Share and Talk buttons remain available in messages and appointments.

<a id="verwaltete-anmeldeart"></a>
<a id="managed-sign-in-method"></a>

### Sign in when Settings is hidden

If credentials are missing, **Insert Nextcloud share** or **Insert Talk link** opens the sign-in dialog. With `AuthMode=LoginFlow` and a matching registry URL, browser sign-in starts automatically.

When `AuthMode=LoginFlow` is a valid managed value, a successful sign-in and connection test saves the credentials and closes the dialog, including when the user started browser sign-in manually. Without that managed value, the user saves the settings. With `AuthMode=Manual`, the user enters a username and app password.

The original action then resumes if the message or appointment is still open. A failed sign-in can be retried in the open dialog.

Unless `ShowMainRibbonTab=0`, the main tab and Settings remain visible. A managed URL alone hides nothing.

<a id="registry-vorgaben-zurücknehmen"></a>
<a id="removing-registry-overrides"></a>

### Change or remove managed values

Registry changes take effect after an Outlook restart. To show Settings again, set `ShowMainRibbonTab` to `1` or remove the value.

To return a setting to local control, remove its managed value from every registry path and view in use. For TLS, logging, and IFB, remove the entire group. Otherwise, a lower-priority value can become active.

Enterprise Rollout ends only when none of the listed values remains. Stored credentials and local settings remain. Removing the URL value does not restore an earlier server address automatically.

Enterprise Rollout limits NC Connector functions, not ordinary Outlook mail. The explicit `SendPolicyFailureMode=failclosed` setting additionally prevents sending while central sending requirements are still unknown. A confirmed missing, paused, or invalid Seat disables protected functions but does not block ordinary mail. This is not a data-loss-prevention policy.

<a id="registry-übersicht"></a>

## Registry reference

Set values under one of these paths:

```text
HKLM\Software\Policies\NC Connector
HKCU\Software\Policies\NC Connector
```

Priority for each value: **HKLM 64-bit → HKLM 32-bit → HKCU 64-bit → HKCU 32-bit**. The first present value wins, even when it is invalid. The URL lock must be in the same path and registry view as its URL.

For switches, use `REG_DWORD` with `0` for off and `1` for on. The strings `false` and `true` are also accepted. Selections require `REG_SZ`; numeric IFB values require `REG_DWORD`.

<a id="verbindung-und-anmeldung"></a>
<a id="connection-and-sign-in"></a>
<a id="oberfläche-und-standardwerte"></a>
<a id="interface-and-defaults"></a>
<a id="verwaltete-nextcloud-url"></a>
<a id="managed-nextcloud-url"></a>

### Connection and interface

| Value | Type / values | Effect | Without a managed value |
| --- | --- | --- | --- |
| `NextcloudUrl` | `REG_SZ`, HTTPS URL | Fills an empty server address. Replaces a stored address only together with the URL lock. | Stored address; empty for new profiles |
| `NextcloudUrlLocked` | `REG_DWORD`, `0` / `1` | Locks the address supplied by `NextcloudUrl`. Without a valid URL in the same path, no URL lock is applied. | URL is editable |
| `AuthMode` | `REG_SZ`, `LoginFlow` / `Manual` | Selects the sign-in method and locks the selection. | Stored selection; `LoginFlow` for new profiles |
| `ShowMainRibbonTab` | `REG_DWORD`, `0` / `1` | `0` hides the main tab and Settings button. | Main tab is visible |
| `DefaultsSource` | `REG_SZ`, `local` / `backend` | Selects and locks the default-values source. An explicit backend setting takes priority. | Backend setting, then user selection, then `local` |

Automatic sign-in starts only from Share or Talk when credentials are incomplete. The active URL must match the registry URL. Opening Settings normally does not start sign-in or close the dialog automatically.

<a id="versand-bei-dienstausfällen"></a>
<a id="sending-during-service-outages"></a>

### Sending during service outages

| Value | Type / values | Effect | Without a managed value |
| --- | --- | --- | --- |
| `SendPolicyFailureMode` | `REG_SZ`, `failopen` / `failclosed` | Determines whether sending can continue when Nextcloud or the backend is unavailable and a signature or attachment rule applies. Users cannot change this setting. | `failopen` |

An invalid value uses `failopen` and produces a configuration notice in Settings and the log. Presence activates Enterprise Rollout even with `failopen` or invalid data. Registry changes require an Outlook restart.

| Situation for the current message | `failopen` | `failclosed` |
| --- | --- | --- |
| Confirmed no applicable signature or attachment rule | Send normally without a policy warning. | Send normally without a policy warning. |
| Policy state is unknown, for example on the first start while the server is unavailable | Send normally; do not invent a signature or attachment warning. | Block only when Send is clicked, until the central requirements can be checked. Drafts can still be saved. |
| An applicable rule is known, but Nextcloud or the backend is unavailable | Send with a non-modal warning when the rule cannot be fulfilled; unshared files remain normal attachments. A correctly usable cached signature is applied without an unnecessary warning. | Block sending until availability is confirmed and the rule is fulfilled. Cached policy alone does not replace this check. |
| A known applicable rule is not fulfilled and a service check is still pending | Block until the rule is fulfilled or an outage is confirmed. A fulfilled cached signature permits sending without a warning. | Block until the check succeeds and the rule is fulfilled. |
| Services are available, but a required signature cannot be inserted or attachments still need sharing | Block sending with an explanation and a corrective action. | Same behavior. |

Rules are evaluated for each message, not for the whole Outlook session. A signature applies only to its matching sender and enabled message type. An attachment rule applies only when **Always use NC Connector** is effective or a centrally mandatory threshold is exceeded. A local optional upload offer is not a mandatory threshold.

No sending-policy dialog appears at Outlook startup. Relevant notices appear while composing, replying, forwarding, adding affected attachments, or using Share or Talk; blocking errors appear on Send. A failed Share or Talk action does not impose requirements on later ordinary messages. Authentication rejection, missing permissions, insufficient storage, and local processing errors are not outage exceptions. HTTP `429` is a temporary request limit; observe any server-provided waiting time. A blocked message is never resent automatically.

<a id="verwaltete-transportsicherheit-tls"></a>
<a id="managed-transport-security-tls"></a>

### TLS

If at least one TLS value is present, all three TLS settings are locked. Missing group values receive the defaults from the table. Without managed TLS values, stored local values remain active; a new profile uses the same defaults.

| Value | Type / values | Effect | Default |
| --- | --- | --- | --- |
| `TransportTlsUseSystemDefault` | `REG_DWORD`, `0` / `1` | At `1`, Windows selects the TLS version and the two version switches are ignored. | `0` |
| `TransportTlsEnable12` | `REG_DWORD`, `0` / `1` | Allows TLS 1.2 when the system default is off. | `1` |
| `TransportTlsEnable13` | `REG_DWORD`, `0` / `1` | Allows TLS 1.3 when Windows and the runtime support it. | `0` |

Do not set all three values to `0`: no server connection will be possible. Setting only `TransportTlsEnable12=0` has the same result because of the group defaults.

<a id="logging-und-update-benachrichtigungen"></a>
<a id="verwaltetes-logging"></a>
<a id="managed-logging"></a>
<a id="verwaltete-update-benachrichtigungen"></a>
<a id="managed-update-notifications"></a>

### Logging and update notifications

The two logging values form one group. If either value is present, both controls are locked and a missing value receives its default. Without managed logging values, local settings apply; a new profile has debug logging off and anonymization on.

| Value | Type / values | Effect | Default |
| --- | --- | --- | --- |
| `DebugLoggingEnabled` | `REG_DWORD`, `0` / `1` | Enables detailed logs. Errors are logged even without debug logging. | `0` |
| `LogAnonymizationEnabled` | `REG_DWORD`, `0` / `1` | Anonymizes personal data in logs. Passwords and tokens are masked regardless of this setting. | `1` |
| `UpdateNotifyEnabled` | `REG_DWORD`, `0` / `1` | Controls and locks update notices; independent of the logging group. | Local selection; `0` for a new profile |

`UpdateNotifyEnabled=0` disables only the notice, not the daily version check or **Check now**. The add-in does not install updates.

<a id="ifb"></a>
<a id="verwaltetes-internet-freebusy-ifb"></a>
<a id="managed-internet-freebusy-ifb"></a>

### Internet Free/Busy

If at least one IFB value is present, all four values, including the address-book cache duration under **Advanced**, are managed and locked. Missing group values receive the defaults from the table. Without managed values, local settings apply.

| Value | Type / values | Effect | Default |
| --- | --- | --- | --- |
| `IfbEnabled` | `REG_DWORD`, `0` / `1` | Enables Nextcloud availability in Outlook. | `0` |
| `IfbDays` | `REG_DWORD`, `10`, `30`, `60`, `90` | Number of availability days requested | `30` |
| `IfbCacheHours` | `REG_DWORD`, `1`–`24` | System-address-book cache lifetime in hours; also used by Talk. | `24` |
| `IfbPort` | `REG_DWORD`, `1024`–`49151` | Local port; ports other than `7777` need a separate URL reservation. | `7777` |

Setting only a port does not enable IFB. During initial setup without managed values, the sign-in dialog may preselect IFB; check the intended setting before saving.

<a id="optionales-nc-connector-backend"></a>
<a id="optional-nc-connector-backend"></a>
<a id="policy-rollout"></a>

## Backend defaults and signatures

The backend manages Share and Talk settings, templates, signatures, and separate password delivery. Policies apply to users with a valid assigned seat. Administrator rights do not replace that assignment.

Leave **Editable in add-on** enabled when users may change a default. Disable it when the value must be mandatory. Select at least one day for expiration settings. Disable an attachment threshold with its switch rather than by entering zero; enabled thresholds support 1–10240 MB.

### Default values source

Under **Group settings → Default settings → General**, a Nextcloud administrator selects where Outlook obtains starting values for new actions. This setting cannot be delegated to group administrators.

It applies to sharing, Talk, attachment automation, the language of generated text, and signature switches. The signature template itself always comes from the backend.

| Backend setting | Result in Outlook |
| --- | --- |
| **Local** or **Backend**, not editable in the add-in | The backend selection applies and overrides a different registry value. |
| **Local** or **Backend**, editable in the add-in | Users can choose under **Advanced → Default values source**. Until then, the backend selection applies. |
| **No default**; also an older backend without this setting | Registry `DefaultsSource` applies. If it is also absent, the user selection applies, otherwise **Local**. |

With **Local**, stored local values are used first. If no local selection exists, backend values and then product defaults follow. With **Backend**, the order is backend, local selection, and product default.

Individually locked backend values apply regardless of this selection. Editable fields can still be changed in the wizard for the current action.

With **Backend** as the source, the **Share**, **Talk link**, and **Signature** settings tabs are disabled. Previously stored local values remain. The source selection is unavailable without a valid seat. Older add-in versions do not apply the new backend source setting.

After a backend change, reopen Settings or the affected wizard. To release the source selection, enable **Editable in add-on** or select **No default** in the backend, and remove the registry value as well.

<a id="verwaltete-signaturen-einrichten"></a>
<a id="configure-managed-signatures"></a>

### Signatures

1. Assign a signature template to the user in the backend.
2. Verify that the assigned email address matches the actual **From** address in Outlook. This also applies to shared mailboxes and delegated senders.
3. Enable the signature for new messages, replies, and forwards as required.
4. Open a message with the intended sender address and check the rendered result.

If duplicate signatures appear, check Outlook's own signature configuration. Select the language of the share block under **Share** and the language of the Talk description under **Talk link**.

<a id="vorlagen-erstellen"></a>
<a id="template-authoring"></a>

### Custom templates

Outlook renders simple HTML with tables and inline styles most reliably. Use absolute HTTPS links and avoid web layouts based on Flexbox or Grid. Test custom templates in Outlook before deployment.

In share templates, `{LINK_INTRO}` and `{LINK_LABEL}` adapt to the selected link target. Place the label and placeholder of an optional field, such as `{PASSWORD}`, in the same HTML paragraph or table row. This prevents a lone label when the value is empty.

<a id="funktionsbetrieb"></a>
<a id="feature-operation"></a>

## Operating notes

<a id="freigaben-und-uploads"></a>
<a id="sharing-and-uploads"></a>

### Files and attachments

**My Nextcloud** copies selected files and folders to a new share folder. The originals remain unchanged. This requires read access to the source, write access to the destination, and enough storage. Symbolic links and junctions are not supported.

<a id="anhangsautomatisierung"></a>
<a id="attachment-automation"></a>

Configure attachment automation under **Share → Attachments**. It can always route attachments through NC Connector or offer NC Connector above a size threshold. The link target can be ZIP download or share page; ZIP download is the default when no value is set. Manual shares always link to the share page.

Outlook or Exchange can reject an attachment before NC Connector receives it. In that case, select the file directly through **Insert Nextcloud share**. A locked central size threshold makes sharing mandatory when exceeded; a locally configured upload offer does not. While services are available, a mandatory attachment rule blocks sending while an affected file remains attached normally. During service outages, [SendPolicyFailureMode](#sending-during-service-outages) applies.

Original attachments remain until their files have been shared successfully and the share link has been inserted into the message. Cancelling the wizard or a failed upload or insertion preserves them. If files are removed from the wizard selection, only the originals actually shared are removed; later additions and unrelated same-name files remain.

<a id="ungesendete-mail-und-freigabebereinigung"></a>
<a id="unsent-mail-and-share-cleanup"></a>

When an unsaved message is discarded, NC Connector tries to remove its newly created share folder. Saved drafts retain their shares. If a saved draft is deleted later, remove unused shares manually in Nextcloud.

### Separate password delivery

Separate password delivery starts when the user clicks **Send** and uses the same Outlook account as the primary message. Note these limits:

- **Keep the message open until sending.** After closing and reopening a draft or restarting Outlook, the share must be created again. A visible share block does not restore password delivery. This also applies to templates that already contain a share block. Saving and AutoSave are supported while the message remains open.
- **With delayed delivery, the password may arrive first.** The password message does not wait for the primary message in the Outbox. It may already have been sent when Outlook subsequently rejects the primary message.
- If delivery clearly fails, a prepared password message opens for manual sending. An ambiguous result is not retried automatically to avoid duplicates.
- With Secrets, each recipient receives an individual one-time link. **If secret creation fails, the password is sent as plain text in the separate message and a warning is shown.**

Explain these limits to users when introducing the function.

<a id="talk-raum-lebenszyklus"></a>
<a id="talk-room-lifecycle"></a>

### Talk rooms and appointments

Deleting a saved Outlook appointment deletes its Talk room only when this is explicitly enabled under **Talk link**. It is disabled by default. Deleting the room affects every participant.

For recurring appointments, deleting one occurrence does not delete the shared room. A Talk link copied into an appointment does not trigger room deletion. After temporary connection errors, queued room deletions are retried later, including after an Outlook restart.

<a id="internet-freebusy-gateway-ifb"></a>
<a id="zweck-und-aktivierung"></a>
<a id="purpose-and-activation"></a>

## Configure Internet Free/Busy

IFB displays Nextcloud availability in Outlook's Scheduling Assistant. It needs the system address book, a signed-in account, and a running Outlook instance.

1. Expose the [system address book](#system-address-book) for the user.
2. Enable IFB under **Settings → IFB** and select the range and port; the cache duration is under **Advanced**. Alternatively, use the registry group.
3. Save and restart Outlook.
4. Add a known Nextcloud email address to an appointment and check its availability in the Scheduling Assistant.

<a id="standard-reservierung-prüfen"></a>
<a id="verify-the-default-reservation"></a>

The MSI configures the default port `7777`. With Outlook running, IFB enabled, and sign-in complete, check the local connection with:

```powershell
netsh http show urlacl url=http://127.0.0.1:7777/nc-ifb/
Test-NetConnection 127.0.0.1 -Port 7777
```

IFB is reachable only from the local computer. The Outlook URL contains an automatically managed secret path. A direct browser request without that path therefore returns `404` and is not a valid function test.

<a id="eigener-ifb-port"></a>
<a id="custom-ifb-port"></a>

### Use a different port

For another port between `1024` and `49151`, an administrator must create the URL reservation from an elevated PowerShell. Example for port `8888`:

```powershell
netsh http add urlacl url=http://127.0.0.1:8888/nc-ifb/ sddl="D:(A;;GX;;;AU)"
```

Then configure the same port in the add-in or through `IfbPort` and restart Outlook. The permission applies to authenticated Windows users; do not replace it with `Everyone`.

When the custom port is no longer used, remove the reservation you created:

```powershell
netsh http delete urlacl url=http://127.0.0.1:8888/nc-ifb/
```

## Update, backup, and uninstall

<a id="upgrade-oder-rückkehr-auf-eine-ältere-version"></a>
<a id="upgrade-or-return-to-an-older-version"></a>

### Update or return to the previous version

1. Save open work and close Outlook in every Windows session when possible. If it is still open, Windows Installer can offer to close it cleanly; an unattended installation attempts this automatically.
2. [Back up the profile data](#back-up-and-restore-profile-data).
3. Install the required MSI over the existing installation. Keep the previous MSI for recovery.
4. Start Outlook and test sign-in and the functions in use.

If Outlook remains active, for example in another Windows session, setup stops before removing the old version and cleaning up IFB. Close Outlook manually and run setup again. It does not force termination or promise to restart Outlook automatically.

User settings remain. To return to an older release, install its MSI in the same way and restore its matching backup if needed. Software deployment must not disable Windows Installer rollback.

<a id="profildaten"></a>
<a id="profile-data"></a>
<a id="sicherung-und-wiederherstellung"></a>
<a id="backup-and-restore"></a>

### Back up and restore profile data

Data is stored under `%LOCALAPPDATA%\NC4OL\`:

| Files | Contents |
| --- | --- |
| `settings_<OutlookProfile>.xml` and `.xml.bak` | Settings and the previous valid version; when no profile name can be determined, the file is named `settings_default.xml`. |
| `talk-room-lifecycle-*.dat*` | Pending Talk room deletions |
| `ifb-registry-state-*.dat*` | Previous registry values saved for IFB |
| `addin-runtime.log_YYYYMMDD` | Operational logs |

To create a backup, close Outlook and copy the settings and state files. Record the Windows user and Outlook profile. To restore, close Outlook and copy back the matching files.

The app password and state files are protected for the Windows user. Do not use them as templates for other accounts or computers. If the password cannot be read after restoration, sign in again. Logs and the address-book cache do not need to be restored.

If you distribute an XML template instead of registry values, create only profile files that do not yet exist and remove `AppPasswordProtected` first. Let each user sign in. An XML template alone does not activate Enterprise Rollout.

<a id="deinstallation"></a>
<a id="uninstall"></a>

### Uninstall and reinstall

Close Outlook in every session and uninstall the add-in through **Windows Settings → Apps**. The MSI removes program files, add-in registration, the default IFB reservation, and owned Free/Busy remnants. Installation, upgrade, and repair also remove such old IFB values even if the data directory was deleted earlier.

For an unattended uninstall with a log:

```powershell
msiexec.exe /x "NCConnectorForOutlook-3.4.3.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-uninstall.log"
```

User settings, caches, and logs under `%LOCALAPPDATA%\NC4OL\` remain. The add-in can therefore still be signed in after reinstalling. For a completely fresh setup, rename the directory while Outlook is closed and after taking a backup. This resets the configuration of every Outlook profile stored there; pending Talk room deletions will no longer run.

To restore an earlier Free/Busy provider, disable IFB before uninstalling while the saved previous values are still available. They cannot be reconstructed from a deleted data directory. Current third-party values and administrative policies are not removed. Remove custom port reservations [separately](#use-a-different-port).

<a id="runbooks-zur-störungsbehebung"></a>
<a id="troubleshooting-runbooks"></a>

## Troubleshooting

### Setup still reports that Outlook is open

1. Save any unsaved work in Outlook.
2. Close Outlook in every Windows session and check whether an `OUTLOOK.EXE` process is still running.
3. Close any remaining instance cleanly and run setup again. Do not force-terminate Outlook while unsaved content may remain.

Windows Installer can offer to close Outlook and attempts this automatically during unattended setup. If Outlook remains active, NC Connector blocks replacement of the old version; Outlook is not guaranteed to restart automatically.

<a id="add-in-wird-nicht-geladen"></a>
<a id="add-in-does-not-load"></a>
<a id="installationspfade-und-registrierungsprüfung"></a>
<a id="installed-paths-and-registration-checks"></a>

### Add-in or main tab is missing

1. Verify that Outlook classic is in use.
2. If only the main tab is missing, check `ShowMainRibbonTab`. At `0`, this is intended; Share and Talk must remain available in messages and appointments.
3. Under **File → Options → Add-ins**, check both **COM Add-ins** and **Disabled Items**. The add-in is named `NcTalkOutlook.AddIn`.
4. Check that `C:\Program Files\NC4OL\NcTalkOutlookAddIn.dll` exists and that `LoadBehavior` is `3` in the path for the Outlook bitness:

```text
64-bit Outlook: HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
32-bit Outlook: HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

If files or registration are missing, close Outlook and repair the MSI. If the problem remains, review the installation log and the Windows Event Viewer entries for Outlook/.NET.

<a id="verbindungs--oder-tls-test-schlägt-fehl"></a>
<a id="connection-or-tls-test-fails"></a>

### Sign-in or connection fails

1. Open the configured Nextcloud URL in a browser on the workstation. Check the URL, DNS, system time, certificate, and proxy.
2. For a TLS configuration notice, correct the registry values first and restart Outlook. Different credentials do not resolve a TLS error.
3. For HTTP `401`, sign in again; for `403`, check permissions and upstream access controls.
4. Run the connection test again. If only one workstation is affected, compare its proxy, certificate, and endpoint-security configuration with a working workstation.

Do not disable certificate validation or change computer-wide TLS settings as an experiment.

### A registry setting is not applied as expected

First check the path, registry view, data type, and spelling against the tables. Then restart Outlook. A higher-priority invalid entry is not replaced by a valid entry lower in the order.

| Affected setting | Behavior for an invalid value / action |
| --- | --- |
| URL, URL lock, main tab | An invalid URL is not used; an invalid lock does not lock the URL; an invalid ribbon value leaves the main tab visible. Correct the value in the active path. |
| Sign-in method | The selection remains locked to Login Flow and automatic start is disabled. Set `LoginFlow` or `Manual` as `REG_SZ`. |
| Default-values source | Without a higher-priority backend setting, locked `local` is used. Set `local` or `backend` as `REG_SZ`. |
| Sending during outages | An invalid `SendPolicyFailureMode` uses `failopen` and produces a configuration notice. Set `failopen` or `failclosed` as `REG_SZ`. |
| TLS | Invalid values or three disabled switches block server requests, including sign-in and update checks. Correct the values or selected TLS mode. |
| Logging | The affected field uses its default; the connection is not blocked. |
| Update notices | Notices remain off. The version check continues. |
| IFB | IFB is disabled. Invalid cache hours use 24 hours; other valid cache values remain active. |

<a id="backend-policy-oder-seat-wird-nicht-angewendet"></a>
<a id="backend-policy-or-seat-is-not-applied"></a>
<a id="voraussetzungen-und-betriebszustände"></a>
<a id="prerequisites-and-operating-states"></a>
<a id="lizenzhinweise-in-outlook"></a>
<a id="license-notices-in-outlook"></a>

### Backend notice, missing seat, or unexpected settings

| Notice / problem | What should you check? |
| --- | --- |
| Backend missing | Install or enable the Nextcloud app `ncc_backend_4mc`. Check access to `/apps/ncc_backend_4mc/api/v1/status`. |
| Seat missing, paused, or invalid | Check the assignment for the affected user and the license state in the backend. For a paused assignment, adjust the seat count or license capacity. |
| Check unavailable | Check the connection between Outlook and Nextcloud; this alone is not a license rejection. |
| License synchronization failed | In the backend, check the license-server connection and the last successful synchronization. |
| Grace period or activation problem | Open the backend license overview and address the action shown there. |
| Value or settings tab is locked | Check registry values, the default-values source, and **Editable in add-on** in the backend. |
| Old or unexpected starting value | Check the default-values source and local selection; reopen the dialog after changes. |

When licensed capacity is exceeded, first check which assignments are paused. Seats that remain active do not need to be reassigned.

Without Enterprise Rollout, Share and Talk remain available with local settings when the user has no valid seat; managed signatures and separate password delivery do not. With Enterprise Rollout, the backend and seat assignment must be available to use NC Connector. A short connection failure can be bridged by the most recently confirmed state.

<a id="benutzersuche-oder-verfügbarkeiten-fehlen"></a>
<a id="benutzersuche-oder-moderatorauswahl-ist-deaktiviert"></a>
<a id="user-search-or-moderator-selection-is-disabled"></a>

### System address book

If users are missing from search, moderator selection remains disabled, or Nextcloud availability is missing, check address-book access first:

1. In Nextcloud, enable exposure under **Administration settings → Groupware → System address book**. Also check the sharing and autocompletion rules under **Sharing** for the affected user.
2. Rebuild the system address book on the Nextcloud server. Run the command from the Nextcloud directory; adapt the HTTP user and command for container installations:

```bash
sudo -E -u www-data php occ dav:sync-system-addressbook
```

3. Test the export with the affected account:

```text
https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system?export
```

4. Test the connection in Outlook again and repeat the search. For IFB, also check the local connection below.

Expect a vCard address book, not an HTML error page. For `401`, check sign-in; for `403`, check access rights. Users without an email address can appear in search but cannot be matched by email address.

If Nextcloud shows the system address book as enabled but the export remains unavailable, inspect the stored server setting:

```bash
sudo -E -u www-data php occ config:app:get dav system_addressbook_exposed
```

When the system address book should be exposed to clients, the value must be `yes`. Correct a different value deliberately and rebuild the address book:

```bash
sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"
sudo -E -u www-data php occ dav:sync-system-addressbook
```

<a id="nextcloud-pretty-urls"></a>
<a id="kurztest"></a>
<a id="quick-test"></a>
<a id="pretty-url-oder-talk-link-liefert-404"></a>
<a id="pretty-url-or-talk-link-returns-404"></a>

### Talk link returns 404 in the browser

Compare `https://<cloud>/login` with `https://<cloud>/index.php/login`. If only the second address works, correct the rewrite rules in the web server or reverse proxy using the official Nextcloud configuration. For installations under `/nextcloud`, retain the subpath in both URLs. Do not add `/index.php` to the add-in server address as a workaround.

Then open the login page and a newly created Talk link from a workstation again. Back up the server configuration before making changes and validate it with the web server before reloading.

<a id="nginx"></a>

For **Nginx**, use the [official Nextcloud configuration](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html) instead of copying individual rewrite rules from unrelated configurations.

<a id="apache"></a>

For **Apache**, check the rewrite modules, `AllowOverride`, and the rewrite base; regenerate Nextcloud's `.htaccess` after changes. The steps are in the [Nextcloud guide to Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls).

<a id="upload-bleibt-bei-null-oder-schlägt-fehl"></a>
<a id="upload-remains-at-zero-or-fails"></a>
<a id="anhangsautomatisierung-startet-nicht"></a>
<a id="attachment-automation-does-not-start"></a>

### Upload or attachment automation fails

1. Check whether the wizard still shows the local file scan or has reached upload. Large folders can take time to scan before upload starts.
2. Check source-file readability, write access, and free Nextcloud storage. HTTP `507` indicates insufficient storage.
3. If only large files fail, check reverse-proxy size limits and timeouts.
4. If automation does not start, check **Share → Attachments**. If Outlook has already rejected the attachment, select the file directly in the sharing wizard.
5. Narrow down the problem with a small file and review the matching time range in the `FILELINK` log.

If sending is blocked because central requirements could not be checked, verify `SendPolicyFailureMode`, restore service availability, keep the message open, and send it again after the check succeeds. This notice does not mean that an upload failed. Original attachments remain after a cancelled or failed sharing attempt; review them before retrying.

<a id="verwaltete-signatur-fehlt-oder-steht-falsch"></a>
<a id="managed-signature-is-missing-or-misplaced"></a>

### Signature is missing or blocks sending

Check the seat, signature assignment, actual **From** address, and the switches for new messages, replies, and forwards. For duplicate signatures, also check Outlook's own signature configuration.

During an outage, sending follows [SendPolicyFailureMode](#sending-during-service-outages). A warning about a missing signature is shown only for a signature known to apply to that message, not merely because a policy lookup failed. With `failclosed`, an unknown policy also blocks Send until checked. If the required signature cannot be inserted while services are available, sending remains blocked in both modes. Keep the message open and provide the error message and matching log period to support.

<a id="einstellungen-können-nicht-geladen-oder-gespeichert-werden"></a>

### Settings cannot be loaded or saved

1. Close Outlook and back up the affected `settings_*.xml` and `.xml.bak` files.
2. Restart Outlook. If the primary file is damaged, the add-in tries to load the previous valid version.
3. If only the app password is missing, sign in again. If both files are damaged, enter the settings again and save them explicitly.
4. If saving still fails, check free disk space, access rights, and endpoint security, then review the `CORE` entries.

<a id="ifb-antwortet-nicht"></a>
<a id="ifb-does-not-respond"></a>

### IFB does not start or shows no data

1. Check activation, sign-in, and the registry group. For a centrally managed installation, also check the backend state.
2. Check the URL reservation and port as described under [Configure IFB](#configure-internet-freebusy). Specifying a custom port alone is not enough.
3. For a port conflict, identify the process using the port, for example for the default port:

```powershell
netstat -ano | Select-String ":7777"
```

4. If IFB starts but shows no availability, check the system address book and the email address in use. The `IFB` log contains the request and CalDAV result.
5. If **An unowned NC Connector IFB value already exists** appears after reinstalling, close Outlook in every session and repair the current MSI. Repair removes old owned IFB entries; do not manually delete values belonging to another provider.

<a id="monitoring-und-support"></a>
<a id="monitoring-and-support"></a>
<a id="logs"></a>

## Logs and support

Enable detailed logging under **Settings → Debug** and leave **Anonymize logs** enabled. When the main tab is hidden, administrators can use the two logging registry values.

Files are stored under `%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD`. Errors are recorded even without debug logging. Retention is limited to the latest seven daily files; files older than 30 days are also removed.

<a id="supportpaket"></a>
<a id="support-package"></a>

For a support request:

1. Record the time, add-in version, Outlook version and bitness, and Nextcloud and relevant app versions.
2. Reproduce the problem once with debug logging enabled.
3. Collect the action, error message, and matching log period. Include the MSI log for installation problems.
4. Before sharing, check for credentials, private links, recipient data, and customer data, even when anonymization is enabled.

`CORE` covers settings and startup, `API` server communication, `FILELINK` sharing, `TALK` meetings, and `IFB` availability. Return debug logging to the intended operating value after diagnosis.

<a id="sicherheit-und-datenverarbeitung"></a>
<a id="security-and-data-handling"></a>

Treat credentials and complete share or Secret links as confidential. Limit changes to backend templates and registry settings to authorized administrators. The daily version check transmits the product, version, channel, and a rotating anonymous client hash, but no Nextcloud credentials or message content.

Further information:

- [Support form](https://nc-connector.de/support/)
- [Nextcloud system address book](https://docs.nextcloud.com/server/32/admin_manual/groupware/contacts.html#system-address-book)
- [Nextcloud configuration for Nginx](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html)
- [Nextcloud configuration for Apache and Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls)
