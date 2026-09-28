# Operations Guide — NC Connector for Outlook

This guide is for administrators and operations teams that deploy and run **NC Connector for Outlook**. It covers prerequisites, rollout, managed configuration, server preparation, operating checks, logging, and incident handling.

Source layout, internal processing, protocol implementation, builds, and developer tests are documented in [DEVELOPMENT.md](DEVELOPMENT.md).

## Contents

- [Scope and responsibilities](#scope-and-responsibilities)
- [Requirements and rollout planning](#requirements-and-rollout-planning)
- [Deployment and application lifecycle](#deployment-and-application-lifecycle)
- [Managed configuration and user data](#managed-configuration-and-user-data)
- [Nextcloud server preparation](#nextcloud-server-preparation)
- [Optional NC Connector Backend](#optional-nc-connector-backend)
- [Feature operation](#feature-operation)
- [Internet Free/Busy Gateway (IFB)](#internet-freebusy-gateway-ifb)
- [Security and data handling](#security-and-data-handling)
- [Monitoring and support](#monitoring-and-support)
- [Troubleshooting runbooks](#troubleshooting-runbooks)

## Scope and responsibilities

NC Connector adds these functions to Outlook classic:

- Nextcloud file and folder shares from new mails, replies, and forwards
- attachment routing through Nextcloud
- Talk room creation and maintenance from Outlook appointments
- optional centrally managed email signatures
- optional Internet Free/Busy (IFB) through a local Nextcloud proxy

The usual division of responsibility is:

- **Windows/Outlook administration:** MSI rollout, add-in registration, managed registry values, client proxy and certificate trust, IFB URL reservations, and client logs
- **Nextcloud administration:** supported server version, required apps, public routing, Files Sharing, Talk, system address book, storage, and reverse-proxy limits
- **NC Connector Backend administration:** seat assignment, centrally managed defaults and locks, templates, signature assignments, and separate password delivery

## Requirements and rollout planning

### Client requirements

- 64-bit Windows 10 or Windows 11
- Outlook classic 2019 or newer; the new Outlook is not supported
- .NET Framework 4.7.2
- administrator rights for MSI installation

The MSI registers both 64-bit and 32-bit Outlook views. A 32-bit Outlook installation on 64-bit Windows is supported.

### Nextcloud requirements

- Nextcloud 32 or newer for every add-in function
- Files Sharing for uploads and public shares
- Talk for meeting functions
- Nextcloud Secrets plus NC Connector Backend for one-time Secret-link password delivery
- the Nextcloud system address book for user search, participant defaults, and moderator selection

The optional NC Connector Backend is not required for local sharing, Talk, or IFB. It is required for central policies, managed signatures, and separate password delivery.

The Nextcloud Password Policy app is optional. When available, NC Connector reads its password requirements; otherwise it creates passwords with its local generator.

### Network requirements

Clients need HTTPS access to the configured public Nextcloud base URL, including its OCS and WebDAV paths. Preserve a public subpath such as `/nextcloud`, but do not add `/index.php` to the URL configured in NC Connector.

NC Connector rejects explicit HTTP URLs, credentials embedded in a URL, and base URLs with a query or fragment. Dynamic login and password-policy endpoints must resolve to HTTPS on the same host, port, and scheme as the configured Nextcloud base URL.

Optional outbound destinations:

- `https://nc-connector.de/wp-json/ncc/v1/update-check` for daily release metadata
- GitHub release assets when an administrator or user opens a download link

IFB listens only on `127.0.0.1` and does not require an inbound firewall rule from other computers.

### Pre-deployment checklist

Before a broad rollout:

1. Record the intended Nextcloud base URL and whether it contains a subpath.
2. Verify Nextcloud 32 or newer and Files Sharing.
3. Verify the public certificate chain, DNS, proxy path, and TLS inspection policy from a representative workstation.
4. Complete the [Pretty URL test](#nextcloud-pretty-urls).
5. Enable Talk, Secrets, the system address book, and NC Connector Backend only where the corresponding functions are planned.
6. Assign a backend seat to a pilot user before testing centrally managed functions.
7. Keep the current and previous MSI available for rollout and recovery.
8. Define a pilot acceptance test covering installation, connection, one share, one Talk meeting, and every optional function in use.

## Deployment and application lifecycle

### Install

1. Close Outlook and wait until no `OUTLOOK.EXE` process remains.
2. Install the MSI with administrator rights.
3. Start Outlook.
4. Open **NC Connector -> Settings**, configure the Nextcloud connection, run the connection test, and save.

Interactive installation:

```powershell
msiexec.exe /i "NCConnectorForOutlook-<version>.msi"
```

Silent installation with an MSI log:

```powershell
msiexec.exe /i "NCConnectorForOutlook-<version>.msi" /qn /norestart /L*v "$env:TEMP\NCConnectorForOutlook-install.log"
```

Expected result:

- **NC Connector** appears in Outlook appointment and mail compose ribbons.
- `C:\Program Files\NC4OL\` exists.
- The add-in is listed under **File -> Options -> Add-ins -> COM Add-ins**.

If any check fails, use [Add-in does not load](#add-in-does-not-load).

### Pilot acceptance test

Run this test with a normal user account:

1. In **Settings**, test and save the Nextcloud connection.
2. Create a new mail and insert a small Nextcloud share.
3. Create a new appointment and insert a Talk link when Talk is in scope.
4. Test a reply or forward when sharing or managed signatures are in scope.
5. Test IFB from [the local endpoint](#internet-freebusy-gateway-ifb) when enabled.
6. Close and reopen Outlook, then repeat the connection test.

Record the add-in version, Outlook bitness, Nextcloud version, and result of each step.

### Upgrade or return to an older version

The MSI replaces an installed newer, equal, or older release. Per-user settings remain in the user profile.

At the first enabled IFB start after an upgrade, NC Connector migrates only its legacy `/nc-ifb/freebusy/...` Outlook values to the protected current path. Existing external Outlook Free/Busy values remain unchanged. This migration concerns per-user Outlook values; it does not require deleting the HTTP URL reservation with `netsh`.

The add-in reports release metadata but does not install updates. **Settings -> Advanced -> Inform me about new versions** controls the popup; the daily metadata check still runs when the popup is disabled. MSI approval and deployment remain administrator tasks.

1. Close Outlook.
2. Back up `%LOCALAPPDATA%\NC4OL\settings_*.xml`.
3. Install the target MSI over the existing installation.
4. Start Outlook and repeat the pilot acceptance test for the functions in use.

To return to the previous add-in release, repeat the same procedure with the previous MSI. If settings must also be restored, close Outlook first and restore only the files belonging to the same Windows user. A protected app password copied to another Windows account or computer may not be readable; authenticate again in that case.

### Uninstall

1. While Outlook is still installed, open **NC Connector -> Settings -> IFB**.
2. Disable IFB and save. This restores the previously recorded Outlook Free/Busy path.
3. Close Outlook and wait until no `OUTLOOK.EXE` process remains.
4. Use **Windows Settings -> Apps -> Installed apps** or:

```powershell
msiexec.exe /x "NCConnectorForOutlook-<version>.msi" /qn /norestart
```

The MSI removes installed files, add-in registration, and the default IFB URL reservation. Per-user settings, caches, and logs under `%LOCALAPPDATA%\NC4OL\` remain so that reinstalling does not erase user configuration.

Delete that profile directory only when its settings and logs are no longer needed. A URL reservation created manually for a custom IFB port is not owned by the MSI; remove it separately as described under [Custom IFB port](#custom-ifb-port).

### Installed paths and registration checks

Default installation directory:

```text
C:\Program Files\NC4OL\
```

Primary add-in registration:

```text
HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

32-bit Outlook on 64-bit Windows reads:

```text
HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn
```

`LoadBehavior` should be `3`. The MSI also writes `HKLM\Software\NC4OL\HttpUrl` as an installation marker for the default IFB reservation.

## Managed configuration and user data

### Profile data

Settings are stored per Windows user and Outlook profile:

```text
%LOCALAPPDATA%\NC4OL\settings_<OutlookProfile>.xml
```

If Outlook does not expose a profile name, the add-in uses:

```text
%LOCALAPPDATA%\NC4OL\settings_default.xml
```

The app password is stored as `AppPasswordProtected` with Windows Data Protection for the current user. It is not a portable credential.

Every successful settings save is written to a validated temporary file in the same directory and then replaces the primary file. The previous valid file remains as `settings_<OutlookProfile>.xml.bak`. If only the protected password is unreadable, NC Connector keeps the other settings, clears the password, and blocks automatic settings writes until the user explicitly saves a corrected configuration.

Pending Talk room deletions and IFB registry ownership are stored separately per Outlook profile below `%LOCALAPPDATA%\NC4OL\`. These state files use Windows Data Protection and keep a backup next to the primary file. Do not copy protected state to another Windows account.

Older `settings.ini` files below these directories are migrated on first start and removed only after a successful migration:

```text
%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\
%LOCALAPPDATA%\NextcloudTalkOutlookAddIn\
```

### Backup and restore

To back up a client profile:

1. Close Outlook.
2. Copy `settings_*.xml` and `settings_*.xml.bak` from `%LOCALAPPDATA%\NC4OL\`.
3. If pending Talk room deletions must survive the backup, copy `talk-room-lifecycle-*.dat*` and `ifb-registry-state-*.dat*` too. They can be restored only for the same Windows user.
4. Record the Windows user and Outlook profile name.

To restore it:

1. Close Outlook.
2. Restore the matching primary and backup files for the same Windows user.
3. Start Outlook and run the connection test.
4. If authentication fails, use the Nextcloud login flow again.

Logs and the IFB address-book cache are operating data, not required for configuration recovery. Omitting the Talk lifecycle files abandons pending Talk-room deletion retries.

### Rollout and pre-seeding

Use the managed registry policy below for the Nextcloud URL. Use the optional backend for central feature defaults and locks.

If a profile XML must be pre-seeded:

- deploy it only when no profile file exists
- use a file created by the same add-in release as the template
- include only stable defaults needed by the organization
- remove `AppPasswordProtected` before distribution
- let each user authenticate through the Nextcloud login flow
- validate the result against every Outlook profile naming pattern in the rollout

Do not copy a protected app password between users or computers.

### Managed Nextcloud URL

Supported policy locations:

```text
HKLM\Software\Policies\NC Connector
HKCU\Software\Policies\NC Connector
```

Values:

- `NextcloudUrl` (`REG_SZ`): full public Nextcloud URL
- `NextcloudUrlLocked` (`REG_DWORD` or string, optional): `1` or `true` locks the field

Example for a machine policy:

```powershell
$policyPath = "HKLM:\Software\Policies\NC Connector"
New-Item -Path $policyPath -Force | Out-Null
New-ItemProperty -Path $policyPath -Name NextcloudUrl -PropertyType String -Value "https://cloud.example.com" -Force | Out-Null
New-ItemProperty -Path $policyPath -Name NextcloudUrlLocked -PropertyType DWord -Value 1 -Force | Out-Null
```

Priority and result:

- `HKLM` takes priority over `HKCU`.
- Both 64-bit and 32-bit registry views are read.
- An unlocked value fills only an empty profile.
- A locked value is used for every profile and disables the URL field.
- Credentials remain user-specific.

After policy deployment, restart Outlook and verify the URL and lock state in **NC Connector -> Settings**.

## Nextcloud server preparation

### Base service checks

For a pilot user:

1. Open **NC Connector -> Settings**.
2. Enter the public Nextcloud URL.
3. Authenticate through the login flow or with an app password.
4. Run the connection test.

The test must report a supported Nextcloud version. An older server or a response without a usable version is rejected.

Verify the optional functions separately:

- create a public share for Files Sharing
- create a Talk room for Talk
- search for a user after enabling the system address book
- open backend-managed settings with an assigned seat

### Nextcloud Pretty URLs

Pretty URLs are a server-wide Nextcloud routing requirement. They affect authentication, files, apps, Talk, and other routes; they are not a Talk-only or add-in setting.

NC Connector creates a public Talk URL in this form:

```text
https://cloud.example.com/call/<TOKEN>
```

If the room works only as `https://cloud.example.com/index.php/call/<TOKEN>`, the web server or reverse proxy does not route Pretty URLs correctly. Fix the public route; do not add `/index.php` to the URL configured in NC Connector.

For a Nextcloud installation below `/nextcloud`, the expected URL is:

```text
https://cloud.example.com/nextcloud/call/<TOKEN>
```

#### Quick test

At the web root, open:

```text
https://cloud.example.com/index.php/login
https://cloud.example.com/login
```

Below `/nextcloud`, open:

```text
https://cloud.example.com/nextcloud/index.php/login
https://cloud.example.com/nextcloud/login
```

The URL without `/index.php` must reach Nextcloud or redirect to its login page. A web-server 404 means that the rewrite is not active.

#### Nginx

Use the complete Nextcloud Nginx configuration as the baseline. The following snippets belong in the matching existing `server` and PHP/FastCGI locations; do not create duplicate locations.

At the web root:

```nginx
location / {
    try_files $uri $uri/ /index.php$request_uri;
}
```

In the PHP/FastCGI location:

```nginx
fastcgi_param front_controller_active true;
```

Below `/nextcloud`, the fallback must retain that path:

```nginx
location /nextcloud {
    try_files $uri $uri/ /nextcloud/index.php$request_uri;
}
```

Validate and reload:

```bash
sudo nginx -t
sudo systemctl reload nginx
```

#### Apache

Apache must load `mod_rewrite` and `mod_env`. The HTTP user must be able to write Nextcloud's `.htaccess`, and the matching `<Directory>` block must allow those rules with `AllowOverride All`.

On Debian or Ubuntu:

```bash
sudo a2enmod rewrite env
sudo systemctl reload apache2
```

For Nextcloud at the web root, set in `config/config.php`:

```php
'overwrite.cli.url' => 'https://cloud.example.com/',
'htaccess.RewriteBase' => '/',
```

For Nextcloud below `/nextcloud`:

```php
'overwrite.cli.url' => 'https://cloud.example.com/nextcloud',
'htaccess.RewriteBase' => '/nextcloud',
```

Behind a reverse proxy, `htaccess.RewriteBase` is relative to the backend Apache `DocumentRoot` after proxy mapping. If the proxy removes `/nextcloud` before forwarding, use `/`.

Regenerate `.htaccess` with the real installation path:

```bash
cd /var/www/nextcloud
sudo -E -u www-data php occ maintenance:update:htaccess
sudo systemctl reload apache2
```

Only after checking the modules, `AllowOverride`, rewrite base, and regenerated `.htaccess`, the following Nextcloud fallback can be tested:

```php
'htaccess.IgnoreFrontController' => true,
```

Run `maintenance:update:htaccess` again and reload Apache.

Repeat the login test and open a newly created `/call/<TOKEN>` link from a client outside the server network.

Official Nextcloud references:

- [Nginx configuration](https://docs.nextcloud.com/server/32/admin_manual/installation/nginx.html)
- [Apache installation and Pretty URLs](https://docs.nextcloud.com/server/32/admin_manual/installation/source_installation.html#pretty-urls)
- [`maintenance:update:htaccess`](https://docs.nextcloud.com/server/32/admin_manual/occ_command.html#maintenance-commands)

### System address book

The system address book is required for:

- moderator selection in the Talk wizard
- the **Add users** default
- the **Add guests** default
- IFB address resolution

Enable it in **Nextcloud Administration settings -> Groupware -> System Address Book**. Also enable username autocompletion or system address-book access under **Administration settings -> Sharing**.

If the administration page shows it as enabled but clients still cannot use it:

```bash
sudo -E -u www-data php occ config:app:delete dav system_addressbook_exposed
sudo -E -u www-data php occ config:app:set dav system_addressbook_exposed --value="yes"
sudo -E -u www-data php occ dav:sync-system-addressbook
```

Then verify the generated address book for a test user:

```text
https://<cloud>/remote.php/dav/addressbooks/users/<user>/z-server-generated--system?export
```

Expected result: user search and moderator controls become available after Outlook reconnects. A complete, valid address-book export also works if a reverse proxy incorrectly returns HTTP 404 or an incorrect content type. Other HTTP errors, such as 401 or 403, still require fixing access to the endpoint. An empty HTTP 404 response is not an address book.

If a refresh fails or the response is damaged, Outlook shows an address-book error and keeps the affected controls unavailable until a successful retry. Previously cached contacts are retained, but are not presented as a successful refresh. Participant synchronization stops instead of treating unresolved internal users as guests. Users with a valid Nextcloud UID but no email address remain available in user search and moderator selection; matching an email recipient still requires an email address.

Official Nextcloud references:

- [System address book](https://docs.nextcloud.com/server/32/admin_manual/groupware/contacts.html#system-address-book)
- [`dav:sync-system-addressbook`](https://docs.nextcloud.com/server/32/admin_manual/occ_command.html#sync-system-address-book)

## Optional NC Connector Backend

### Prerequisites and operating states

Backend-managed functions require:

- the `ncc_backend_4mc` app installed and enabled
- client access to `/apps/ncc_backend_4mc/api/v1/status`
- an active seat assigned to the current Nextcloud user
- a policy domain supported by the installed backend version

Observed behavior by state:

- **No backend configuration:** Sharing, Talk, and IFB use local settings. Central signatures and separate password delivery are unavailable.
- **Reachable backend with active seat:** Saved Outlook values control editable fields. Locked backend values override the local value and cannot be changed in Outlook.
- **Reachable backend without a usable seat:** Sharing and Talk use local settings; Outlook displays the seat or license state. Central signatures and separate password delivery are unavailable.
- **Backend temporarily unreachable:** Sharing and Talk use saved local settings. A matching message that requires a central signature can remain open and unsent until the signature policy can be checked.
- **Backend lacks the signature domain:** Share and Talk policies continue to work. Central signatures stay disabled and Outlook displays an update notice.

### License notices in Outlook

Settings, the sharing wizard and the Talk dialog use the same status notice. An active license has no license warning. During the grace period, a yellow notice shows how long Pro features remain available; it does not disable them for users with an active assigned seat.

Once grace has ended, Outlook distinguishes an expired license from an inactive or invalid license, an activation problem, and a missed offline verification deadline. A suspended seat is reported separately. Users without a seat continue to see the seat-assignment hint. Basic sharing and Talk remain usable with local settings.

Full Nextcloud administrators see license notices even without their own seat and can open **Manage license in backend**. The link opens NC Connector's administration page on the configured Nextcloud. Other users are directed to their Nextcloud administrator. Older backends that do not report these details receive a general notice instead of a guessed expiry reason or an administration link.

An administrator without a personal seat also sees that sharing and Talk remain usable with local settings. A tooltip on a blocked seat feature names the missing assignment, even when the banner also reports grace or a synchronization problem. Administrative rights do not unlock personal features.

An active assigned Community seat has the same existing functions as an active assigned Pro seat. When assignments exceed capacity, only the excess suspended seats lose seat features; the remaining active users keep their policies and functions.

A failed license synchronization is not itself labelled as an invalid license: the cause can be a connection problem, an unusable server response, or unavailable local activation data. The notice includes the last successful synchronization and the offline deadline when supplied by the backend. A failure to retrieve the Nextcloud backend status has its own connection notice. Disabled-feature tooltips use the corresponding reason; there is no additional license pop-up on every action.

License activation remains part of the backend's normal synchronization. Outlook does not activate licenses or contact the license server directly. After correcting a license or seat assignment, reopen the affected dialog or refresh the connection in Settings to load the current backend status.

### Policy rollout

The backend can manage:

- Talk defaults and saved-appointment room deletion
- sharing defaults, password rules, and attachment link target
- share, password-mail, and Talk invitation templates
- separate password delivery and optional Secret links
- central signature assignment and separate switches for new mail, reply, and forward

Keep values editable when users may retain or change their saved Outlook setting. Lock only settings that users must not change. Test both a user with an active seat and a user without one before broad rollout.

### Template authoring

For custom share templates:

- use `{LINK_INTRO}` and `{LINK_LABEL}` where wording must follow the effective link target
- manual shares always use share-page wording
- attachment automation can use ZIP-download or share-page wording
- templates without these variables keep their existing wording
- place an optional field's fixed label and placeholder, such as `{PASSWORD}`, in the same `tr`, `p`, `li`, or `div`; Outlook removes that complete block when the value is empty
- use absolute `https://` links

For Talk appointment HTML:

- use tables (`table`, `tbody`, `tr`, `td`) for layout
- use simple inline styles
- avoid `flex`, `grid`, `border-radius`, `overflow`, `object-fit`, and `user-select`
- use explicit full `https://` links

Unsupported or unsafe HTML can be removed or the template can be rejected. Test templates in HTML/RTF Outlook appointments and in every supported Office theme.

### Managed signature acceptance test

A central signature is applied only when the effective Outlook **From** address matches the email address assigned by the backend. A shared mailbox or delegated **From** address must resolve to the same SMTP address.

Test this matrix before rollout:

1. New HTML mail with the matching identity.
2. New plain-text mail with the matching identity.
3. Reply and forward with the matching identity selected from the start.
4. Reply and forward after changing from a non-matching identity to the matching identity.
5. A non-matching identity.
6. A shared or delegated mailbox, if used.
7. A separate password follow-up mail.

Expected result:

- new mail places the signature after user-written text
- replies and forwards place it above the quoted message
- switching away from the matching identity removes only the managed signature
- a non-matching identity and its own signature remain unchanged
- separate new-mail, reply, and forward policy switches are followed
- a locked backend value cannot be changed in Outlook
- the follow-up password mail receives the signature only when its effective sender also matches

If a required final signature check cannot complete, Outlook keeps the message open instead of sending it with an unverified signature state.

The message distinguishes an unavailable signature policy from a signature that could not be updated safely. Check the Nextcloud connection if policy is unavailable, then retry sending. A previously confirmed policy for the same account remains usable after a failed refresh; a newly received refusal takes effect.

## Feature operation

### Sharing and uploads

The sharing wizard accepts local files and folders as well as existing content from the configured user's own Nextcloud. The queue shows the complete destination path, groups both sources, and reports entries, source count, total size, and the user's Nextcloud storage state. The **My Nextcloud** picker includes Explorer-style back, forward, up and refresh controls, clickable path segments, filtering, familiar file icons, file details, multi-selection feedback, storage information, and on-demand previews for images and documents supported by the Nextcloud server. Document previews such as PDF or Office files depend on the preview providers enabled on that server. This source works without NC Connector Backend.

Selected Nextcloud content is copied within the same account into the new share folder. The original remains unchanged and is not downloaded to Outlook for the transfer. For a preview, Outlook first requests a size-limited image generated by Nextcloud. If the server has no generated preview for a supported image file, Outlook can temporarily load that original image up to 5 MiB. Other original files are not downloaded for previews. Local content still uses the upload method supported by the server and selected files. Implementation details are in [DEVELOPMENT.md](DEVELOPMENT.md#sharing-flow-mail-compose).

Operating limits and error behavior:

- symbolic links and junctions are rejected
- a source file changed after the initial scan stops the upload
- the configured account needs read access to selected Nextcloud content and write access to the destination folder
- an existing manual share-root name stops that share; attachment automation can select a numbered name
- HTTP `507` means that Nextcloud has insufficient storage
- proxy timeouts and request-size limits can affect uploads even when the client and Nextcloud are otherwise healthy

For a large-folder incident, collect the selected item count, total size, timestamp, displayed phase, Nextcloud storage state, reverse-proxy limits, and `FILELINK` log entries.

### Attachment automation

In **Settings -> Sharing -> Attachments**, administrators or backend policy can select:

- always route attachments through NC Connector
- offer NC Connector above a size threshold
- `ZIP download` or `Nextcloud share page` as the attachment link target

`ZIP download` is the default when neither a local value nor a backend value is available. The link-target setting applies only to attachment automation; manually created shares always link to the Nextcloud share page.

Attachment rules older than five minutes continue to apply while they are refreshed in the background. Sending is not interrupted solely because of that age. If no rules have been loaded yet, or settings have just changed, an information message asks the user to try sending again in a moment. The mail remains open and its attachments are unchanged. A separate warning explains when the effective settings actually require NC Connector sharing; neither notice reports an upload failure.

Both attachment targets remain read-only shares. If a valid ZIP-download URL cannot be derived from the public share, insertion stops with an error. NC Connector does not label a normal share-page URL as a ZIP download.

Outlook or Exchange can reject a large attachment before an add-in event runs. In that case, users must select **Insert Nextcloud share** and add the file directly in the sharing wizard.

### Unsent mail and share cleanup

After a share block is inserted, NC Connector tracks the newly created server share until Outlook successfully saves the message after that insertion. If the compose window is discarded before a successful save, NC Connector removes the exact server folder with the account context captured when the share was created.

- Saving or automatically saving a draft keeps the share. Sending the message, delayed delivery, and offline Outbox delivery also keep it.
- Cancelling the close action or moving an inline reply into its own window does not remove the share.
- Deleting a draft that Outlook has already saved is not tracked. Remove its unused share manually in Nextcloud.
- If the share block cannot be inserted, the wizard reports failure and immediately attempts to remove the newly created server folder with the captured account context.
- An enforcing attachment policy blocks sending while a normal attachment that should have been routed through NC Connector remains in the message.
- Separate password delivery is submitted directly when the user clicks **Send**, as described below.

### Separate password delivery

Separate password delivery requires NC Connector Backend and an active seat.

- The primary mail contains no plain password.
- Separate password delivery is tied to the open compose session in which the share was created. The user must click **Send** from that same, still-open message. Saving or AutoSave does not interrupt the flow while the compose window remains open.
- Saving the primary mail as a draft or `.oft` template and then closing it is not supported. Reopening that draft, restarting Outlook before the first send attempt, or creating a message from that template restores the visible share block, but not the password follow-up state. No password follow-up is created; create a new share in the final message before sending it.
- Clicking **Send** finalizes and immediately submits the password follow-up with the primary mail's effective Outlook account. NC Connector does not save an intermediate password draft or inspect Drafts, Outbox, or Sent folders.
- This direct boundary does not wait for delayed or offline primary delivery. The password follow-up can therefore be submitted while the primary mail is still in the Outbox or if Outlook rejects it after NC Connector's send callback.
- If sender verification or automatic submission definitely fails, Outlook opens a fully prepared message for manual sending. An ambiguous Outlook submission is not repeated automatically.
- With the Secrets mode, one one-time Secret link is created per final recipient.
- Equal SMTP addresses across To, Cc, and Bcc receive only one Secret follow-up.
- If Secrets creation fails, Outlook uses the plain password follow-up and displays a warning.
- A matching backend signature is included only when the follow-up sender matches the assigned signature address.

### Talk room lifecycle

Deleting a saved Outlook appointment removes its remote Talk room only when the setting is explicitly enabled and the appointment contains NC Connector room metadata. The setting is disabled by default. A Talk URL copied into a location or body is not sufficient for remote deletion.

Deleting one occurrence or an exception from a recurring appointment does not remove the shared room. Only a non-recurring appointment or the series master can queue room deletion.

NC Connector reacts only to Outlook events for the individual Talk appointment. Opening or selecting a Talk appointment binds that appointment to the deletion event, including the current calendar selection restored after an Outlook restart. Deletion from an open appointment and from the calendar view therefore uses the same checks and queue. Outlook startup does not enumerate stores or calendar folders and does not scan calendar items.

Queued room deletions are stored per Outlook profile and retried in the background after temporary Nextcloud failures or an Outlook restart. The cleanup of a newly created room from an unsaved, discarded appointment remains active.

Moderator delegation rejects the current Nextcloud user when the wizard input matches the canonical user ID, the configured login, or the known primary email address. A successful handoff to another user still causes the original moderator to leave the room.

Before enabling saved-appointment room deletion across an organization:

1. Create and save a Talk appointment with a pilot account.
2. Restart Outlook, then delete the still-unopened appointment directly from the calendar view.
3. Verify that the remote room is removed.
4. Repeat with deletion from the open appointment, then document room recovery or recreation procedures for users.

## Internet Free/Busy Gateway (IFB)

### Purpose and activation

IFB lets Outlook request Nextcloud free/busy data through a local HTTP endpoint.

1. Verify the Nextcloud system address book.
2. Open **NC Connector -> Settings -> IFB**.
3. Enable IFB and select the number of days, cache duration, and local port. Defaults are 30 days, 24 cache hours, and port `7777`.
4. Save and restart Outlook.

Reserved listener namespace:

```text
http://127.0.0.1:7777/nc-ifb/
```

The MSI reserves the default URL namespace for authenticated Windows users. NC Connector adds a random path segment to Outlook's Free/Busy URL; requests without that segment return `404`. The secret path is managed internally and is intentionally not shown in this guide.

Enabling IFB updates only Outlook's per-user Free/Busy values. Existing values and their types are recorded separately. Disabling IFB restores a value only while it still contains the value written by NC Connector; later administrator or application changes are left untouched. Values below `Software\Policies` are read for conflicts but are never written. The address-book cache is separated by Outlook profile, full Nextcloud base URL including any subpath, and configured login. The address-book request itself uses the canonical Nextcloud user ID.

The listener runs only while Outlook is running, IFB is enabled, and the stored Nextcloud credentials are complete.

### Verify the default reservation

```powershell
netsh http show urlacl | Select-String -Pattern "127.0.0.1:7777/nc-ifb"
Test-NetConnection 127.0.0.1 -Port 7777
```

Then create a test meeting in Outlook, add an address from the Nextcloud system address book, and open **Scheduling Assistant**. Expected result: the reservation exists, the TCP test succeeds, and Outlook displays free/busy data. A direct request to the public `/nc-ifb/freebusy/...` path must return `404`.

### Custom IFB port

Valid configured ports are `1024` through `49151`. The MSI creates a reservation only for port `7777`. For another port, open an elevated PowerShell window and add a reservation for authenticated users:

```powershell
netsh http add urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/ sddl="D:(A;;GX;;;AU)"
```

Verify it:

```powershell
netsh http show urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/
```

If the port is changed again or the add-in is removed, delete the manually created reservation:

```powershell
netsh http delete urlacl url=http://127.0.0.1:<ifb-port>/nc-ifb/
```

Do not grant the reservation to `Everyone` (`S-1-1-0`).

## Security and data handling

- Use HTTPS for the Nextcloud base URL.
- Keep the workstation certificate store, proxy trust, and Windows TLS policy current.
- App passwords are protected for the current Windows user; never distribute protected credential blobs.
- Backend templates are treated as managed content. Review them before rollout and limit edit rights in Nextcloud.
- Secret-link keys remain in the URL fragment. Treat a full Secret link as confidential.
- IFB binds to loopback only. A custom URL reservation grants execute rights to authenticated local users, not to anonymous or remote users.
- Attachment cleanup can delete server data created for an unsent mail. It does not treat arbitrary public URLs as deletion targets.
- Log anonymization is enabled by default. Review every log before sharing it outside the organization.

The daily update request sends product, installed version, channel, and a rotating anonymous client hash. It does not send the Nextcloud URL, email address, username, app password, license key, or tenant content. Downloads link directly to GitHub release assets.

## Monitoring and support

### Routine operating checks

After an add-in, Outlook, proxy, certificate, or Nextcloud update:

1. Run the Settings connection test.
2. Create and open a small share.
3. Create and open a Talk link when Talk is enabled.
4. Test user search when the system address book is used.
5. Test every locked backend setting with a seat-assigned user.
6. Run the signature acceptance cases in use.
7. Test the local IFB endpoint when enabled.
8. Open **Settings -> Advanced -> Check now** and review the displayed release metadata.

### Logs

Enable logging under **NC Connector -> Settings -> Debug**. Keep **Anonymize logs** enabled unless support specifically requests other data.

Daily files:

```text
%LOCALAPPDATA%\NC4OL\addin-runtime.log_YYYYMMDD
```

Common categories:

- `CORE`: startup, settings, and registration
- `API`: Nextcloud requests and status codes
- `TALK`: room and appointment operations
- `FILELINK`: scan, upload, share, cleanup, and password follow-up
- `IFB`: listener, cache, and Free/Busy requests

Runtime errors are written even when debug logging is disabled. With debug logging active, the file also contains operating decisions and periodic upload progress. Logs retain the latest seven daily files and also remove files older than 30 days when possible.

### Support package

For a reproducible incident:

1. Enable debug logging.
2. Record local time, add-in version, Outlook version and bitness, Windows version, Nextcloud version, and relevant app versions.
3. Reproduce the problem once.
4. Copy only the affected time window from the latest log.
5. Include the user-visible message, HTTP status if shown, and exact operating step.
6. Remove app passwords, authorization values, private links, full message bodies, recipient lists, and customer data before sharing.

## Troubleshooting runbooks

### Add-in does not load

1. Open **Outlook -> File -> Options -> Add-ins**.
2. Check **COM Add-ins** for `NcTalkOutlook.AddIn`.
3. Check **Disabled Items** and re-enable the add-in if Outlook disabled it after a crash.
4. Verify `LoadBehavior=3` in the registry path matching Outlook bitness.
5. Verify `C:\Program Files\NC4OL\NcTalkOutlookAddIn.dll` exists.
6. Close Outlook and repair or reinstall the MSI.

If the add-in still does not load, collect the MSI log, Windows Event Viewer entries for Outlook/.NET, and the matching registry path.

### Connection or TLS test fails

1. Open the configured base URL from the affected workstation.
2. Check DNS, system time, certificate trust, proxy authentication, and TLS inspection.
3. Confirm that the URL contains the public subpath but not `/index.php`.
4. In **Settings -> Advanced -> Transport security (TLS)**, test the organization-approved mode.
5. Run the connection test again.
6. Compare the result from a workstation outside the affected proxy segment.

Do not apply machine-wide TLS registry changes until the certificate, proxy, and Windows Schannel policies have been reviewed.

### Pretty URL or Talk link returns 404

Run the [`/login` comparison](#quick-test). If only the `/index.php/login` URL works, repair the Nginx, Apache, or reverse-proxy rewrite and test again from outside the server network.

### Upload remains at zero or fails

1. Note whether the wizard displays scanning, folder preparation, or upload.
2. Wait for the local scan to finish before judging network throughput; a large source tree can spend time in the scan phase.
3. Check for symbolic links, junctions, inaccessible files, or files being modified during upload.
4. Check Nextcloud free storage; HTTP `507` means insufficient storage.
5. Review proxy request-body limits, timeouts, and WebDAV handling.
6. Reproduce with one small file, one large file, and then the original folder.
7. Collect `FILELINK` entries for the affected time window.

### Attachment automation does not start

1. Verify the configured **Attachments** mode and threshold.
2. Check whether Outlook or Exchange rejected the file before it appeared in the compose window.
3. Check for another Outlook add-in that handles large attachments.
4. Use **Insert Nextcloud share** as the supported path when the host blocks the attachment before NC Connector receives it.

### Backend policy or seat is not applied

1. Confirm that `ncc_backend_4mc` is installed and enabled.
2. Confirm that the affected Nextcloud user has an active assigned seat.
3. Check client access to `/apps/ncc_backend_4mc/api/v1/status`.
4. Open Settings and review the displayed backend or seat state.
5. Confirm whether the setting is a default or a locked value.
6. Compare with another seat-assigned pilot user.

Share and Talk can use local settings during a backend outage. Managed signatures and separate password delivery require a valid backend state.

### Managed signature is missing or misplaced

1. Confirm the backend seat and signature assignment.
2. Compare the effective Outlook **From** SMTP address with the assigned backend email address.
3. Check the separate switches for new mail, reply, and forward.
4. Repeat the matching cases in the [signature acceptance test](#managed-signature-acceptance-test).
5. Test once with Outlook's own signature enabled and once without it.
6. Collect `CORE` and relevant compose log entries without sharing the signature HTML.

Do not work around a blocked final signature check by copying unknown HTML into the message. Restore backend access or correct the sender/policy assignment.

### User search or moderator selection is disabled

1. Verify the system address-book and sharing/autocomplete settings.
2. Run `occ dav:sync-system-addressbook`.
3. Test the generated address-book URL for the affected user.
4. Restart Outlook and run the connection test.

### Settings cannot be loaded or saved

1. Close Outlook and copy `settings_*.xml` plus `settings_*.xml.bak` from `%LOCALAPPDATA%\NC4OL\` to a support directory.
2. Start Outlook and check whether the values were recovered from the backup.
3. If only the app password is empty, authenticate again; the other readable settings remain available.
4. If both files are invalid, enter the settings again and use the explicit **Save** action. Background writes remain blocked until that save succeeds.
5. If saving fails, check free disk space, access rights, endpoint-security blocks, and the `CORE` log entries. The dialog remains open and the active runtime configuration is not replaced.

Expected result: a successful explicit save creates a valid primary file and preserves the previous valid version as `.bak`.

### IFB does not respond

1. Confirm that IFB is enabled and credentials are complete.
2. Check the configured port.
3. Check the matching URL reservation.
4. Check whether another process owns the port:

```powershell
netstat -ano | Select-String ":<ifb-port>"
```

5. Run `Test-NetConnection 127.0.0.1 -Port <ifb-port>`.
6. Create a test meeting, add a known address from the Nextcloud system address book, and open **Scheduling Assistant**.
7. Review `IFB` log entries for the Outlook request and the upstream CalDAV result. A direct request without the internally managed path segment returning `404` is expected.

If a custom reservation has the wrong principal, delete it and recreate it with `D:(A;;GX;;;AU)`.
