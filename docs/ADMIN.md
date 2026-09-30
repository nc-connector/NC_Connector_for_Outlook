# Operations Guide — NC Connector for Outlook

This guide is for administrators and operations teams that deploy and run **NC Connector for Outlook**. It covers prerequisites, rollout, managed configuration, server preparation, operating checks, logging, and incident handling.

Source layout, internal processing, protocol implementation, builds, and developer tests are documented in [DEVELOPMENT.md](DEVELOPMENT.md).

## Contents

- [Scope and responsibilities](#scope-and-responsibilities)
- [Requirements and rollout planning](#requirements-and-rollout-planning)
- [Deployment and application lifecycle](#deployment-and-application-lifecycle)
- [Managed configuration and user data](#managed-configuration-and-user-data)
  - [Registry reference](#registry-reference)
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

The optional NC Connector Backend is not required for local sharing, Talk, or IFB in an unmanaged installation. It is required for central policies, managed signatures, separate password delivery, and Enterprise Rollout.

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
6. Assign an active backend Seat to the intended users before enabling Enterprise Rollout.
7. Keep the current and previous MSI available for rollout and recovery.
8. Plan initial setup with a representative user account and the functions the organization will use.

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

### Initial setup

1. Sign in to Nextcloud under the intended user account, run the connection test in **Settings**, and save.
2. Check the functions that account will use, such as inserting a share or Talk link. For centrally managed functions, assign an active Seat first.
3. If a function is unavailable or fails, follow the matching [troubleshooting runbook](#troubleshooting-runbooks).

### Upgrade or return to an older version

The MSI replaces an installed newer, equal, or older release. Per-user settings remain in the user profile.

Installation, direct upgrade and MSI repair remove obsolete NC Connector Free/Busy paths from local Windows user profiles. This also works if `%LOCALAPPDATA%\NC4OL\` was deleted or renamed. On the next start with IFB enabled, NC Connector registers the current path again. Other Free/Busy providers and administrative policies remain unchanged; no manual registry cleanup or `netsh` command is required.

Close Outlook in **all Windows sessions** before running setup. Setup stops if Outlook is still running; it does not terminate Outlook. Windows Installer rollback must be enabled.

The add-in reports release metadata but does not install updates. **Settings -> Advanced -> Inform me about new versions** controls the popup; the daily metadata check still runs when the popup is disabled. MSI approval and deployment remain administrator tasks.

1. Close Outlook.
2. Back up `%LOCALAPPDATA%\NC4OL\settings_*.xml`.
3. Install the target MSI over the existing installation.
4. Start Outlook, check the connection, and verify the functions affected by the update.

To return to the previous add-in release, repeat the same procedure with the previous MSI. If settings must also be restored, close Outlook first and restore only the files belonging to the same Windows user. A protected app password copied to another Windows account or computer may not be readable; authenticate again in that case.

### Uninstall

1. Close Outlook in all Windows sessions and wait until no `OUTLOOK.EXE` process remains.
2. Use **Windows Settings -> Apps -> Installed apps** or:

```powershell
msiexec.exe /x "NCConnectorForOutlook-<version>.msi" /qn /norestart
```

The MSI removes installed files, add-in registration, the default IFB URL reservation, and any remaining own Outlook Free/Busy paths. This cleanup does not require disabling IFB beforehand or retaining its data directory. Other providers' current values and administrative policies remain unchanged. Per-user settings, caches, and logs under `%LOCALAPPDATA%\NC4OL\` remain so that reinstalling does not erase user configuration.

If you want to restore a previously used external Free/Busy provider before uninstalling, disable IFB in Settings while its saved configuration still exists. The installer removes own leftovers; it cannot reconstruct an external provider's address from a deleted data directory.

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

The previous valid settings file remains as `settings_<OutlookProfile>.xml.bak` and can be used for recovery. If only the protected password is unreadable, NC Connector keeps the other settings. Authenticate again and save the corrected configuration to resume normal operation.

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

Use the managed registry policy below for Enterprise Rollout. Deploy and configure the backend and assign each user an active Seat before enabling the rollout. Community and Pro Seats provide the same access.

If a profile XML must be pre-seeded:

- deploy it only when no profile file exists
- use a file created by the same add-in release as the template
- include only stable defaults needed by the organization
- remove `AppPasswordProtected` before distribution
- let each user authenticate through the Nextcloud login flow
- deploy the file under the name of the intended Outlook profile

Do not copy a protected app password between users or computers.

### Registry reference

All values below belong under `HKLM\Software\Policies\NC Connector` or `HKCU\Software\Policies\NC Connector`.

- **Precedence:** HKLM 64-bit → HKLM 32-bit → HKCU 64-bit → HKCU 32-bit. For each value, the first present entry wins, even if invalid. Exception: `NextcloudUrlLocked` belongs to the URL in the same path and registry view.
- **Managed mode:** Any of these 15 values enables [Enterprise Rollout](#enterprise-rollout), including `0`, empty, or invalid values. Prepare the backend and assign active Seats before deployment. Community and Pro Seats provide the same access.
- **Local values:** “Local; new: …” means the saved user preference applies; the stated factory default applies only without a saved preference.
- **Group rule:** Any TLS, logging, or IFB value configures and locks its **whole group**. Missing values within that group use the listed factory defaults, **not** local preferences. Other groups are unaffected.
- **Changes:** Restart Outlook to apply registry changes. Only `ShowMainRibbonTab=0` hides the main tab and Settings button.

Use `REG_DWORD` with `0` / `1` for on/off values. Boolean strings such as `false` / `true` are also accepted. Numeric IFB settings require actual `REG_DWORD` values; `AuthMode` and `DefaultsSource` require `REG_SZ`.

| Key | Type / values | Effect and dependencies | When absent |
| --- | --- | --- | --- |
| `NextcloudUrl` | `REG_SZ`: public HTTPS Nextcloud base URL, including any installation subpath | Sets the server; replaces an existing local URL only when the URL is locked | Local URL; new profile: empty |
| `NextcloudUrlLocked` | On/off | `1`: locks the URL field; requires a valid `NextcloudUrl` in the same path and registry view | `0`: URL editable |
| `ShowMainRibbonTab` | On/off | `0`: hides the main tab and Settings; Share/Talk action buttons remain accessible | `1`: visible |
| `AuthMode` | `REG_SZ`: `LoginFlow` / `Manual` | Selects and locks the sign-in method; automatic start requires a matching registry URL, see [sign-in](#managed-sign-in-method) | Local; new: `LoginFlow`; no managed automatic start |
| `DefaultsSource` | `REG_SZ`: `local` / `backend` | Sets the source of editable defaults; an explicit backend setting wins, see [source selection](#default-values-source) | Backend setting, otherwise user choice, otherwise `local` |
| `TransportTlsUseSystemDefault` | On/off; TLS group | `1`: Windows selects TLS; both version switches have no effect | Local; new: `0` |
| `TransportTlsEnable12` | On/off; TLS group | Allows TLS 1.2 when system-default mode is off | Local; new: `1` |
| `TransportTlsEnable13` | On/off; TLS group | Allows TLS 1.3 when system-default mode is off; requires runtime/Windows support | Local; new: `0` |
| `DebugLoggingEnabled` | On/off; logging group | Enables detailed debug logging; errors are also logged without debug enabled | Local; new: `0` |
| `LogAnonymizationEnabled` | On/off; logging group | Anonymizes personal log values; password/token masking always remains active | Local; new: `1` |
| `UpdateNotifyEnabled` | On/off | Controls update notifications and locks the switch; does not disable daily retrieval or **Check now** | Local; new: `0` |
| `IfbEnabled` | On/off; IFB group | Enables IFB; requires sign-in, a valid active Seat, and a matching URL reservation | Local; new: `0` |
| `IfbDays` | `REG_DWORD`: `10`, `30`, `60`, `90`; IFB group | Free/busy time horizon in days | Local; new: `30` |
| `IfbCacheHours` | `REG_DWORD`: `1`–`24`; IFB group | Cache lifetime in hours; also applies to the shared Talk address-book cache | Local; new: `24` |
| `IfbPort` | `REG_DWORD`: `1024`–`49151`; IFB group | Local IFB port; another port needs its own [URL reservation](#custom-ifb-port) | Local; new: `7777` |

For example, setting only `TransportTlsEnable12=0` produces the invalid TLS combination `0 / 0 / 0`. Setting only `IfbPort` does **not** enable IFB: `IfbEnabled` remains `0`.

### Managed Nextcloud URL

Without a URL lock, the policy fills only an empty URL field. With the lock, it replaces each profile's URL. Credentials remain user-specific. `NextcloudUrlLocked` alone enables Enterprise Rollout but neither supplies a URL nor locks the field.

An invalid selected URL is not used; an invalid URL lock is treated as off. An invalid ribbon value leaves the main tab and Settings visible. Lower-priority registry entries do not replace an invalid higher-priority value.

### Managed sign-in method

With incomplete credentials, **Insert Nextcloud share** or **Insert Talk link** starts browser sign-in directly when `AuthMode=LoginFlow` and a valid registry `NextcloudUrl` are set. The actual URL must match that registry URL. A local URL alone, a standalone URL-lock value, or `Manual` does not trigger automatic login.

Opening ordinary Settings and having complete credentials do not trigger automatic login. Each sign-in dialog makes at most one automatic attempt; after failure or cancellation, the login button remains available for an explicit retry.

With a valid managed `AuthMode=LoginFlow`, initial Share or Talk sign-in saves automatically and closes the dialog after successful login and connection verification. This also applies after an explicit login retry. Only successful saving resumes the original action while its message or appointment is still open. Normally opened Settings never close automatically; `Manual` still requires **Save**.

Invalid `AuthMode` locks the selection to `LoginFlow` with a configuration hint but no automatic start. Explicitly clicking the login button remains possible.

### Default values source

`DefaultsSource` covers sharing, Talk, attachment automation, generated text languages, and signature switches. **local** prefers saved local choices, then editable backend defaults. **backend** prefers available backend values, then local choices. Individually enforced policies always win; editable values remain adjustable in the wizard for the current action. The signature template still comes from the backend.

Precedence with a valid active Seat:

1. An explicit backend source under **Group settings → Default settings → General**; full Nextcloud administrators only, not delegable. With **Editable in add-on**, a saved user source choice wins over this suggestion.
2. With **No preset**, or an older backend without this field: registry `DefaultsSource`, locking the source selector.
3. Without a registry value: user choice under **Settings → Advanced → Default values source**, otherwise `local`.

Without a valid active Seat, the selector is disabled and the source is local; Enterprise Rollout still requires valid access. With effective source `backend`, the **Sharing**, **Talk Link**, and **Signature** settings tabs are greyed out. An invalid registry source uses locked `local` with a configuration hint unless an explicit backend setting wins.

After backend changes, refresh the connection or reopen Settings. To return source selection to users, set the backend source to **No preset** and remove the registry override. Older clients ignore the new backend source setting.

### Managed transport security (TLS)

All three fields under **Settings → Advanced → Transport security (TLS)** are locked together. Without the TLS registry group, local preferences apply; a new profile uses system default off / TLS 1.2 on / TLS 1.3 off.

Invalid values or `0 / 0 / 0` block server HTTP requests, including sign-in and update checks, with a configuration hint. There is no silent TLS 1.2 activation or fallback when the runtime rejects a selected TLS version. Correct the policy and restart Outlook; changing credentials or Seats does not fix this error.

### Managed logging

Both fields under **Settings → Debug** are locked together. An invalid value uses the table's default for that field only and shows a configuration hint; valid values in the other field remain effective. The hint is logged even without debug enabled. Logging errors do not block connections; see [Logs](#logs).

### Managed update notifications

The switch under **Settings → Advanced** is locked. An invalid value uses `0` with a configuration hint. Update retrieval is unchanged; the add-in never installs updates automatically.

### Managed Internet Free/Busy (IFB)

Activation, days, and port under **Settings → IFB**, plus the cache lifetime under **Advanced**, are locked together. Sign-in and port preparation are described under [IFB](#internet-freebusy-gateway-ifb). Without a managed IFB group, initial setup can preselect IFB once when credentials become available and no user decision has been saved; the factory default is off.

An invalid value disables IFB with a configuration hint, not Share or Talk. Invalid cache hours make the shared address-book cache use 24 hours; a valid cache lifetime remains effective even when another IFB value is invalid.

### Removing registry overrides

Remove the relevant value from every applicable path and registry view; for TLS, logging, and IFB, remove the **whole group**. Restart Outlook. Removing only a higher-priority entry can expose a lower-priority setting. Saved local sign-in, source, TLS, logging, notification, and IFB choices remain intact during overrides. URL prefill, however, is not a backup of an earlier server address.

Other remaining keys keep Enterprise Rollout active. To leave it completely, remove all values listed in the table.

### Enterprise Rollout

Enterprise Rollout applies as soon as any value in the [registry reference](#registry-reference) is present. With all fifteen values absent, existing local behavior remains unchanged.

**Upgrade impact:** An existing registry-based URL deployment also enables this mode after upgrading. Assign Seats and prepare the backend before rollout; an XML-preseeded URL alone does not enable it.

- Only `ShowMainRibbonTab=false` hides the main Explorer tab together with its Settings button. With the value absent or `true`, both remain visible and the full settings dialog is available, including with a managed URL or URL lock.
- The Share and Talk buttons remain in their existing mail and appointment tabs, including inline replies. Their visibility does not depend on `ShowMainRibbonTab`.
- A confirmed, valid, personally active Seat enables NC Connector. A Community Seat is fully equivalent to a Pro Seat. Global overcapacity alone does not block an active Seat.
- A confirmed missing backend shows an installation/configuration message. A missing, paused or invalid Seat shows the managed-installation Seat message. An unconfirmed connection or response error shows a verification message, not a false missing-backend or missing-Seat message.
- Share, Talk, attachment automation, managed signatures and IFB require rollout access. Ordinary Outlook mail remains usable; this mode is not an Outlook send restriction or a data-loss prevention system.
- Previously confirmed access can remain available during a temporary connection failure. A newly confirmed refusal takes effect; an already-open wizard continues with the status available when it was opened.
- Losing a Seat does not delete existing files, shares or appointments. Cleanup of newly created, abandoned shares or rooms and already accepted password follow-ups continues.

First sign-in and verification:

1. Restart Outlook after deploying the registry values.
2. Without credentials, click **Insert Nextcloud share** in a message or **Insert Talk link** in an appointment. The existing settings dialog opens directly with a blue **Connect to Nextcloud** invitation, without a preliminary error. This also applies to unmanaged installations. Only with `ShowMainRibbonTab=false` are the other settings tabs unavailable; otherwise the full dialog is also accessible through **NC Connector -> Settings**.
3. Complete sign-in. With a valid managed `AuthMode=LoginFlow`, verified credentials are saved automatically and the dialog closes; otherwise click **Save**. The [managed sign-in method](#managed-sign-in-method) section explains automatic browser login. The managed URL and its lock remain effective; existing preferences are preserved. After saving successfully, the original action continues if its message or appointment is still open (an inline reply must still be active). Cancelling ends the action quietly. Signing in alone does not upload files or create a Talk room.
4. If access is unavailable after signing in, check the backend and the user's active Seat assignment. Rejected credentials reopen sign-in with a reminder; connection failures have a separate message. If the main tab is hidden, normal settings remain hidden too.

To show the tab and Settings again, remove the effective `ShowMainRibbonTab=false` policy or set it to `true`, then restart Outlook. This does not bypass managed access checks. To leave Enterprise Rollout, remove all fifteen trigger values from all applicable policy locations and registry views, then restart Outlook. Saved credentials and preferences remain.

Registry policy is an administrator deployment control, not protection against a user who can change that policy or replace the add-in. Protect the policy keys and deployment permissions accordingly.

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

The local fallbacks described in this section apply to unmanaged installations. With [Enterprise Rollout](#enterprise-rollout), the backend and an active assigned Seat are required; the managed-installation messages take precedence.

### Prerequisites and operating states

Backend-managed functions require:

- the `ncc_backend_4mc` app installed and enabled
- client access to `/apps/ncc_backend_4mc/api/v1/status`
- an active seat assigned to the current Nextcloud user
- a policy domain supported by the installed backend version

Observed behavior by state:

- **No backend configuration:** Sharing, Talk, and IFB use local settings. Central signatures and separate password delivery are unavailable.
- **Reachable backend with active seat:** The [default values source](#default-values-source) controls whether local or backend defaults take precedence for editable fields. Locked backend values always take precedence and cannot be changed in Outlook.
- **Reachable backend without a usable seat:** Sharing and Talk use local settings; Outlook displays the seat or license state. Central signatures and separate password delivery are unavailable.
- **Backend temporarily unreachable:** A previously confirmed status for the same account can retain the defaults source and policies. Without a usable confirmed status, Sharing and Talk use local settings in unmanaged installations. A matching message that requires a central signature can remain open and unsent until the signature policy can be checked.
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

Keep values editable when users may adjust them for an individual action. Use the [default values source](#default-values-source) to choose between local and backend defaults; lock individual settings only when users must not change them. Assign active Seats to the users who should receive central policies.

Settings, the Sharing and Talk wizards, attachment automation, and signature switches use the selected source without changing individual policy locks. An explicit local `false` or a saved product-default value remains a user choice. Saving credentials alone does not select untouched options. Administrative overrides never overwrite the saved local choice; it returns when local defaults apply again and the field is unlocked. Without usable Seat access, local settings remain available in unmanaged installations; Seat-only functions remain restricted.

New backend share lifetimes start at one day. A zero-day value from an older backend is interpreted consistently as one day. This does not alter existing shares or a locally saved choice to disable expiration. Attachment thresholds use 1–10240 MB; an explicit backend `null` disables the threshold, while a legacy backend zero retains the established 5 MB behavior.

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

Unsupported or unsafe HTML can be removed or the template can be rejected. Preview custom templates in the message formats and Office themes used by your organization before deploying them.

### Configure managed signatures

1. Assign an active Seat and a signature to the user in the backend.
2. Match the assigned email address to the effective Outlook **From** address. For shared mailboxes or delegated senders, the actual sender SMTP address must also match; the signed-in Nextcloud account alone is not sufficient.
3. Enable the signature for new messages, replies, and forwards as required. Lock the corresponding backend settings only when users must not change them.

In new messages, the signature appears after the user's text. In replies and forwards, it appears above the quoted message. Changing to a non-matching sender removes the managed signature; other senders and their own signatures remain unchanged. A separate password mail receives the managed signature only when its sender also matches.

Before rollout, open a message with the intended sender and check that the assigned signature appears in the expected position.

If a required final signature check cannot complete, Outlook keeps the message open instead of sending it with an unverified signature state.

The message distinguishes an unavailable signature policy from a signature that could not be updated safely. Check the Nextcloud connection if policy is unavailable, then retry sending. A previously confirmed policy for the same account remains usable after a failed refresh; a newly received refusal takes effect.

## Feature operation

### Sharing and uploads

Select the language of the sharing HTML block under **Settings -> Sharing**, below the sharing defaults. Existing choices remain unchanged; a locked backend language remains read-only.

The sharing wizard accepts local files and folders as well as existing content from the configured user's own Nextcloud. **My Nextcloud** shows files, folders, storage information, and previews. Document previews such as PDF or Office files depend on the preview providers enabled on that server. This source works without NC Connector Backend unless Enterprise Rollout is enabled.

Selected Nextcloud content is copied within the same account into the new share folder. The original remains unchanged and is not downloaded to Outlook for the transfer. For a preview, Outlook first requests a size-limited image generated by Nextcloud. If the server has no generated preview for a supported image file, Outlook can temporarily load that original image up to 5 MiB. Other original files are not downloaded for previews.

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

Outlook or Exchange can reject a large attachment before NC Connector can process it. In that case, users must select **Insert Nextcloud share** and add the file directly in the sharing wizard.

### Unsent mail and share cleanup

If a message is discarded before Outlook saves it after inserting a share, NC Connector removes the newly created share folder from the Nextcloud account used to create it.

- Saving or automatically saving a draft keeps the share. Sending the message, delayed delivery, and offline Outbox delivery also keep it.
- Cancelling the close action or moving an inline reply into its own window does not remove the share.
- Deleting a draft that Outlook has already saved is not tracked. Remove its unused share manually in Nextcloud.
- If the share block cannot be inserted, the wizard reports failure and attempts to remove the newly created share folder.
- An enforcing attachment policy blocks sending while a normal attachment that should have been routed through NC Connector remains in the message.
- Separate password delivery is submitted directly when the user clicks **Send**, as described below.

### Separate password delivery

Separate password delivery requires NC Connector Backend and an active seat.

- The primary mail contains no plain password.
- Separate password delivery is tied to the open compose session in which the share was created. The user must click **Send** from that same, still-open message. Saving or AutoSave does not interrupt the flow while the compose window remains open.
- Saving the primary mail as a draft or `.oft` template and then closing it is not supported. Reopening that draft, restarting Outlook before the first send attempt, or creating a message from that template restores the visible share block, but not the password follow-up state. No password follow-up is created; create a new share in the final message before sending it.
- Clicking **Send** immediately submits the password mail using the primary mail's Outlook account.
- The password mail does not wait for delayed or offline delivery of the primary mail. It may therefore be sent before the primary mail leaves the Outbox, or even if Outlook subsequently rejects the primary mail.
- If sender verification or automatic submission definitely fails, Outlook opens a fully prepared message for manual sending. An ambiguous Outlook submission is not repeated automatically.
- With the Secrets mode, one one-time Secret link is created per final recipient.
- Equal SMTP addresses across To, Cc, and Bcc receive only one Secret follow-up.
- If Secrets creation fails, Outlook uses the plain password follow-up and displays a warning.
- A matching backend signature is included only when the follow-up sender matches the assigned signature address.

### Talk room lifecycle

Select the language of the Talk description under **Settings -> Talk link**, below the Talk defaults. This option is no longer under Advanced. Existing choices and backend locks remain unchanged.

Deleting a saved Outlook appointment removes its remote Talk room only when the setting is explicitly enabled and the appointment contains NC Connector room metadata. The setting is disabled by default. A Talk URL copied into a location or body is not sufficient for remote deletion.

Deleting one occurrence or an exception from a recurring appointment does not remove the shared room. Room deletion applies only to a non-recurring appointment or the entire series, whether deleted from the calendar view or an open appointment.

Pending room deletions are retained per Outlook profile and retried after temporary Nextcloud failures or an Outlook restart. A newly created room from an unsaved, discarded appointment is still cleaned up even when deletion of saved appointments is disabled.

Moderator delegation requires another Nextcloud user. After a successful handoff, the original moderator leaves the room.

Before enabling deletion of saved appointments, inform users that deleting the appointment also deletes its Talk room, and document how to create a replacement room when needed.

## Internet Free/Busy Gateway (IFB)

### Purpose and activation

IFB lets Outlook request Nextcloud free/busy data through a local HTTP endpoint. It is off by default. For locked settings, use the [managed IFB policy](#managed-internet-freebusy-ifb); otherwise configure it locally:

1. Verify the Nextcloud system address book.
2. Open **NC Connector -> Settings -> IFB**.
3. Enable IFB and select the number of days and local port. Set the shared address-book cache duration under **Settings -> Advanced**. Defaults are 30 days, 24 cache hours, and port `7777`.
4. Save and restart Outlook.

Reserved listener namespace:

```text
http://127.0.0.1:7777/nc-ifb/
```

The MSI reserves the default URL namespace for authenticated Windows users. NC Connector adds a random path segment to Outlook's Free/Busy URL; requests without that segment return `404`. The secret path is managed internally and is intentionally not shown in this guide.

Enabling IFB updates only Outlook's per-user Free/Busy values. Disabling IFB restores the previous configuration unless another application or administrator has since changed it. Group-policy values below `Software\Policies` are not overwritten. The address-book cache is separate for each Outlook profile and Nextcloud account.

The listener runs only while Outlook is running, IFB is effectively enabled, and the stored Nextcloud credentials are complete. Invalid managed IFB settings prevent listener startup; Enterprise Rollout also requires confirmed backend access and an active assigned Seat for requests.

### Verify the default reservation

```powershell
netsh http show urlacl | Select-String -Pattern "127.0.0.1:7777/nc-ifb"
Test-NetConnection 127.0.0.1 -Port 7777
```

Then create a test meeting in Outlook, add an address from the Nextcloud system address book, and open **Scheduling Assistant**. Expected result: the reservation exists, the TCP test succeeds, and Outlook displays free/busy data. A direct request to the public `/nc-ifb/freebusy/...` path must return `404`.

### Custom IFB port

Valid configured ports are `1024` through `49151`. The MSI creates a reservation only for port `7777`. This applies to local settings and managed IFB policy alike; the add-in does not elevate or create a custom reservation. For another port, an administrator must open an elevated PowerShell window and add a reservation for authenticated users:

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

After a change to the add-in, Outlook, proxy, certificates, or Nextcloud, run the Settings connection test and check the functions affected by that change under a representative user account. For example, check a share after changing upload limits, or the assigned signature after changing its template. Use the relevant section of this guide for configuration and troubleshooting.

### Logs

Enable logging under **NC Connector -> Settings -> Debug**. Keep **Anonymize logs** enabled unless support specifically requests other data. Organization policies can lock both controls; see [Managed logging](#managed-logging).

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
4. In **Settings -> Advanced -> Transport security (TLS)**, test the organization-approved mode. If the TLS group is locked, review the [managed TLS policy](#managed-transport-security-tls) instead; correct invalid values at their effective registry location and restart Outlook. A TLS policy configuration error must be resolved before authentication or Seat checks can run.
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
4. Open Settings and review the displayed backend or seat state. If `ShowMainRibbonTab=false` hides Settings, use the Share or Talk action to check the access message.
5. Confirm whether the setting is a default or a locked value.
6. For user-specific problems, compare with a working account that has an assigned Seat.

In unmanaged installations, Share and Talk can use local settings during a backend outage. Managed signatures and separate password delivery require a valid backend state. Enterprise Rollout requires confirmed access; previously confirmed access may remain available during a temporary connection failure as described above.

### Managed signature is missing or misplaced

1. Confirm the backend seat and signature assignment.
2. Compare the effective Outlook **From** SMTP address with the assigned backend email address.
3. Check the separate switches for new mail, reply, and forward.
4. Check the sender and message type described under [Configure managed signatures](#configure-managed-signatures).
5. If signatures are duplicated or overlap, check whether Outlook also inserts its own signature for that sender.
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

1. Confirm that IFB is enabled and credentials are complete. If the settings are locked, check the [managed IFB policy](#managed-internet-freebusy-ifb), correct any configuration error, and restart Outlook. Enterprise Rollout also requires confirmed backend access and an active assigned Seat.
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
