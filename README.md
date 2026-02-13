<div align="center" style="background:#0082C9; padding:1px 0;"><img src="assets/header-solid-blue-1920x480.png" alt="Add-in" height="80"></div>

[English](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.md) | [Deutsch](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/README.de.md)
[Admin Guide](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/ADMIN.md) | [Development Guide](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/docs/DEVELOPMENT.md) | [Translations](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md)

# NC Connector for Outlook

NC Connector for Outlook connects Outlook seamlessly with your Nextcloud. The add-in automates Talk rooms for appointments, provides a local free/busy proxy, and ships a powerful filelink wizard for emails. The goal is a professional workflow from calendar to file storage — without context switching and with clear admin controls.

This is a community project and is not an official Nextcloud GmbH product.

## Highlights

- **One-click Nextcloud Talk**  
Open an appointment, choose Nextcloud Talk, configure the room, and select a moderator. Optionally, invited attendees can be added to the room automatically (separately for internal Nextcloud users and external email guests). The wizard writes title, location, and a description block (including help link) into the appointment.
- **Sharing deluxe**  
The compose button “Insert Nextcloud share” starts the sharing wizard with upload queue, password generator, expiration date, and note field. The finished share is inserted as formatted HTML directly into the email.
- **Enterprise-grade security**  
Lobby until start time, moderator delegation, automatic cleanup of discarded appointments, mandatory passwords, and expiration policies help protect sensitive meetings and files.
- **Internet Free/Busy Gateway (IFB)**  
A local HTTP listener answers Outlook free/busy requests directly from Nextcloud. The installer configures registry values for search path and read URL. If the direct fetch returns HTTP 404, the add-in falls back to a scheduling POST so availability data is still provided.
- **Debug logging at the press of a button**  
Enable it in the Debug tab. Writes structured logs (authentication, appointment and filelink flows, IFB) to `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`. The path is displayed in the UI.

## Changelog

See [`CHANGELOG.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/CHANGELOG.md).

## Feature overview

### Nextcloud Talk directly from the appointment
- Talk dialog with lobby, password, listable scope, room type, and moderator search.
- Automatically writes title, location and a description block (incl. help link and password) into the appointment.
- Room tracking, lobby updates, delegation workflow, and cleanup when an appointment is discarded or moved.
- Calendar changes (drag & drop or dialog edits) keep the Talk room lobby/start time in sync.
- Optional participant sync after saving the appointment:
  - **Users:** internal Nextcloud users are added to the room.
  - **Guests:** external email addresses are invited as guests (Nextcloud may also send an additional invitation email).

### Nextcloud Sharing in the compose window
- Four steps (share, expiration date, files, note) with a password-protected upload folder.
- Upload queue with duplicate checks, progress display and optional share creation.
- Automatic HTML block with link, password, expiration date and optional note.

### Administration & compliance
- Login Flow V2 (app password is created automatically) and central options (base URL, debug mode, sharing paths, defaults for Sharing/Talk).
- Full localization (see [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md)) and structured debug logs for support cases.

## Language & translations

- The UI language follows the Outlook/Office UI language. If Outlook is set to **Use system settings**, this usually matches the Windows display language.
- Supported languages are documented in [`Translations.md`](https://github.com/nc-connector/NC_Connector_for_Outlook/blob/main/Translations.md). Fallback is `de`, then `en`.

### Language overrides (text blocks)

In Settings under **Advanced**, you can choose the language for inserted text blocks independently of the UI language:

- **Sharing HTML block** (email): language of the formatted HTML block inserted when sharing.
- **Talk description text** (appointment): language of the inserted text (e.g., password line / help link).

Option `Default (UI)` uses the current UI language (including fallbacks).

## System requirements

- Windows 10 or Windows 11 (64-bit)  
- Microsoft Outlook classic >= 2019  
- .NET Framework 4.7.2 Runtime  
- Nextcloud Server with Talk and Files Sharing apps

## Installation and updates

1. Close Outlook.  
2. Run the latest MSI (for example `NCConnectorForOutlook-2.2.7.msi`) and confirm the UAC prompt (administrator rights are required). The setup configures URLACL and all required registry keys for IFB.  
3. Start Outlook and click **NC Connector → Settings** in the ribbon.  
4. Choose the login mode, run the connection test, and save. If the test succeeds, IFB is active automatically.  
5. Verify the filelink base directory and enable debug logging if needed.

Updates are applied by installing a higher MSI version. Personal settings (`settings.ini`) are kept. Uninstall removes the add-in, stops the IFB listener, and resets the registry values.

## Troubleshooting

- **Debug log**: enable it in the *Debug* tab. Log file: `%LOCALAPPDATA%\NextcloudTalkOutlookAddInData\addin-runtime.log`.  
- **Add-in not visible**: installation must be run with admin rights. Check `HKLM\Software\Microsoft\Office\Outlook\Addins\NcTalkOutlook.AddIn` and optionally run a repair from an elevated prompt: `msiexec /i "NCConnectorForOutlook-2.2.7.msi" ADDLOCAL=ALL`.  
- **Test IFB**: `powershell -Command "Invoke-WebRequest http://127.0.0.1:7777/nc-ifb/freebusy/<mail>.vfb -UseBasicParsing"`. If behavior differs, verify the registry under `HKCU\Software\Microsoft\Office\<Version>\Outlook\Options\Calendar`.  
- **Check TLS/proxy**: `powershell -Command "Test-NetConnection <your-domain> -Port 443"`. If you see SSL warnings, verify certificates/proxy settings.  
- **Sharing errors**: the debug log includes HTTP status codes and exception details. Required wizard fields are validated.

## Screenshots

<details>
<summary><strong>Settings</strong></summary>

| <a href="Screenshots/settings.jpg"><img src="Screenshots/settings.jpg" alt="Settings dialog" width="230"></a> |
| --- |

</details>

<details>
<summary><strong>Talk link workflow</strong></summary>

| <a href="Screenshots/1_talk.jpg"><img src="Screenshots/1_talk.jpg" alt="Talk step 1" width="230"></a> | <a href="Screenshots/2_talk.jpg"><img src="Screenshots/2_talk.jpg" alt="Talk step 2" width="230"></a> |
| --- | --- |

</details>

<details open>
<summary><strong>Sharing wizard</strong></summary>

| <a href="Screenshots/1_filelink.jpg"><img src="Screenshots/1_filelink.jpg" alt="Sharing step 1" width="230"></a> | <a href="Screenshots/2_filelink.jpg"><img src="Screenshots/2_filelink.jpg" alt="Sharing step 2" width="230"></a> |
| --- | --- |
| <a href="Screenshots/3_filelink.jpg"><img src="Screenshots/3_filelink.jpg" alt="Sharing step 3" width="230"></a> | <a href="Screenshots/4_filelink.jpg"><img src="Screenshots/4_filelink.jpg" alt="Sharing step 4" width="230"></a> |
| <a href="Screenshots/5_filelink.jpg"><img src="Screenshots/5_filelink.jpg" alt="Sharing step 5" width="230"></a> | |

</details>

<details>
<summary><strong>Internet Free/Busy</strong></summary>

| <a href="Screenshots/ifb.jpg"><img src="Screenshots/ifb.jpg" alt="IFB settings" width="230"></a> |
| --- |

</details>
