# Translations — NC Connector for Outlook

This project uses WebExtension-style locale files (`messages.json`) embedded in the add-in assembly.

## Supported languages

Language codes (must match folder names under `src/NcTalkOutlookAddIn/Resources/_locales/`):

- `de` (default)
- `en`
- `fr`
- `cs`
- `es`
- `hu`
- `it`
- `ja`
- `nl`
- `pl`
- `pt_BR`
- `pt_PT`
- `ru`
- `zh_CN`
- `zh_TW`

## Where translations live

- `src/NcTalkOutlookAddIn/Resources/_locales/<lang>/messages.json`

Strings are loaded at runtime by:

- `src/NcTalkOutlookAddIn/Utilities/Strings.cs`

## Placeholders

Use `$1`, `$2`, ... in `messages.json`.
They are automatically converted to `.NET` `string.Format` placeholders (`{0}`, `{1}`, ...).

## Adding / updating strings

1. Add a new key to **all** `messages.json` files.
2. Add a corresponding property to `src/NcTalkOutlookAddIn/Utilities/Strings.cs` (or reuse an existing one).
3. Build the add-in and verify the UI.

## Adding a new language (future)

1. Create `src/NcTalkOutlookAddIn/Resources/_locales/<new_code>/messages.json`.
2. Add `<new_code>` to `SupportedLanguages` and `SupportedLanguageOverrides` in `src/NcTalkOutlookAddIn/Utilities/Strings.cs`.
3. Ensure every key used by `Strings.cs` exists and is translated.

