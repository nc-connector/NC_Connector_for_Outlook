# Third-Party Dependencies

## HtmlSanitizer

- Package: `HtmlSanitizer`
- Version: `8.1.870`
- Source: https://www.nuget.org/packages/HtmlSanitizer/8.1.870
- Upstream repository: https://github.com/mganss/HtmlSanitizer
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/HtmlSanitizer.dll`
- License: MIT
- Usage in this add-in:
  - Sanitization of backend-provided Share/Talk HTML templates
  - Runtime consumers:
    - `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs`
    - `src/NcTalkOutlookAddIn/Utilities/FileLinkHtmlBuilder.cs`
    - `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs`

## AngleSharp

- Package: `AngleSharp`
- Version: `0.17.1`
- Source: https://www.nuget.org/packages/AngleSharp/0.17.1
- Upstream repository: https://github.com/AngleSharp/AngleSharp
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/AngleSharp.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependency of `HtmlSanitizer`

## AngleSharp.Css

- Package: `AngleSharp.Css`
- Version: `0.17.0`
- Source: https://www.nuget.org/packages/AngleSharp.Css/0.17.0
- Upstream repository: https://github.com/AngleSharp/AngleSharp.Css
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/AngleSharp.Css.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependency of `HtmlSanitizer`

## .NET Runtime Dependencies (vendored with sanitizer stack)

- Packages:
  - `System.Buffers` (`4.6.28619.1` assembly version)
  - `System.Collections.Immutable` (`8.0.23.53103` assembly version)
  - `System.Memory` (`4.6.31308.1` assembly version)
  - `System.Runtime.CompilerServices.Unsafe` (`6.0.21.52210` assembly version)
  - `System.Text.Encoding.CodePages` (`6.0.21.52210` assembly version)
- Sources:
  - https://www.nuget.org/packages/System.Buffers
  - https://www.nuget.org/packages/System.Collections.Immutable
  - https://www.nuget.org/packages/System.Memory
  - https://www.nuget.org/packages/System.Runtime.CompilerServices.Unsafe
  - https://www.nuget.org/packages/System.Text.Encoding.CodePages
- Included files:
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Buffers.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Collections.Immutable.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Memory.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Runtime.CompilerServices.Unsafe.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Text.Encoding.CodePages.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependencies required by `HtmlSanitizer`/`AngleSharp`
