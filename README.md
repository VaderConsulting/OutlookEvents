# OutlookEvents

This is a Visual Studio 2008 VB.NET WinForms program that starts an Outlook Application COM object and listens for `NewMailEx`. For each new item it writes sender, display name, CC, subject, body, and attachment count to the console. It is Dave Robinson's working copy from Historical Dev, kept as a COM-automation counterpart to the VSTO mail add-ins.

**Source last updated:** 2008-02-26  
**Language:** VB.NET  
**Target:** .NET 3.5  
**Output:** WinForms exe

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `OutlookEvents` | VB.NET | WinForms | Form1 registers `EventClassModule` for Outlook `NewMailEx` logging |

## How to open

Open `OutlookEvents.sln` in Visual Studio 2008. Outlook should be installed so the Microsoft.Office.Interop.Outlook reference resolves.

## Attribution and provenance

From Dave Robinson's Historical Dev archive (OneDrive folder `OutlookEvents`). Assembly copyright 2008.

## License

MIT License. Copyright (c) 2026 VaderConsulting.
