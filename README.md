# Microsoft365-4D

![Delphi Version](https://img.shields.io/badge/Delphi-11%2C%2012%2C%2013-red)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux-blue)
![License](https://img.shields.io/badge/license-MIT-green)

Pure Delphi library for Microsoft Graph API integration. Provides OAuth2 authentication with PKCE and typed clients for Mail, Calendar, Contacts, and SharePoint. No external dependencies beyond the Delphi RTL.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Project Structure](#project-structure)
- [API Reference](#api-reference)
- [Demo Application](#demo-application)
- [Azure AD App Registration](#azure-ad-app-registration)
- [Token Storage](#token-storage)
- [License](#license)
- [About GDK Software](#about-gdk-software)

## Features

- **OAuth2 + PKCE** authentication flow for Microsoft identity platform
- **Mail** -search, read, draft, send, delete, move messages, list folders, attachments
- **Calendar** -list, create, update, delete events, check schedule availability
- **Contacts** -search, create, update, delete contacts
- **SharePoint** -browse sites, list/search drive items, get file content
- **Zero dependencies** -uses only `System.Net.HttpClient` (Delphi RTL)
- **Pluggable logging** -`TLogProc` callback, no global logger
- **Shared HTTP client** -all Graph clients can share a single `TGraphHttpClient`
- **Shared mailbox support** -access shared/delegated mailboxes via `MailboxAddress` property
- **Typed responses** -all API calls return strongly-typed records (`TMailMessage`, `TCalendarEvent`, etc.)
- **Interface-based** -all clients implement interfaces (`IMailClient`, `ICalendarClient`, etc.) for dependency injection

## Requirements

- Delphi 11 Alexandria or later (RAD Studio 11.x+)
- Azure AD App Registration with appropriate API permissions

## Installation

### Using as a Library

Add the `Source/OAuth2` and `Source/Graph` directories to your project's search path:

```
Source\OAuth2;Source\Graph
```

Then add the units you need to your uses clause:

```pascal
uses
  MSGraph.OAuth2.Types,
  MSGraph.OAuth2.PKCE,
  MSGraph.OAuth2.Client,
  MSGraph.OAuth2.TokenStore,
  MSGraph.Graph.Http,
  MSGraph.Graph.Mail.Interfaces,
  MSGraph.Graph.Mail,
  MSGraph.Graph.Calendar.Interfaces,
  MSGraph.Graph.Calendar,
  MSGraph.Graph.Contacts.Interfaces,
  MSGraph.Graph.Contacts,
  MSGraph.Graph.SharePoint.Interfaces,
  MSGraph.Graph.SharePoint;
```

## Quick Start

### 1. Configure OAuth2

```pascal
var Config: TOAuth2Config;
Config.ClientId := 'your-client-id';
Config.ClientSecret := 'your-client-secret';
Config.TenantId := 'your-tenant-id';
Config.RedirectUri := 'http://localhost:8080/oauth/callback';
Config.Scopes := TArray<string>.Create(
  'openid', 'offline_access',
  'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
  'Calendars.ReadWrite', 'Contacts.ReadWrite',
  'Sites.Read.All', 'User.Read'
);
```

### 2. Authenticate with PKCE

```pascal
var PKCESession := TOAuth2PKCE.Generate;
var OAuthClient := TOAuth2Client.Create(Config);

var AuthUrl := OAuthClient.GenerateAuthorizationUrl(PKCESession);
// Open AuthUrl in browser, handle callback to receive authorization code

var Tokens := OAuthClient.ExchangeCodeForToken(Code, PKCESession.CodeVerifier);
```

### 3. Use Graph Clients

```pascal
var Mail: IMailClient := TMailClient.Create(Tokens.AccessToken);
var SearchResult := Mail.SearchMessages('*', '', 10, 0);
for var Msg in SearchResult.Messages do
  WriteLn(Msg.Subject, ' - ', Msg.From.Address);
```

### 4. Refresh Tokens

```pascal
if Tokens.IsExpiringSoon(300) then
begin
  var NewTokens := OAuthClient.RefreshAccessToken(Tokens.RefreshToken);
  if NewTokens.RefreshToken.IsEmpty then
    NewTokens.RefreshToken := Tokens.RefreshToken;
end;
```

### 5. Share a Single HTTP Client

All Graph clients accept an existing `TGraphHttpClient`, allowing you to share one connection across multiple services:

```pascal
var Http := TGraphHttpClient.Create(Tokens.AccessToken);
var Mail: IMailClient := TMailClient.Create(Http);
var Calendar: ICalendarClient := TCalendarClient.Create(Http);
```

### 6. Access Shared Mailboxes

Set `MailboxAddress` on `TGraphHttpClient` to target a shared or delegated mailbox. When empty (default), endpoints use `/me`. When set, endpoints use `/users/{address}`:

```pascal
var Http := TGraphHttpClient.Create(Tokens.AccessToken);
Http.MailboxAddress := 'projects@company.com';
var Mail: IMailClient := TMailClient.Create(Http);
var Messages := Mail.SearchMessages('*', '', 10, 0);
```

This requires the `Mail.Read.Shared` and/or `Mail.Send.Shared` delegated permissions in your Azure AD app registration.

One restriction applies to large attachments: with delegated permissions Microsoft Graph refuses to attach a file larger than 3 MB to a message in a shared or delegated mailbox and answers `HTTP 403 Forbidden`. Only the signed-in user's own mailbox works that way. Use application permissions when a shared mailbox has to send large attachments. When `AddAttachment` gets a `403` on a request against another mailbox, it appends this as a hint to the error message. The hint is only a hint: `MailboxAddress` being set does not prove the mailbox is shared, so Microsoft's own message is always reported alongside it and never replaced.

### 7. Add Custom Internet Message Headers

`CreateDraft` has an overload that takes custom headers. They are written to the message as `internetMessageHeaders` and travel with it when the draft is sent:

```pascal
var Draft := Mail.CreateDraft('Subject', 'Body', ['recipient@company.com'], [], [], False,
  [TMailHeader.Create('x-example-id', '42')]);
Mail.SendDraft(Draft.Id);
```

Microsoft Graph requires a custom header name to start with `x-`, and accepts at most five custom headers per message. The library validates the headers before sending anything and raises `EInvalidMailHeaderException` with a readable message when a rule is broken:

| Rule | Rejected example |
|------|------------------|
| At most five custom headers per message | six headers |
| The name must not be empty | `''` |
| The name must start with `x-` (case insensitive) | `Example-Id` |
| The name may contain printable ASCII only, and no `:` | `x-example-id:42`, `x-example-id ` |
| The value must not contain control characters or line separators | `42<CR><LF>x-injected: yes` |
| The same name must not be supplied twice (case insensitive) | `x-example-id` and `X-Example-Id` |

Names are sent exactly as supplied — the library does not change their casing.

Graph accepts `internetMessageHeaders` only when a message is created, so `UpdateDraft` does not offer this parameter.

Two things to know about the receiving side. Exchange Online maps a custom header name to a named property on first use, and does not store the value on that very first message — only later messages carrying the same header name keep their value. A transport rule or the header firewall can also strip custom headers in transit.

### 8. Attach Files of Any Supported Size

`AddAttachment` picks its own route based on the size of the file. Callers pass the bytes and do not choose:

```pascal
var Draft := Mail.CreateDraft('Report', 'See attached', ['recipient@company.com'], [], [], False);
Mail.AddAttachment(Draft.Id, 'report.pdf', 'application/pdf', TFile.ReadAllBytes(ReportPath));
Mail.SendDraft(Draft.Id);
```

A caller that needs the attachment id afterwards — to remove or inspect the attachment later — calls `AddAttachmentAndGetId` instead, which returns the id in place of the plain success flag:

```pascal
var AttachmentId := Mail.AddAttachmentAndGetId(Draft.Id, 'report.pdf', 'application/pdf',
  TFile.ReadAllBytes(ReportPath));
```

Both methods raise the same exceptions on failure; `AddAttachmentAndGetId` only ever returns once the upload succeeded and its id is known.

| Size | Route |
|------|-------|
| Below 3 MB | The file is posted as part of the message itself |
| 3 MB up to 25 MB | An upload session is created on the existing draft and the file is uploaded in chunks of 3 MiB |

Microsoft does not document the exact byte count behind "3 MB", so both routes can refuse a file that sits right on the boundary. The library handles that itself: a message-body upload rejected with `HTTP 413` is retried once through an upload session, and an upload session rejected with `ErrorAttachmentSizeShouldNotBeLessThanMinimumSize` is retried once through the message body. It never switches more than once.

An empty file and a file over 25 MB are rejected with `EInvalidAttachmentException` before any request is sent.

When a chunk fails, the library cancels the upload session and raises `EGraphApiException` naming the file, its size, the byte range that failed and the HTTP status. The draft is never sent and is left untouched, so the caller can retry or call `DeleteDraft`. There is no automatic retry: every non-success status ends the upload.

A connection that drops halfway is treated the same way. The transport error is caught, the upload session is cancelled, and an `EGraphApiException` naming the file, its size, the byte range that was in flight and the original error message is raised in its place, so a caller never sees a bare socket error and never leaves an upload session behind.

Cancelling that session can itself fail. When the `DELETE` answers with a non-success status or does not answer at all, that is logged and appended to the exception the caller already gets, rather than replacing it: the reason the upload failed stays the headline, and the fact that a session was left behind at Microsoft is visible right after it.

The last `PUT` of a successful upload answers `HTTP 201 Created` with a `Location` header holding the attachment id. If that header is missing or holds no id, `EGraphApiException` is raised as well. The upload itself did succeed in that case, and the message says so: the attachment is on the draft, but its id could not be determined, so `AddAttachmentAndGetId` tells the caller instead of handing back an empty string. The upload session is not cancelled, because there is nothing left to cancel.

Graph itself accepts attachments up to 150 MB, but the library caps at 25 MB. That stays under the message size limit of the mailbox, 35 MB by default in Exchange Online, so an upload accepted by the library is not expected to fail later at `SendDraft` for size reasons alone.

`Mail.ReadWrite` is required to create the upload session; `Mail.Send` remains required to send the draft.

## Project Structure

```
Source/
  OAuth2/
    MSGraph.OAuth2.Types.pas              -Config records, token response, PKCE session, exceptions
    MSGraph.OAuth2.PKCE.pas               -PKCE code_verifier + code_challenge generation
    MSGraph.OAuth2.Client.pas             -OAuth2 flow: auth URL, token exchange, refresh
    MSGraph.OAuth2.TokenStore.pas         -Thread-safe in-memory token + PKCE storage
  Graph/
    MSGraph.Graph.Http.pas                -Graph API HTTP client with error handling
    MSGraph.Graph.Http.Redaction.pas      -Masks sensitive query parameters before logging a URL
    MSGraph.Graph.JsonHelper.pas          -JSON parsing utilities (TGraphJson)
    MSGraph.Graph.Mail.Types.pas          -Mail record types (TMailMessage, TMailFolder, etc.)
    MSGraph.Graph.Mail.Interfaces.pas     -IMailClient and IAttachmentUploader interfaces
    MSGraph.Graph.Mail.Attachments.pas    -Attachment upload, small route and upload session
    MSGraph.Graph.Mail.pas                -TMailClient implementation
    MSGraph.Graph.Calendar.Types.pas      -Calendar record types (TCalendarEvent, TAttendee, etc.)
    MSGraph.Graph.Calendar.Interfaces.pas -ICalendarClient interface
    MSGraph.Graph.Calendar.pas            -TCalendarClient implementation
    MSGraph.Graph.Contacts.Types.pas      -Contact record types (TContact, TPostalAddress, etc.)
    MSGraph.Graph.Contacts.Interfaces.pas -IContactsClient interface
    MSGraph.Graph.Contacts.pas            -TContactsClient implementation
    MSGraph.Graph.SharePoint.Types.pas    -SharePoint record types (TSite, TDriveItem)
    MSGraph.Graph.SharePoint.Interfaces.pas -ISharePointClient interface
    MSGraph.Graph.SharePoint.pas          -TSharePointClient implementation
Examples/
    Microsoft365Demo.dpr                  -Console demo with interactive menu
    Microsoft365Demo.App.pas              -Demo application logic
    Microsoft365Demo.CallbackServer.pas   -Indy HTTP server for OAuth callback
```

## API Reference

### Exception Hierarchy

All library exceptions inherit from `EMSGraphException`:

| Exception | Inherits from | Raised by |
|-----------|---------------|-----------|
| `EMSGraphException` | `Exception` | Base exception for all Microsoft365-4D errors |
| `EOAuth2Exception` | `EMSGraphException` | Token exchange failures, invalid responses |
| `EGraphApiException` | `EMSGraphException` | Anything that stops a Graph call from succeeding: an HTTP error, a missing access token, an interrupted upload, or a request the library refuses to send |
| `ETokenStoreException` | `EMSGraphException` | Missing tokens, expired PKCE sessions |
| `EInvalidMailHeaderException` | `EGraphApiException` | Invalid custom mail header supplied by the caller |
| `EInvalidAttachmentException` | `EGraphApiException` | Attachment that is empty or larger than the supported maximum |
| `EDeltaLinkExpiredException` | `EGraphApiException` | An expired delta link, so a full resynchronisation is needed |

The two validation exceptions sit under `EGraphApiException` on purpose: they report a Graph call that will not succeed, in the same way as a rejected access token does. A caller that handles `EGraphApiException` therefore catches every reason a mail operation can fail, and can still catch the specific class when it wants to tell the cases apart.

### TMailClient

| Method | Description |
|--------|-------------|
| `SearchMessages(Query, FolderId, Top, Skip)` | Search or list messages |
| `GetMessage(MessageId)` | Get full message by ID |
| `GetMessageAttachments(MessageId)` | List attachments |
| `GetAttachmentContent(MessageId, AttachmentId)` | Get attachment content |
| `CreateDraft(Subject, Body, To, Cc, Bcc, IsHtml)` | Create draft |
| `CreateDraft(Subject, Body, To, Cc, Bcc, IsHtml, Headers)` | Create draft with custom internet message headers |
| `UpdateDraft(MessageId, Subject, Body, To, Cc, Bcc, IsHtml)` | Update existing draft |
| `AddAttachment(MessageId, FileName, ContentType, ContentBytes)` | Attach a file to a draft, picking the message-body route or an upload session by size |
| `AddAttachmentAndGetId(MessageId, FileName, ContentType, ContentBytes)` | Same as `AddAttachment`, but returns the attachment id instead of a plain success flag |
| `SendDraft(MessageId)` | Send a draft message |
| `DeleteDraft(MessageId)` | Delete a draft |
| `MoveMessage(MessageId, FolderId)` | Move message to folder |
| `ListMailFolders(ParentFolderId)` | List mail folders |
| `GetMailboxSignature` | Get HTML signature |

### TCalendarClient

| Method | Description |
|--------|-------------|
| `ListEvents(Start, End, Top, Timezone)` | List calendar events in range |
| `GetEvent(EventId)` | Get event details |
| `CreateEvent(Subject, Start, End, Location, Body, Attendees, IsAllDay)` | Create event |
| `UpdateEvent(EventId, Subject, Start, End, Location, Body, Attendees, IsAllDay)` | Update event |
| `DeleteEvent(EventId)` | Delete event |
| `GetScheduleAvailability(Schedules, Start, End, Timezone)` | Check availability |

### TContactsClient

| Method | Description |
|--------|-------------|
| `SearchContacts(Query, Top)` | Search or list contacts |
| `GetContact(ContactId)` | Get contact details |
| `CreateContact(GivenName, Surname, Email, Phone, Company, JobTitle)` | Create contact |
| `UpdateContact(ContactId, GivenName, Surname, Email, Phone, Company, JobTitle)` | Update contact |
| `DeleteContact(ContactId)` | Delete contact |

### TSharePointClient

| Method | Description |
|--------|-------------|
| `ListSites(Query, Top)` | Search SharePoint sites |
| `GetSite(SiteId)` | Get site details |
| `ListDriveItems(SiteId, FolderId, Top)` | List items in drive/folder |
| `SearchDriveItems(SiteId, Query, Top)` | Search drive items |
| `GetDriveItemContent(SiteId, ItemId)` | Get item details + download URL |

## Demo Application

The `Examples/` folder contains a console application demonstrating the full OAuth2 flow and all Graph API operations.

### Running the Demo

```bash
Microsoft365Demo.exe --client-id "your-id" --client-secret "your-secret" --tenant-id "your-tenant" --redirect-uri "http://localhost:8080/oauth/callback"
```

All parameters can also be provided interactively when omitted. Available options:

| Parameter | Description | Default |
|-----------|-------------|---------|
| `--client-id` | Azure AD application ID | *(prompted)* |
| `--client-secret` | Azure AD client secret | *(prompted)* |
| `--tenant-id` | Azure AD tenant ID | *(prompted)* |
| `--redirect-uri` | OAuth2 redirect URI | `http://localhost:8080/oauth/callback` |
| `--port` | Local callback server port | `8080` |

### Demo Menu

The demo provides an interactive menu with the following options:

1. Authenticate (opens browser for Microsoft login)
2. List messages
3. Read message by ID
4. Send email (create draft + send)
5. List mail folders
6. List calendar events
7. Create calendar event
8. Search contacts
9. List SharePoint sites
10. Refresh token manually

## Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations** > **New registration**
2. Set **Redirect URI** to your callback URL (Web platform), e.g. `http://localhost:8080/oauth/callback`
3. Create a **Client secret** under **Certificates & secrets**
4. Add **API permissions** (Microsoft Graph, Delegated):

| Permission | Description |
|---|---|
| `openid` | Sign-in |
| `profile` | User profile |
| `offline_access` | Refresh tokens |
| `Mail.Read` | Read mail |
| `Mail.ReadWrite` | Create/edit drafts |
| `Mail.Send` | Send mail |
| `MailboxSettings.Read` | Read mailbox signature |
| `Calendars.ReadWrite` | Calendar access |
| `Contacts.ReadWrite` | Contacts access |
| `Sites.Read.All` | SharePoint read access |
| `User.Read` | User info |
| `Mail.Read.Shared` | Read shared/delegated mailboxes |
| `Mail.Send.Shared` | Send from shared/delegated mailboxes |

5. Click **Grant admin consent** if you have admin rights

## Token Storage

The included `TTokenStore` stores tokens **in-memory only** and is intended for demo/development use. For production:

- Persist tokens to a database, file, or OS-level credential store
- **Encrypt** refresh tokens and access tokens before storage
- Implement your own storage by subclassing or replacing `TTokenStore`

## License

MIT License. See [LICENSE](LICENSE) for details.

## Commercial Support

This library is MIT licensed and free to use. For companies that depend on it commercially we offer support and maintenance agreements with guaranteed response times, and sponsored development of features you need. Contact us at [gdksoftware.com/contact-us](https://gdksoftware.com/contact-us) or open an issue to get in touch.

## About GDK Software

Microsoft365-4D is developed by [GDK Software](https://gdksoftware.com), a Delphi-focused software company building developer tools, MCP integrations, and enterprise applications.
