unit Microsoft365Demo.App;

interface

uses
  MSGraph.OAuth2.Types,
  MSGraph.OAuth2.Client,
  MSGraph.OAuth2.TokenStore;

type
  TDemoApp = class
  strict private
    FConfig: TOAuth2Config;
    FOAuthClient: TOAuth2Client;
    FTokenStore: TTokenStore;

    procedure LogHandler(const Level: string; const Message: string);
    procedure EnsureValidToken;
    procedure PrintSeparator;

    procedure DoAuthenticate;
    procedure DoListMessages;
    procedure DoReadMessage;
    procedure DoSendEmail;
    procedure DoListFolders;
    procedure DoListEvents;
    procedure DoCreateEvent;
    procedure DoSearchContacts;
    procedure DoListSites;
    procedure DoRefreshToken;
  public
    constructor Create(const Config: TOAuth2Config);
    destructor Destroy; override;

    procedure Run;
  end;

implementation

uses
  System.SysUtils,
  System.DateUtils,
{$IFDEF MSWINDOWS}
  Winapi.ShellAPI,
{$ENDIF}
{$IFDEF POSIX}
  Posix.Stdlib,
{$ENDIF}
  MSGraph.OAuth2.PKCE,
  MSGraph.Graph.Mail,
  MSGraph.Graph.Calendar,
  MSGraph.Graph.Contacts,
  MSGraph.Graph.SharePoint,
  Microsoft365Demo.CallbackServer;

constructor TDemoApp.Create(const Config: TOAuth2Config);
begin
  inherited Create;
  FConfig := Config;
  FOAuthClient := TOAuth2Client.Create(FConfig, LogHandler);
  FTokenStore := TTokenStore.Create;
end;

destructor TDemoApp.Destroy;
begin
  FTokenStore.Free;
  FOAuthClient.Free;
  inherited;
end;

procedure TDemoApp.LogHandler(const Level: string; const Message: string);
begin
  if Level = 'ERROR' then
    Writeln('[ERROR] ', Message)
  else if Level = 'WARNING' then
    Writeln('[WARN]  ', Message)
  else if Level = 'INFO' then
    Writeln('[INFO]  ', Message);
end;

procedure TDemoApp.PrintSeparator;
begin
  Writeln(StringOfChar('-', 60));
end;

procedure TDemoApp.EnsureValidToken;
begin
  if not FTokenStore.HasValidTokens then
    raise EMSGraphException.Create('Not authenticated. Please authenticate first (option 1).');

  var Tokens := FTokenStore.GetTokens;
  if not Tokens.IsExpiringSoon(300) then
    Exit;

  Writeln('Token expiring soon, refreshing...');
  var NewTokens := FOAuthClient.RefreshAccessToken(Tokens.RefreshToken);
  if NewTokens.RefreshToken.IsEmpty then
    NewTokens.RefreshToken := Tokens.RefreshToken;
  FTokenStore.StoreTokens(NewTokens);
  Writeln('Token refreshed successfully.');
end;

procedure TDemoApp.DoAuthenticate;
begin
  Writeln;
  Writeln('Starting OAuth2 authentication flow...');

  var PKCESession := TOAuth2PKCE.Generate;
  FTokenStore.StorePKCESession(PKCESession);

  var AuthUrl := FOAuthClient.GenerateAuthorizationUrl(PKCESession);

  var CallbackServer := TCallbackServer.Create(FConfig.Port);
  try
    CallbackServer.Start;
    Writeln('Callback server listening on http://localhost:', FConfig.Port, '/oauth/callback');
    Writeln('Opening browser for Microsoft login...');

{$IFDEF MSWINDOWS}
    ShellExecute(0, 'open', PChar(AuthUrl), nil, nil, 1);
{$ENDIF}
{$IFDEF POSIX}
    _system(PAnsiChar(AnsiString('xdg-open "' + AuthUrl + '" &')));
{$ENDIF}
    Writeln('Waiting for callback (timeout: 2 minutes)...');

    const Received = CallbackServer.WaitForCallback(120000);
    if not Received then
    begin
      Writeln('ERROR: Timed out waiting for callback.');
      Exit;
    end;

    const HasError = not CallbackServer.Error.IsEmpty;
    if HasError then
    begin
      Writeln('ERROR: ', CallbackServer.Error);
      Exit;
    end;

    var StoredSession := FTokenStore.RetrievePKCESession(CallbackServer.State);
    FTokenStore.DeletePKCESession(CallbackServer.State);

    Writeln('Exchanging code for tokens...');
    var Tokens := FOAuthClient.ExchangeCodeForToken(CallbackServer.Code, StoredSession.CodeVerifier);
    FTokenStore.StoreTokens(Tokens);

    Writeln;
    Writeln('Authentication successful!');
    Writeln('  Expires in:  ', Tokens.ExpiresIn, 's');
    Writeln('  Scopes:      ', Tokens.Scope);
  finally
    CallbackServer.Stop;
    CallbackServer.Free;
  end;
end;

procedure TDemoApp.DoListMessages;
begin
  EnsureValidToken;

  var Query := '';
  Write('Search query (Enter for all): ');
  Readln(Query);
  if Query.IsEmpty then
    Query := '*';

  var Mail := TMailClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var SearchResult := Mail.SearchMessages(Query, '', 20, 0);

    if Length(SearchResult.Messages) = 0 then
    begin
      Writeln('No messages found.');
      Exit;
    end;

    Writeln(Format('Found %d message(s):', [Length(SearchResult.Messages)]));
    PrintSeparator;
    var Counter := 0;
    for var Msg in SearchResult.Messages do
    begin
      Inc(Counter);
      var ReadMarker := ' ';
      if not Msg.IsRead then
        ReadMarker := '*';

      Writeln(Format('%s [%d] %s', [ReadMarker, Counter, Msg.Subject]));
      Writeln(Format('       From: %s  |  %s', [Msg.From.Address, Msg.ReceivedDateTime]));
      Writeln(Format('       ID: %s', [Msg.Id]));
      PrintSeparator;
    end;
  finally
    Mail.Free;
  end;
end;

procedure TDemoApp.DoReadMessage;
begin
  EnsureValidToken;

  var MessageId := '';
  Write('Message ID: ');
  Readln(MessageId);
  if MessageId.Trim.IsEmpty then
    Exit;

  var Mail := TMailClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var Msg := Mail.GetMessage(MessageId);

    PrintSeparator;
    Writeln('Subject: ', Msg.Subject);
    Writeln('Date:    ', Msg.ReceivedDateTime);
    Writeln('Body:    ', Msg.Body.Substring(0, 500));
    PrintSeparator;
  finally
    Mail.Free;
  end;
end;

procedure TDemoApp.DoSendEmail;
begin
  EnsureValidToken;

  var ToAddress := '';
  Write('To: ');
  Readln(ToAddress);
  if ToAddress.Trim.IsEmpty then
    Exit;

  var Subject := '';
  Write('Subject: ');
  Readln(Subject);

  var Body := '';
  Write('Body: ');
  Readln(Body);

  var Mail := TMailClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    Writeln('Creating draft...');
    var Draft := Mail.CreateDraft(Subject, Body, TArray<string>.Create(ToAddress), nil, False);

    Writeln('Sending...');
    if Mail.SendDraft(Draft.Id) then
      Writeln('Email sent!')
    else
      Writeln('ERROR: Failed to send.');
  finally
    Mail.Free;
  end;
end;

procedure TDemoApp.DoListFolders;
begin
  EnsureValidToken;

  var Mail := TMailClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var Folders := Mail.ListMailFolders;

    if Length(Folders) = 0 then
    begin
      Writeln('No folders found.');
      Exit;
    end;

    PrintSeparator;
    for var Folder in Folders do
      Writeln(Format('  %-30s  Total: %d  Unread: %d', [Folder.DisplayName, Folder.TotalItemCount, Folder.UnreadItemCount]));
    PrintSeparator;
  finally
    Mail.Free;
  end;
end;

procedure TDemoApp.DoListEvents;
begin
  EnsureValidToken;

  var DaysStr := '';
  Write('Days ahead (default 7): ');
  Readln(DaysStr);
  var Days := StrToIntDef(DaysStr, 7);

  var Calendar := TCalendarClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var Events := Calendar.ListEvents(Now, Now + Days, 20, 'Europe/Amsterdam');

    if Length(Events) = 0 then
    begin
      Writeln('No events found.');
      Exit;
    end;

    Writeln(Format('Found %d event(s):', [Length(Events)]));
    PrintSeparator;
    var Counter := 0;
    for var CalEvent in Events do
    begin
      Inc(Counter);
      Writeln(Format('  [%d] %s', [Counter, CalEvent.Subject]));
      Writeln(Format('       Start: %s  Location: %s', [CalEvent.StartDateTime, CalEvent.Location]));
    end;
    PrintSeparator;
  finally
    Calendar.Free;
  end;
end;

procedure TDemoApp.DoCreateEvent;
begin
  EnsureValidToken;

  var Subject := '';
  Write('Subject: ');
  Readln(Subject);
  if Subject.Trim.IsEmpty then
    Exit;

  var Location := '';
  Write('Location: ');
  Readln(Location);

  Writeln('Start time will be tomorrow at 10:00, duration 1 hour.');

  var StartDt := EncodeDate(YearOf(Tomorrow), MonthOf(Tomorrow), DayOf(Tomorrow)) + EncodeTime(10, 0, 0, 0);
  var EndDt := IncHour(StartDt, 1);

  var Calendar := TCalendarClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var EventResult := Calendar.CreateEvent(Subject, StartDt, EndDt, Location, '', nil, False);
    Writeln('Event created! ID: ', EventResult.Id);
  finally
    Calendar.Free;
  end;
end;

procedure TDemoApp.DoSearchContacts;
begin
  EnsureValidToken;

  var Query := '';
  Write('Search contacts (Enter for all): ');
  Readln(Query);

  var Contacts := TContactsClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var ContactList := Contacts.SearchContacts(Query, 20);

    if Length(ContactList) = 0 then
    begin
      Writeln('No contacts found.');
      Exit;
    end;

    Writeln(Format('Found %d contact(s):', [Length(ContactList)]));
    PrintSeparator;
    var Counter := 0;
    for var Contact in ContactList do
    begin
      Inc(Counter);
      Writeln(Format('  [%d] %s', [Counter, Contact.DisplayName]));
      Writeln(Format('       Company: %s  Title: %s', [Contact.Company, Contact.JobTitle]));
    end;
    PrintSeparator;
  finally
    Contacts.Free;
  end;
end;

procedure TDemoApp.DoListSites;
begin
  EnsureValidToken;

  var Query := '';
  Write('Search sites (Enter for all): ');
  Readln(Query);

  var SP := TSharePointClient.Create(FTokenStore.GetTokens.AccessToken, LogHandler);
  try
    var Sites := SP.ListSites(Query, 20);

    if Length(Sites) = 0 then
    begin
      Writeln('No sites found.');
      Exit;
    end;

    Writeln(Format('Found %d site(s):', [Length(Sites)]));
    PrintSeparator;
    var Counter := 0;
    for var Site in Sites do
    begin
      Inc(Counter);
      Writeln(Format('  [%d] %s', [Counter, Site.DisplayName]));
      Writeln(Format('       URL: %s', [Site.WebUrl]));
    end;
    PrintSeparator;
  finally
    SP.Free;
  end;
end;

procedure TDemoApp.DoRefreshToken;
begin
  if not FTokenStore.HasValidTokens then
  begin
    Writeln('Not authenticated.');
    Exit;
  end;

  var Tokens := FTokenStore.GetTokens;
  if Tokens.RefreshToken.IsEmpty then
  begin
    Writeln('No refresh token available.');
    Exit;
  end;

  var NewTokens := FOAuthClient.RefreshAccessToken(Tokens.RefreshToken);
  if NewTokens.RefreshToken.IsEmpty then
    NewTokens.RefreshToken := Tokens.RefreshToken;
  FTokenStore.StoreTokens(NewTokens);
  Writeln('Token refreshed! Expires in ', NewTokens.ExpiresIn, 's');
end;

procedure TDemoApp.Run;
begin
  Writeln('=== Microsoft365-4D Demo ===');
  Writeln;
  Writeln('Client ID:    ', FConfig.ClientId);
  Writeln('Tenant ID:    ', FConfig.TenantId);
  Writeln('Redirect URI: ', FConfig.RedirectUri);
  Writeln;

  var Running := True;
  while Running do
  begin
    PrintSeparator;
    Writeln('  1.  Authenticate');
    Writeln('  2.  List messages');
    Writeln('  3.  Read message');
    Writeln('  4.  Send email');
    Writeln('  5.  List mail folders');
    Writeln('  6.  List calendar events');
    Writeln('  7.  Create event');
    Writeln('  8.  Search contacts');
    Writeln('  9.  List SharePoint sites');
    Writeln('  10. Refresh token');
    Writeln('  0.  Exit');
    PrintSeparator;

    var Choice := '';
    Write('Choose: ');
    Readln(Choice);

    try
      if Choice = '1' then DoAuthenticate
      else if Choice = '2' then DoListMessages
      else if Choice = '3' then DoReadMessage
      else if Choice = '4' then DoSendEmail
      else if Choice = '5' then DoListFolders
      else if Choice = '6' then DoListEvents
      else if Choice = '7' then DoCreateEvent
      else if Choice = '8' then DoSearchContacts
      else if Choice = '9' then DoListSites
      else if Choice = '10' then DoRefreshToken
      else if Choice = '0' then Running := False
      else Writeln('Invalid option.');
    except
      on E: Exception do
        Writeln('ERROR: ', E.Message);
    end;
  end;

  Writeln('Goodbye!');
end;

end.
