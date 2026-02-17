program Microsoft365Demo;

{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  MSGraph.OAuth2.Types in '..\Source\OAuth2\MSGraph.OAuth2.Types.pas',
  MSGraph.OAuth2.PKCE in '..\Source\OAuth2\MSGraph.OAuth2.PKCE.pas',
  MSGraph.OAuth2.Client in '..\Source\OAuth2\MSGraph.OAuth2.Client.pas',
  MSGraph.OAuth2.TokenStore in '..\Source\OAuth2\MSGraph.OAuth2.TokenStore.pas',
  MSGraph.Graph.Http in '..\Source\Graph\MSGraph.Graph.Http.pas',
  MSGraph.Graph.JsonHelper in '..\Source\Graph\MSGraph.Graph.JsonHelper.pas',
  MSGraph.Graph.Mail.Types in '..\Source\Graph\MSGraph.Graph.Mail.Types.pas',
  MSGraph.Graph.Mail.Interfaces in '..\Source\Graph\MSGraph.Graph.Mail.Interfaces.pas',
  MSGraph.Graph.Mail in '..\Source\Graph\MSGraph.Graph.Mail.pas',
  MSGraph.Graph.Calendar.Types in '..\Source\Graph\MSGraph.Graph.Calendar.Types.pas',
  MSGraph.Graph.Calendar.Interfaces in '..\Source\Graph\MSGraph.Graph.Calendar.Interfaces.pas',
  MSGraph.Graph.Calendar in '..\Source\Graph\MSGraph.Graph.Calendar.pas',
  MSGraph.Graph.Contacts.Types in '..\Source\Graph\MSGraph.Graph.Contacts.Types.pas',
  MSGraph.Graph.Contacts.Interfaces in '..\Source\Graph\MSGraph.Graph.Contacts.Interfaces.pas',
  MSGraph.Graph.Contacts in '..\Source\Graph\MSGraph.Graph.Contacts.pas',
  MSGraph.Graph.SharePoint.Types in '..\Source\Graph\MSGraph.Graph.SharePoint.Types.pas',
  MSGraph.Graph.SharePoint.Interfaces in '..\Source\Graph\MSGraph.Graph.SharePoint.Interfaces.pas',
  MSGraph.Graph.SharePoint in '..\Source\Graph\MSGraph.Graph.SharePoint.pas',
  Microsoft365Demo.CallbackServer in 'Microsoft365Demo.CallbackServer.pas',
  Microsoft365Demo.App in 'Microsoft365Demo.App.pas';

function FindParamValue(const ParamName: string): string;
begin
  Result := '';
  for var Index := 1 to ParamCount do
  begin
    if ParamStr(Index) = ParamName then
    begin
      if Index < ParamCount then
        Result := ParamStr(Index + 1);
      Exit;
    end;
  end;
end;

function PromptForValue(const Prompt: string): string;
begin
  Write(Prompt);
  Readln(Result);
end;

begin
  try
    var Config: TOAuth2Config;
    Config.ClientId := FindParamValue('--client-id');
    Config.ClientSecret := FindParamValue('--client-secret');
    Config.TenantId := FindParamValue('--tenant-id');
    Config.RedirectUri := FindParamValue('--redirect-uri');
    var PortStr := FindParamValue('--port');

    if Config.ClientId.IsEmpty then
      Config.ClientId := PromptForValue('Client ID: ');

    if Config.ClientSecret.IsEmpty then
      Config.ClientSecret := PromptForValue('Client Secret: ');

    if Config.TenantId.IsEmpty then
      Config.TenantId := PromptForValue('Tenant ID: ');

    if PortStr.IsEmpty then
      Config.Port := 8080
    else
      Config.Port := StrToIntDef(PortStr, 8080);

    if Config.RedirectUri.IsEmpty then
      Config.RedirectUri := PromptForValue('Redirect URI (Enter for http://localhost:' + IntToStr(Config.Port) + '/oauth/callback): ');

    if Config.RedirectUri.IsEmpty then
      Config.RedirectUri := 'http://localhost:' + IntToStr(Config.Port) + '/oauth/callback';
    Config.Scopes := TArray<string>.Create(
      'openid', 'profile', 'offline_access',
      'Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'MailboxSettings.Read',
      'Calendars.ReadWrite', 'Contacts.ReadWrite',
      'Sites.Read.All', 'User.Read'
    );

    var App := TDemoApp.Create(Config);
    try
      App.Run;
    finally
      App.Free;
    end;
  except
    on E: Exception do
    begin
      Writeln('Fatal error: ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
