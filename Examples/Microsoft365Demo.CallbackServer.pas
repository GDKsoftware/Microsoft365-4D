unit Microsoft365Demo.CallbackServer;

interface

uses
  System.SysUtils,
  System.Classes,
  System.SyncObjs,
  IdHTTPServer,
  IdContext,
  IdCustomHTTPServer;

type
  TCallbackServer = class
  strict private
    FHttpServer: TIdHTTPServer;
    FPort: Integer;
    FCallbackEvent: TEvent;
    FCode: string;
    FState: string;
    FError: string;

    procedure HandleRequest(AContext: TIdContext;
      ARequestInfo: TIdHTTPRequestInfo;
      AResponseInfo: TIdHTTPResponseInfo);
    function ParseQueryParam(const Query: string; const ParamName: string): string;
  public
    constructor Create(const Port: Integer);
    destructor Destroy; override;

    procedure Start;
    procedure Stop;
    function WaitForCallback(const TimeoutMs: Cardinal = 120000): Boolean;

    property Code: string read FCode;
    property State: string read FState;
    property Error: string read FError;
    property Port: Integer read FPort;
  end;

implementation

constructor TCallbackServer.Create(const Port: Integer);
begin
  inherited Create;
  FPort := Port;
  FCallbackEvent := TEvent.Create(nil, True, False, '');
  FHttpServer := TIdHTTPServer.Create(nil);
  FHttpServer.DefaultPort := FPort;
  FHttpServer.OnCommandGet := HandleRequest;
end;

destructor TCallbackServer.Destroy;
begin
  if FHttpServer.Active then
    FHttpServer.Active := False;
  FHttpServer.Free;
  FCallbackEvent.Free;
  inherited;
end;

procedure TCallbackServer.Start;
begin
  FCallbackEvent.ResetEvent;
  FCode := '';
  FState := '';
  FError := '';
  FHttpServer.Active := True;
end;

procedure TCallbackServer.Stop;
begin
  if FHttpServer.Active then
    FHttpServer.Active := False;
end;

function TCallbackServer.WaitForCallback(const TimeoutMs: Cardinal): Boolean;
begin
  Result := (FCallbackEvent.WaitFor(TimeoutMs) = wrSignaled);
end;

function TCallbackServer.ParseQueryParam(const Query: string; const ParamName: string): string;
begin
  Result := '';
  var Pairs := Query.Split(['&']);
  for var Pair in Pairs do
  begin
    var KeyValue := Pair.Split(['=']);
    if (Length(KeyValue) = 2) and (KeyValue[0] = ParamName) then
    begin
      Result := KeyValue[1];
      Exit;
    end;
  end;
end;

procedure TCallbackServer.HandleRequest(AContext: TIdContext;
  ARequestInfo: TIdHTTPRequestInfo;
  AResponseInfo: TIdHTTPResponseInfo);
const
  SuccessHtml =
    '<!DOCTYPE html><html><head><title>Authentication Successful</title>' +
    '<style>body{font-family:Segoe UI,sans-serif;display:flex;justify-content:center;' +
    'align-items:center;min-height:100vh;margin:0;background:#f5f5f5}' +
    '.card{background:white;padding:40px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.1);' +
    'text-align:center;max-width:500px}h1{color:#4caf50}p{color:#666}</style></head>' +
    '<body><div class="card"><h1>Authentication Successful!</h1>' +
    '<p>You can close this window and return to the console application.</p></div></body></html>';
begin
  const IsCallbackPath = ARequestInfo.Document.StartsWith('/oauth/callback');
  if not IsCallbackPath then
  begin
    AResponseInfo.ResponseNo := 404;
    AResponseInfo.ContentText := 'Not found';
    Exit;
  end;

  var QueryString := ARequestInfo.QueryParams;
  FCode := ParseQueryParam(QueryString, 'code');
  FState := ParseQueryParam(QueryString, 'state');
  FError := ParseQueryParam(QueryString, 'error');

  AResponseInfo.ResponseNo := 200;
  AResponseInfo.ContentType := 'text/html; charset=utf-8';

  const HasError = not FError.IsEmpty;
  if HasError then
  begin
    var ErrorDesc := ParseQueryParam(QueryString, 'error_description');
    AResponseInfo.ContentText :=
      '<!DOCTYPE html><html><head><title>Authentication Failed</title>' +
      '<style>body{font-family:Segoe UI,sans-serif;display:flex;justify-content:center;' +
      'align-items:center;min-height:100vh;margin:0;background:#f5f5f5}' +
      '.card{background:white;padding:40px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.1);' +
      'text-align:center;max-width:500px}h1{color:#d32f2f}p{color:#666}</style></head>' +
      '<body><div class="card"><h1>Authentication Failed</h1>' +
      '<p>' + FError + ': ' + ErrorDesc + '</p>' +
      '<p>You can close this window.</p></div></body></html>';
  end
  else
    AResponseInfo.ContentText := SuccessHtml;

  FCallbackEvent.SetEvent;
end;

end.
