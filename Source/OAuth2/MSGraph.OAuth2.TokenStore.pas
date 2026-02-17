unit MSGraph.OAuth2.TokenStore;

interface

uses
  System.SysUtils,
  System.Generics.Collections,
  System.SyncObjs,
  MSGraph.OAuth2.Types;

type
  TTokenStore = class
  strict private
    FLock: TCriticalSection;
    FTokens: TOAuth2TokenResponse;
    FHasTokens: Boolean;
    FPKCESessions: TDictionary<string, TPKCESession>;

    procedure CleanupExpiredSessions;
  public
    constructor Create;
    destructor Destroy; override;

    procedure StoreTokens(const Tokens: TOAuth2TokenResponse);
    function GetTokens: TOAuth2TokenResponse;
    function HasValidTokens: Boolean;
    procedure ClearTokens;

    procedure StorePKCESession(const Session: TPKCESession);
    function RetrievePKCESession(const State: string): TPKCESession;
    procedure DeletePKCESession(const State: string);
  end;

implementation

constructor TTokenStore.Create;
begin
  inherited Create;
  FLock := TCriticalSection.Create;
  FPKCESessions := TDictionary<string, TPKCESession>.Create;
  FHasTokens := False;
end;

destructor TTokenStore.Destroy;
begin
  FPKCESessions.Free;
  FLock.Free;
  inherited;
end;

procedure TTokenStore.StoreTokens(const Tokens: TOAuth2TokenResponse);
begin
  FLock.Enter;
  try
    FTokens := Tokens;
    FHasTokens := True;
  finally
    FLock.Leave;
  end;
end;

function TTokenStore.GetTokens: TOAuth2TokenResponse;
begin
  FLock.Enter;
  try
    if not FHasTokens then
      raise ETokenStoreException.Create('No tokens stored');
    Result := FTokens;
  finally
    FLock.Leave;
  end;
end;

function TTokenStore.HasValidTokens: Boolean;
begin
  FLock.Enter;
  try
    Result := FHasTokens and (not FTokens.IsExpired);
  finally
    FLock.Leave;
  end;
end;

procedure TTokenStore.ClearTokens;
begin
  FLock.Enter;
  try
    FTokens := Default(TOAuth2TokenResponse);
    FHasTokens := False;
  finally
    FLock.Leave;
  end;
end;

procedure TTokenStore.StorePKCESession(const Session: TPKCESession);
begin
  FLock.Enter;
  try
    CleanupExpiredSessions;
    FPKCESessions.AddOrSetValue(Session.State, Session);
  finally
    FLock.Leave;
  end;
end;

function TTokenStore.RetrievePKCESession(const State: string): TPKCESession;
begin
  FLock.Enter;
  try
    if not FPKCESessions.TryGetValue(State, Result) then
      raise ETokenStoreException.Create('PKCE session not found for state: ' + State);

    if Now > Result.ExpiresAt then
    begin
      FPKCESessions.Remove(State);
      raise ETokenStoreException.Create('PKCE session expired');
    end;
  finally
    FLock.Leave;
  end;
end;

procedure TTokenStore.DeletePKCESession(const State: string);
begin
  FLock.Enter;
  try
    FPKCESessions.Remove(State);
  finally
    FLock.Leave;
  end;
end;

procedure TTokenStore.CleanupExpiredSessions;
begin
  var ExpiredKeys: TArray<string>;
  SetLength(ExpiredKeys, 0);

  for var Pair in FPKCESessions do
  begin
    if Now > Pair.Value.ExpiresAt then
    begin
      SetLength(ExpiredKeys, Length(ExpiredKeys) + 1);
      ExpiredKeys[High(ExpiredKeys)] := Pair.Key;
    end;
  end;

  for var Key in ExpiredKeys do
    FPKCESessions.Remove(Key);
end;

end.
