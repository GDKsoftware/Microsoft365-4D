unit MSGraph.Graph.People;

interface

uses
  System.JSON,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http,
  MSGraph.Graph.People.Types,
  MSGraph.Graph.People.Interfaces;

type
  TPeopleClient = class(TInterfacedObject, IPeopleClient)
  strict private
    FGraphClient: TGraphHttpClient;
    FOwnsClient: Boolean;

    function EndpointPeople: string;

    class function ParsePerson(const PersonObj: TJSONObject): TPerson; static;
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

    function GetRelevantPeople(const Top: Integer = 20): TArray<TPerson>;
    function SearchPeople(const Query: string; const Top: Integer = 10): TArray<TPerson>;

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.SysUtils,
  System.NetEncoding,
  MSGraph.Graph.JsonHelper;

constructor TPeopleClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FGraphClient := TGraphHttpClient.Create(AccessToken, LogProc);
  FOwnsClient := True;
end;

constructor TPeopleClient.Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean);
begin
  inherited Create;
  FGraphClient := GraphClient;
  FOwnsClient := OwnsClient;
end;

destructor TPeopleClient.Destroy;
begin
  if FOwnsClient then
    FGraphClient.Free;
  inherited;
end;

function TPeopleClient.EndpointPeople: string;
begin
  Result := FGraphClient.GetUserPrefix + '/people';
end;

class function TPeopleClient.ParsePerson(const PersonObj: TJSONObject): TPerson;
begin
  Result := Default(TPerson);
  if not Assigned(PersonObj) then
    Exit;

  Result.Id := TGraphJson.GetString(PersonObj, 'id');
  Result.DisplayName := TGraphJson.GetString(PersonObj, 'displayName');

  var EmailsArray := TGraphJson.GetArray(PersonObj, 'scoredEmailAddresses');
  if Assigned(EmailsArray) and (EmailsArray.Count > 0) then
  begin
    var FirstEmail := TGraphJson.ArrayItem(EmailsArray, 0);
    if Assigned(FirstEmail) then
      Result.EmailAddress := TGraphJson.GetString(FirstEmail, 'address');
  end;
end;

function TPeopleClient.GetRelevantPeople(const Top: Integer): TArray<TPerson>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 20
  else if ActualTop > 100 then
    ActualTop := 100;

  var QueryParams := Format(
    '$top=%d&$select=id,displayName,scoredEmailAddresses',
    [ActualTop]);

  var Response := FGraphClient.Get(EndpointPeople, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParsePerson(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TPeopleClient.SearchPeople(const Query: string; const Top: Integer): TArray<TPerson>;
begin
  Result := nil;

  const HasQuery = not Query.Trim.IsEmpty;
  if not HasQuery then
    Exit;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 10
  else if ActualTop > 50 then
    ActualTop := 50;

  var QueryParams := Format(
    '$top=%d&$select=id,displayName,scoredEmailAddresses&$search="%s"',
    [ActualTop, TNetEncoding.URL.Encode(Query)]);

  var Response := FGraphClient.Get(EndpointPeople, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParsePerson(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

end.
