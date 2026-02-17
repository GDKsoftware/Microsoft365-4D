unit MSGraph.Graph.Contacts;

interface

uses
  System.JSON,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http,
  MSGraph.Graph.Contacts.Types,
  MSGraph.Graph.Contacts.Interfaces;

type
  TContactsClient = class(TInterfacedObject, IContactsClient)
  strict private
    FGraphClient: TGraphHttpClient;
    FOwnsClient: Boolean;

    function BuildContactBody(const GivenName: string; const Surname: string;
      const Email: string; const Phone: string; const Company: string;
      const JobTitle: string): TJSONObject;

    function EndpointContacts: string;

    class function ParseAddress(const AddressObj: TJSONObject): TPostalAddress; static;
    class function ParseContact(const ContactObj: TJSONObject): TContact; static;
    class function ParseCreateContactResult(const ResponseObj: TJSONObject): TCreateContactResult; static;
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

    function SearchContacts(const Query: string; const Top: Integer = 50): TArray<TContact>;
    function GetContact(const ContactId: string): TContact;
    function CreateContact(const GivenName: string; const Surname: string;
      const Email: string; const Phone: string; const Company: string;
      const JobTitle: string): TCreateContactResult;
    function UpdateContact(const ContactId: string; const GivenName: string;
      const Surname: string; const Email: string; const Phone: string;
      const Company: string; const JobTitle: string): TCreateContactResult;
    function DeleteContact(const ContactId: string): Boolean;

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.SysUtils,
  System.NetEncoding,
  System.Generics.Collections,
  MSGraph.Graph.JsonHelper;

constructor TContactsClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FGraphClient := TGraphHttpClient.Create(AccessToken, LogProc);
  FOwnsClient := True;
end;

constructor TContactsClient.Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean);
begin
  inherited Create;
  FGraphClient := GraphClient;
  FOwnsClient := OwnsClient;
end;

destructor TContactsClient.Destroy;
begin
  if FOwnsClient then
    FGraphClient.Free;
  inherited;
end;

function TContactsClient.EndpointContacts: string;
begin
  Result := FGraphClient.GetUserPrefix + '/contacts';
end;

function TContactsClient.BuildContactBody(const GivenName: string; const Surname: string;
  const Email: string; const Phone: string; const Company: string;
  const JobTitle: string): TJSONObject;
begin
  Result := TJSONObject.Create;

  if not GivenName.Trim.IsEmpty then
    Result.AddPair('givenName', GivenName);

  if not Surname.Trim.IsEmpty then
    Result.AddPair('surname', Surname);

  if not Email.Trim.IsEmpty then
  begin
    var EmailAddresses := TJSONArray.Create;
    var EmailObj := TJSONObject.Create;
    EmailObj.AddPair('address', Email);
    EmailObj.AddPair('name', GivenName + ' ' + Surname);
    EmailAddresses.Add(EmailObj);
    Result.AddPair('emailAddresses', EmailAddresses);
  end;

  if not Phone.Trim.IsEmpty then
  begin
    var BusinessPhones := TJSONArray.Create;
    BusinessPhones.Add(Phone);
    Result.AddPair('businessPhones', BusinessPhones);
  end;

  if not Company.Trim.IsEmpty then
    Result.AddPair('companyName', Company);

  if not JobTitle.Trim.IsEmpty then
    Result.AddPair('jobTitle', JobTitle);
end;

class function TContactsClient.ParseAddress(const AddressObj: TJSONObject): TPostalAddress;
begin
  Result := Default(TPostalAddress);
  if not Assigned(AddressObj) then
    Exit;
  Result.Street := TGraphJson.GetString(AddressObj, 'street');
  Result.City := TGraphJson.GetString(AddressObj, 'city');
  Result.State := TGraphJson.GetString(AddressObj, 'state');
  Result.PostalCode := TGraphJson.GetString(AddressObj, 'postalCode');
  Result.Country := TGraphJson.GetString(AddressObj, 'countryOrRegion');
end;

class function TContactsClient.ParseContact(const ContactObj: TJSONObject): TContact;
begin
  Result := Default(TContact);
  if not Assigned(ContactObj) then
    Exit;

  Result.Id := TGraphJson.GetString(ContactObj, 'id');
  Result.GivenName := TGraphJson.GetString(ContactObj, 'givenName');
  Result.Surname := TGraphJson.GetString(ContactObj, 'surname');
  Result.DisplayName := TGraphJson.GetString(ContactObj, 'displayName');
  Result.Company := TGraphJson.GetString(ContactObj, 'companyName');
  Result.JobTitle := TGraphJson.GetString(ContactObj, 'jobTitle');
  Result.Department := TGraphJson.GetString(ContactObj, 'department');
  Result.OfficeLocation := TGraphJson.GetString(ContactObj, 'officeLocation');
  Result.MobilePhone := TGraphJson.GetString(ContactObj, 'mobilePhone');
  Result.Birthday := TGraphJson.GetString(ContactObj, 'birthday');
  Result.PersonalNotes := TGraphJson.GetString(ContactObj, 'personalNotes');

  var EmailsArray := TGraphJson.GetArray(ContactObj, 'emailAddresses');
  if Assigned(EmailsArray) and (EmailsArray.Count > 0) then
  begin
    var FirstEmail := TGraphJson.ArrayItem(EmailsArray, 0);
    Result.Email := TGraphJson.GetString(FirstEmail, 'address');
  end;

  var BusinessPhones := TGraphJson.GetArray(ContactObj, 'businessPhones');
  if Assigned(BusinessPhones) and (BusinessPhones.Count > 0) then
    Result.BusinessPhone := BusinessPhones.Items[0].Value;

  var HomePhones := TGraphJson.GetArray(ContactObj, 'homePhones');
  if Assigned(HomePhones) and (HomePhones.Count > 0) then
    Result.HomePhone := HomePhones.Items[0].Value;

  Result.BusinessAddress := ParseAddress(TGraphJson.GetObject(ContactObj, 'businessAddress'));
  Result.HomeAddress := ParseAddress(TGraphJson.GetObject(ContactObj, 'homeAddress'));
end;

class function TContactsClient.ParseCreateContactResult(const ResponseObj: TJSONObject): TCreateContactResult;
begin
  Result := Default(TCreateContactResult);
  if not Assigned(ResponseObj) then
    Exit;
  Result.Id := TGraphJson.GetString(ResponseObj, 'id');
  Result.DisplayName := TGraphJson.GetString(ResponseObj, 'displayName');
end;

function TContactsClient.SearchContacts(const Query: string; const Top: Integer): TArray<TContact>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 50
  else if ActualTop > 100 then
    ActualTop := 100;

  var QueryParams := Format('$top=%d&$select=id,givenName,surname,displayName,emailAddresses,businessPhones,companyName,jobTitle',
    [ActualTop]);

  const HasSearchQuery = not Query.Trim.IsEmpty;
  if HasSearchQuery then
    QueryParams := QueryParams + '&$search="' + TNetEncoding.URL.Encode(Query) + '"'
  else
    QueryParams := QueryParams + '&$orderby=displayName';

  var Response := FGraphClient.Get(EndpointContacts, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseContact(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TContactsClient.GetContact(const ContactId: string): TContact;
begin
  var Response := FGraphClient.Get(EndpointContacts + '/' + ContactId,
    '$select=id,givenName,surname,displayName,emailAddresses,businessPhones,mobilePhone,homePhones,companyName,jobTitle,department,officeLocation,businessAddress,homeAddress,birthday,personalNotes');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseContact(Response);
  finally
    Response.Free;
  end;
end;

function TContactsClient.CreateContact(const GivenName: string; const Surname: string;
  const Email: string; const Phone: string; const Company: string;
  const JobTitle: string): TCreateContactResult;
begin
  var ContactObj := BuildContactBody(GivenName, Surname, Email, Phone, Company, JobTitle);
  try
    var Response := FGraphClient.Post(EndpointContacts, ContactObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result := ParseCreateContactResult(Response);
    finally
      Response.Free;
    end;
  finally
    ContactObj.Free;
  end;
end;

function TContactsClient.UpdateContact(const ContactId: string; const GivenName: string;
  const Surname: string; const Email: string; const Phone: string;
  const Company: string; const JobTitle: string): TCreateContactResult;
begin
  var ContactObj := BuildContactBody(GivenName, Surname, Email, Phone, Company, JobTitle);
  try
    var Response := FGraphClient.Patch(EndpointContacts + '/' + ContactId, ContactObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result := ParseCreateContactResult(Response);
    finally
      Response.Free;
    end;
  finally
    ContactObj.Free;
  end;
end;

function TContactsClient.DeleteContact(const ContactId: string): Boolean;
begin
  var Response := FGraphClient.Delete(EndpointContacts + '/' + ContactId);
  try
    Result := not TGraphJson.HasError(Response);
  finally
    Response.Free;
  end;
end;

end.
