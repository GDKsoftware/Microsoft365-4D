unit MSGraph.Graph.Mail;

interface

uses
  System.JSON,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http,
  MSGraph.Graph.Mail.Types,
  MSGraph.Graph.Mail.Interfaces;

type
  TMailClient = class(TInterfacedObject, IMailClient)
  strict private
    FGraphClient: TGraphHttpClient;
    FOwnsClient: Boolean;

    function BuildRecipientArray(const Recipients: TArray<string>): TJSONArray;
    function BuildMessageBody(const Subject: string; const Body: string;
      const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
      const BccRecipients: TArray<string>; const IsHtml: Boolean): TJSONObject;
    function MessageEndpoint(const MessageId: string): string;

    function EndpointMessages: string;
    function EndpointMailFolders: string;
    function EndpointMailboxSettings: string;

    class function ParseEmailAddress(const Obj: TJSONObject): TEmailAddress; static;
    class function ParseRecipients(const MsgObj: TJSONObject; const FieldName: string): TArray<TEmailAddress>; static;
    class function ParseMessage(const MsgObj: TJSONObject): TMailMessage; static;
    class function ParseFolder(const FolderObj: TJSONObject): TMailFolder; static;
    class function ParseAttachment(const AttachObj: TJSONObject): TMailAttachment; static;
    class function BuildSearchQueryParams(const SearchQuery: string; const UseSearch: Boolean;
      const FilterUnread: Boolean; const ActualTop: Integer; const Skip: Integer): string; static;

    const
      ContentTypeHtml = 'HTML';
      ContentTypeText = 'Text';
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

    function SearchMessages(const Query: string; const FolderId: string;
      const Top: Integer; const Skip: Integer): TSearchMessagesResult;
    function GetMessage(const MessageId: string; const IncludeBody: Boolean = True): TMailMessage;
    function GetMessageAttachments(const MessageId: string): TArray<TMailAttachment>;
    function GetAttachmentContent(const MessageId: string; const AttachmentId: string): TMailAttachment;
    function CreateDraft(const Subject: string; const Body: string;
      const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
      const BccRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;
    function UpdateDraft(const MessageId: string; const Subject: string; const Body: string;
      const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
      const BccRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;
    function SendDraft(const MessageId: string): Boolean;
    function DeleteDraft(const MessageId: string): Boolean;
    function GetMailboxSignature: string;
    function CreateReplyDraft(const MessageId: string; const Body: string;
      const CcRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;
    function MoveMessage(const MessageId: string; const DestinationFolderId: string): TMoveMessageResult;
    function ListMailFolders(const ParentFolderId: string = ''): TArray<TMailFolder>;
    function ListFolderMessages(const FolderId: string; const Top: Integer = 50;
      const Skip: Integer = 0): TSearchMessagesResult;

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.SysUtils,
  System.NetEncoding,
  MSGraph.Graph.JsonHelper;

constructor TMailClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FGraphClient := TGraphHttpClient.Create(AccessToken, LogProc);
  FOwnsClient := True;
end;

constructor TMailClient.Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean);
begin
  inherited Create;
  FGraphClient := GraphClient;
  FOwnsClient := OwnsClient;
end;

destructor TMailClient.Destroy;
begin
  if FOwnsClient then
    FGraphClient.Free;
  inherited;
end;

function TMailClient.BuildRecipientArray(const Recipients: TArray<string>): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var Recipient in Recipients do
  begin
    var RecipientObj := TJSONObject.Create;
    var EmailObj := TJSONObject.Create;
    EmailObj.AddPair('address', Recipient);
    RecipientObj.AddPair('emailAddress', EmailObj);
    Result.Add(RecipientObj);
  end;
end;

function TMailClient.BuildMessageBody(const Subject: string; const Body: string;
  const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
  const BccRecipients: TArray<string>; const IsHtml: Boolean): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('subject', Subject);

  var BodyObj := TJSONObject.Create;
  if IsHtml then
    BodyObj.AddPair('contentType', ContentTypeHtml)
  else
    BodyObj.AddPair('contentType', ContentTypeText);
  BodyObj.AddPair('content', Body);
  Result.AddPair('body', BodyObj);

  Result.AddPair('toRecipients', BuildRecipientArray(ToRecipients));

  const HasCcRecipients = (Length(CcRecipients) > 0);
  if HasCcRecipients then
    Result.AddPair('ccRecipients', BuildRecipientArray(CcRecipients));

  const HasBccRecipients = (Length(BccRecipients) > 0);
  if HasBccRecipients then
    Result.AddPair('bccRecipients', BuildRecipientArray(BccRecipients));
end;

function TMailClient.EndpointMessages: string;
begin
  Result := FGraphClient.GetUserPrefix + '/messages';
end;

function TMailClient.EndpointMailFolders: string;
begin
  Result := FGraphClient.GetUserPrefix + '/mailFolders';
end;

function TMailClient.EndpointMailboxSettings: string;
begin
  Result := FGraphClient.GetUserPrefix + '/mailboxSettings';
end;

function TMailClient.MessageEndpoint(const MessageId: string): string;
begin
  Result := EndpointMessages + '/' + TNetEncoding.URL.Encode(MessageId);
end;

class function TMailClient.ParseEmailAddress(const Obj: TJSONObject): TEmailAddress;
begin
  Result := Default(TEmailAddress);
  var EmailAddr := TGraphJson.GetObject(Obj, 'emailAddress');
  if not Assigned(EmailAddr) then
    Exit;
  Result.Name := TGraphJson.GetString(EmailAddr, 'name');
  Result.Address := TGraphJson.GetString(EmailAddr, 'address');
end;

class function TMailClient.ParseRecipients(const MsgObj: TJSONObject; const FieldName: string): TArray<TEmailAddress>;
begin
  Result := nil;
  var Arr := TGraphJson.GetArray(MsgObj, FieldName);
  if not Assigned(Arr) then
    Exit;
  SetLength(Result, Arr.Count);
  for var Index := 0 to Arr.Count - 1 do
    Result[Index] := ParseEmailAddress(TGraphJson.ArrayItem(Arr, Index));
end;

class function TMailClient.ParseMessage(const MsgObj: TJSONObject): TMailMessage;
begin
  Result := Default(TMailMessage);
  if not Assigned(MsgObj) then
    Exit;
  Result.Id := TGraphJson.GetString(MsgObj, 'id');
  Result.Subject := TGraphJson.GetString(MsgObj, 'subject');
  Result.From := ParseEmailAddress(TGraphJson.GetObject(MsgObj, 'from'));
  Result.ToRecipients := ParseRecipients(MsgObj, 'toRecipients');
  Result.CcRecipients := ParseRecipients(MsgObj, 'ccRecipients');
  Result.ReceivedDateTime := TGraphJson.GetString(MsgObj, 'receivedDateTime');
  Result.IsRead := TGraphJson.GetBool(MsgObj, 'isRead');
  Result.HasAttachments := TGraphJson.GetBool(MsgObj, 'hasAttachments');
  Result.BodyPreview := TGraphJson.GetString(MsgObj, 'bodyPreview');
  Result.Importance := TGraphJson.GetString(MsgObj, 'importance');
  Result.ParentFolderId := TGraphJson.GetString(MsgObj, 'parentFolderId');

  var BodyObj := TGraphJson.GetObject(MsgObj, 'body');
  if not Assigned(BodyObj) then
    Exit;
  Result.Body := TGraphJson.GetString(BodyObj, 'content');
  Result.BodyType := TGraphJson.GetString(BodyObj, 'contentType');
end;

class function TMailClient.ParseFolder(const FolderObj: TJSONObject): TMailFolder;
begin
  Result := Default(TMailFolder);
  if not Assigned(FolderObj) then
    Exit;
  Result.Id := TGraphJson.GetString(FolderObj, 'id');
  Result.DisplayName := TGraphJson.GetString(FolderObj, 'displayName');
  Result.ParentFolderId := TGraphJson.GetString(FolderObj, 'parentFolderId');
  Result.ChildFolderCount := TGraphJson.GetInt(FolderObj, 'childFolderCount');
  Result.TotalItemCount := TGraphJson.GetInt(FolderObj, 'totalItemCount');
  Result.UnreadItemCount := TGraphJson.GetInt(FolderObj, 'unreadItemCount');
end;

class function TMailClient.ParseAttachment(const AttachObj: TJSONObject): TMailAttachment;
begin
  Result := Default(TMailAttachment);
  if not Assigned(AttachObj) then
    Exit;
  Result.Id := TGraphJson.GetString(AttachObj, 'id');
  Result.Name := TGraphJson.GetString(AttachObj, 'name');
  Result.ContentType := TGraphJson.GetString(AttachObj, 'contentType');
  Result.Size := TGraphJson.GetInt64(AttachObj, 'size');
  Result.IsInline := TGraphJson.GetBool(AttachObj, 'isInline');
  Result.ContentBytes := TGraphJson.GetString(AttachObj, 'contentBytes');
end;

class function TMailClient.BuildSearchQueryParams(const SearchQuery: string; const UseSearch: Boolean;
  const FilterUnread: Boolean; const ActualTop: Integer; const Skip: Integer): string;
begin
  Result := '$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,importance,parentFolderId';

  const IncludeBody = (ActualTop <= 5);
  if IncludeBody then
    Result := Result + ',body';

  if not UseSearch then
    Result := Result + '&$orderby=receivedDateTime desc';

  Result := Result + '&$top=' + ActualTop.ToString;

  const HasSkip = (Skip > 0) and (not UseSearch);
  if HasSkip then
    Result := Result + '&$skip=' + Skip.ToString;

  const ApplyUnreadFilter = FilterUnread and (not UseSearch);
  if ApplyUnreadFilter then
    Result := Result + '&$filter=isRead eq false'
  else if UseSearch then
    Result := Result + '&$search="' + TNetEncoding.URL.Encode(SearchQuery) + '"';
end;

function TMailClient.SearchMessages(const Query: string; const FolderId: string;
  const Top: Integer; const Skip: Integer): TSearchMessagesResult;
begin
  Result := Default(TSearchMessagesResult);

  const FilterUnread = (Query.ToLower.Contains('is:unread') or Query.ToLower.Contains('isread:false'));
  var SearchQuery := Query;

  if FilterUnread then
  begin
    SearchQuery := SearchQuery.Replace('is:unread', '', [rfReplaceAll, rfIgnoreCase]);
    SearchQuery := SearchQuery.Replace('isread:false', '', [rfReplaceAll, rfIgnoreCase]);
    SearchQuery := SearchQuery.Trim;
  end;

  const UseSearch = ((not SearchQuery.IsEmpty) and (SearchQuery <> '*'));

  var Endpoint: string;
  const HasFolderId = not FolderId.IsEmpty;
  if HasFolderId then
    Endpoint := EndpointMailFolders + '/' + TNetEncoding.URL.Encode(FolderId) + '/messages'
  else if FGraphClient.IsSharedMailbox then
    Endpoint := EndpointMailFolders + '/Inbox/messages'
  else
    Endpoint := EndpointMessages;

  var ActualTop := Top;
  if ActualTop > 50 then
    ActualTop := 50;
  if ActualTop < 1 then
    ActualTop := 20;

  var QueryParams := BuildSearchQueryParams(SearchQuery, UseSearch, FilterUnread, ActualTop, Skip);

  var Response := FGraphClient.Get(Endpoint, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result.Messages, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result.Messages[Index] := ParseMessage(TGraphJson.ArrayItem(ValueArray, Index));

    Result.HasMore := (Length(Result.Messages) >= ActualTop) or TGraphJson.HasNextPage(Response);
  finally
    Response.Free;
  end;
end;

function TMailClient.GetMessage(const MessageId: string; const IncludeBody: Boolean): TMailMessage;
begin
  var Response := FGraphClient.Get(MessageEndpoint(MessageId),
    '$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,body,bodyPreview,importance,parentFolderId');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseMessage(Response);
  finally
    Response.Free;
  end;
end;

function TMailClient.GetMessageAttachments(const MessageId: string): TArray<TMailAttachment>;
begin
  Result := nil;

  var Response := FGraphClient.Get(MessageEndpoint(MessageId) + '/attachments',
    '$select=id,name,contentType,size,isInline');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseAttachment(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TMailClient.GetAttachmentContent(const MessageId: string; const AttachmentId: string): TMailAttachment;
begin
  var Endpoint := MessageEndpoint(MessageId) + '/attachments/' + TNetEncoding.URL.Encode(AttachmentId);

  var Response := FGraphClient.Get(Endpoint);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseAttachment(Response);
  finally
    Response.Free;
  end;
end;

function TMailClient.GetMailboxSignature: string;
begin
  Result := '';
  var Response := FGraphClient.Get(EndpointMailboxSettings);
  if not Assigned(Response) then
    Exit;

  try
    if TGraphJson.HasError(Response) then
      Exit;

    Result := TGraphJson.GetString(Response, 'signatureHtml');
  finally
    Response.Free;
  end;
end;

function TMailClient.CreateDraft(const Subject: string; const Body: string;
  const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
  const BccRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;
begin
  Result := Default(TDraftResult);
  var Signature := GetMailboxSignature;
  var FinalBody := Body;

  const HasSignature = not Signature.IsEmpty;
  if HasSignature then
  begin
    if IsHtml then
      FinalBody := FinalBody + '<br><br>' + Signature
    else
      FinalBody := FinalBody + #13#10#13#10 + Signature;
  end;

  var MessageObj := BuildMessageBody(Subject, FinalBody, ToRecipients, CcRecipients, BccRecipients, IsHtml);
  try
    var DraftEndpoint: string;
    if FGraphClient.IsSharedMailbox then
      DraftEndpoint := EndpointMailFolders + '/Drafts/messages'
    else
      DraftEndpoint := EndpointMessages;

    var Response := FGraphClient.Post(DraftEndpoint, MessageObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.Id := TGraphJson.GetString(Response, 'id');
    finally
      Response.Free;
    end;
  finally
    MessageObj.Free;
  end;
end;

function TMailClient.UpdateDraft(const MessageId: string; const Subject: string; const Body: string;
  const ToRecipients: TArray<string>; const CcRecipients: TArray<string>;
  const BccRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;
begin
  Result := Default(TDraftResult);
  var MessageObj := BuildMessageBody(Subject, Body, ToRecipients, CcRecipients, BccRecipients, IsHtml);
  try
    var Response := FGraphClient.Patch(MessageEndpoint(MessageId), MessageObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.Id := TGraphJson.GetString(Response, 'id');
    finally
      Response.Free;
    end;
  finally
    MessageObj.Free;
  end;
end;

function TMailClient.SendDraft(const MessageId: string): Boolean;
begin
  var Response := FGraphClient.Post(MessageEndpoint(MessageId) + '/send');
  try
    Result := not TGraphJson.HasError(Response);
  finally
    Response.Free;
  end;
end;

function TMailClient.DeleteDraft(const MessageId: string): Boolean;
begin
  var Response := FGraphClient.Delete(MessageEndpoint(MessageId));
  try
    Result := not TGraphJson.HasError(Response);
  finally
    Response.Free;
  end;
end;

function TMailClient.CreateReplyDraft(const MessageId: string; const Body: string;
  const CcRecipients: TArray<string>; const IsHtml: Boolean): TDraftResult;

  function FormatRecipients(const Recipients: TArray<TEmailAddress>): string;
  begin
    Result := '';
    for var Index := 0 to High(Recipients) do
    begin
      if Index > 0 then
        Result := Result + '; ';
      if not Recipients[Index].Name.IsEmpty then
        Result := Result + Recipients[Index].Name + ' &lt;' + Recipients[Index].Address + '&gt;'
      else
        Result := Result + Recipients[Index].Address;
    end;
  end;

  function BuildQuotedOriginal(const Original: TMailMessage): string;
  begin
    Result :=
      '<hr style="display:inline-block;width:98%">' +
      '<div id="divRplyFwdMsg" dir="ltr">' +
      '<font face="Calibri, sans-serif" color="#000000" style="font-size:11pt">' +
      '<b>From:</b> ' + FormatRecipients([Original.From]) + '<br>' +
      '<b>Sent:</b> ' + Original.ReceivedDateTime + '<br>' +
      '<b>To:</b> ' + FormatRecipients(Original.ToRecipients) + '<br>';
    const HasCc = (Length(Original.CcRecipients) > 0);
    if HasCc then
      Result := Result + '<b>Cc:</b> ' + FormatRecipients(Original.CcRecipients) + '<br>';
    Result := Result +
      '<b>Subject:</b> ' + Original.Subject +
      '</font><br><br>' + Original.Body + '</div>';
  end;

  function BuildCreateReplyBody(const CombinedBody: string; const ContentType: string): TJSONObject;
  begin
    Result := TJSONObject.Create;
    var MessageObj := TJSONObject.Create;
    var BodyObj := TJSONObject.Create;
    BodyObj.AddPair('contentType', ContentType);
    BodyObj.AddPair('content', CombinedBody);
    MessageObj.AddPair('body', BodyObj);
    Result.AddPair('message', MessageObj);
  end;

begin
  Result := Default(TDraftResult);

  var OriginalMessage := GetMessage(MessageId);

  var CombinedBody: string;
  var ContentType: string;
  if IsHtml then
  begin
    CombinedBody := Body + BuildQuotedOriginal(OriginalMessage);
    ContentType := ContentTypeHtml;
  end
  else
  begin
    CombinedBody := Body + #13#10#13#10 + '---' + #13#10 +
      'From: ' + OriginalMessage.From.Address + #13#10 +
      'Subject: ' + OriginalMessage.Subject + #13#10#13#10 +
      OriginalMessage.Body;
    ContentType := ContentTypeText;
  end;

  var RequestBody := BuildCreateReplyBody(CombinedBody, ContentType);
  try
    var Response := FGraphClient.Post(MessageEndpoint(MessageId) + '/createReply', RequestBody.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.Id := TGraphJson.GetString(Response, 'id');
      Result.Subject := TGraphJson.GetString(Response, 'subject');
    finally
      Response.Free;
    end;
  finally
    RequestBody.Free;
  end;

  const HasAdditionalCc = (Length(CcRecipients) > 0);
  if HasAdditionalCc then
  begin
    var DraftMessage := GetMessage(Result.Id);

    var MergedCc: TArray<string>;
    SetLength(MergedCc, Length(DraftMessage.CcRecipients));
    for var Index := 0 to High(DraftMessage.CcRecipients) do
      MergedCc[Index] := DraftMessage.CcRecipients[Index].Address;

    for var NewCc in CcRecipients do
    begin
      var AlreadyExists := False;
      for var Existing in MergedCc do
      begin
        if Existing.ToLower = NewCc.ToLower then
        begin
          AlreadyExists := True;
          Break;
        end;
      end;
      if not AlreadyExists then
      begin
        SetLength(MergedCc, Length(MergedCc) + 1);
        MergedCc[High(MergedCc)] := NewCc;
      end;
    end;

    var PatchObj := TJSONObject.Create;
    try
      PatchObj.AddPair('ccRecipients', BuildRecipientArray(MergedCc));
      var PatchResponse := FGraphClient.Patch(MessageEndpoint(Result.Id), PatchObj.ToJSON);
      try
        if TGraphJson.HasError(PatchResponse) then
          raise EGraphApiException.Create(TGraphJson.GetErrorMessage(PatchResponse));
      finally
        PatchResponse.Free;
      end;
    finally
      PatchObj.Free;
    end;
  end;
end;

function TMailClient.MoveMessage(const MessageId: string; const DestinationFolderId: string): TMoveMessageResult;
begin
  Result := Default(TMoveMessageResult);
  var RequestBody := TJSONObject.Create;
  try
    RequestBody.AddPair('destinationId', DestinationFolderId);
    var Response := FGraphClient.Post(MessageEndpoint(MessageId) + '/move', RequestBody.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.NewMessageId := TGraphJson.GetString(Response, 'id');
    finally
      Response.Free;
    end;
  finally
    RequestBody.Free;
  end;
end;

function TMailClient.ListMailFolders(const ParentFolderId: string): TArray<TMailFolder>;

  function FetchWellKnownFolders: TArray<TMailFolder>;
  const
    WellKnownFolderNames: array[0..5] of string = (
      'Inbox', 'SentItems', 'Drafts', 'DeletedItems', 'Archive', 'JunkEmail'
    );
    SelectFields = '$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount';
  begin
    Result := nil;
    var Folders: TArray<TMailFolder>;
    SetLength(Folders, 0);
    for var FolderName in WellKnownFolderNames do
    begin
      var Response := FGraphClient.Get(EndpointMailFolders + '/' + FolderName, SelectFields);
      try
        if TGraphJson.HasError(Response) then
          Continue;
        SetLength(Folders, Length(Folders) + 1);
        Folders[High(Folders)] := ParseFolder(Response);
      finally
        Response.Free;
      end;
    end;
    Result := Folders;
  end;

begin
  Result := nil;

  if ParentFolderId.IsEmpty and FGraphClient.IsSharedMailbox then
  begin
    Result := FetchWellKnownFolders;
    Exit;
  end;

  var Endpoint: string;
  if ParentFolderId.IsEmpty then
    Endpoint := EndpointMailFolders
  else
    Endpoint := EndpointMailFolders + '/' + ParentFolderId + '/childFolders';

  var Response := FGraphClient.Get(Endpoint,
    '$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseFolder(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TMailClient.ListFolderMessages(const FolderId: string;
  const Top: Integer; const Skip: Integer): TSearchMessagesResult;
const
  SelectFields = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,' +
    'isRead,hasAttachments,bodyPreview,body,importance,parentFolderId';
begin
  Result := Default(TSearchMessagesResult);

  var Endpoint: string;
  if not FolderId.IsEmpty then
    Endpoint := EndpointMailFolders + '/' + TNetEncoding.URL.Encode(FolderId) + '/messages'
  else
    Endpoint := EndpointMessages;

  var QueryParams := '$select=' + SelectFields +
    '&$orderby=receivedDateTime desc' +
    '&$top=' + Top.ToString;

  const HasSkip = (Skip > 0);
  if HasSkip then
    QueryParams := QueryParams + '&$skip=' + Skip.ToString;

  var Response := FGraphClient.Get(Endpoint, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result.Messages, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result.Messages[Index] := ParseMessage(TGraphJson.ArrayItem(ValueArray, Index));

    Result.HasMore := (Length(Result.Messages) >= Top) or TGraphJson.HasNextPage(Response);
  finally
    Response.Free;
  end;
end;

end.
