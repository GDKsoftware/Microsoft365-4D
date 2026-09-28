unit MSGraph.Graph.Mail.Types;

interface

uses
  MSGraph.OAuth2.Types;

type
  EInvalidMailHeaderException = class(EGraphApiException);
  EInvalidAttachmentException = class(EGraphApiException);
  EAttachmentContentUnavailableException = class(EGraphApiException);
  EInvalidRecipientException = class(EGraphApiException);
  EDeltaLinkExpiredException = class(EGraphApiException);

{$SCOPEDENUMS ON}
  TMailLastVerb = (ReplyToSender, ReplyToAll, Forwarded);
  TMailBodyFormat = (Html, Text);
  TMailAttachmentKind = (&File, Item, Reference, Unknown);
{$SCOPEDENUMS OFF}

  TMailLastVerbHelper = record helper for TMailLastVerb
  public
    function ToMapiValue: Integer;
    function ToIconIndex: Integer;
  end;

  TMailBodyFormatHelper = record helper for TMailBodyFormat
  public
    function ToPreferHeader: string;
  end;

  TMailAttachmentKindHelper = record helper for TMailAttachmentKind
  public
    class function FromODataType(const ODataType: string): TMailAttachmentKind; static;
  end;

  TEmailAddress = record
  public
    Name: string;
    Address: string;
  end;

  TMailHeader = record
  public
    Name: string;
    Value: string;

    constructor Create(const HeaderName: string; const HeaderValue: string);
  end;

  TMailMessage = record
  public
    Id: string;
    ConversationId: string;
    Subject: string;
    From: TEmailAddress;
    ToRecipients: TArray<TEmailAddress>;
    CcRecipients: TArray<TEmailAddress>;
    ReceivedDateTime: string;
    IsRead: Boolean;
    HasAttachments: Boolean;
    Body: string;
    BodyType: string;
    UniqueBody: string;
    BodyPreview: string;
    Importance: string;
    ParentFolderId: string;
    MeetingMessageType: string;
  end;

  TMailAttachment = record
  public
    Id: string;
    Name: string;
    ContentType: string;
    Size: Int64;
    IsInline: Boolean;
    ContentId: string;
    ContentBytes: string;
    Kind: TMailAttachmentKind;
  end;

  TMailFolder = record
  public
    Id: string;
    DisplayName: string;
    ParentFolderId: string;
    ChildFolderCount: Integer;
    TotalItemCount: Integer;
    UnreadItemCount: Integer;
  end;

  TSearchMessagesResult = record
  public
    Messages: TArray<TMailMessage>;
    HasMore: Boolean;
  end;

  TDraftResult = record
  public
    Id: string;
    Subject: string;
  end;

  TMoveMessageResult = record
  public
    NewMessageId: string;
  end;

  TDeltaMessageChange = record
  public
    Message: TMailMessage;
    IsRemoved: Boolean;
  end;

  TDeltaSyncResult = record
  public
    Changes: TArray<TDeltaMessageChange>;
    DeltaLink: string;
  end;

implementation

uses
  System.SysUtils;

const
  MapiVerbReplyToSender = 102;
  MapiVerbReplyToAll    = 103;
  MapiVerbForward       = 104;

  MapiIconReplied   = 261;
  MapiIconForwarded = 262;

  PreferBodyHtml = 'Prefer: outlook.body-content-type="html"';
  PreferBodyText = 'Prefer: outlook.body-content-type="text"';

  ODataTypeFileAttachment      = '#microsoft.graph.fileAttachment';
  ODataTypeItemAttachment      = '#microsoft.graph.itemAttachment';
  ODataTypeReferenceAttachment = '#microsoft.graph.referenceAttachment';

function TMailLastVerbHelper.ToMapiValue: Integer;
begin
  case Self of
    TMailLastVerb.ReplyToSender : Result := MapiVerbReplyToSender;
    TMailLastVerb.ReplyToAll    : Result := MapiVerbReplyToAll;
    TMailLastVerb.Forwarded     : Result := MapiVerbForward;
  else
    raise ENotSupportedException.CreateFmt('Unsupported mail verb: %d', [Ord(Self)]);
  end;
end;

function TMailLastVerbHelper.ToIconIndex: Integer;
begin
  case Self of
    TMailLastVerb.ReplyToSender,
    TMailLastVerb.ReplyToAll : Result := MapiIconReplied;
    TMailLastVerb.Forwarded  : Result := MapiIconForwarded;
  else
    raise ENotSupportedException.CreateFmt('Unsupported mail verb: %d', [Ord(Self)]);
  end;
end;

function TMailBodyFormatHelper.ToPreferHeader: string;
begin
  case Self of
    TMailBodyFormat.Html : Result := PreferBodyHtml;
    TMailBodyFormat.Text : Result := PreferBodyText;
  else
    raise ENotSupportedException.CreateFmt('Unsupported mail body format: %d', [Ord(Self)]);
  end;
end;

class function TMailAttachmentKindHelper.FromODataType(const ODataType: string): TMailAttachmentKind;
begin
  const IsFileAttachment = (ODataType.IsEmpty or SameText(ODataType, ODataTypeFileAttachment));
  if IsFileAttachment then
    Exit(TMailAttachmentKind.&File);

  if SameText(ODataType, ODataTypeItemAttachment) then
    Exit(TMailAttachmentKind.Item);

  if SameText(ODataType, ODataTypeReferenceAttachment) then
    Exit(TMailAttachmentKind.Reference);

  Result := TMailAttachmentKind.Unknown;
end;

constructor TMailHeader.Create(const HeaderName: string; const HeaderValue: string);
begin
  Name := HeaderName;
  Value := HeaderValue;
end;

end.
