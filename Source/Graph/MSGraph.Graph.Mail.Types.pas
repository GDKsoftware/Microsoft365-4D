unit MSGraph.Graph.Mail.Types;

interface

uses
  MSGraph.OAuth2.Types;

type
  EInvalidMailHeaderException = class(EGraphApiException);
  EInvalidAttachmentException = class(EGraphApiException);
  EDeltaLinkExpiredException = class(EGraphApiException);

{$SCOPEDENUMS ON}
  TMailLastVerb = (ReplyToSender, ReplyToAll, Forwarded);
{$SCOPEDENUMS OFF}

  TMailLastVerbHelper = record helper for TMailLastVerb
  public
    function ToMapiValue: Integer;
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

constructor TMailHeader.Create(const HeaderName: string; const HeaderValue: string);
begin
  Name := HeaderName;
  Value := HeaderValue;
end;

end.
