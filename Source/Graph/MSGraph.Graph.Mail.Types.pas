unit MSGraph.Graph.Mail.Types;

interface

type
  TEmailAddress = record
    Name: string;
    Address: string;
  end;

  TMailHeader = record
  public
    Name: string;
    Value: string;

    constructor Create(const AName: string; const AValue: string);
  end;

  TMailMessage = record
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
    Id: string;
    Name: string;
    ContentType: string;
    Size: Int64;
    IsInline: Boolean;
    ContentId: string;
    ContentBytes: string;
  end;

  TMailFolder = record
    Id: string;
    DisplayName: string;
    ParentFolderId: string;
    ChildFolderCount: Integer;
    TotalItemCount: Integer;
    UnreadItemCount: Integer;
  end;

  TSearchMessagesResult = record
    Messages: TArray<TMailMessage>;
    HasMore: Boolean;
  end;

  TDraftResult = record
    Id: string;
    Subject: string;
  end;

  TMoveMessageResult = record
    NewMessageId: string;
  end;

  TDeltaMessageChange = record
    Message: TMailMessage;
    IsRemoved: Boolean;
  end;

  TDeltaSyncResult = record
    Changes: TArray<TDeltaMessageChange>;
    DeltaLink: string;
  end;

implementation

constructor TMailHeader.Create(const AName: string; const AValue: string);
begin
  Name := AName;
  Value := AValue;
end;

end.
