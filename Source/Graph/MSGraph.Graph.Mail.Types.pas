unit MSGraph.Graph.Mail.Types;

interface

type
  TEmailAddress = record
    Name: string;
    Address: string;
  end;

  TMailMessage = record
    Id: string;
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
  end;

  TMailAttachment = record
    Id: string;
    Name: string;
    ContentType: string;
    Size: Int64;
    IsInline: Boolean;
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
  end;

  TMoveMessageResult = record
    NewMessageId: string;
  end;

implementation

end.
