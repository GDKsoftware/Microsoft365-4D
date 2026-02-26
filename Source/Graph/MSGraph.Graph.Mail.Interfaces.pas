unit MSGraph.Graph.Mail.Interfaces;

interface

uses
  MSGraph.Graph.Mail.Types;

type
  IMailClient = interface
    ['{B2C3D4E5-F6A7-4B5C-9D0E-1F2A3B4C5D6E}']
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
  end;

implementation

end.
