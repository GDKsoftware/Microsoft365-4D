unit MSGraph.Graph.Contacts.Interfaces;

interface

uses
  MSGraph.Graph.Contacts.Types;

type
  IContactsClient = interface
    ['{D4E5F6A7-B8C9-4D5E-1F2A-3B4C5D6E7F8A}']
    function SearchContacts(const Query: string; const Top: Integer = 50): TArray<TContact>;
    function GetContact(const ContactId: string): TContact;
    function CreateContact(const GivenName: string; const Surname: string;
      const Email: string; const Phone: string; const Company: string;
      const JobTitle: string): TCreateContactResult;
    function UpdateContact(const ContactId: string; const GivenName: string;
      const Surname: string; const Email: string; const Phone: string;
      const Company: string; const JobTitle: string): TCreateContactResult;
    function DeleteContact(const ContactId: string): Boolean;
  end;

implementation

end.
