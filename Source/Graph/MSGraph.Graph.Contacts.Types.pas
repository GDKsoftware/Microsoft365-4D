unit MSGraph.Graph.Contacts.Types;

interface

type
  TPostalAddress = record
    Street: string;
    City: string;
    State: string;
    PostalCode: string;
    Country: string;
  end;

  TContact = record
    Id: string;
    GivenName: string;
    Surname: string;
    DisplayName: string;
    Email: string;
    BusinessPhone: string;
    MobilePhone: string;
    HomePhone: string;
    Company: string;
    JobTitle: string;
    Department: string;
    OfficeLocation: string;
    BusinessAddress: TPostalAddress;
    HomeAddress: TPostalAddress;
    Birthday: string;
    PersonalNotes: string;
  end;

  TCreateContactResult = record
    Id: string;
    DisplayName: string;
  end;

implementation

end.
