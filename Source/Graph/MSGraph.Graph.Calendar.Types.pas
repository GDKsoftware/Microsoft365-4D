unit MSGraph.Graph.Calendar.Types;

interface

type
  TAttendee = record
    Name: string;
    Email: string;
    Response: string;
  end;

  TCalendarEvent = record
    Id: string;
    Subject: string;
    StartDateTime: string;
    EndDateTime: string;
    Location: string;
    Organizer: string;
    Attendees: TArray<TAttendee>;
    IsAllDay: Boolean;
    IsCancelled: Boolean;
    Body: string;
    BodyPreview: string;
    WebLink: string;
    ShowAs: string;
  end;

  TScheduleItemEntry = record
    Status: string;
    Subject: string;
    StartDateTime: string;
    EndDateTime: string;
  end;

  TScheduleResult = record
    Email: string;
    AvailabilityView: string;
    Items: TArray<TScheduleItemEntry>;
  end;

  TCreateEventResult = record
    Id: string;
    WebLink: string;
  end;

implementation

end.
