unit MSGraph.Graph.Calendar.Types;

interface

type
  {$SCOPEDENUMS ON}
  TEventSensitivity = (Normal, Personal, Private, Confidential);
  {$SCOPEDENUMS OFF}

  TEventSensitivityHelper = record helper for TEventSensitivity
    class function FromGraphValue(const Value: string): TEventSensitivity; static;
    function GraphValue: string;
  end;

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
    Sensitivity: TEventSensitivity;
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

  {$SCOPEDENUMS ON}
  TEventResponseType = (None, Organizer, TentativelyAccepted, Accepted, Declined, NotResponded);
  {$SCOPEDENUMS OFF}

  TProposedNewTime = record
    StartDateTime: string;
    EndDateTime: string;
    TimeZone: string;
  end;

implementation

uses
  System.SysUtils;

const
  SensitivityNormal = 'normal';
  SensitivityPersonal = 'personal';
  SensitivityPrivate = 'private';
  SensitivityConfidential = 'confidential';
  UnsupportedSensitivity = 'Unsupported event sensitivity: %d';

class function TEventSensitivityHelper.FromGraphValue(const Value: string): TEventSensitivity;
begin
  for var Candidate := Low(TEventSensitivity) to High(TEventSensitivity) do
  begin
    const IsMatch = SameText(Candidate.GraphValue, Value);
    if IsMatch then
      Exit(Candidate);
  end;

  Result := TEventSensitivity.Normal;
end;

function TEventSensitivityHelper.GraphValue: string;
begin
  case Self of
    TEventSensitivity.Normal       : Result := SensitivityNormal;
    TEventSensitivity.Personal     : Result := SensitivityPersonal;
    TEventSensitivity.Private      : Result := SensitivityPrivate;
    TEventSensitivity.Confidential : Result := SensitivityConfidential;
  else
    raise ENotSupportedException.CreateFmt(UnsupportedSensitivity, [Ord(Self)]);
  end;
end;

end.
