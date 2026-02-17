unit MSGraph.Graph.Calendar.Interfaces;

interface

uses
  MSGraph.Graph.Calendar.Types;

const
  DefaultCalendarTimeZone = 'Europe/Amsterdam';

type
  ICalendarClient = interface
    ['{C3D4E5F6-A7B8-4C5D-0E1F-2A3B4C5D6E7F}']
    function ListEvents(const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const Top: Integer = 50; const Timezone: string = ''): TArray<TCalendarEvent>;
    function GetEvent(const EventId: string): TCalendarEvent;
    function CreateEvent(const Subject: string; const StartDateTime: TDateTime;
      const EndDateTime: TDateTime; const Location: string; const Body: string;
      const Attendees: TArray<string>; const IsAllDay: Boolean;
      const TimeZone: string = DefaultCalendarTimeZone): TCreateEventResult;
    function UpdateEvent(const EventId: string; const Subject: string;
      const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const Location: string; const Body: string; const Attendees: TArray<string>;
      const IsAllDay: Boolean; const TimeZone: string = DefaultCalendarTimeZone): TCreateEventResult;
    function DeleteEvent(const EventId: string): Boolean;
    function GetScheduleAvailability(const Schedules: TArray<string>;
      const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const TimeZone: string = DefaultCalendarTimeZone): TArray<TScheduleResult>;
  end;

implementation

end.
