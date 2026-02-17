unit MSGraph.Graph.Calendar;

interface

uses
  System.JSON,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http,
  MSGraph.Graph.Calendar.Types,
  MSGraph.Graph.Calendar.Interfaces;

type
  TCalendarClient = class(TInterfacedObject, ICalendarClient)
  strict private
    FGraphClient: TGraphHttpClient;
    FOwnsClient: Boolean;

    function DateTimeToISO8601(const Value: TDateTime): string;
    function BuildAttendeesArray(const Attendees: TArray<string>): TJSONArray;
    function BuildEventBody(const Subject: string; const StartDateTime: TDateTime;
      const EndDateTime: TDateTime; const Location: string; const Body: string;
      const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string): TJSONObject;

    function EndpointCalendarView: string;
    function EndpointEvents: string;
    function EndpointGetSchedule: string;

    class function ParseAttendee(const AttendeeObj: TJSONObject): TAttendee; static;
    class function ParseAttendees(const EventObj: TJSONObject): TArray<TAttendee>; static;
    class function ParseEvent(const EventObj: TJSONObject): TCalendarEvent; static;
    class function ParseScheduleItem(const ItemObj: TJSONObject): TScheduleItemEntry; static;
    class function ParseScheduleResult(const ScheduleObj: TJSONObject): TScheduleResult; static;
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

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

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.SysUtils,
  System.NetEncoding,
  MSGraph.Graph.JsonHelper;

constructor TCalendarClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FGraphClient := TGraphHttpClient.Create(AccessToken, LogProc);
  FOwnsClient := True;
end;

constructor TCalendarClient.Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean);
begin
  inherited Create;
  FGraphClient := GraphClient;
  FOwnsClient := OwnsClient;
end;

destructor TCalendarClient.Destroy;
begin
  if FOwnsClient then
    FGraphClient.Free;
  inherited;
end;

function TCalendarClient.EndpointCalendarView: string;
begin
  Result := FGraphClient.GetUserPrefix + '/calendarView';
end;

function TCalendarClient.EndpointEvents: string;
begin
  Result := FGraphClient.GetUserPrefix + '/events';
end;

function TCalendarClient.EndpointGetSchedule: string;
begin
  Result := FGraphClient.GetUserPrefix + '/calendar/getSchedule';
end;

function TCalendarClient.DateTimeToISO8601(const Value: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-mm-dd"T"HH:nn:ss', Value);
end;

function TCalendarClient.BuildAttendeesArray(const Attendees: TArray<string>): TJSONArray;
begin
  Result := TJSONArray.Create;
  for var Attendee in Attendees do
  begin
    if Attendee.Trim.IsEmpty then
      Continue;

    var AttendeeObj := TJSONObject.Create;
    var EmailObj := TJSONObject.Create;
    EmailObj.AddPair('address', Attendee.Trim);
    AttendeeObj.AddPair('emailAddress', EmailObj);
    AttendeeObj.AddPair('type', 'required');
    Result.Add(AttendeeObj);
  end;
end;

function TCalendarClient.BuildEventBody(const Subject: string; const StartDateTime: TDateTime;
  const EndDateTime: TDateTime; const Location: string; const Body: string;
  const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('subject', Subject);

  var StartObj := TJSONObject.Create;
  StartObj.AddPair('dateTime', DateTimeToISO8601(StartDateTime));
  StartObj.AddPair('timeZone', TimeZone);
  Result.AddPair('start', StartObj);

  var EndObj := TJSONObject.Create;
  EndObj.AddPair('dateTime', DateTimeToISO8601(EndDateTime));
  EndObj.AddPair('timeZone', TimeZone);
  Result.AddPair('end', EndObj);

  Result.AddPair('isAllDay', TJSONBool.Create(IsAllDay));

  if not Location.Trim.IsEmpty then
  begin
    var LocationObj := TJSONObject.Create;
    LocationObj.AddPair('displayName', Location);
    Result.AddPair('location', LocationObj);
  end;

  if not Body.Trim.IsEmpty then
  begin
    var BodyObj := TJSONObject.Create;
    BodyObj.AddPair('contentType', 'HTML');
    BodyObj.AddPair('content', Body);
    Result.AddPair('body', BodyObj);
  end;

  if Length(Attendees) > 0 then
    Result.AddPair('attendees', BuildAttendeesArray(Attendees));
end;

class function TCalendarClient.ParseAttendee(const AttendeeObj: TJSONObject): TAttendee;
begin
  Result := Default(TAttendee);
  if not Assigned(AttendeeObj) then
    Exit;

  var EmailAddr := TGraphJson.GetObject(AttendeeObj, 'emailAddress');
  if Assigned(EmailAddr) then
  begin
    Result.Name := TGraphJson.GetString(EmailAddr, 'name');
    Result.Email := TGraphJson.GetString(EmailAddr, 'address');
  end;

  var StatusObj := TGraphJson.GetObject(AttendeeObj, 'status');
  if Assigned(StatusObj) then
    Result.Response := TGraphJson.GetString(StatusObj, 'response');
end;

class function TCalendarClient.ParseAttendees(const EventObj: TJSONObject): TArray<TAttendee>;
begin
  Result := nil;
  var Arr := TGraphJson.GetArray(EventObj, 'attendees');
  if not Assigned(Arr) then
    Exit;
  SetLength(Result, Arr.Count);
  for var Index := 0 to Arr.Count - 1 do
    Result[Index] := ParseAttendee(TGraphJson.ArrayItem(Arr, Index));
end;

class function TCalendarClient.ParseEvent(const EventObj: TJSONObject): TCalendarEvent;
begin
  Result := Default(TCalendarEvent);
  if not Assigned(EventObj) then
    Exit;

  Result.Id := TGraphJson.GetString(EventObj, 'id');
  Result.Subject := TGraphJson.GetString(EventObj, 'subject');
  Result.IsAllDay := TGraphJson.GetBool(EventObj, 'isAllDay');
  Result.IsCancelled := TGraphJson.GetBool(EventObj, 'isCancelled');
  Result.WebLink := TGraphJson.GetString(EventObj, 'webLink');
  Result.BodyPreview := TGraphJson.GetString(EventObj, 'bodyPreview');
  Result.ShowAs := TGraphJson.GetString(EventObj, 'showAs');

  var StartObj := TGraphJson.GetObject(EventObj, 'start');
  if Assigned(StartObj) then
    Result.StartDateTime := TGraphJson.GetString(StartObj, 'dateTime');

  var EndObj := TGraphJson.GetObject(EventObj, 'end');
  if Assigned(EndObj) then
    Result.EndDateTime := TGraphJson.GetString(EndObj, 'dateTime');

  var LocationObj := TGraphJson.GetObject(EventObj, 'location');
  if Assigned(LocationObj) then
    Result.Location := TGraphJson.GetString(LocationObj, 'displayName');

  var OrganizerObj := TGraphJson.GetObject(EventObj, 'organizer');
  if Assigned(OrganizerObj) then
  begin
    var EmailAddr := TGraphJson.GetObject(OrganizerObj, 'emailAddress');
    if Assigned(EmailAddr) then
      Result.Organizer := TGraphJson.GetString(EmailAddr, 'name');
  end;

  var BodyObj := TGraphJson.GetObject(EventObj, 'body');
  if Assigned(BodyObj) then
    Result.Body := TGraphJson.GetString(BodyObj, 'content');

  Result.Attendees := ParseAttendees(EventObj);
end;

class function TCalendarClient.ParseScheduleItem(const ItemObj: TJSONObject): TScheduleItemEntry;
begin
  Result := Default(TScheduleItemEntry);
  if not Assigned(ItemObj) then
    Exit;
  Result.Status := TGraphJson.GetString(ItemObj, 'status');
  Result.Subject := TGraphJson.GetString(ItemObj, 'subject');

  var StartObj := TGraphJson.GetObject(ItemObj, 'start');
  if Assigned(StartObj) then
    Result.StartDateTime := TGraphJson.GetString(StartObj, 'dateTime');

  var EndObj := TGraphJson.GetObject(ItemObj, 'end');
  if Assigned(EndObj) then
    Result.EndDateTime := TGraphJson.GetString(EndObj, 'dateTime');
end;

class function TCalendarClient.ParseScheduleResult(const ScheduleObj: TJSONObject): TScheduleResult;
begin
  Result := Default(TScheduleResult);
  if not Assigned(ScheduleObj) then
    Exit;
  Result.Email := TGraphJson.GetString(ScheduleObj, 'scheduleId');
  Result.AvailabilityView := TGraphJson.GetString(ScheduleObj, 'availabilityView');

  var ItemsArr := TGraphJson.GetArray(ScheduleObj, 'scheduleItems');
  if not Assigned(ItemsArr) then
    Exit;
  SetLength(Result.Items, ItemsArr.Count);
  for var Index := 0 to ItemsArr.Count - 1 do
    Result.Items[Index] := ParseScheduleItem(TGraphJson.ArrayItem(ItemsArr, Index));
end;

function TCalendarClient.ListEvents(const StartDateTime: TDateTime;
  const EndDateTime: TDateTime; const Top: Integer; const Timezone: string): TArray<TCalendarEvent>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 50
  else if ActualTop > 100 then
    ActualTop := 100;

  var StartISO := DateTimeToISO8601(StartDateTime);
  var EndISO := DateTimeToISO8601(EndDateTime);

  var QueryParams := Format(
    'startDateTime=%s&endDateTime=%s&$top=%d&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,webLink,bodyPreview',
    [TNetEncoding.URL.Encode(StartISO), TNetEncoding.URL.Encode(EndISO), ActualTop]);

  var Response: TJSONObject;
  if not Timezone.Trim.IsEmpty then
    Response := FGraphClient.GetWithHeaders(EndpointCalendarView, QueryParams,
      TArray<string>.Create('Prefer: outlook.timezone="' + Timezone + '"'))
  else
    Response := FGraphClient.Get(EndpointCalendarView, QueryParams);

  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseEvent(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TCalendarClient.GetEvent(const EventId: string): TCalendarEvent;
begin
  var Response := FGraphClient.Get(EndpointEvents + '/' + EventId,
    '$select=id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,webLink,body,bodyPreview,recurrence,sensitivity,showAs');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseEvent(Response);
  finally
    Response.Free;
  end;
end;

function TCalendarClient.CreateEvent(const Subject: string; const StartDateTime: TDateTime;
  const EndDateTime: TDateTime; const Location: string; const Body: string;
  const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string): TCreateEventResult;
begin
  Result := Default(TCreateEventResult);
  var EventObj := BuildEventBody(Subject, StartDateTime, EndDateTime, Location, Body, Attendees, IsAllDay, TimeZone);
  try
    var Response := FGraphClient.Post(EndpointEvents, EventObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.Id := TGraphJson.GetString(Response, 'id');
      Result.WebLink := TGraphJson.GetString(Response, 'webLink');
    finally
      Response.Free;
    end;
  finally
    EventObj.Free;
  end;
end;

function TCalendarClient.UpdateEvent(const EventId: string; const Subject: string;
  const StartDateTime: TDateTime; const EndDateTime: TDateTime;
  const Location: string; const Body: string; const Attendees: TArray<string>;
  const IsAllDay: Boolean; const TimeZone: string): TCreateEventResult;
begin
  Result := Default(TCreateEventResult);
  var EventObj := BuildEventBody(Subject, StartDateTime, EndDateTime, Location, Body, Attendees, IsAllDay, TimeZone);
  try
    var Response := FGraphClient.Patch(EndpointEvents + '/' + EventId, EventObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      Result.Id := TGraphJson.GetString(Response, 'id');
      Result.WebLink := TGraphJson.GetString(Response, 'webLink');
    finally
      Response.Free;
    end;
  finally
    EventObj.Free;
  end;
end;

function TCalendarClient.DeleteEvent(const EventId: string): Boolean;
begin
  var Response := FGraphClient.Delete(EndpointEvents + '/' + EventId);
  try
    Result := not TGraphJson.HasError(Response);
  finally
    Response.Free;
  end;
end;

function TCalendarClient.GetScheduleAvailability(const Schedules: TArray<string>;
  const StartDateTime: TDateTime; const EndDateTime: TDateTime;
  const TimeZone: string): TArray<TScheduleResult>;
begin
  Result := nil;
  var RequestObj := TJSONObject.Create;
  try
    var SchedulesArray := TJSONArray.Create;
    for var Schedule in Schedules do
      SchedulesArray.Add(Schedule);
    RequestObj.AddPair('schedules', SchedulesArray);

    var StartObj := TJSONObject.Create;
    StartObj.AddPair('dateTime', DateTimeToISO8601(StartDateTime));
    StartObj.AddPair('timeZone', TimeZone);
    RequestObj.AddPair('startTime', StartObj);

    var EndObj := TJSONObject.Create;
    EndObj.AddPair('dateTime', DateTimeToISO8601(EndDateTime));
    EndObj.AddPair('timeZone', TimeZone);
    RequestObj.AddPair('endTime', EndObj);

    RequestObj.AddPair('availabilityViewInterval', TJSONNumber.Create(30));

    var Response := FGraphClient.Post(EndpointGetSchedule, RequestObj.ToJSON);
    try
      if TGraphJson.HasError(Response) then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

      var ValueArray := TGraphJson.GetArray(Response, 'value');
      if not Assigned(ValueArray) then
        Exit;

      SetLength(Result, ValueArray.Count);
      for var Index := 0 to ValueArray.Count - 1 do
        Result[Index] := ParseScheduleResult(TGraphJson.ArrayItem(ValueArray, Index));
    finally
      Response.Free;
    end;
  finally
    RequestObj.Free;
  end;
end;

end.
