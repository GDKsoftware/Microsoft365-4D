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

    const
      SensitivityKey = 'sensitivity';

    function DateTimeToISO8601(const Value: TDateTime): string;
    function BuildAttendeesArray(const Attendees: TArray<string>): TJSONArray;
    function BuildEventBody(const Subject: string; const StartDateTime: TDateTime;
      const EndDateTime: TDateTime; const Location: string; const Body: string;
      const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string;
      const Sensitivity: TEventSensitivity): TJSONObject;

    function EndpointCalendarView: string;
    function EndpointEvents: string;
    function EndpointGetSchedule: string;

    class function ParseAttendee(const AttendeeObj: TJSONObject): TAttendee; static;
    class function ParseAttendees(const EventObj: TJSONObject): TArray<TAttendee>; static;
    class function ParseEvent(const EventObj: TJSONObject): TCalendarEvent; static;
    class function ParseScheduleItem(const ItemObj: TJSONObject): TScheduleItemEntry; static;
    class function ParseScheduleResult(const ScheduleObj: TJSONObject): TScheduleResult; static;

    // PreferTokens are tokens that go in the HTTP "Prefer:" header
    // (e.g. 'outlook.timezone="Europe/Amsterdam"', 'outlook.body-content-type="text"').
    // The helper joins them comma-separated into a single Prefer header line.
    function FetchCalendarPage(const QueryParams: string; const PreferTokens: TArray<string>;
      out NextLink: string): TArray<TCalendarEvent>;
    function FetchCalendarPageFromUrl(const FullUrl: string; const PreferTokens: TArray<string>;
      out NextLink: string): TArray<TCalendarEvent>;
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

    function ListEvents(const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const Top: Integer = 50; const Timezone: string = ''): TArray<TCalendarEvent>;
    function SearchEvents(const Query: string;
      const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const Timezone: string = ''; const BodyAsText: Boolean = True): TArray<TCalendarEvent>;
    function GetEvent(const EventId: string): TCalendarEvent;
    function CreateEvent(const Subject: string; const StartDateTime: TDateTime;
      const EndDateTime: TDateTime; const Location: string; const Body: string;
      const Attendees: TArray<string>; const IsAllDay: Boolean;
      const TimeZone: string = DefaultCalendarTimeZone;
      const Sensitivity: TEventSensitivity = TEventSensitivity.Normal): TCreateEventResult;
    function UpdateEvent(const EventId: string; const Subject: string;
      const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const Location: string; const Body: string; const Attendees: TArray<string>;
      const IsAllDay: Boolean; const TimeZone: string = DefaultCalendarTimeZone;
      const Sensitivity: TEventSensitivity = TEventSensitivity.Normal): TCreateEventResult;
    function DeleteEvent(const EventId: string): Boolean;
    function GetScheduleAvailability(const Schedules: TArray<string>;
      const StartDateTime: TDateTime; const EndDateTime: TDateTime;
      const TimeZone: string = DefaultCalendarTimeZone): TArray<TScheduleResult>;

    function AcceptEvent(const EventId: string; const Comment: string = ''; const SendResponse: Boolean = True): Boolean;
    function DeclineEvent(const EventId: string; const Comment: string = ''; const SendResponse: Boolean = True): Boolean;
    function TentativelyAcceptEvent(const EventId: string; const Comment: string = ''; const SendResponse: Boolean = True): Boolean;
    function ProposeNewTime(const EventId: string; const NewTime: TProposedNewTime; const Comment: string = ''; const SendResponse: Boolean = True): Boolean;

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.SysUtils,
  System.NetEncoding,
  System.Generics.Collections,
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
  const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string;
  const Sensitivity: TEventSensitivity): TJSONObject;
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

  const HasNonDefaultSensitivity = (Sensitivity <> TEventSensitivity.Normal);
  const SensitivityValue = Sensitivity.GraphValue;
  if HasNonDefaultSensitivity then
    Result.AddPair(SensitivityKey, SensitivityValue);

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

  const SensitivityValue = TGraphJson.GetString(EventObj, SensitivityKey);
  Result.Sensitivity := TEventSensitivity.FromGraphValue(SensitivityValue);

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

function BuildPreferHeaders(const PreferTokens: TArray<string>): TArray<string>;
begin
  Result := nil;
  if Length(PreferTokens) = 0 then
    Exit;

  var Combined := '';
  for var Token in PreferTokens do
  begin
    if Token.Trim.IsEmpty then
      Continue;
    if not Combined.IsEmpty then
      Combined := Combined + ', ';
    Combined := Combined + Token;
  end;

  if Combined.IsEmpty then
    Exit;

  Result := TArray<string>.Create('Prefer: ' + Combined);
end;

function TCalendarClient.FetchCalendarPage(const QueryParams: string;
  const PreferTokens: TArray<string>; out NextLink: string): TArray<TCalendarEvent>;
begin
  Result := nil;
  NextLink := '';

  const Headers = BuildPreferHeaders(PreferTokens);

  var Response: TJSONObject;
  if Length(Headers) > 0 then
    Response := FGraphClient.GetWithHeaders(EndpointCalendarView, QueryParams, Headers)
  else
    Response := FGraphClient.Get(EndpointCalendarView, QueryParams);

  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if Assigned(ValueArray) then
    begin
      SetLength(Result, ValueArray.Count);
      for var Index := 0 to ValueArray.Count - 1 do
        Result[Index] := ParseEvent(TGraphJson.ArrayItem(ValueArray, Index));
    end;

    var NextLinkValue := Response.GetValue('@odata.nextLink');
    if Assigned(NextLinkValue) then
      NextLink := NextLinkValue.Value;
  finally
    Response.Free;
  end;
end;

function TCalendarClient.FetchCalendarPageFromUrl(const FullUrl: string;
  const PreferTokens: TArray<string>; out NextLink: string): TArray<TCalendarEvent>;
begin
  Result := nil;
  NextLink := '';

  const Headers = BuildPreferHeaders(PreferTokens);

  var Response: TJSONObject;
  if Length(Headers) > 0 then
    Response := FGraphClient.GetAbsoluteUrlWithHeaders(FullUrl, Headers)
  else
    Response := FGraphClient.GetAbsoluteUrl(FullUrl);

  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if Assigned(ValueArray) then
    begin
      SetLength(Result, ValueArray.Count);
      for var Index := 0 to ValueArray.Count - 1 do
        Result[Index] := ParseEvent(TGraphJson.ArrayItem(ValueArray, Index));
    end;

    var NextLinkValue := Response.GetValue('@odata.nextLink');
    if Assigned(NextLinkValue) then
      NextLink := NextLinkValue.Value;
  finally
    Response.Free;
  end;
end;

function TCalendarClient.ListEvents(const StartDateTime: TDateTime;
  const EndDateTime: TDateTime; const Top: Integer; const Timezone: string): TArray<TCalendarEvent>;
begin
  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 50
  else if ActualTop > 100 then
    ActualTop := 100;

  var StartISO := DateTimeToISO8601(StartDateTime);
  var EndISO := DateTimeToISO8601(EndDateTime);

  var QueryParams := Format(
    'startDateTime=%s&endDateTime=%s&$top=%d&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,webLink,bodyPreview,sensitivity',
    [TNetEncoding.URL.Encode(StartISO), TNetEncoding.URL.Encode(EndISO), ActualTop]);

  var PreferTokens: TArray<string> := nil;
  if not Timezone.Trim.IsEmpty then
    PreferTokens := TArray<string>.Create('outlook.timezone="' + Timezone + '"');

  var IgnoredNextLink := '';
  Result := FetchCalendarPage(QueryParams, PreferTokens, IgnoredNextLink);
end;

function TCalendarClient.SearchEvents(const Query: string;
  const StartDateTime: TDateTime; const EndDateTime: TDateTime;
  const Timezone: string; const BodyAsText: Boolean): TArray<TCalendarEvent>;
const
  GraphMaxPageSize = 1000;
begin
  Result := nil;
  if Query.Trim.IsEmpty then
    Exit;

  var StartISO := DateTimeToISO8601(StartDateTime);
  var EndISO := DateTimeToISO8601(EndDateTime);

  var QueryParams := Format(
    'startDateTime=%s&endDateTime=%s&$top=%d&$orderby=start/dateTime&$select=id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,webLink,body,bodyPreview,sensitivity',
    [TNetEncoding.URL.Encode(StartISO), TNetEncoding.URL.Encode(EndISO), GraphMaxPageSize]);

  var PreferTokens := TList<string>.Create;
  try
    if not Timezone.Trim.IsEmpty then
      PreferTokens.Add('outlook.timezone="' + Timezone + '"');
    if BodyAsText then
      PreferTokens.Add('outlook.body-content-type="text"');

    const PreferArray = PreferTokens.ToArray;
    const QueryLower = Query.ToLower;
    var Matches := TList<TCalendarEvent>.Create;
    try
      var NextLink := '';
      var Page := FetchCalendarPage(QueryParams, PreferArray, NextLink);
      while True do
      begin
        for var Event in Page do
        begin
          const SubjectHit = Event.Subject.ToLower.Contains(QueryLower);
          const BodyHit = Event.Body.ToLower.Contains(QueryLower);
          if SubjectHit or BodyHit then
            Matches.Add(Event);
        end;

        if NextLink.IsEmpty then
          Break;

        Page := FetchCalendarPageFromUrl(NextLink, PreferArray, NextLink);
      end;

      Result := Matches.ToArray;
    finally
      Matches.Free;
    end;
  finally
    PreferTokens.Free;
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
  const Attendees: TArray<string>; const IsAllDay: Boolean; const TimeZone: string;
  const Sensitivity: TEventSensitivity): TCreateEventResult;
begin
  Result := Default(TCreateEventResult);
  var EventObj := BuildEventBody(Subject, StartDateTime, EndDateTime, Location, Body, Attendees, IsAllDay, TimeZone,
                                 Sensitivity);
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
  const IsAllDay: Boolean; const TimeZone: string;
  const Sensitivity: TEventSensitivity): TCreateEventResult;
begin
  Result := Default(TCreateEventResult);
  var EventObj := BuildEventBody(Subject, StartDateTime, EndDateTime, Location, Body, Attendees, IsAllDay, TimeZone,
                                 Sensitivity);
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

function TCalendarClient.AcceptEvent(const EventId: string; const Comment: string; const SendResponse: Boolean): Boolean;
begin
  var RequestObj := TJSONObject.Create;
  try
    RequestObj.AddPair('comment', Comment);
    RequestObj.AddPair('sendResponse', TJSONBool.Create(SendResponse));

    var Response := FGraphClient.Post(EndpointEvents + '/' + EventId + '/accept', RequestObj.ToJSON);
    try
      Result := not TGraphJson.HasError(Response);
      if not Result then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));
    finally
      Response.Free;
    end;
  finally
    RequestObj.Free;
  end;
end;

function TCalendarClient.DeclineEvent(const EventId: string; const Comment: string; const SendResponse: Boolean): Boolean;
begin
  var RequestObj := TJSONObject.Create;
  try
    RequestObj.AddPair('comment', Comment);
    RequestObj.AddPair('sendResponse', TJSONBool.Create(SendResponse));

    var Response := FGraphClient.Post(EndpointEvents + '/' + EventId + '/decline', RequestObj.ToJSON);
    try
      Result := not TGraphJson.HasError(Response);
      if not Result then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));
    finally
      Response.Free;
    end;
  finally
    RequestObj.Free;
  end;
end;

function TCalendarClient.TentativelyAcceptEvent(const EventId: string; const Comment: string; const SendResponse: Boolean): Boolean;
begin
  var RequestObj := TJSONObject.Create;
  try
    RequestObj.AddPair('comment', Comment);
    RequestObj.AddPair('sendResponse', TJSONBool.Create(SendResponse));

    var Response := FGraphClient.Post(EndpointEvents + '/' + EventId + '/tentativelyAccept', RequestObj.ToJSON);
    try
      Result := not TGraphJson.HasError(Response);
      if not Result then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));
    finally
      Response.Free;
    end;
  finally
    RequestObj.Free;
  end;
end;

function TCalendarClient.ProposeNewTime(const EventId: string; const NewTime: TProposedNewTime; const Comment: string; const SendResponse: Boolean): Boolean;
begin
  var RequestObj := TJSONObject.Create;
  try
    RequestObj.AddPair('comment', Comment);
    RequestObj.AddPair('sendResponse', TJSONBool.Create(SendResponse));

    var ProposedObj := TJSONObject.Create;
    var StartObj := TJSONObject.Create;
    StartObj.AddPair('dateTime', NewTime.StartDateTime);
    StartObj.AddPair('timeZone', NewTime.TimeZone);
    ProposedObj.AddPair('start', StartObj);
    var EndObj := TJSONObject.Create;
    EndObj.AddPair('dateTime', NewTime.EndDateTime);
    EndObj.AddPair('timeZone', NewTime.TimeZone);
    ProposedObj.AddPair('end', EndObj);
    RequestObj.AddPair('proposedNewTime', ProposedObj);

    var Response := FGraphClient.Post(EndpointEvents + '/' + EventId + '/tentativelyAccept', RequestObj.ToJSON);
    try
      Result := not TGraphJson.HasError(Response);
      if not Result then
        raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));
    finally
      Response.Free;
    end;
  finally
    RequestObj.Free;
  end;
end;

end.
