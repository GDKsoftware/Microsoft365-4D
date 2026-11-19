unit MSGraph.Graph.Calendar.Tests;

interface

uses
  System.JSON,
  DUnitX.TestFramework,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Transport.Fake,
  MSGraph.Graph.Calendar.Interfaces,
  MSGraph.Graph.Calendar.Types;

const
  NormalValue = 'normal';
  PersonalValue = 'personal';
  PrivateValue = 'private';
  ConfidentialValue = 'confidential';

type
  [TestFixture]
  TEventSensitivityTests = class
  public
    [Test]
    [TestCase('Normal', NormalValue)]
    [TestCase('Personal', PersonalValue)]
    [TestCase('Private', PrivateValue)]
    [TestCase('Confidential', ConfidentialValue)]
    procedure FromGraphValue_KnownValue_RoundTripsThroughGraphValue(const Value: string);
    [Test]
    procedure FromGraphValue_MixedCase_IsAccepted;
    [Test]
    procedure FromGraphValue_UnknownValue_FallsBackToNormal;
  end;

  [TestFixture]
  TCalendarClientTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FGraphClient: TGraphHttpClient;
    FCalendarClient: ICalendarClient;
    FEventStart: TDateTime;
    FEventEnd: TDateTime;

    function CreateEventJson(const Sensitivity: TEventSensitivity): TJSONObject;
    function PostedEventJson: TJSONObject;
    procedure AssertSensitivityOmitted(const Body: TJSONObject);
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    procedure CreateEvent_PrivateSensitivity_PostsSensitivity;
    [Test]
    procedure CreateEvent_NormalSensitivity_OmitsSensitivity;
    [Test]
    procedure CreateEvent_WithoutSensitivityArgument_OmitsSensitivity;
    [Test]
    procedure UpdateEvent_ConfidentialSensitivity_PatchesSensitivity;
    [Test]
    procedure GetEvent_ResponseWithSensitivity_ParsesSensitivity;
    [Test]
    procedure GetEvent_ResponseWithoutSensitivity_DefaultsToNormal;
    [Test]
    procedure ListEvents_SelectsSensitivity;
    [Test]
    procedure CreateEvent_OnlineMeeting_PostsTeamsProvider;
    [Test]
    procedure CreateEvent_WithoutOnlineMeetingArgument_OmitsOnlineMeeting;
    [Test]
    procedure CreateEvent_OnlineMeetingResponse_ReturnsJoinUrl;
    [Test]
    procedure GetEvent_OnlineMeetingResponse_ParsesJoinUrl;
    [Test]
    procedure ChangeEvent_AttendeesOnly_PatchesOnlyAttendees;
    [Test]
    procedure ChangeEvent_NewPeriod_PatchesStartAndEnd;
    [Test]
    procedure ChangeEvent_OnlineMeeting_PatchesTeamsProvider;
    [Test]
    procedure ChangeEvent_NoChanges_RaisesWithoutRequest;
  end;

implementation

uses
  System.SysUtils,
  System.DateUtils,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Calendar;

const
  DummyAccessToken = 'unit-test-token';
  EventSubject = 'Quarterly review';
  EventId = 'AAMkEvent';
  EventCreatedResponse = '{"id":"AAMkEvent","webLink":"https://outlook.office365.com/owa/?itemid=AAMkEvent"}';
  ConfidentialEventResponse = '{"id":"AAMkEvent","subject":"Quarterly review","sensitivity":"confidential"}';
  PlainEventResponse = '{"id":"AAMkEvent","subject":"Quarterly review"}';
  EmptyListResponse = '{"value":[]}';
  SensitivityKey = 'sensitivity';
  UnknownValue = 'secret';
  MixedCaseValue = 'Private';
  SingleRequest = 1;
  FirstRequest = 0;
  NoRequests = 0;
  PatchMethod = 'PATCH';
  JoinUrl = 'https://teams.microsoft.com/l/meetup-join/unit-test';
  OnlineEventResponse = '{"id":"AAMkEvent","isOnlineMeeting":true,"onlineMeeting":{"joinUrl":"' + JoinUrl + '"}}';
  IsOnlineMeetingKey = 'isOnlineMeeting';
  OnlineMeetingProviderKey = 'onlineMeetingProvider';
  TeamsProvider = 'teamsForBusiness';
  AttendeesKey = 'attendees';
  SubjectKey = 'subject';
  StartKey = 'start';
  EndKey = 'end';
  DateTimeKey = 'dateTime';
  AttendeeAddress = 'attendee@example.com';
  EventStartIso = '2026-09-20T09:00:00';
  EventEndIso = '2026-09-20T09:30:00';
  AsOnlineMeeting = True;
  SingleField = 1;
  SingleAttendee = 1;

procedure TEventSensitivityTests.FromGraphValue_KnownValue_RoundTripsThroughGraphValue(const Value: string);
begin
  const Sensitivity = TEventSensitivity.FromGraphValue(Value);

  const Mapped = Sensitivity.GraphValue;

  Assert.AreEqual(Value, Mapped);
end;

procedure TEventSensitivityTests.FromGraphValue_MixedCase_IsAccepted;
begin
  const Sensitivity = TEventSensitivity.FromGraphValue(MixedCaseValue);

  Assert.AreEqual<TEventSensitivity>(TEventSensitivity.Private, Sensitivity);
end;

procedure TEventSensitivityTests.FromGraphValue_UnknownValue_FallsBackToNormal;
begin
  const Sensitivity = TEventSensitivity.FromGraphValue(UnknownValue);

  Assert.AreEqual<TEventSensitivity>(TEventSensitivity.Normal, Sensitivity);
end;

procedure TCalendarClientTests.Setup;
begin
  FFake           := TFakeGraphHttpTransport.Create;
  FGraphClient    := TGraphHttpClient.Create(DummyAccessToken, FFake);
  FCalendarClient := TCalendarClient.Create(FGraphClient, True);
  FEventStart     := EncodeDateTime(2026, 9, 20, 9, 0, 0, 0);
  FEventEnd       := IncMinute(FEventStart, 30);
end;

procedure TCalendarClientTests.TearDown;
begin
  FCalendarClient := nil;
  FGraphClient := nil;
  FFake := nil;
end;

procedure TCalendarClientTests.CreateEvent_PrivateSensitivity_PostsSensitivity;
begin
  const Body = CreateEventJson(TEventSensitivity.Private);
  try
    const PostedSensitivity = TGraphJson.GetString(Body, SensitivityKey);

    Assert.AreEqual(PrivateValue, PostedSensitivity);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.CreateEvent_NormalSensitivity_OmitsSensitivity;
begin
  const Body = CreateEventJson(TEventSensitivity.Normal);
  try
    AssertSensitivityOmitted(Body);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.CreateEvent_WithoutSensitivityArgument_OmitsSensitivity;
begin
  FFake.EnqueueResponse(201, EventCreatedResponse);

  FCalendarClient.CreateEvent(EventSubject, FEventStart, FEventEnd, '', '', [], False);

  const Body = PostedEventJson;
  try
    AssertSensitivityOmitted(Body);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.UpdateEvent_ConfidentialSensitivity_PatchesSensitivity;
begin
  FFake.EnqueueResponse(200, EventCreatedResponse);

  FCalendarClient.UpdateEvent(EventId, EventSubject, FEventStart, FEventEnd, '', '', [], False,
                              DefaultCalendarTimeZone, TEventSensitivity.Confidential);

  const Patched = FFake.RequestAt(FirstRequest);
  Assert.AreEqual(PatchMethod, Patched.Method);

  const Body = PostedEventJson;
  try
    const PostedSensitivity = TGraphJson.GetString(Body, SensitivityKey);

    Assert.AreEqual(ConfidentialValue, PostedSensitivity);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.GetEvent_ResponseWithSensitivity_ParsesSensitivity;
begin
  FFake.EnqueueResponse(200, ConfidentialEventResponse);

  const CalendarEvent = FCalendarClient.GetEvent(EventId);

  Assert.AreEqual<TEventSensitivity>(TEventSensitivity.Confidential, CalendarEvent.Sensitivity);
end;

procedure TCalendarClientTests.GetEvent_ResponseWithoutSensitivity_DefaultsToNormal;
begin
  FFake.EnqueueResponse(200, PlainEventResponse);

  const CalendarEvent = FCalendarClient.GetEvent(EventId);

  Assert.AreEqual<TEventSensitivity>(TEventSensitivity.Normal, CalendarEvent.Sensitivity);
end;

procedure TCalendarClientTests.ListEvents_SelectsSensitivity;
begin
  FFake.EnqueueResponse(200, EmptyListResponse);

  FCalendarClient.ListEvents(FEventStart, FEventEnd);

  const Requested = FFake.RequestAt(FirstRequest);
  const SelectsSensitivity = Requested.Url.Contains(SensitivityKey);
  Assert.IsTrue(SelectsSensitivity, Format('the calendar view must select sensitivity: %s', [Requested.Url]));
end;

procedure TCalendarClientTests.CreateEvent_OnlineMeeting_PostsTeamsProvider;
begin
  FFake.EnqueueResponse(201, OnlineEventResponse);

  FCalendarClient.CreateEvent(EventSubject, FEventStart, FEventEnd, '', '', [], False,
                              DefaultCalendarTimeZone, TEventSensitivity.Normal, AsOnlineMeeting);

  const Body = PostedEventJson;
  try
    const IsOnlineMeeting = TGraphJson.GetBool(Body, IsOnlineMeetingKey);
    const Provider = TGraphJson.GetString(Body, OnlineMeetingProviderKey);

    Assert.IsTrue(IsOnlineMeeting, 'an online meeting must post isOnlineMeeting');
    Assert.AreEqual(TeamsProvider, Provider);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.CreateEvent_WithoutOnlineMeetingArgument_OmitsOnlineMeeting;
begin
  FFake.EnqueueResponse(201, EventCreatedResponse);

  FCalendarClient.CreateEvent(EventSubject, FEventStart, FEventEnd, '', '', [], False);

  const Body = PostedEventJson;
  try
    const HasOnlineMeeting = Assigned(Body.Values[IsOnlineMeetingKey]);
    const HasProvider = Assigned(Body.Values[OnlineMeetingProviderKey]);

    Assert.IsFalse(HasOnlineMeeting, 'a plain event must not post isOnlineMeeting');
    Assert.IsFalse(HasProvider, 'a plain event must not post an online meeting provider');
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.CreateEvent_OnlineMeetingResponse_ReturnsJoinUrl;
begin
  FFake.EnqueueResponse(201, OnlineEventResponse);

  const Created = FCalendarClient.CreateEvent(EventSubject, FEventStart, FEventEnd, '', '', [], False,
                                              DefaultCalendarTimeZone, TEventSensitivity.Normal, AsOnlineMeeting);

  Assert.AreEqual(JoinUrl, Created.OnlineMeetingUrl);
end;

procedure TCalendarClientTests.GetEvent_OnlineMeetingResponse_ParsesJoinUrl;
begin
  FFake.EnqueueResponse(200, OnlineEventResponse);

  const CalendarEvent = FCalendarClient.GetEvent(EventId);

  Assert.IsTrue(CalendarEvent.IsOnlineMeeting, 'the event must be marked as an online meeting');
  Assert.AreEqual(JoinUrl, CalendarEvent.OnlineMeetingUrl);
end;

procedure TCalendarClientTests.ChangeEvent_AttendeesOnly_PatchesOnlyAttendees;
begin
  FFake.EnqueueResponse(200, EventCreatedResponse);
  var Changes := Default(TEventChanges);
  Changes.Attendees := [AttendeeAddress];

  FCalendarClient.ChangeEvent(EventId, Changes);

  const Patched = FFake.RequestAt(FirstRequest);
  Assert.AreEqual(PatchMethod, Patched.Method);

  const Body = PostedEventJson;
  try
    const Attendees = TGraphJson.GetArray(Body, AttendeesKey);

    Assert.AreEqual(SingleField, Body.Count, Format('only the attendees may be patched: %s', [Patched.Body]));
    Assert.IsNotNull(Attendees, 'the patch must carry the attendees');
    Assert.AreEqual(SingleAttendee, Attendees.Count);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.ChangeEvent_NewPeriod_PatchesStartAndEnd;
begin
  FFake.EnqueueResponse(200, EventCreatedResponse);
  var Changes := Default(TEventChanges);
  Changes.StartDateTime := FEventStart;
  Changes.EndDateTime   := FEventEnd;

  FCalendarClient.ChangeEvent(EventId, Changes);

  const Body = PostedEventJson;
  try
    const StartObj = TGraphJson.GetObject(Body, StartKey);
    const EndObj = TGraphJson.GetObject(Body, EndKey);
    const PatchedStart = TGraphJson.GetString(StartObj, DateTimeKey);
    const PatchedEnd = TGraphJson.GetString(EndObj, DateTimeKey);
    const HasSubject = Assigned(Body.Values[SubjectKey]);

    Assert.AreEqual(EventStartIso, PatchedStart);
    Assert.AreEqual(EventEndIso, PatchedEnd);
    Assert.IsFalse(HasSubject, 'an unchanged subject must not be patched');
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.ChangeEvent_OnlineMeeting_PatchesTeamsProvider;
begin
  FFake.EnqueueResponse(200, OnlineEventResponse);
  var Changes := Default(TEventChanges);
  Changes.IsOnlineMeeting := AsOnlineMeeting;

  const Changed = FCalendarClient.ChangeEvent(EventId, Changes);

  const Body = PostedEventJson;
  try
    const Provider = TGraphJson.GetString(Body, OnlineMeetingProviderKey);

    Assert.AreEqual(TeamsProvider, Provider);
    Assert.AreEqual(JoinUrl, Changed.OnlineMeetingUrl);
  finally
    Body.Free;
  end;
end;

procedure TCalendarClientTests.ChangeEvent_NoChanges_RaisesWithoutRequest;
begin
  const NoChanges = Default(TEventChanges);

  Assert.WillRaise(
    procedure
    begin
      FCalendarClient.ChangeEvent(EventId, NoChanges);
    end,
    EArgumentException);

  Assert.AreEqual(NoRequests, FFake.RequestCount, 'an empty change must not reach Graph');
end;

function TCalendarClientTests.CreateEventJson(const Sensitivity: TEventSensitivity): TJSONObject;
begin
  FFake.EnqueueResponse(201, EventCreatedResponse);

  FCalendarClient.CreateEvent(EventSubject, FEventStart, FEventEnd, '', '', [], False,
                              DefaultCalendarTimeZone, Sensitivity);

  Result := PostedEventJson;
end;

function TCalendarClientTests.PostedEventJson: TJSONObject;
begin
  Assert.AreEqual(SingleRequest, FFake.RequestCount, 'the event call must send exactly one request');

  const Posted = FFake.RequestAt(FirstRequest);
  const Parsed = TJSONObject.ParseJSONValue(Posted.Body);
  const IsObject = (Parsed is TJSONObject);
  if IsObject then
  begin
    Result := TJSONObject(Parsed);
  end
  else
  begin
    Parsed.Free;
    Result := nil;
  end;

  Assert.IsNotNull(Result, Format('body is not a JSON object: %s', [Posted.Body]));
end;

procedure TCalendarClientTests.AssertSensitivityOmitted(const Body: TJSONObject);
begin
  const HasSensitivity = Assigned(Body.Values[SensitivityKey]);

  Assert.IsFalse(HasSensitivity, 'a normal event must not carry a sensitivity');
end;

end.
