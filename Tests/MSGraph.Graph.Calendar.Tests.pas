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
  Assert.AreEqual('PATCH', Patched.Method);

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
