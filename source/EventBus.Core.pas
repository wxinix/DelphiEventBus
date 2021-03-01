{*******************************************************************************
Copyright 2016-2020 Daniele Spinetti

Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License.  You  may obtain a copy of the
License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed
under  the  License  is  distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR
CONDITIONS OF ANY KIND, either express or implied.

See the License for the specific language governing permissions and limitations
under the License.
********************************************************************************}

unit EventBus.Core;

interface

uses
  EventBus;

type
  TEventBusFactory = class
  strict private
    class var FGlobalEventBus: IEventBus;
    class constructor Create;
  public
    function CreateEventBus: IEventBus;
    class property GlobalEventBus: IEventBus read FGlobalEventBus;
  end;

implementation

uses
  System.Classes,
  System.Generics.Collections,
  System.Generics.Defaults,
  System.Rtti,
  System.SyncObjs,
  System.SysUtils,
  System.Threading,
  EventBus.Helpers,
  EventBus.Subscribers;

type
  {$REGION 'Type aliases to improve readability'}
  TSubscriptions = TObjectList<TSubscription>;
  TMethodCategory = string;
  TMethodCategories = TList<TMethodCategory>;
  TMethodCategoryToSubscriptionsMap = TObjectDictionary<TMethodCategory, TSubscriptions>;
  TSubscriberToMethodCategoriesMap = TObjectDictionary<TObject, TMethodCategories>;

  TAttributeName = string;
  TMethodCategoryToSubscriptionsByAttributeName = TObjectDictionary<TAttributeName, TMethodCategoryToSubscriptionsMap>;
  TSubscriberToMethodCategoriesByAttributeName = TObjectDictionary<TAttributeName, TSubscriberToMethodCategoriesMap>;
  {$ENDREGION}

  TEventBus = class(TInterfacedObject, IEventBus)
  strict private
    FCategoryToSubscriptionsByAttrName: TMethodCategoryToSubscriptionsByAttributeName;
    FMrewSync: TLightweightMREW;
    FSubscriberToCategoriesByAttrName: TSubscriberToMethodCategoriesByAttributeName;
    procedure DeleteSubscriber<T: TSubscriberMethodAttribute>(ASubscriber: TObject);
    function GetCreateSubscriberCategories<T: TSubscriberMethodAttribute>(const ASubscriber: TObject): TMethodCategories;
    function GetCreateCategorizedSubscriptions<T: TSubscriberMethodAttribute>(const ACategory: string): TSubscriptions;
    procedure InvokeSubscriber(ASubscription: TSubscription; const Args: array of TValue);
    function IsRegistered<T: TSubscriberMethodAttribute>(ASubscriber: TObject): Boolean;
    procedure RegisterSubscriber<T: TSubscriberMethodAttribute>(ASubscriber: TObject; ARaiseExcIfEmpty: Boolean);
    function RemoveSubscription<T: TSubscriberMethodAttribute>(ASubscriber: TObject; const ACategory: string): TSubscription;
    procedure Subscribe<T: TSubscriberMethodAttribute>(ASubscriber: TObject; ASubscriberMethod: TSubscriberMethod);
    function TryGetSubscriberCategories<T: TSubscriberMethodAttribute>(const ASubscriber: TObject; out ACategories: TMethodCategories): Boolean;
    function TryGetCategorizedSubscriptions<T: TSubscriberMethodAttribute>(const ACategory: string; out ASubscriptions: TSubscriptions): Boolean;
    procedure UnregisterSubscriber<T: TSubscriberMethodAttribute>(ASubscriber: TObject);
    procedure Unsubscribe<T: TSubscriberMethodAttribute>(ASubscriber: TObject; const AMethodCategory: TMethodCategory);
    {$REGION'IEventBus interface methods'}
    function IsRegisteredForChannels(ASubscriber: TObject): Boolean;
    function IsRegisteredForEvents(ASubscriber: TObject): Boolean;
    procedure Post(const AChannel: string; const AMessage: string); overload;
    procedure Post(const AEvent: IInterface; const AContext: string = ''); overload;
    procedure RegisterNewContext(ASubscriber: TObject; AEvent: IInterface; const AOldContext: string; const ANewContext: string);
    procedure RegisterSubscriberForChannels(ASubscriber: TObject);
    procedure SilentRegisterSubscriberForChannels(ASubscriber: TObject);
    procedure RegisterSubscriberForEvents(ASubscriber: TObject);
    procedure SilentRegisterSubscriberForEvents(ASubscriber: TObject);
    procedure UnregisterForChannels(ASubscriber: TObject);
    procedure UnregisterForEvents(ASubscriber: TObject);
    {$ENDREGION}
  strict protected
    procedure PostToChannel(ASubscription: TSubscription; const AMessage: string; AIsMainThread: Boolean); virtual;
    procedure PostToSubscription(ASubscription: TSubscription; const AEvent: IInterface; AIsMainThread: Boolean); virtual;
  public
    constructor Create; virtual;
    destructor Destroy; override;
  end;

constructor TEventBus.Create;
begin
  inherited Create;
  FCategoryToSubscriptionsByAttrName := TMethodCategoryToSubscriptionsByAttributeName.Create([doOwnsValues]);
  FSubscriberToCategoriesByAttrName := TSubscriberToMethodCategoriesByAttributeName.Create([doOwnsValues]);
end;

destructor TEventBus.Destroy;
begin
  FCategoryToSubscriptionsByAttrName.Free;
  FSubscriberToCategoriesByAttrName.Free;
  inherited;
end;

procedure TEventBus.DeleteSubscriber<T>(ASubscriber: TObject);
begin
  var LAttrName := T.ClassName;
  var LSubscriberToCategoriesMap: TSubscriberToMethodCategoriesMap;

  if FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
  begin
    LSubscriberToCategoriesMap := FSubscriberToCategoriesByAttrName[LAttrName];

    if LSubscriberToCategoriesMap.ContainsKey(ASubscriber) then
      LSubscriberToCategoriesMap.Remove(ASubscriber);
  end;
end;

function TEventBus.GetCreateSubscriberCategories<T>(const ASubscriber: TObject): TMethodCategories;
begin
  var LAttrName := T.ClassName;
  var LSubsToCatsMap: TSubscriberToMethodCategoriesMap;

  if not FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
  begin
    LSubsToCatsMap := TSubscriberToMethodCategoriesMap.Create([doOwnsValues]);
    FSubscriberToCategoriesByAttrName.Add(LAttrName, LSubsToCatsMap);
  end else
  begin
    LSubsToCatsMap := FSubscriberToCategoriesByAttrName[LAttrName];
  end;

  if (not LSubsToCatsMap.TryGetValue(ASubscriber, Result)) then
  begin
    Result := TMethodCategories.Create;
    LSubsToCatsMap.Add(ASubscriber, Result);
  end;
end;

function TEventBus.GetCreateCategorizedSubscriptions<T>(const ACategory: string): TSubscriptions;
begin
  var LAttrName := T.ClassName;
  var LCatToSubsMap: TMethodCategoryToSubscriptionsMap;

  if not FCategoryToSubscriptionsByAttrName.ContainsKey(LAttrName) then
  begin
    LCatToSubsMap := TMethodCategoryToSubscriptionsMap.Create([doOwnsValues]);
    FCategoryToSubscriptionsByAttrName.Add(LAttrName, LCatToSubsMap);
  end else
  begin
    LCatToSubsMap := FCategoryToSubscriptionsByAttrName[LAttrName];
  end;

  if (not LCatToSubsMap.ContainsKey(ACategory)) then
  begin
    Result := TSubscriptions.Create(
      TComparer<TSubscription>.Construct(
        function(const Left, Right: TSubscription): Integer
        begin
          if Left.Equals(Right) then
            Result := 0
          else
            Result := Left.GetHashCode - Right.GetHashCode;
        end)
      ,
      True // Owns the object for its life cycle.
    );

    LCatToSubsMap.Add(ACategory, Result);
  end else
  begin
    Result := LCatToSubsMap[ACategory];
  end;
end;

procedure TEventBus.InvokeSubscriber(ASubscription: TSubscription; const Args: array of TValue);
begin
  try
    if not ASubscription.Active then
      Exit;

    ASubscription.SubscriberMethod.Method.Invoke(ASubscription.Subscriber, Args);
  except
    on E: Exception do
    begin
      raise EInvokeSubscriberError.CreateFmt(
        'Error invoking subscriber method. Subscriber class: %s. Event type: %s. Original exception %s: %s.',
        [
          ASubscription.Subscriber.ClassName,
          ASubscription.SubscriberMethod.EventType,
          E.ClassName,
          E.Message
        ]);
    end;
  end;
end;

function TEventBus.IsRegistered<T>(ASubscriber: TObject): Boolean;
begin
  FMrewSync.BeginRead;

  try
    var LAttrName := T.ClassName;
    var LSubscriberToCategoriesMap: TSubscriberToMethodCategoriesMap;

    if not FSubscriberToCategoriesByAttrName.TryGetValue(LAttrName, LSubscriberToCategoriesMap) then
      Exit(False);

    Result := LSubscriberToCategoriesMap.ContainsKey(ASubscriber);
  finally
    FMrewSync.EndRead;
  end;
end;

function TEventBus.IsRegisteredForChannels(ASubscriber: TObject): Boolean;
begin
  Result := IsRegistered<ChannelAttribute>(ASubscriber);
end;

function TEventBus.IsRegisteredForEvents(ASubscriber: TObject): Boolean;
begin
  Result := IsRegistered<SubscribeAttribute>(ASubscriber);
end;

procedure TEventBus.Post(const AChannel, AMessage: string);
begin
  FMrewSync.BeginRead;

  try
    var LSubscriptions: TSubscriptions;
    if not TryGetCategorizedSubscriptions<ChannelAttribute>(TSubscriberMethod.EncodeCategory(AChannel), LSubscriptions) then
      Exit;

    for var LSubscription in LSubscriptions do
    begin
      if (LSubscription.Context <> AChannel) or (not LSubscription.Active) then
        Continue;

      var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;
      PostToChannel(LSubscription, AMessage, LIsMainThread);
    end;
  finally
    FMrewSync.EndRead;
  end;
end;

procedure TEventBus.Post(const AEvent: IInterface; const AContext: string = '');
begin
  FMrewSync.BeginRead;

  try
    var LEventType:= TInterfaceHelper.GetQualifiedName(AEvent);
    var LSubscriptions: TSubscriptions;

    if not TryGetCategorizedSubscriptions<SubscribeAttribute>(TSubscriberMethod.EncodeCategory(AContext, LEventType), LSubscriptions) then
      Exit;

    for var LSubscription in LSubscriptions do
    begin
      if not LSubscription.Active then
        Continue;

      var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;
      PostToSubscription(LSubscription, AEvent, LIsMainThread);
    end;
  finally
    FMrewSync.EndRead;
  end;
end;

procedure TEventBus.PostToChannel(ASubscription: TSubscription; const AMessage: string; AIsMainThread: Boolean);
begin
  if not Assigned(ASubscription.Subscriber) then
    Exit;

  var LProc: TProc :=
    procedure
    begin
      InvokeSubscriber(ASubscription, [AMessage]);
    end;

  case ASubscription.SubscriberMethod.ThreadMode of
    Posting:
      LProc();

    Main:
      if (AIsMainThread) then
        LProc()
      else
        TThread.Queue(nil, TThreadProcedure(LProc));

    Background:
      if (AIsMainThread) then
        TTask.Run(LProc)
      else
        LProc();

    Async:
      TTask.Run(LProc);
  else
    raise EUnknownThreadMode.CreateFmt('Unknown thread mode: %s.', [Ord(ASubscription.SubscriberMethod.ThreadMode)]);
  end;
end;

procedure TEventBus.PostToSubscription(ASubscription: TSubscription; const AEvent: IInterface; AIsMainThread: Boolean);
begin
  if not Assigned(ASubscription.Subscriber) then
    Exit;

  var LProc: TProc :=
    procedure
    begin
     InvokeSubscriber(ASubscription, [AEvent as TObject]);
    end;

  case ASubscription.SubscriberMethod.ThreadMode of
    Posting:
      LProc();

    Main:
      if (AIsMainThread) then
        LProc()
      else
        TThread.Queue(nil, TThreadProcedure(LProc));

    Background:
      if (AIsMainThread) then
        TTask.Run(LProc)
      else
        LProc();

    Async:
      TTask.Run(LProc);
  else
    raise EUnknownThreadMode.CreateFmt('Unknown thread mode: %s.', [Ord(ASubscription.SubscriberMethod.ThreadMode)]);
  end;
end;

procedure TEventBus.RegisterNewContext(ASubscriber: TObject; AEvent: IInterface; const AOldContext: string; const ANewContext: string);
const
  sSubscriberHasNRE = 'Subscruber has null reference.';
  sNoMatchedSubscription = 'No existing subscription that matches the old context [%s] and the event type [%s].';
begin
  FMrewSync.BeginWrite;

  try
    if not Assigned(ASubscriber) then
      raise EArgumentException.Create(sSubscriberHasNRE);

    var LEventName := TInterfaceHelper.GetQualifiedName(AEvent);
    var LCategory := TSubscriberMethod.EncodeCategory(AOldContext, LEventName);
    var LRemovedSubscription := RemoveSubscription<SubscribeAttribute>(ASubscriber, LCategory);

    if LRemovedSubscription = nil then
      raise EArgumentException.CreateFmt(sNoMatchedSubscription, [AOldContext, LEventName]);

    try
      with LRemovedSubscription.SubscriberMethod do
      begin
        var LNewSubMethod := TSubscriberMethod.Create(Method, EventType, ThreadMode, ANewContext, Priority);
        Subscribe<SubscribeAttribute>(ASubscriber, LNewSubMethod);
      end;
    finally
      LRemovedSubscription.Free;
    end;
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.RegisterSubscriber<T>(ASubscriber: TObject; ARaiseExcIfEmpty: Boolean);
begin
  FMrewSync.BeginWrite;

  try
    var LSubscriberClass := ASubscriber.ClassType;
    var LSubscriberMethods := TSubscribersFinder.FindSubscriberMethods<T>(LSubscriberClass, ARaiseExcIfEmpty);

    for var LSubscriberMethod in LSubscriberMethods do
      Subscribe<T>(ASubscriber, LSubscriberMethod);
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.RegisterSubscriberForChannels(ASubscriber: TObject);
begin
  RegisterSubscriber<ChannelAttribute>(ASubscriber, True);
end;

procedure TEventBus.RegisterSubscriberForEvents(ASubscriber: TObject);
begin
  RegisterSubscriber<SubscribeAttribute>(ASubscriber, True);
end;

function TEventBus.RemoveSubscription<T>(ASubscriber: TObject; const ACategory: string): TSubscription;
begin
  Result := nil;
  var LSubscriptions: TSubscriptions;

  if not TryGetCategorizedSubscriptions<T>(ACategory, LSubscriptions) then
    Exit;

  for var LSubscription in LSubscriptions do
  begin
    if LSubscription.Subscriber = ASubscriber then
    begin
      Result := LSubscriptions.Extract(LSubscription); // Found!
      Break;
    end
  end;

  if Assigned(Result) then
  begin
    Unsubscribe<T>(ASubscriber, ACategory);
    Result.Active := False;
  end;

  var LCategories: TMethodCategories;
  if TryGetSubscriberCategories<T>(ASubscriber, LCategories) then
    LCategories.Remove(ACategory);
end;

procedure TEventBus.SilentRegisterSubscriberForChannels(ASubscriber: TObject);
begin
  RegisterSubscriber<ChannelAttribute>(ASubscriber, False);
end;

procedure TEventBus.SilentRegisterSubscriberForEvents(ASubscriber: TObject);
begin
  RegisterSubscriber<SubscribeAttribute>(ASubscriber, False);
end;

procedure TEventBus.Subscribe<T>(ASubscriber: TObject; ASubscriberMethod: TSubscriberMethod);
begin
  var LCategory := ASubscriberMethod.Category; // Category = Context:EventType
  var LSubscriptions := GetCreateCategorizedSubscriptions<T>(LCategory);
  var LNewSubscription := TSubscription.Create(ASubscriber, ASubscriberMethod);

  if not LSubscriptions.Contains(LNewSubscription) then
  begin
    LSubscriptions.Add(LNewSubscription);
  end else
  begin
    LNewSubscription.Free;
    raise ESubscriberMethodAlreadyRegistered.CreateFmt('Subscriber [%s] already registered to %s.', [ASubscriber.ClassName, LCategory]);
  end;

  GetCreateSubscriberCategories<T>(ASubscriber).Add(LCategory);
end;

function TEventBus.TryGetSubscriberCategories<T>(const ASubscriber: TObject; out ACategories: TMethodCategories): Boolean;
begin
  ACategories := nil;
  var LAttrName := T.ClassName;
  var LSubscriberToCategoriesMap: TSubscriberToMethodCategoriesMap;

  if FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
  begin
    LSubscriberToCategoriesMap := FSubscriberToCategoriesByAttrName[LAttrName];

    if LSubscriberToCategoriesMap.ContainsKey(ASubscriber) then
      ACategories := LSubscriberToCategoriesMap[ASubscriber];
  end;

  Result := Assigned(ACategories);
end;

function TEventBus.TryGetCategorizedSubscriptions<T>(const ACategory: string; out ASubscriptions: TSubscriptions): Boolean;
begin
  var LAttrName := T.ClassName;
  var LCatToSubsMap: TMethodCategoryToSubscriptionsMap;
  ASubscriptions := nil;

  if FCategoryToSubscriptionsByAttrName.ContainsKey(LAttrName) then
  begin
    LCatToSubsMap := FCategoryToSubscriptionsByAttrName[LAttrName];

    if LCatToSubsMap.ContainsKey(ACategory) then
      ASubscriptions := LCatToSubsMap[ACategory];
  end;

  Result := Assigned(ASubscriptions);
end;

procedure TEventBus.UnregisterForChannels(ASubscriber: TObject);
begin
  UnregisterSubscriber<ChannelAttribute>(ASubscriber);
end;

procedure TEventBus.UnregisterForEvents(ASubscriber: TObject);
begin
  UnregisterSubscriber<SubscribeAttribute>(ASubscriber);
end;

procedure TEventBus.UnregisterSubscriber<T>(ASubscriber: TObject);
begin
  FMrewSync.BeginWrite;

  try
    var LCategories: TMethodCategories;

    if TryGetSubscriberCategories<T>(ASubscriber, LCategories) then
      for var LCategory in LCategories do
        Unsubscribe<T>(ASubscriber, LCategory);

    DeleteSubscriber<T>(ASubscriber);
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.Unsubscribe<T>(ASubscriber: TObject; const AMethodCategory: TMethodCategory);
begin
  var LSubscriptions: TObjectList<TSubscription>;
  if not TryGetCategorizedSubscriptions<T>(AMethodCategory, LSubscriptions) then
    Exit;

  if (LSubscriptions.Count < 1) then
    Exit;

  for var I := LSubscriptions.Count - 1 downto 0 do
  begin
    var LSubscription := LSubscriptions[I];
    // Note - If the subscriber has been freed without unregistering itself, calling
    // LSubscription.Subscriber.Equals() will cause Access Violation, hence use '=' instead.
    if LSubscription.Subscriber = ASubscriber then
    begin
      LSubscription.Active := False;
      LSubscriptions.Delete(I);
    end;
  end;
end;

class constructor TEventBusFactory.Create;
begin
  FGlobalEventBus := TEventBus.Create;
end;

function TEventBusFactory.CreateEventBus: IEventBus;
begin
  Result := TEventBus.Create;
end;

end.
