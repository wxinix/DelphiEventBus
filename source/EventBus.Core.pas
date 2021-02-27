{*******************************************************************************
  Copyright 2016-2020 Daniele Spinetti

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

  http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
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

    procedure InvokeSubscriber(ASubscription: TSubscription; const Args: array of TValue);
    function IsRegistered<T: TSubscriberMethodAttribute>(ASubscriber: TObject): Boolean;
    procedure RegisterSubscriber<T: TSubscriberMethodAttribute>(ASubscriber: TObject; ARaiseExcIfEmpty: Boolean);
    function RemoveSubscription<T: TSubscriberMethodAttribute>(ASubscriber: TObject; const ACategory: string): TSubscription;
    procedure Subscribe<T: TSubscriberMethodAttribute>(ASubscriber: TObject; ASubscriberMethod: TSubscriberMethod);
    procedure UnregisterSubscriber<T: TSubscriberMethodAttribute>(ASubscriber: TObject);
    procedure Unsubscribe<T: TSubscriberMethodAttribute>(ASubscriber: TObject; const AMethodCategory: TMethodCategory);
  protected
    procedure PostToChannel(ASubscription: TSubscription; const AMessage: string; AIsMainThread: Boolean); virtual;
    procedure PostToSubscription(ASubscription: TSubscription; const AEvent: IInterface; AIsMainThread: Boolean); virtual;
  public
    constructor Create; virtual;
    destructor Destroy; override;

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

procedure TEventBus.InvokeSubscriber(ASubscription: TSubscription; const Args: array of TValue);
begin
  try
    if not ASubscription.Active then
      Exit;

    ASubscription.SubscriberMethod.Method.Invoke(ASubscription.Subscriber, Args);
  except
    on E: Exception do begin
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
    var LAttrName := ChannelAttribute.ClassName;
    var LCategoryToSubscriptionsMap: TMethodCategoryToSubscriptionsMap;

    if not FCategoryToSubscriptionsByAttrName.TryGetValue(LAttrName, LCategoryToSubscriptionsMap) then
      Exit;

    var LSubscriptions: TSubscriptions;
    if not LCategoryToSubscriptionsMap.TryGetValue(TSubscriberMethod.EncodeCategory(AChannel), LSubscriptions) then
      Exit;

    var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;

    for var LSubscription in LSubscriptions do begin
      if (LSubscription.Context <> AChannel) or (not LSubscription.Active) then Continue;
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
    var LAttrName := SubscribeAttribute.ClassName;
    var LCategoryToSubscriptionsMap: TMethodCategoryToSubscriptionsMap;

    if not FCategoryToSubscriptionsByAttrName.TryGetValue(LAttrName, LCategoryToSubscriptionsMap) then
      Exit;

    var LEventType:= TInterfaceHelper.GetQualifiedName(AEvent);
    var LSubscriptions: TSubscriptions;

    if not LCategoryToSubscriptionsMap.TryGetValue(TSubscriberMethod.EncodeCategory(AContext, LEventType), LSubscriptions) then
      Exit;

    var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;

    for var LSubscription in LSubscriptions do begin
      if not LSubscription.Active then
        Continue;

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
  sNoMatchedSubscription = 'There is no existing subscription that matches the event type [%s] and old context [%s].';
begin
  FMrewSync.BeginWrite;

  try
    if not Assigned(ASubscriber) then
      raise EArgumentException.Create(sSubscriberHasNRE);

    var LEventName := TInterfaceHelper.GetQualifiedName(AEvent);
    var LMethodCategory := TSubscriberMethod.EncodeCategory(AOldContext, LEventName);
    var LRemovedSubscription := RemoveSubscription<SubscribeAttribute>(ASubscriber, LMethodCategory);

    if LRemovedSubscription = nil then
      raise EArgumentException.CreateFmt(sNoMatchedSubscription, [LEventName, AOldContext]);

    try
      var LOldSubMethod := LRemovedSubscription.SubscriberMethod;
      var LNewSubMethod := TSubscriberMethod.Create(LOldSubMethod.Method, LOldSubMethod.EventType, LOldSubMethod.ThreadMode, ANewContext, LOldSubMethod.Priority);
      Subscribe<SubscribeAttribute>(ASubscriber, LNewSubMethod);
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

  var LAttrName := T.ClassName;
  var LCategoryToSubscriptionsMap: TMethodCategoryToSubscriptionsMap;

  if (not FCategoryToSubscriptionsByAttrName.TryGetValue(LAttrName, LCategoryToSubscriptionsMap)) then
    Exit;

  var LSubscriptions: TSubscriptions;
  if (not LCategoryToSubscriptionsMap.TryGetValue(ACategory, LSubscriptions)) then
    Exit;

  for var LSubscription in LSubscriptions do begin
    if LSubscription.Subscriber = ASubscriber then begin
      Result := LSubscriptions.Extract(LSubscription);
      Break;
    end
  end;

  if Assigned(Result) then begin
    Unsubscribe<T>(ASubscriber, ACategory);
    Result.Active := False;
  end
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
var
  LSubscriptions: TSubscriptions;
  LCategories: TMethodCategories;
  LCategoryToSubscriptionsMap: TMethodCategoryToSubscriptionsMap;
  LSubscriberToCategoriesMap: TSubscriberToMethodCategoriesMap;
begin
  var LAttrName := T.ClassName;

  if not FCategoryToSubscriptionsByAttrName.ContainsKey(LAttrName) then begin
    LCategoryToSubscriptionsMap := TMethodCategoryToSubscriptionsMap.Create([doOwnsValues]);
    FCategoryToSubscriptionsByAttrName.Add(LAttrName, LCategoryToSubscriptionsMap);
  end else begin
    LCategoryToSubscriptionsMap := FCategoryToSubscriptionsByAttrName[LAttrName];
  end;

  var LCategory := ASubscriberMethod.Category; // Category = Context:EventType
  var LNewSubscription := TSubscription.Create(ASubscriber, ASubscriberMethod);

  if (not LCategoryToSubscriptionsMap.ContainsKey(LCategory)) then begin

    LSubscriptions := TSubscriptions.Create(
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

    LCategoryToSubscriptionsMap.Add(LCategory, LSubscriptions);
  end else begin
    LSubscriptions := LCategoryToSubscriptionsMap[LCategory];
    if (LSubscriptions.Contains(LNewSubscription)) then begin
      LNewSubscription.Free;
      raise ESubscriberMethodAlreadyRegistered.CreateFmt('Subscriber [%s] already registered to %s.', [ASubscriber.ClassName, LCategory]);
    end;
  end;

  LSubscriptions.Add(LNewSubscription);

  if not FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then begin
    LSubscriberToCategoriesMap := TSubscriberToMethodCategoriesMap.Create([doOwnsValues]);
    FSubscriberToCategoriesByAttrName.Add(LAttrName, LSubscriberToCategoriesMap);
  end else begin
    LSubscriberToCategoriesMap := FSubscriberToCategoriesByAttrName[LAttrName];
  end;

  if (not LSubscriberToCategoriesMap.TryGetValue(ASubscriber, LCategories)) then begin
    LCategories := TMethodCategories.Create;
    LSubscriberToCategoriesMap.Add(ASubscriber, LCategories);
  end;

  LCategories.Add(LCategory);
end;

procedure TEventBus.UnregisterSubscriber<T>(ASubscriber: TObject);
begin
  FMrewSync.BeginWrite;

  try
    var LAttrName := T.ClassName;
    var LSubscriberToCategoriesMap: TSubscriberToMethodCategoriesMap;

    if not FSubscriberToCategoriesByAttrName.TryGetValue(LAttrName, LSubscriberToCategoriesMap) then
      Exit;

    var LCategories: TMethodCategories;
    if LSubscriberToCategoriesMap.TryGetValue(ASubscriber, LCategories) then begin

      for var LCategory in LCategories do
        Unsubscribe<T>(ASubscriber, LCategory);

      LSubscriberToCategoriesMap.Remove(ASubscriber);
    end;
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.UnregisterForChannels(ASubscriber: TObject);
begin
  UnregisterSubscriber<ChannelAttribute>(ASubscriber);
end;

procedure TEventBus.UnregisterForEvents(ASubscriber: TObject);
begin
  UnregisterSubscriber<SubscribeAttribute>(ASubscriber);
end;

procedure TEventBus.Unsubscribe<T>(ASubscriber: TObject; const AMethodCategory: TMethodCategory);
begin
  var LAttrName := T.ClassName;
  var LCategoryToSubscriptionsMap: TMethodCategoryToSubscriptionsMap;

  if not FCategoryToSubscriptionsByAttrName.TryGetValue(LAttrName, LCategoryToSubscriptionsMap) then
    Exit;

  var LSubscriptions: TObjectList<TSubscription>;
  if not LCategoryToSubscriptionsMap.TryGetValue(AMethodCategory, LSubscriptions) then
    Exit;

  if (LSubscriptions.Count < 1) then
    Exit;

  for var I := LSubscriptions.Count - 1 downto 0 do begin
    var LSubscription := LSubscriptions[I];
    // Note - If the subscriber has been freed without unregistering itself, calling
    // LSubscription.Subscriber.Equals() will cause Access Violation, hence use '=' instead.
    if LSubscription.Subscriber = ASubscriber then begin
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
