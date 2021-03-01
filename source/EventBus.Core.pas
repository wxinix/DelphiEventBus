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
  TCategory = string;
  TCategories = TList<TCategory>;
  TCategoryToSubscriptionsMap = TObjectDictionary<TCategory, TSubscriptions>;
  TSubscriberToCategoriesMap = TObjectDictionary<TObject, TCategories>;

  TAttributeName = string;
  TCategoryToSubscriptionsByAttributeName = TObjectDictionary<TAttributeName, TCategoryToSubscriptionsMap>;
  TSubscriberToCategoriesByAttributeName = TObjectDictionary<TAttributeName, TSubscriberToCategoriesMap>;
  {$ENDREGION}

  TEventBus = class(TInterfacedObject, IEventBus)
  strict private
    FCategoryToSubscriptionsByAttrName: TCategoryToSubscriptionsByAttributeName;
    FMrewSync: TLightweightMREW;
    FSubscriberToCategoriesByAttrName: TSubscriberToCategoriesByAttributeName;
  strict private

    /// <summary>
    ///   Retrieves the subscriptions that belong to the specified category, while matching the specified
    ///   subscriber attribute. If the category currently does not have any subscriptions belonging to to
    ///   it, a new subscribers list will be created and returned.
    /// </summary>
    /// <typeparam name="T">
    ///   The subscriber attribute
    /// </typeparam>
    /// <param name="ACategory">
    ///   The category.
    /// </param>
    /// <returns>
    ///   The subscriptions belonging to the category and matching the subscriber attribute.
    /// </returns>
    function GetCreateSubscriptions<T: TSubscriberAttribute>(const ACategory: string): TSubscriptions;

    /// <summary>
    ///   Retrieves the categories associated with the subscriber object, while matching the specified
    ///   subscriber attribute. If there is no existing categories for the subscriber, a new categories
    ///   list will be created and returned.
    /// </summary>
    /// <typeparam name="T">
    ///   The subscriber attribute
    /// </typeparam>
    /// <param name="ASubscriber">
    ///   The subscriber object
    /// </param>
    /// <returns>
    ///   The categories associated with the subscriber object, while matching the attribute.
    /// </returns>
    function GetCreateCategories<T: TSubscriberAttribute>(const ASubscriber: TObject): TCategories;

    /// <summary>
    ///   Retrieves the categories associated with the subscriber object.
    /// </summary>
    /// <param name="ASubscriber">
    ///   The subscriber object
    /// </param>
    /// <param name="ACategories">
    ///   The retrieved categories. Nil will be returned if the subscriber does not have any existing
    ///   categoreis.
    /// </param>
    /// <returns>
    ///   Returns True if retrieved successfully, False otherwise.
    /// </returns>
    function TryGetCategories<T: TSubscriberAttribute>(const ASubscriber: TObject; out ACategories: TCategories): Boolean;


    /// <summary>
    ///   Retrieves all subscriptions (that may come from different subscriber objects) belonging to the
    ///   specified category.
    /// </summary>
    /// <typeparam name="T">
    ///   The subscriber attribute.
    /// </typeparam>
    /// <param name="ACategory">
    ///   The category for which subscriptions are to be retrieved.
    /// </param>
    /// <param name="ASubscriptions">
    ///   Retrieved subscriptions. Nil will be returned if there is not existing subscriptions for the
    ///   specified category.
    /// </param>
    /// <returns>
    ///   Returns True if retrieves successfully, False otherwise.
    /// </returns>
    function TryGetSubscriptions<T: TSubscriberAttribute>(const ACategory: string; out ASubscriptions: TSubscriptions): Boolean;

    /// <summary>
    ///   Checks if the specified subscriber object has been registered.
    /// </summary>
    /// <typeparam name="T">
    ///   Subscriber attribute.
    /// </typeparam>
    /// <param name="ASubscriber">
    ///   The subscriber object to check.
    /// </param>
    /// <returns>
    ///   Returns True if already registered, False otherwise.
    /// </returns>
    function IsRegistered<T: TSubscriberAttribute>(ASubscriber: TObject): Boolean;

    /// <summary>
    ///   Registers the specified subscriber object.
    /// </summary>
    /// <param name="ASubscriber">
    ///   The subscriber object.
    /// </param>
    /// <param name="ARaiseExcIfEmpty">
    ///   If True, exception will be raised when the subscriber object does not have any subscriber
    ///   methods.
    /// </param>
    /// <exception cref="EObjectHasNoSubscriberMethods">
    ///   If the subscriber object does not have any subscriber methods.
    /// </exception>
    procedure DoRegister<T: TSubscriberAttribute>(ASubscriber: TObject; ARaiseExcIfEmpty: Boolean);

    /// <summary>
    ///   Subscribes the specific method for the subscriber object.
    /// </summary>
    procedure DoSubscribe<T: TSubscriberAttribute>(ASubscriber: TObject; ASubscriberMethod: TSubscriberMethod);

    /// <summary>
    ///   Unregisters the subscriber object.
    /// </summary>
    procedure DoUnregister<T: TSubscriberAttribute>(ASubscriber: TObject);

    /// <summary>
    ///   Unsubscribes the specific event category for the subscriber object.
    /// </summary>
    function DoUnsubscribe<T: TSubscriberAttribute>(ASubscriber: TObject; const ACategory: TCategory): TSubscription;

    procedure InvokeSubscriber(ASubscription: TSubscription; const AParams: array of TValue);
    procedure Post(ASubscription: TSubscription; const AMessage: string; AIsMainThread: Boolean); overload;
    procedure Post(ASubscription: TSubscription; const AEvent: IInterface; AIsMainThread: Boolean); overload;

    {$REGION'IEventBus interface methods'}
    function IsRegisteredForChannels(ASubscriber: TObject): Boolean;
    function IsRegisteredForEvents(ASubscriber: TObject): Boolean;
    procedure Post(const AChannel: string; const AMessage: string); overload;
    procedure Post(const AEvent: IInterface; const AContext: string = ''); overload;
    procedure RegisterNewContext(ASubscriber: TObject; AEvent: IInterface; const AOldContext, ANewContext: string);
    procedure RegisterSubscriberForChannels(ASubscriber: TObject);
    procedure SilentRegisterSubscriberForChannels(ASubscriber: TObject);
    procedure RegisterSubscriberForEvents(ASubscriber: TObject);
    procedure SilentRegisterSubscriberForEvents(ASubscriber: TObject);
    procedure UnregisterForChannels(ASubscriber: TObject);
    procedure UnregisterForEvents(ASubscriber: TObject);
    {$ENDREGION}
  public
    constructor Create; virtual;
    destructor Destroy; override;
  end;

constructor TEventBus.Create;
begin
  inherited Create;
  FCategoryToSubscriptionsByAttrName := TCategoryToSubscriptionsByAttributeName.Create([doOwnsValues]);
  FSubscriberToCategoriesByAttrName := TSubscriberToCategoriesByAttributeName.Create([doOwnsValues]);
end;

destructor TEventBus.Destroy;
begin
  FCategoryToSubscriptionsByAttrName.Free;
  FSubscriberToCategoriesByAttrName.Free;
  inherited;
end;

function TEventBus.GetCreateCategories<T>(const ASubscriber: TObject): TCategories;
begin
  var LAttrName := T.ClassName;
  var LSubsToCatsMap: TSubscriberToCategoriesMap;

  if not FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
  begin
    LSubsToCatsMap := TSubscriberToCategoriesMap.Create([doOwnsValues]);
    FSubscriberToCategoriesByAttrName.Add(LAttrName, LSubsToCatsMap);
  end else
  begin
    LSubsToCatsMap := FSubscriberToCategoriesByAttrName[LAttrName];
  end;

  if (not LSubsToCatsMap.TryGetValue(ASubscriber, Result)) then
  begin
    Result := TCategories.Create;
    LSubsToCatsMap.Add(ASubscriber, Result);
  end;
end;

function TEventBus.GetCreateSubscriptions<T>(const ACategory: string): TSubscriptions;
begin
  var LAttrName := T.ClassName;
  var LCatToSubsMap: TCategoryToSubscriptionsMap;

  if not FCategoryToSubscriptionsByAttrName.ContainsKey(LAttrName) then
  begin
    LCatToSubsMap := TCategoryToSubscriptionsMap.Create([doOwnsValues]);
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

procedure TEventBus.InvokeSubscriber(ASubscription: TSubscription; const AParams: array of TValue);
begin
  try
    if not ASubscription.Active then
      Exit;

    ASubscription.SubscriberMethod.Method.Invoke(ASubscription.Subscriber, AParams);
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
    var LSubscriberToCatsMap: TSubscriberToCategoriesMap;

    if not FSubscriberToCategoriesByAttrName.TryGetValue(LAttrName, LSubscriberToCatsMap) then
      Exit(False);

    Result := LSubscriberToCatsMap.ContainsKey(ASubscriber);
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
    var LCategory := TSubscriberMethod.EncodeCategory(AChannel);

    if not TryGetSubscriptions<ChannelAttribute>(LCategory, LSubscriptions) then
      Exit;

    for var LSubscription in LSubscriptions do
    begin
      if (LSubscription.Context <> AChannel) or (not LSubscription.Active) then
        Continue;

      var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;
      Post(LSubscription, AMessage, LIsMainThread);
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
    var LCategory := TSubscriberMethod.EncodeCategory(AContext, LEventType);
    var LSubscriptions: TSubscriptions;

    if not TryGetSubscriptions<SubscribeAttribute>(LCategory, LSubscriptions) then
      Exit;

    for var LSubscription in LSubscriptions do
    begin
      if not LSubscription.Active then
        Continue;

      var LIsMainThread := MainThreadID = TThread.CurrentThread.ThreadID;
      Post(LSubscription, AEvent, LIsMainThread);
    end;
  finally
    FMrewSync.EndRead;
  end;
end;

procedure TEventBus.Post(ASubscription: TSubscription; const AMessage: string; AIsMainThread: Boolean);
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

procedure TEventBus.Post(ASubscription: TSubscription; const AEvent: IInterface; AIsMainThread: Boolean);
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

procedure TEventBus.RegisterNewContext(ASubscriber: TObject; AEvent: IInterface; const AOldContext, ANewContext: string);
begin
  FMrewSync.BeginWrite;

  try
    if not Assigned(ASubscriber) then
      raise EArgumentException.Create('Subscruber has null reference.');

    var LEventName := TInterfaceHelper.GetQualifiedName(AEvent);
    var LOldCategory := TSubscriberMethod.EncodeCategory(AOldContext, LEventName);
    var LOldSubscription := DoUnsubscribe<SubscribeAttribute>(ASubscriber, LOldCategory);

    if Assigned(LOldSubscription) then
    begin
      var LCategories: TCategories;
      if TryGetCategories<SubscribeAttribute>(ASubscriber, LCategories) then
        LCategories.Remove(LOldCategory);
    end else
    begin
      raise EArgumentException.CreateFmt('Subscriber [%s] does not have subscription belonging to category [%s].',
        [ASubscriber.ClassName, LOldCategory]);
    end;

    try
      with LOldSubscription.SubscriberMethod do
      begin
        var LNewSubMethod := TSubscriberMethod.Create(Method, EventType, ThreadMode, ANewContext, Priority);
        // Possibly throws ESubscriberMethodAlreadyRegistered
        DoSubscribe<SubscribeAttribute>(ASubscriber, LNewSubMethod);
      end;
    finally
      LOldSubscription.Free;
    end;
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.DoRegister<T>(ASubscriber: TObject; ARaiseExcIfEmpty: Boolean);
begin
  FMrewSync.BeginWrite;

  try
    var LSubscriberClass := ASubscriber.ClassType;
    var LSubscriberMethods := TSubscribersFinder.FindSubscriberMethods<T>(LSubscriberClass, ARaiseExcIfEmpty);

    for var LSubscriberMethod in LSubscriberMethods do
      DoSubscribe<T>(ASubscriber, LSubscriberMethod);
  finally
    FMrewSync.EndWrite;
  end;
end;

procedure TEventBus.RegisterSubscriberForChannels(ASubscriber: TObject);
begin
  DoRegister<ChannelAttribute>(ASubscriber, True);
end;

procedure TEventBus.RegisterSubscriberForEvents(ASubscriber: TObject);
begin
  DoRegister<SubscribeAttribute>(ASubscriber, True);
end;

procedure TEventBus.SilentRegisterSubscriberForChannels(ASubscriber: TObject);
begin
  DoRegister<ChannelAttribute>(ASubscriber, False);
end;

procedure TEventBus.SilentRegisterSubscriberForEvents(ASubscriber: TObject);
begin
  DoRegister<SubscribeAttribute>(ASubscriber, False);
end;

procedure TEventBus.DoSubscribe<T>(ASubscriber: TObject; ASubscriberMethod: TSubscriberMethod);
begin
  var LCategory := ASubscriberMethod.Category; // Category = Context:EventType
  var LSubscriptions := GetCreateSubscriptions<T>(LCategory);
  var LNewSubscription := TSubscription.Create(ASubscriber, ASubscriberMethod);

  if not LSubscriptions.Contains(LNewSubscription) then
  begin
    LSubscriptions.Add(LNewSubscription);
  end else
  begin
    LNewSubscription.Free;
    raise ESubscriberMethodAlreadyRegistered.CreateFmt(
      'Subscriber [%s] already registered to %s.',
      [ASubscriber.ClassName, LCategory]);
  end;

  GetCreateCategories<T>(ASubscriber).Add(LCategory);
end;

function TEventBus.TryGetCategories<T>(const ASubscriber: TObject; out ACategories: TCategories): Boolean;
begin
  ACategories := nil;
  var LAttrName := T.ClassName;
  var LSubscriberToCategoriesMap: TSubscriberToCategoriesMap;

  if FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
  begin
    LSubscriberToCategoriesMap := FSubscriberToCategoriesByAttrName[LAttrName];

    if LSubscriberToCategoriesMap.ContainsKey(ASubscriber) then
      ACategories := LSubscriberToCategoriesMap[ASubscriber];
  end;

  Result := Assigned(ACategories);
end;

function TEventBus.TryGetSubscriptions<T>(const ACategory: string; out ASubscriptions: TSubscriptions): Boolean;
begin
  var LAttrName := T.ClassName;
  var LCatToSubsMap: TCategoryToSubscriptionsMap;
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
  DoUnregister<ChannelAttribute>(ASubscriber);
end;

procedure TEventBus.UnregisterForEvents(ASubscriber: TObject);
begin
  DoUnregister<SubscribeAttribute>(ASubscriber);
end;

procedure TEventBus.DoUnregister<T>(ASubscriber: TObject);
begin
  FMrewSync.BeginWrite;

  try
    var LCategories: TCategories;
    if TryGetCategories<T>(ASubscriber, LCategories) then
    begin
      for var LCategory in LCategories do
      begin
        var LSubscription := DoUnsubscribe<T>(ASubscriber, LCategory);
        if Assigned(LSubscription) then
          LSubscription.Free;
      end;
    end;

    {$REGION 'Remove Subscriber'}
    var LAttrName := T.ClassName;
    var LSubscriberToCatsMap: TSubscriberToCategoriesMap;

    if FSubscriberToCategoriesByAttrName.ContainsKey(LAttrName) then
    begin
      LSubscriberToCatsMap := FSubscriberToCategoriesByAttrName[LAttrName];

      if LSubscriberToCatsMap.ContainsKey(ASubscriber) then
        LSubscriberToCatsMap.Remove(ASubscriber);
    end;
    {$ENDREGION}
  finally
    FMrewSync.EndWrite;
  end;
end;

function TEventBus.DoUnsubscribe<T>(ASubscriber: TObject; const ACategory: TCategory): TSubscription;
begin
  Result := nil;

  var LSubscriptions: TObjectList<TSubscription>;
  if not TryGetSubscriptions<T>(ACategory, LSubscriptions) then
    Exit;

  for var LSubscription in LSubscriptions do
  begin
    // If the subscriber has been freed without unregistering itself, calling LSubscription.Subscriber.Equals
    // will cause access violation, hence we use '=' here instead.
    if LSubscription.Subscriber = ASubscriber then
    begin
      Result := LSubscriptions.Extract(LSubscription);
      Result.Active := False;
      Break;
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
