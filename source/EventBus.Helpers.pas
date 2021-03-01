{****************************************************************************************
           DUnitX Copyright (C) 2015 Vincent Parrett & Contributors
           vincent@finalbuilder.com
           http://www.finalbuilder.com

*****************************************************************************************

Licensed under the Apache License Version 2.0 (the "License"); you may not use this  file
except in compliance with the License. You may obtain a copy of the  License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the
License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
either express or implied.See the License for the specific language governing permissions
and limitations under the License.

******************************************************************************************
Portions of the file also fall under the following license  as they were taken from  the
DSharp Project

          https://bitbucket.org/sglienke/dsharp
          Copyright (c) 2011-2012, Stefan Glienke
          All rights reserved.

Redistribution and use in source and binary forms,  with  or without  modification, are
permitted provided that the following conditions are met:

- Redistributions of source code must retain the above copyright notice,  this list  of
  conditions and the following disclaimer.

- Redistributions in binary form must reproduce the above copyright  notice,  this  list
  of conditions and the following disclaimer in the documentation and/or other materials
  provided with the distribution.

- Neither the name of this library nor  the  names of its contributors  may  be used  to
  endorse or promote products derived from this software without specific prior  written
  permission.

THIS SOFTWARE IS PROVIDED  BY THE COPYRIGHT  HOLDERS AND CONTRIBUTORS  "AS IS"  AND  ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE  IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE  ARE DISCLAIMED.  IN NO EVENT SHALL
THE  COPYRIGHT HOLDER  OR  CONTRIBUTORS  BE  LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT
OF SUBSTITUTE GOODS OR SERVICES;LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING  IN  ANY  WAY  OUT OF THE  USE  OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
****************************************************************************************}

unit EventBus.Helpers;

interface

uses
  System.Generics.Collections,
  System.Rtti,
  System.SysUtils,
  System.TimeSpan,
  System.Types,
  System.TypInfo;

type
  TCustomAttributeClass = class of TCustomAttribute;

  TAttributeUtils = class
  public
    class function ContainsAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): Boolean;
    class function FindAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): TCustomAttribute; overload;
    class function FindAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass; var AAttribute: TCustomAttribute; const AStartIndex: Integer = 0): Integer; overload;
    class function FindAttributes(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): TArray<TCustomAttribute>;
  end;

  TStrUtils = class
    class function EncodeWhitespace(const AStr: string): string;
    class function Join(const AValues: TArray<string>; const ADelimiter: string): string; overload;
    class function PadString(const AStr: string; const ATotalLength: Integer; const APadLeft: Boolean = True; APadChar: Char = ' '): string;
    class function SplitString(const AStr, ADelimiters: string): TArray<string>;
  end;

  TListStringUtils = class
    class function ToArray(const AValues: TList<string>): TArray<string>;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.TObject">TObject</see> for easier RTTI use.
  ///	</summary>
  TObjectHelper = class helper for TObject
  public
    ///	<summary>
    ///	  Returns a list of all fields of the object.
    ///	</summary>
    function GetFields: TArray<TRttiField>;

    ///	<summary>
    ///	  Returns the field with the given name; <b>nil</b> if nothing is found.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the field to find
    ///	</param>
    function GetField(const AName: string): TRttiField;

    ///	<summary>
    ///	  Returns the member with the given name; <b>nil</b> if nothing is found.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the member to find
    ///	</param>
    function GetMember(const AName: string): TRttiMember;

    ///	<summary>
    ///	  Returns a list of all methods of the object.
    ///	</summary>
    function GetMethods: TArray<TRttiMethod>;

    ///	<summary>
    ///	  Returns the method at the given code address; <b>nil</b> if nothing
    ///	  is found.
    ///	</summary>
    ///	<param name="ACodeAddress">
    ///	  Code address of the method to find.
    ///	</param>
    function GetMethod(ACodeAddress: Pointer): TRttiMethod; overload;

    ///	<summary>
    ///	  Returns the method with the given name; <b>nil</b> if nothing is
    ///	  found.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the method to find
    ///	</param>
    function GetMethod(const AName: string): TRttiMethod; overload;

    ///	<summary>
    ///	  Returns a list of all properties of the object.
    ///	</summary>
    function GetProperties: TArray<TRttiProperty>;

    ///	<summary>
    ///	  Returns the property with the given name; <b>nil</b> if nothing is
    ///	  found.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the property to find
    ///	</param>
    function GetProperty(const AName: string): TRttiProperty;

    ///	<summary>
    ///	  Returns the type of the object; nil if nothing is found.
    ///	</summary>
    function GetType: TRttiType;

    ///	<summary>
    ///	  Returns if the object contains a field with the given name.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the field to find
    ///	</param>
    function HasField(const AName: string): Boolean;

    ///	<summary>
    ///	  Returns if the object contains a method with the given name.
    ///	</summary>
    function HasMethod(const AName: string): Boolean;

    ///	<summary>
    ///	  Returns if the object contains a property with the given name.
    ///	</summary>
    function HasProperty(const AName: string): Boolean;

    ///	<summary>
    ///	  Retrieves the method with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the field to find
    ///	</param>
    ///	<param name="AField">
    ///	  Field that was found when Result is <b>True</b>
    ///	</param>
    function TryGetField(const AName: string; out AField: TRttiField): Boolean;

    ///	<param name="AName">
    ///	  Name of the member to find
    ///	</param>
    ///	<param name="AMember">
    ///	  Member that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMember(const AName: string; out AMember: TRttiMember): Boolean;

    ///	<summary>
    ///	  Retrieves the method with the given code address and returns if this
    ///	  was successful.
    ///	</summary>
    ///	<param name="ACodeAddress">
    ///	  Code address of the method to find
    ///	</param>
    ///	<param name="AMethod">
    ///	  Method that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMethod(ACodeAddress: Pointer; out AMethod: TRttiMethod): Boolean; overload;

    ///	<summary>
    ///	  Retrieves the method with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the method to find
    ///	</param>
    ///	<param name="AMethod">
    ///	  Method that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMethod(const AName: string; out AMethod: TRttiMethod): Boolean; overload;

    ///	<summary>
    ///	  Retrieves the property with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the property to find
    ///	</param>
    ///	<param name="AProperty">
    ///	  Property that was found when Result is <b>True</b>
    ///	</param>
    function TryGetProperty(const AName: string; out AProperty: TRttiProperty): Boolean;

    ///	<summary>
    ///	  Retrieves the type of the object and returns if this was successful.
    ///	</summary>
    ///	<param name="AType">
    ///	  Type of the object when Result is <b>True</b>
    ///	</param>
    function TryGetType(out AType: TRttiType): Boolean;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiField">TRttiField</see> for easier RTTI
  ///	  use.
  ///	</summary>
  TRttiFieldHelper = class helper for TRttiField
  public
    ///	<summary>
    ///	  Retrieves the AValue of the field and returns if this was successful.
    ///	</summary>
    ///	<param name="AInstance">
    ///	  Pointer to the AInstance of the field
    ///	</param>
    ///	<param name="AValue">
    ///	  AValue of the field when Result is <b>True</b>
    ///	</param>
    function TryGetValue(AInstance: Pointer; out AValue: TValue): Boolean;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiInvokableType">TRttiInvokableType</see>
  ///	  for easier RTTI use.
  ///	</summary>
  TRttiInvokableTypeHelper = class helper for TRttiInvokableType
  private
    function GetParameterCount: Integer;
  public
    property ParameterCount: Integer read GetParameterCount;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiMember">TRttiMember</see> for easier RTTI
  ///	  use.
  ///	</summary>
  TRttiMemberHelper = class helper for TRttiMember
  private
    function GetMemberIsReadable: Boolean;
    function GetMemberIsWritable: Boolean;
    function GetMemberRttiType: TRttiType;
  public
    property IsReadable: Boolean read GetMemberIsReadable;
    property IsWritable: Boolean read GetMemberIsWritable;
    property RttiType: TRttiType read GetMemberRttiType;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiMethod">TRttiMethod</see> for easier RTTI
  ///	  use.
  ///	</summary>
  TRttiMethodHelper = class helper for TRttiMethod
  private
    function GetParameterCount: Integer;
  public
    function Format(const AArgs: array of TValue; ASkipSelf: Boolean = True): string;
    property ParameterCount: Integer read GetParameterCount;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiObject">TRttiObject</see> for easier RTTI
  ///	  use.
  ///	</summary>
  TRttiObjectHelper = class helper for TRttiObject
  public
    function FindAttribute<T: TCustomAttribute>: T;
    function FindAttributes<T: TCustomAttribute>: TArray<T>;
    function HasAttribute<T: TCustomAttribute>: Boolean;
    function TryGetAttribute<T: TCustomAttribute>(out AAttribute: T): Boolean;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiParameter">TRttiParameter</see> for easier
  ///	  RTTI use.
  ///	</summary>
  TRttiParameterHelper = class helper for TRttiParameter
  public
    class function Equals(const Left, Right: TArray<TRttiParameter>): Boolean;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiProperty">TRttiProperty</see> for easier
  ///	  RTTI use.
  ///	</summary>
  TRttiPropertyHelper = class helper for TRttiProperty
  public
    ///	<summary>
    ///	  Retrieves the AValue of the property and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AInstance">
    ///	  Pointer to the AInstance of the field
    ///	</param>
    ///	<param name="AValue">
    ///	  AValue of the field when Result is <b>True</b>
    ///	</param>
    function TryGetValue(AInstance: Pointer; out AValue: TValue): Boolean;

    ///	<summary>
    ///	  Sets the AValue of the property and returns if this was successful.
    ///	</summary>
    ///	<param name="AInstance">
    ///	  Pointer to the AInstance of the field
    ///	</param>
    ///	<param name="AValue">
    ///	  AValue the field should be set to
    ///	</param>
    function TrySetValue(AInstance: Pointer; AValue: TValue): Boolean;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TRttiType">TRttiType</see> for easier RTTI use.
  ///	</summary>
  TRttiTypeHelper = class helper for TRttiType
  private
    function GetAsInterface: TRttiInterfaceType;
    function GetIsInterface: Boolean;
    function GetMethodCount: Integer;
    function InheritsFrom(OtherType: PTypeInfo): Boolean;
  public
    /// <summary>
    ///   Retrieves all attributes that of type T and its derived classes, specified on the current class
    ///   and its immediate base class.
    /// </summary>
    function FindAttributes<T: TCustomAttribute>: TArray<T>;

    /// <summary>
    ///   Retrieves the generic type arguments.
    /// </summary>
    function GetGenericArguments: TArray<TRttiType>;

    /// <summary>
    ///   Get the generic type definition of the type. Generic type parameters will be replaced as T, T0,
    ///   T1, etc.
    /// </summary>
    function GetGenericTypeDefinition(const AIncludeUnitName: Boolean = True): string;

    /// <summary>
    ///   Retrieves a class member (field, property or method) by the specified name. First searches
    ///   properties, then fields, and lastly methods to get the member that matches the name. The first
    ///   found will be returned.
    /// </summary>
    /// <param name="AName">
    ///   Name of the member.
    /// </param>
    /// <returns>
    ///   The member found. If not found, nil will be returned.
    /// </returns>
    function GetMember(const AName: string): TRttiMember;

    ///	<summary>
    ///	  Returns the method at the given code address; <b>nil</b> if nothing
    ///	  is found.
    ///	</summary>
    ///	<param name="ACodeAddress">
    ///	  Code address of the method to find
    ///	</param>
    function GetMethod(ACodeAddress: Pointer): TRttiMethod; overload;

    /// <summary>
    ///   Retrieves the property by the specified name.
    /// </summary>
    /// <param name="AName">
    ///   Property name to search.
    /// </param>
    /// <returns>
    ///   The property found. If not found, nil will be returned.
    /// </returns>
    function GetProperty(const AName: string): TRttiProperty;

    /// <summary>
    ///   Get the declared parameter-less constructor of the current class. The first found will be
    ///   returned.
    /// </summary>
    /// <returns>
    ///   The constructor found. If not found, nil will be returned.
    /// </returns>
    function GetStandardConstructor: TRttiMethod;

    function IsCovariantTo(AOtherClass: TClass): Boolean; overload;
    function IsCovariantTo(AOtherType: PTypeInfo): Boolean; overload;
    function IsGenericTypeDefinition: Boolean;
    function IsGenericTypeOf(const ABaseTypeName: string): Boolean;
    function IsInheritedFrom(AOtherType: TRttiType): Boolean; overload;
    function IsInheritedFrom(const AOtherTypeName: string): Boolean; overload;
    function MakeGenericType(const ATypeArguments: array of PTypeInfo): TRttiType;

    ///	<summary>
    ///	  Retrieves the method with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the field to find
    ///	</param>
    ///	<param name="AField">
    ///	  Field that was found when Result is <b>True</b>
    ///	</param>
    function TryGetField(const AName: string; out AField: TRttiField): Boolean;

    ///	<summary>
    ///	  Retrieves the member with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the member to find
    ///	</param>
    ///	<param name="AMember">
    ///	  Member that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMember(const AName: string; out AMember: TRttiMember): Boolean;

    ///	<summary>
    ///	  Retrieves the method with the given code address and returns if this
    ///	  was successful.
    ///	</summary>
    ///	<param name="ACodeAddress">
    ///	  Code address of the method to find
    ///	</param>
    ///	<param name="AMethod">
    ///	  Method that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMethod(ACodeAddress: Pointer; out AMethod: TRttiMethod): Boolean; overload;

    ///	<summary>
    ///	  Retrieves the method with the given code address and returns if this
    ///	  was successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the method to find
    ///	</param>
    ///	<param name="AMethod">
    ///	  Method that was found when Result is <b>True</b>
    ///	</param>
    function TryGetMethod(const AName: string; out AMethod: TRttiMethod): Boolean; overload;

    /// <summary>
    ///   Retrieves the parameter-less constructor by first searching the subject type then its parent
    ///   classes. The first found will be returned.
    /// </summary>
    /// <param name="AMethod">
    ///   The constructor found.
    /// </param>
    /// <returns>
    ///   True indicates success, False otherwise.
    /// </returns>
    function TryGetConstructor(out AMethod : TRttiMethod) : Boolean;

    /// <summary>
    ///   Retrieves the destructor declared by the current class.
    /// </summary>
    /// <param name="AMethod">
    ///   The destructor found.
    /// </param>
    /// <returns>
    ///   True indicates success, False otherwise.
    /// </returns>
    function TryGetDestructor(out AMethod : TRttiMethod) : Boolean;

    ///	<summary>
    ///	  Retrieves the property with the given name and returns if this was
    ///	  successful.
    ///	</summary>
    ///	<param name="AName">
    ///	  Name of the property to find
    ///	</param>
    ///	<param name="AProperty">
    ///	  Property that was found when Result is <b>True</b>
    ///	</param>
    function TryGetProperty(const AName: string; out AProperty: TRttiProperty): Boolean;

    /// <summary>
    ///   Retrieves the parameter-less constructor declared by the current class. The first one found will
    ///   be returned.
    /// </summary>
    /// <returns>
    ///   True indicates success, False otherwise. <br />
    /// </returns>
    function TryGetStandardConstructor(out AMethod: TRttiMethod): Boolean;
  public
    property AsInterface: TRttiInterfaceType read GetAsInterface;
    property IsInterface: Boolean read GetIsInterface;
    property MethodCount: Integer read GetMethodCount;
  end;

  ///	<summary>
  ///	  Extends <see cref="System.Rtti.TValue">TValue</see> for easier RTTI use.
  ///	</summary>
  TValueHelper = record helper for TValue
  private
    function GetRttiType: TRttiType;
    class function FromFloat(ATypeInfo: PTypeInfo; AValue: Extended): TValue; static;
  public
    // conversion for almost all standard types
    function TryConvert(ATypeInfo: PTypeInfo; out AResult: TValue): Boolean; overload;
    function TryConvert<T>(out AResult: TValue): Boolean; overload;

    function AsByte: Byte;
    function AsCardinal: Cardinal;
    function AsCurrency: Currency;
    function AsDate: TDate;
    function AsDateTime: TDateTime;
    function AsDouble: Double;
    function AsFloat: Extended;
    function AsPointer: Pointer;
    function AsShortInt: ShortInt;
    function AsSingle: Single;
    function AsSmallInt: SmallInt;
    function AsTime: TTime;
    function AsUInt64: UInt64;
    function AsWord: Word;

    function ToObject: TObject;
    function ToVarRec: TVarRec;

    class function ToString(const AValue: TValue): string; overload; static;
    class function ToString(const AValues: array of TValue): string; overload; static;
    class function ToVarRecs(const AValues: array of TValue): TArray<TVarRec>; static;
    class function Equals(const Left, Right: TArray<TValue>): Boolean; overload; static;
    class function Equals<T>(const Left, Right: T): Boolean; overload; static;

    class function From(ABuffer: Pointer; ATypeInfo: PTypeInfo): TValue; overload; static;
    class function From(AValue: NativeInt; ATypeInfo: PTypeInfo): TValue; overload; static;
    class function From(AObject: TObject; AClass: TClass): TValue; overload; static;
    class function FromBoolean(const AValue: Boolean): TValue; static;
    class function FromString(const AValue: string): TValue; static;
    class function FromVarRec(const AValue: TVarRec): TValue; static;

    function IsBoolean: Boolean;
    function IsByte: Boolean;
    function IsCardinal: Boolean;
    function IsCurrency: Boolean;
    function IsDate: Boolean;
    function IsDateTime: Boolean;
    function IsDouble: Boolean;
    function IsFloat: Boolean;
    function IsInstance: Boolean;
    function IsInt64: Boolean;
    function IsInteger: Boolean;
    function IsInterface: Boolean;
    function IsNumeric: Boolean;
    function IsPointer: Boolean;
    function IsShortInt: Boolean;
    function IsSingle: Boolean;
    function IsSmallInt: Boolean;
    function IsString: Boolean;
    function IsTime: Boolean;
    function IsUInt64: Boolean;
    function IsVariant: Boolean;
    function IsWord: Boolean;

    property RttiType: TRttiType read GetRttiType;
  end;

  PPropInfoExt = ^TPropInfoExt;
  TPropInfoExt = record
    PropType: PPTypeInfo;
    GetProc: Pointer;
    SetProc: Pointer;
    StoredProc: Pointer;
    Index: Integer;
    Default: Integer;
    NameIndex: SmallInt;
    NameLength : Byte;
    NameData : array[0..255] of Byte;
    function NameFld: TTypeInfoFieldAccessor; inline;
    function Tail: PPropInfoExt; inline;
  end;

  TRttiPropertyExtension = class(TRttiInstanceProperty)
  private class var
    FRegister: TDictionary<TPair<PTypeInfo, string>, TRttiPropertyExtension>;
    FPatchedClasses: TDictionary<TClass, TClass>;
  private
    FPropInfo: TPropInfoExt;
    FGetter: TFunc<Pointer, TValue>;
    FSetter: TProc<Pointer, TValue>;

    function GetIsReadableStub: Boolean;
    function GetIsWritableStub: Boolean;
    function DoGetValueStub(Instance: Pointer): TValue;
    procedure DoSetValueStub(Instance: Pointer; const AValue: TValue);
    function GetPropInfoStub: PPropInfo;
  protected
    class procedure InitVirtualMethodTable;

    function GetIsReadable: Boolean; virtual;
    function GetIsWritable: Boolean; virtual;
    function DoGetValue(Instance: Pointer): TValue; virtual;
    procedure DoSetValue(Instance: Pointer; const AValue: TValue); virtual;
    function GetPropInfo: PPropInfo; virtual;
  public
    class constructor Create;
    class destructor Destroy;

    constructor Create(AParent: PTypeInfo; const AName: string; APropertyType: PTypeInfo);

    class function FindByName(AParent: TRttiType; const APropertyName: string): TRttiPropertyExtension; overload;
    class function FindByName(const AFullPropertyName: string): TRttiPropertyExtension; overload;

    property Getter: TFunc<Pointer, TValue> read FGetter write FGetter;
    property Setter: TProc<Pointer, TValue> read FSetter write FSetter;
  end;

  TArrayHelper = class
  public
    class function Concat<T>(const AArrays: array of TArray<T>): TArray<T>; static;
    class function Create<T>(const a : T; const b : T): TArray<T>; static;
  end;

type
  /// <summary>
  ///   Provides interface type helper.
  /// </summary>
  /// <remarks>
  ///   TInterfaceHelper borrows the code from the answer to this StackOverflow question:
  ///   <see href="https://stackoverflow.com/questions/39584234/how-to-obtain-rtti-from-an-interface-reference-in-delphi" />
  /// </remarks>
  TInterfaceHelper = record
  strict private type
    TInterfaceTypes = TDictionary<TGUID, TRttiInterfaceType>;
  strict private
    class var FInterfaceTypes: TInterfaceTypes;
    class var FCached: Boolean;  // Boolean in Delphi is atomic
    class var FCaching: Boolean;
    class constructor Create;
    class destructor Destroy;
    class procedure CacheIfNotCachedAndWaitFinish; static;
    class procedure WaitIfCaching; static;
  public

    /// <summary>
    ///   Refreshes the cached RTTI interface types in a background thread (eg.
    ///   when new package is loaded).
    /// </summary>
    /// <remarks>
    ///   RefreshCache is called at program initialization automatically by the
    ///   class constructor. It may also be called as needed when a package is
    ///   loaded. The purpose of the cache is to speed up querying a given
    ///   interface type inside GetType method.
    /// </remarks>
    class procedure RefreshCache; static;

    /// <summary>
    ///   Obtains the RTTI interface type object of the specified interface.
    /// </summary>
    class function GetType(const AIntf: IInterface): TRttiInterfaceType; overload; static;

    /// <summary>
    ///   Obtains the RTTI interface type object of the specified interface GUID.
    /// </summary>
    class function GetType(const AGuid: TGUID): TRttiInterfaceType; overload; static;

    /// <summary>
    ///   Obtains the RTTI interface type object of the specified TValue-boxed interface.
    /// </summary>
    class function GetType(const AIntfInTValue: TValue): TRttiInterfaceType; overload; static;

    /// <summary>
    ///   Obtains the name of the interface type.
    /// </summary>
    class function GetTypeName(const AIntf: IInterface): string; overload; static;

    /// <summary>
    ///   Obtains the name of the interface type identified by a GUID.
    /// </summary>
    class function GetTypeName(const AGuid: TGUID): string; overload; static;

    /// <summary>
    ///   Obtains the qualified name of the interface type. A qualified name
    ///   includes the unit name separated by dot.
    /// </summary>
    class function GetQualifiedName(const AIntf: IInterface): string; overload; static;

    /// <summary>
    ///   Obtains the qualified name of the interface type identified by a
    ///   GUID. A qualified name includes the unit name separated by dot.
    /// </summary>
    class function GetQualifiedName(const AGuid: TGUID): string; overload; static;

    /// <summary>
    ///   Obtains a list of RTTI objects for all the methods that are members of the specified
    ///   interface.
    /// </summary>
    class function GetMethods(const AIntf: IInterface): TArray<TRttiMethod>; static;

    /// <summary>
    ///   Returns an RTTI object for the interface method with the
    ///   specified name.
    /// </summary>
    class function GetMethod(const AIntf: IInterface; const AMethodName: string): TRttiMethod; static;

    /// <summary>
    ///   Performs a call to the described method.
    /// </summary>
    class function InvokeMethod(const AIntf: IInterface; const AMethodName: string; const Args: array of TValue): TValue; overload; static;

    /// <summary>
    ///   Performs a call to the described method.
    /// </summary>
    class function InvokeMethod(const AIntfInTValue: TValue; const AMethodName: string; const Args: array of TValue): TValue; overload; static;
  end;

type
  /// <summary>
  ///   Throws when the method with the specified name is not found.
  /// </summary>
  EMethodNotFound = class(Exception)
  public
    constructor Create(const AMethodName: string);
  end;

function FindType(const AName: string; out AType: TRttiType): Boolean; overload;
function FindType(const AGuid: TGUID; out AType: TRttiType): Boolean; overload;
function GetRttiType(AClass: TClass): TRttiType; overload;
function GetRttiType(ATypeInfo: PTypeInfo): TRttiType; overload;
function GetRttiTypes: TArray<TRttiType>;
function IsClassCovariantTo(AThisClass, AOtherClass: TClass): Boolean;
function IsTypeCovariantTo(AThisType, AOtherType: PTypeInfo): Boolean;
function TryGetRttiType(AClass: TClass; out AType: TRttiType): Boolean; overload;
function TryGetRttiType(ATypeInfo: PTypeInfo; out AType: TRttiType): Boolean; overload;
function CompareValue(const Left, Right: TValue): Integer;
function SameValue(const Left, Right: TValue): Boolean;
function StripUnitName(const AName: string): string;
function Supports(const AInstance: TValue; const IID: TGUID; out AIntf): Boolean; overload;

var
  Context: TRttiContext;

implementation

uses
  System.Classes,
  System.Generics.Defaults,
  System.Math,
  System.StrUtils,
  System.SyncObjs,
  System.SysConst;

var
  Enumerations: TDictionary<PTypeInfo, TStrings>;

class function TAttributeUtils.ContainsAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): Boolean;
begin
  Result := FindAttribute(AAttributes,AAttributeClass) <> nil;
end;

class function TAttributeUtils.FindAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): TCustomAttribute;
begin
  Result := nil;

  for var LAttribute in AAttributes do
    if LAttribute.ClassType = AAttributeClass then Exit(LAttribute);
end;

class function TAttributeUtils.FindAttribute(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass; var AAttribute: TCustomAttribute; const AStartIndex: Integer = 0): Integer;
begin
  Result := -1;
  AAttribute := nil;

  for var I := AStartIndex to Length(AAttributes) -1 do
  begin
    if AAttributes[I].ClassType = AAttributeClass then
    begin
      AAttribute := AAttributes[I];
      Exit(I);
    end;
  end;
end;

class function TAttributeUtils.FindAttributes(const AAttributes: TArray<TCustomAttribute>; const AAttributeClass: TCustomAttributeClass): TArray<TCustomAttribute>;
begin
  SetLength(Result, 0);

  var I := 0;
  for var LAttribute in AAttributes do
  begin
    if LAttribute.ClassType = AAttributeClass then
    begin
      SetLength(Result, I + 1);
      Result[I] := LAttribute;
      Inc(I);
    end;
  end;
end;

// #13Hello#10World  will be encoded as a string "#13'Hello'#10'World'", where whitespace appears as a string.
class function TStrUtils.EncodeWhitespace(const AStr: string): string;
const
  sDelimiter: array[Boolean] of string = (#39, ''); // #39 is single quote.
begin
  Result := '';
  var LIsLastWhitespace := True;

  for var I := 1 to Length(AStr) do
  begin
    if AStr[I] < #32 then
    begin
      Result := Result + sDelimiter[LIsLastWhitespace] + '#' + IntToStr(Ord(AStr[I]));
      LIsLastWhitespace := True;
    end else
    begin
      Result := Result + sDelimiter[not LIsLastWhitespace] + AStr[I];
      LIsLastWhitespace := False;
    end;
  end;

  Result := Result + sDelimiter[LIsLastWhitespace];
end;

class function TStrUtils.Join(const AValues: TArray<string>; const ADelimiter: string): string;
begin
  Result := '';

  for var str in AValues do
  begin
    if Result <> '' then
      Result := Result + ADelimiter;

    Result := Result + str;
  end;
end;

class function TStrUtils.PadString(const AStr: string; const ATotalLength: Integer; const APadLeft: Boolean = True; APadChar: Char = ' '): string;
begin
  Result := AStr;

  while Length(Result) < ATotalLength do
  begin
    if APadLeft then
      Result := APadChar + Result
    else
      Result := Result + APadChar;
  end;
end;

{$REGION 'Conversion functions'}
type
  TConvertFunc = function(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;

function ConvFail(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := False;
end;

function ConvStr2DynArray(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  var LStr := ASource.AsString;

  if StartsStr('[', LStr) and EndsStr(']', LStr) then
    LStr := Copy(LStr, 2, Length(LStr) - 2);

  var LValues := SplitString(LStr, ',');
  var I := Length(LValues);
  var LPtr := nil;
  DynArraySetLength(LPtr, ATarget, 1, @I);
  TValue.MakeWithoutCopy(@LPtr, ATarget, AResult);
  var LElemType := ATarget.TypeData.DynArrElType^;

  for var J := 0 to High(LValues) do
  begin
    var LVal_1 := TValue.FromString(LValues[J]);
    var LVal_2: TValue;

    if not LVal_1.TryConvert(LElemType, LVal_2) then
      Exit(False);

    AResult.SetArrayElement(J, LVal_2);
  end;

  Result := True;
end;

function ConvAny2Nullable(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LBuffer: array of Byte;
  LType: TRttiType;
  LValue: TValue;
begin
  Result := TryGetRttiType(ATarget, LType) and LType.IsGenericTypeOf('Nullable') and ASource.TryConvert(LType.GetGenericArguments[0].Handle, LValue);

  if Result then
  begin
    SetLength(LBuffer, LType.TypeSize);
    Move(LValue.GetReferenceToRawData^, LBuffer[0], LType.TypeSize - SizeOf(string));
    PString(@LBuffer[LType.TypeSize - SizeOf(string)])^ := DefaultTrueBoolStr;
    TValue.Make(LBuffer, LType.Handle, AResult);
    PString(@LBuffer[LType.TypeSize - SizeOf(string)])^ := '';
  end
end;

function ConvClass2Class(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := ASource.TryCast(ATarget, AResult);

  if not Result and IsTypeCovariantTo(ASource.TypeInfo, ATarget) then
  begin
    AResult := TValue.From(ASource.AsObject, GetTypeData(ATarget).ClassType);
    Result := True;
  end;
end;

function ConvClass2Enum(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := ATarget = TypeInfo(Boolean);
  if Result then AResult := ASource.AsObject <> nil;
end;

function ConvEnum2Class(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LType: TRttiType;
  LStrings: TStrings;
begin
  Result := TryGetRttiType(ATarget, LType) and LType.AsInstance.MetaclassType.InheritsFrom(TStrings);

  if Result then
  begin
    if not Enumerations.TryGetValue(ASource.TypeInfo, LStrings) then
    begin
      LStrings := TStringList.Create;

      with TRttiEnumerationType(ASource.RttiType) do
      begin
        for var I := MinValue to MaxValue do
          LStrings.Add(GetEnumName(Handle, I));
      end;

      Enumerations.Add(ASource.TypeInfo, LStrings);
    end;

    AResult := TValue.From(LStrings, TStrings);
    Result := True;
  end;
end;

function ConvFloat2Ord(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := Frac(ASource.AsExtended) = 0;

  if Result then
    AResult := TValue.FromOrdinal(ATarget, Trunc(ASource.AsExtended));
end;

function ConvFloat2Str(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LValue: TValue;
begin
  if ASource.TypeInfo = TypeInfo(TDate) then
    LValue := DateToStr(ASource.AsExtended)
  else if ASource.TypeInfo = TypeInfo(TDateTime) then
    LValue := DateTimeToStr(ASource.AsExtended)
  else if ASource.TypeInfo = TypeInfo(TTime) then
    LValue := TimeToStr(ASource.AsExtended)
  else
    LValue := FloatToStr(ASource.AsExtended);

  Result := LValue.TryCast(ATarget, AResult);
end;

function ConvIntf2Class(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := ConvClass2Class(ASource.AsInterface as TObject, ATarget, AResult);
end;

function ConvIntf2Intf(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := ASource.TryCast(ATarget, AResult);

  if not Result then
  begin
    if IsTypeCovariantTo(ASource.TypeInfo, ATarget) then
    begin
      AResult := TValue.From(ASource.GetReferenceToRawData, ATarget);
      Result := True;
    end else
    begin
      var LType: TRttiType;
      var LMethod: TRttiMethod;
      if TryGetRttiType(ASource.TypeInfo, LType) and (GetTypeName(ATarget) = 'IList') and LType.IsGenericTypeOf('IList') and LType.TryGetMethod('AsList', LMethod) then
      begin
        var LInterface := LMethod.Invoke(ASource, []).AsInterface;
        AResult := TValue.From(@LInterface, ATarget);
        Result := True;
      end;
    end;
  end;
end;

function ConvNullable2Any(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LType: TRttiType;
begin
  Result := TryGetRttiType(ASource.TypeInfo, LType) and LType.IsGenericTypeOf('Nullable');

  if Result then
  begin
    var LValue := TValue.From(ASource.GetReferenceToRawData, LType.GetGenericArguments[0].Handle);
    Result := LValue.TryConvert(ATarget, AResult);
  end
end;

function ConvOrd2Float(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  AResult := TValue.FromFloat(ATarget, ASource.AsOrdinal);
  Result := True;
end;

function ConvOrd2Ord(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  AResult := TValue.FromOrdinal(ATarget, ASource.AsOrdinal);
  Result := True;
end;

function ConvOrd2Str(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LValue: TValue;
begin
  LValue := ASource.ToString;
  Result := LValue.TryCast(ATarget, AResult);
end;

function ConvRec2Meth(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  if ASource.TypeInfo = TypeInfo(TMethod) then
  begin
    AResult := TValue.From(ASource.GetReferenceToRawData, ATarget);
    Result := True;
  end else
  begin
    Result := ConvNullable2Any(ASource, ATarget, AResult);
  end;
end;

function ConvSet2Class(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LType: TRttiType;
  LStrings: TStrings;
begin
  Result := TryGetRttiType(ATarget, LType) and LType.AsInstance.MetaclassType.InheritsFrom(TStrings);

  if Result then
  begin
    var LTypeData := GetTypeData(ASource.TypeInfo);

    if not Enumerations.TryGetValue(LTypeData.CompType^, LStrings) then
    begin
      LStrings := TStringList.Create;

      with TRttiEnumerationType(TRttiSetType(ASource.RttiType).ElementType) do
      begin
        for var I := MinValue to MaxValue do
          LStrings.Add(GetEnumName(Handle, I));
      end;

      Enumerations.Add(LTypeData.CompType^, LStrings);
    end;

    AResult := TValue.From(LStrings, TStrings);
  end
end;

function ConvStr2Enum(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  AResult := TValue.FromOrdinal(ATarget, GetEnumValue(ATarget, ASource.AsString));
  Result := True;
end;

function ConvStr2Float(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
var
  LFormatSettings : TFormatSettings;
begin
  LFormatSettings.DecimalSeparator := '.';
  var LValue := StringReplace(ASource.AsString, ',', '.', [rfReplaceAll]);

  if ATarget = TypeInfo(TDate) then
    AResult := TValue.From<TDate>(StrToDateDef(LValue, 0))
  else if ATarget = TypeInfo(TDateTime) then
    AResult := TValue.From<TDateTime>(StrToDateTimeDef(LValue, 0))
  else if ATarget = TypeInfo(TTime) then
    AResult := TValue.From<TTime>(StrToTimeDef(LValue, 0))
  else
    AResult := TValue.FromFloat(ATarget, StrToFloatDef(LValue, 0, LFormatSettings));

  Result := True;
end;

function ConvStr2Ord(const ASource: TValue; ATarget: PTypeInfo; out AResult: TValue): Boolean;
begin
  AResult := TValue.FromOrdinal(ATarget, StrToInt64Def(ASource.AsString, 0));
  Result := True;
end;

{$ENDREGION}

{$REGION 'Conversions'}
const
  Conversions: array[TTypeKind, TTypeKind] of TConvertFunc = (
    // tkUnknown
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail , ConvFail
    ),
    // tkInteger
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Float, ConvOrd2Str,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvOrd2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvOrd2Str, ConvFail, ConvFail, ConvFail , ConvFail
    ),
    // tkChar
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Float, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvOrd2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkEnumeration
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Float, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvEnum2Class, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvOrd2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvOrd2Str, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkFloat
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFloat2Ord, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFloat2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFloat2Str, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkString
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkSet
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvSet2Class, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkClass
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvClass2Enum, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvClass2Class, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkMethod
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkWChar
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Float, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvOrd2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkLString
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkWString
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkVariant
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkArray
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkRecord
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvRec2Meth, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkInterface
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvIntf2Class, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvIntf2Intf, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkInt64
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Ord, ConvOrd2Float, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvOrd2Ord, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvOrd2Str, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkDynArray
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkUString
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvStr2Ord, ConvFail, ConvStr2Enum, ConvStr2Float, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvStr2Ord, ConvStr2DynArray,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkClassRef
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkPointer
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    ),
    // tkProcedure
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    )
	,
    // tkMRecord
    (
      // tkUnknown, tkInteger, tkChar, tkEnumeration, tkFloat, tkString,
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkSet, tkClass, tkMethod, tkWChar, tkLString, tkWString
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkVariant, tkArray, tkRecord, tkInterface, tkInt64, tkDynArray
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail, ConvFail,
      // tkUString, tkClassRef, tkPointer, tkProcedure, tkMRecord
      ConvFail, ConvFail, ConvFail, ConvFail, ConvFail
    )
  );
{$ENDREGION}

function ExtractGenericArguments(ATypeInfo: PTypeInfo): string;
begin
  var LStr := UTF8ToString(ATypeInfo.Name);
  var I := LStr.IndexOf('<');

  if I > -1 then
    Result := LStr.SubString(Succ(I), LStr.Length - (Succ(I) + 1))
  else
    Result := string.Empty;
end;

function FindType(const AName: string; out AType: TRttiType): Boolean;
begin
  AType := Context.FindType(AName);

  if not Assigned(AType) then
  begin
    for var LType in Context.GetTypes do
    begin
      if SameText(LType.Name, AName) then
      begin
        AType := LType;
        Break;
      end;
    end;
  end;
  Result := Assigned(AType);
end;

function FindType(const AGuid: TGUID; out AType: TRttiType): Boolean;
begin
  AType := nil;

  for var LType in Context.GetTypes do
  begin
    if (LType is TRttiInterfaceType) and IsEqualGUID(TRttiInterfaceType(LType).GUID, AGuid) then
    begin
      AType := LType;
      Break;
    end;
  end;

  Result := Assigned(AType);
end;

function GetRttiType(AClass: TClass): TRttiType;
begin
  Result := Context.GetType(AClass);
end;

function GetRttiType(ATypeInfo: PTypeInfo): TRttiType;
begin
  Result := Context.GetType(ATypeInfo);
end;

function GetRttiTypes: TArray<TRttiType>;
begin
  Result := Context.GetTypes();
end;

function IsClassCovariantTo(AThisClass, AOtherClass: TClass): Boolean;
begin
  var LType := Context.GetType(AThisClass);
  Result := Assigned(LType) and LType.IsCovariantTo(AOtherClass.ClassInfo);
end;

function IsTypeCovariantTo(AThisType, AOtherType: PTypeInfo): Boolean;
begin
  var LType := Context.GetType(AThisType);
  Result := Assigned(LType) and LType.IsCovariantTo(AOtherType);
end;

function MergeStrings(AValues: TStringDynArray; const ADelimiter: string): string;
begin
  Result := '';

  for var I := Low(AValues) to High(AValues) do
  begin
    if I = 0 then
      Result := AValues[I]
    else
      Result := Result + ADelimiter + AValues[I];
  end;
end;

function TryGetRttiType(AClass: TClass; out AType: TRttiType): Boolean; overload;
begin
  AType := Context.GetType(AClass);
  Result := Assigned(AType);
end;

function TryGetRttiType(ATypeInfo: PTypeInfo; out AType: TRttiType): Boolean; overload;
begin
  AType := Context.GetType(ATypeInfo);
  Result := Assigned(AType);
end;

function StripUnitName(const AName: string): string;
begin
  Result := ReplaceText(AName, 'System.', '');
end;

function Supports(const AInstance: TValue; const IID: TGUID; out AIntf): Boolean;
begin
  if AInstance.Kind in [tkClass, tkInterface] then
    Result := Supports(AInstance.ToObject, IID, AIntf)
  else
    Result := False;
end;

function CompareValue(const Left, Right: TValue): Integer;
begin
  if Left.IsOrdinal and Right.IsOrdinal then
    Result := System.Math.CompareValue(Left.AsOrdinal, Right.AsOrdinal)
  else if Left.IsFloat and Right.IsFloat then
    Result := System.Math.CompareValue(Left.AsFloat, Right.AsFloat)
  else if Left.IsString and Right.IsString then
    Result := System.SysUtils.CompareStr(Left.AsString, Right.AsString)
  else
    Result := 0;
end;

function SameValue(const Left, Right: TValue): Boolean;
begin
  if Left.IsNumeric and Right.IsNumeric then
  begin
    if Left.IsOrdinal then
    begin
      if Right.IsOrdinal then
        Result := Left.AsOrdinal = Right.AsOrdinal
      else if Right.IsSingle then
        Result := System.Math.SameValue(Left.AsOrdinal, Right.AsSingle)
      else if Right.IsDouble then
        Result := System.Math.SameValue(Left.AsOrdinal, Right.AsDouble)
      else
        Result := System.Math.SameValue(Left.AsOrdinal, Right.AsExtended);
    end else if Left.IsSingle then
    begin
      if Right.IsOrdinal then
        Result := System.Math.SameValue(Left.AsSingle, Right.AsOrdinal)
      else if Right.IsSingle then
        Result := System.Math.SameValue(Left.AsSingle, Right.AsSingle)
      else if Right.IsDouble then
        Result := System.Math.SameValue(Left.AsSingle, Right.AsDouble)
      else
        Result := System.Math.SameValue(Left.AsSingle, Right.AsExtended);
    end  else if Left.IsDouble then
    begin
      if Right.IsOrdinal then
        Result := System.Math.SameValue(Left.AsDouble, Right.AsOrdinal)
      else if Right.IsSingle then
        Result := System.Math.SameValue(Left.AsDouble, Right.AsSingle)
      else if Right.IsDouble then
        Result := System.Math.SameValue(Left.AsDouble, Right.AsDouble)
      else
        Result := System.Math.SameValue(Left.AsDouble, Right.AsExtended);
    end else
    begin
      if Right.IsOrdinal then
        Result := System.Math.SameValue(Left.AsExtended, Right.AsOrdinal)
      else if Right.IsSingle then
        Result := System.Math.SameValue(Left.AsExtended, Right.AsSingle)
      else if Right.IsDouble then
        Result := System.Math.SameValue(Left.AsExtended, Right.AsDouble)
      else
        Result := System.Math.SameValue(Left.AsExtended, Right.AsExtended);
    end;
  end
  else if Left.IsString and Right.IsString then
    Result := Left.AsString = Right.AsString
  else if Left.IsClass and Right.IsClass then
    Result := Left.AsClass = Right.AsClass
  else if Left.IsObject and Right.IsObject then
    Result := Left.AsObject = Right.AsObject
  else if Left.IsPointer and Right.IsPointer then
    Result := Left.AsPointer = Right.AsPointer
  else if Left.IsVariant and Right.IsVariant then
    Result := Left.AsVariant = Right.AsVariant
  else if Left.TypeInfo = Right.TypeInfo then
    Result := Left.AsPointer = Right.AsPointer
  else
    Result := False;
end;

class function TArrayHelper.Concat<T>(const AArrays: array of TArray<T>): TArray<T>;
begin
  var LLength := 0;

  for var I := 0 to High(AArrays) do
    Inc(LLength, Length(AArrays[I]));

  SetLength(Result, LLength);
  var LIndex := 0;

  for var I := 0 to High(AArrays) do
  begin
    for var K := 0 to High(AArrays[I]) do
    begin
      Result[LIndex] := AArrays[I][K];
      Inc(LIndex);
    end;
  end;
end;

class function TArrayHelper.Create<T>(const a : T; const b : T): TArray<T>;
begin
  SetLength(Result,2);
  Result[0] := a;
  Result[1] := b;
end;

function TObjectHelper.GetField(const AName: string): TRttiField;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetField(AName);
end;

function TObjectHelper.GetFields: TArray<TRttiField>;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetFields;
end;

function TObjectHelper.GetMember(const AName: string): TRttiMember;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetMember(AName);
end;

function TObjectHelper.GetMethod(const AName: string): TRttiMethod;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetMethod(AName);
end;

function TObjectHelper.GetMethod(ACodeAddress: Pointer): TRttiMethod;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetMethod(ACodeAddress);
end;

function TObjectHelper.GetMethods: TArray<TRttiMethod>;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetMethods();
end;

function TObjectHelper.GetProperties: TArray<TRttiProperty>;
begin
  Result := nil;
  var LType: TRttiType;

  if TryGetType(LType) then
    Result := LType.GetProperties();
end;

function TObjectHelper.GetProperty(const AName: string): TRttiProperty;
var
  LType: TRttiType;
begin
  if TryGetType(LType) then
    Result := LType.GetProperty(AName)
  else
    Result := nil;
end;

function TObjectHelper.GetType: TRttiType;
begin
  TryGetType(Result);
end;

function TObjectHelper.HasField(const AName: string): Boolean;
begin
  Result := GetField(AName) <> nil;
end;

function TObjectHelper.HasMethod(const AName: string): Boolean;
begin
  Result := GetMethod(AName) <> nil;
end;

function TObjectHelper.HasProperty(const AName: string): Boolean;
begin
  Result := GetProperty(AName) <> nil;
end;

function TObjectHelper.TryGetField(const AName: string; out AField: TRttiField): Boolean;
begin
  AField := GetField(AName);
  Result := Assigned(AField);
end;

function TObjectHelper.TryGetMember(const AName: string; out AMember: TRttiMember): Boolean;
begin
  AMember := GetMember(AName);
  Result := Assigned(AMember);
end;

function TObjectHelper.TryGetMethod(ACodeAddress: Pointer; out AMethod: TRttiMethod): Boolean;
begin
  AMethod := GetMethod(ACodeAddress);
  Result := Assigned(AMethod);
end;

function TObjectHelper.TryGetMethod(const AName: string; out AMethod: TRttiMethod): Boolean;
begin
  AMethod := GetMethod(AName);
  Result := Assigned(AMethod);
end;

function TObjectHelper.TryGetProperty(const AName: string; out AProperty: TRttiProperty): Boolean;
begin
  AProperty := GetProperty(AName);
  Result := Assigned(AProperty);
end;

function TObjectHelper.TryGetType(out AType: TRttiType): Boolean;
begin
  Result := False;

  if Assigned(Self) then
  begin
    AType := Context.GetType(ClassInfo);
    Result := Assigned(AType);
  end;
end;

function TRttiFieldHelper.TryGetValue(AInstance: Pointer; out AValue: TValue): Boolean;
begin
  try
    AValue := GetValue(AInstance);
    Result := True;
  except
    AValue := TValue.Empty;
    Result := False;
  end;
end;

function TRttiInvokableTypeHelper.GetParameterCount: Integer;
begin
  Result := Length(GetParameters());
end;

function TRttiMemberHelper.GetMemberIsReadable: Boolean;
begin
  if Self is TRttiField then
    Result := True
  else if Self is TRttiProperty then
    Result := TRttiProperty(Self).IsReadable
  else if Self is TRttiMethod then
    Result := True
  else
    Result := True;
end;

function TRttiMemberHelper.GetMemberIsWritable: Boolean;
begin
  if Self is TRttiField then
    Result := True
  else if Self is TRttiProperty then
    Result := TRttiProperty(Self).IsWritable
  else
    Result := False;
end;

function TRttiMemberHelper.GetMemberRttiType: TRttiType;
begin
  Result := nil;

  if Self is TRttiField then
    Result := TRttiField(Self).FieldType
  else if Self is TRttiProperty then
    Result := TRttiProperty(Self).PropertyType
  else if Self is TRttiMethod then
    Result := TRttiMethod(Self).ReturnType;
end;

function TRttiMethodHelper.Format(const AArgs: array of TValue; ASkipSelf: Boolean = True): string;
begin
  Result := StripUnitName(Parent.Name) + '.' + Name + '(';

  if ASkipSelf then
    if Length(AArgs) > 1 then Result := Result + TValue.ToString(AArgs[1])
  else
    Result := Result + TValue.ToString(AArgs[0]);

  Result := Result + ')';
end;

function TRttiMethodHelper.GetParameterCount: Integer;
begin
  Result := Length(GetParameters());
end;

function TRttiObjectHelper.FindAttribute<T>: T;
begin
  Result := Default(T);

  for var LAttr in GetAttributes do
  begin
    if LAttr.InheritsFrom(T) then
    begin
      Result := T(LAttr);
      Break;
    end;
  end;
end;

function TRttiObjectHelper.FindAttributes<T>: TArray<T>;
begin
  SetLength(Result, 0);

  for var LAttr in GetAttributes do
  begin
    if LAttr.InheritsFrom(T) then
    begin
      SetLength(Result, Length(Result) + 1);
      Result[High(Result)] := T(LAttr);
    end;
  end;
end;

function TRttiObjectHelper.HasAttribute<T>: Boolean;
begin
  Result := FindAttribute<T> <> nil;
end;

function TRttiObjectHelper.TryGetAttribute<T>(out AAttribute: T): Boolean;
begin
  AAttribute := FindAttribute<T>;
  Result := Assigned(AAttribute);
end;

class function TRttiParameterHelper.Equals(const Left, Right: TArray<TRttiParameter>): Boolean;
begin
  Result := Length(Left) = Length(Right);

  if Result then
  begin
    for var I := Low(Left) to High(Left) do
    begin
      if Left[I].ParamType <> Right[I].ParamType then
      begin
        Result := False;
        Break;
      end;
    end
  end;
end;

function TRttiPropertyHelper.TryGetValue(AInstance: Pointer; out AValue: TValue): Boolean;
begin
  try
    if IsReadable then
    begin
      AValue := GetValue(AInstance);
      Result := True;
    end else
    begin
      Result := False;
    end;
  except
    AValue := TValue.Empty;
    Result := False;
  end;
end;

function TRttiPropertyHelper.TrySetValue(AInstance: Pointer; AValue: TValue): Boolean;
begin
  var LValue: TValue;
  Result := AValue.TryConvert(PropertyType.Handle, LValue);

  if Result then
    SetValue(AInstance, LValue);
end;

function TRttiTypeHelper.GetAsInterface: TRttiInterfaceType;
begin
  Result := Self as TRttiInterfaceType;
end;

function TRttiTypeHelper.FindAttributes<T>: TArray<T>;
begin
  SetLength(Result, 0);

  for var LAttribute in GetAttributes do
  begin
    if LAttribute.InheritsFrom(T) then
    begin
      SetLength(Result, Length(Result) + 1);
      Result[High(Result)] := T(LAttribute);
    end;
  end;

  if Assigned(BaseType) then
  begin
    for var LAttribute in BaseType.FindAttributes<T> do
    begin
      if LAttribute.InheritsFrom(T) then
      begin
        SetLength(Result, Length(Result) + 1);
        Result[High(Result)] := T(LAttribute);
      end;
    end;
  end;
end;

function TRttiTypeHelper.GetGenericArguments: TArray<TRttiType>;
begin
  var LArgs := SplitString(ExtractGenericArguments(Handle), ',');
  SetLength(Result, Length(LArgs));
  
  for var I := 0 to Pred(Length(LArgs)) do
    FindType(LArgs[I], Result[I]);
end;

function TRttiTypeHelper.GetGenericTypeDefinition(const AIncludeUnitName: Boolean = True): string;
begin
  var LArgs := SplitString(ExtractGenericArguments(Handle), ',');

  for var I := Low(LArgs) to High(LArgs) do
  begin
    // naive implementation - but will work in most cases
    if (I = 0) and (Length(LArgs) = 1) then
      LArgs[I] := 'T'
    else
      LArgs[I] := 'T' + IntToStr(Succ(I));
  end;

  var LStr: string;
  if IsPublicType and AIncludeUnitName then
    LStr := QualifiedName
  else
    LStr := Name;

  Result := LStr.SubString(0, LStr.IndexOf('<') + 1) + MergeStrings(LArgs, ',');
  if Result <> EmptyStr then
    Result := Result + '>';
end;

function TRttiTypeHelper.GetIsInterface: Boolean;
begin
  Result := Self is TRttiInterfaceType;
end;

function TRttiTypeHelper.GetMember(const AName: string): TRttiMember;
var
  LProperty: TRttiProperty;
  LField: TRttiField;
  LMethod: TRttiMethod;
begin
  if TryGetProperty(AName, LProperty) then
    Result := LProperty
  else if TryGetField(AName, LField) then
    Result := LField
  else if TryGetMethod(AName, LMethod) then
    Result := LMethod
  else
    Result := nil;
end;

function TRttiTypeHelper.GetMethod(ACodeAddress: Pointer): TRttiMethod;
begin
  Result := nil;

  for var LMethod in GetMethods do
  begin
    if LMethod.CodeAddress = ACodeAddress then
    begin
      Result := LMethod;
      Break;
    end;
  end;
end;

function TRttiTypeHelper.GetMethodCount: Integer;
begin
  Result := Length(GetMethods);
end;

function TRttiTypeHelper.GetProperty(const AName: string): TRttiProperty;
begin
  Result := inherited GetProperty(AName);

  if not Assigned(Result) then
    Result := TRttiPropertyExtension.FindByName(Self, AName);
end;

function TRttiTypeHelper.GetStandardConstructor: TRttiMethod;
begin
  Result := nil;

  for var LMethod in GetMethods do
  begin
    if LMethod.IsConstructor and (LMethod.ParameterCount = 0) then
    begin
      Result := LMethod;
      Break;
    end;
  end;
end;

function TRttiTypeHelper.InheritsFrom(OtherType: PTypeInfo): Boolean;
begin
  Result := Handle = OtherType;

  if not Result then
  begin
    var LType := BaseType;

    while Assigned(LType) and not Result do
    begin
      Result := LType.Handle = OtherType;
      LType := LType.BaseType;
    end;
  end;
end;

function TRttiTypeHelper.IsCovariantTo(AOtherType: PTypeInfo): Boolean;
begin
  Result := False;
  var LRttiType := Context.GetType(AOtherType);

  if Assigned(LRttiType) and IsGenericTypeDefinition then
  begin
    if SameText(GetGenericTypeDefinition, LRttiType.GetGenericTypeDefinition) or SameText(GetGenericTypeDefinition(False), LRttiType.GetGenericTypeDefinition(False)) then
    begin
      Result := True;
      var LArgs := GetGenericArguments;
      var LOtherArgs := LRttiType.GetGenericArguments;

      for var I := Low(LArgs) to High(LArgs) do
      begin
        if LArgs[I].IsInterface and LArgs[I].IsInterface and LArgs[I].InheritsFrom(LOtherArgs[I].Handle) then
          Continue;

        if LArgs[I].IsInstance and LOtherArgs[I].IsInstance and LArgs[I].InheritsFrom(LOtherArgs[I].Handle) then
          Continue;

        Result := False;
        Break;
      end;
    end else
    begin
      if Assigned(BaseType) then
        Result := BaseType.IsCovariantTo(AOtherType);
    end;
  end else
  begin
    Result := InheritsFrom(AOtherType);
  end;
end;

function TRttiTypeHelper.IsCovariantTo(AOtherClass: TClass): Boolean;
begin
  Result := Assigned(AOtherClass) and IsCovariantTo(AOtherClass.ClassInfo);
end;

function TRttiTypeHelper.IsGenericTypeDefinition: Boolean;
begin
  Result := Length(GetGenericArguments) > 0;

  if not Result and Assigned(BaseType) then
  begin
    Result := BaseType.IsGenericTypeDefinition;
  end;
end;

function TRttiTypeHelper.IsGenericTypeOf(const ABaseTypeName: string): Boolean;
begin
  var s := Name;
  Result := (s.SubString(0, Succ(ABaseTypeName.Length)) = (ABaseTypeName + '<')) and (s.SubString(s.Length-1, 1) = '>');
end;

function TRttiTypeHelper.IsInheritedFrom(const AOtherTypeName: string): Boolean;
begin
  Result := SameText(Name, AOtherTypeName) or (IsPublicType and SameText(QualifiedName, AOtherTypeName));

  if not Result then
  begin
    var LType := BaseType;

    while Assigned(LType) and not Result do
    begin
      Result := SameText(LType.Name, AOtherTypeName) or (LType.IsPublicType and SameText(LType.QualifiedName, AOtherTypeName));
      LType := LType.BaseType;
    end;
  end;
end;

function TRttiTypeHelper.IsInheritedFrom(AOtherType: TRttiType): Boolean;
begin
  Result := Self.Handle = AOtherType.Handle;

  if not Result then
  begin
    var LType := BaseType;

    while Assigned(LType) and not Result do
    begin
      Result := LType.Handle = AOtherType.Handle;
      LType := LType.BaseType;
    end;
  end;
end;

function TRttiTypeHelper.MakeGenericType(const ATypeArguments: array of PTypeInfo): TRttiType;
begin
  Result := nil;

  if IsPublicType then
  begin
    var LArgs := SplitString(ExtractGenericArguments(Handle), ',');

    for var I := Low(LArgs) to High(LArgs) do
      LArgs[I] := Context.GetType(ATypeArguments[I]).QualifiedName;

    var LStr := QualifiedName.SubString(0, QualifiedName.IndexOf('<') + 1) + MergeStrings(LArgs, ',') + '>';
    Result := Context.FindType(LStr);
  end
end;

function TRttiTypeHelper.TryGetConstructor(out AMethod: TRttiMethod): Boolean;
begin
  Result := False;
  var LRttiType := Self;

  while (LRttiType <> nil) and (LRttiType.Handle <> TypeInfo(TObject)) do
  begin
    var LMethods := LRttiType.GetDeclaredMethods;

    for var Lethod in LMethods do
    begin
      if Lethod.IsConstructor and (Length(Lethod.GetParameters) = 0) then
      begin
        AMethod := Lethod;
        Exit(True);
      end;
    end;

    LRttiType := LRttiType.BaseType;
  end;
end;

function TRttiTypeHelper.TryGetDestructor(out AMethod: TRttiMethod): Boolean;
begin
  Result := False;
  var LMethods := GetDeclaredMethods;

  for var LMethod in LMethods do
  begin
    if LMethod.IsDestructor then
    begin
      AMethod := LMethod;
      Exit(True);
    end;
  end;
end;

function TRttiTypeHelper.TryGetField(const AName: string; out AField: TRttiField): Boolean;
begin
  AField := GetField(AName);
  Result := Assigned(AField);
end;

function TRttiTypeHelper.TryGetMethod(ACodeAddress: Pointer; out AMethod: TRttiMethod): Boolean;
begin
  AMethod := GetMethod(ACodeAddress);
  Result := Assigned(AMethod);
end;

function TRttiTypeHelper.TryGetMember(const AName: string; out AMember: TRttiMember): Boolean;
begin
  AMember := GetMember(AName);
  Result := Assigned(AMember);
end;

function TRttiTypeHelper.TryGetMethod(const AName: string;
  out AMethod: TRttiMethod): Boolean;
begin
  AMethod := GetMethod(AName);
  Result := Assigned(AMethod);
end;

function TRttiTypeHelper.TryGetProperty(const AName: string; out AProperty: TRttiProperty): Boolean;
begin
  AProperty := GetProperty(AName);
  Result := Assigned(AProperty);
end;

function TRttiTypeHelper.TryGetStandardConstructor(out AMethod: TRttiMethod): Boolean;
begin
  AMethod := GetStandardConstructor();
  Result := Assigned(AMethod);
end;

function TValueHelper.AsByte: Byte;
begin
  Result := AsType<Byte>;
end;

function TValueHelper.AsCardinal: Cardinal;
begin
  Result := AsType<Cardinal>;
end;

function TValueHelper.AsCurrency: Currency;
begin
  Result := AsType<Currency>;
end;

function TValueHelper.AsDate: TDate;
begin
  Result := AsType<TDate>;
end;

function TValueHelper.AsDateTime: TDateTime;
begin
  Result := AsType<TDateTime>;
end;

function TValueHelper.AsDouble: Double;
begin
  Result := AsType<Double>;
end;

function TValueHelper.AsFloat: Extended;
begin
  Result := AsType<Extended>;
end;

function TValueHelper.AsPointer: Pointer;
begin
  if Kind in [tkClass, tkInterface] then
    Result := ToObject
  else
    Result := GetReferenceToRawData;
end;

function TValueHelper.AsShortInt: ShortInt;
begin
  Result := AsType<ShortInt>;
end;

function TValueHelper.AsSingle: Single;
begin
  Result := AsType<Single>;
end;

function TValueHelper.AsSmallInt: SmallInt;
begin
  Result := AsType<SmallInt>;
end;

function TValueHelper.AsTime: TTime;
begin
  Result := AsType<TTime>;
end;

function TValueHelper.AsUInt64: UInt64;
begin
  Result := AsType<UInt64>;
end;

function TValueHelper.AsWord: Word;
begin
  Result := AsType<Word>;
end;

class function TValueHelper.Equals(const Left, Right: TArray<TValue>): Boolean;
begin
  Result := Length(Left) = Length(Right);

  if Result then
  begin
    for var I := Low(Left) to High(Left) do
    begin
      if not SameValue(Left[I], Right[I]) then
      begin
        Result := False;
        Break;
      end;
    end
  end;
end;

class function TValueHelper.Equals<T>(const Left, Right: T): Boolean;
begin
  Result := TEqualityComparer<T>.Default.Equals(Left, Right);
end;

class function TValueHelper.From(ABuffer: Pointer; ATypeInfo: PTypeInfo): TValue;
begin
  TValue.Make(ABuffer, ATypeInfo, Result);
end;

class function TValueHelper.From(AValue: NativeInt; ATypeInfo: PTypeInfo): TValue;
begin
  TValue.Make(AValue, ATypeInfo, Result);
end;

class function TValueHelper.From(AObject: TObject; AClass: TClass): TValue;
begin
  TValue.Make(NativeInt(AObject), AClass.ClassInfo, Result);
end;

class function TValueHelper.FromBoolean(const AValue: Boolean): TValue;
begin
  Result := TValue.From<Boolean>(AValue);
end;

class function TValueHelper.FromFloat(ATypeInfo: PTypeInfo; AValue: Extended): TValue;
begin
  case GetTypeData(ATypeInfo).FloatType of
    ftSingle:
      Result := TValue.From<Single>(AValue);
    ftDouble:
      Result := TValue.From<Double>(AValue);
    ftExtended:
      Result := TValue.From<Extended>(AValue);
    ftComp:
      Result := TValue.From<Comp>(AValue);
    ftCurr:
      Result := TValue.From<Currency>(AValue);
  end;
end;

class function TValueHelper.FromString(const AValue: string): TValue;
begin
  Result := TValue.From<string>(AValue);
end;

class function TValueHelper.FromVarRec(const AValue: TVarRec): TValue;
begin
  case AValue.VType of
    vtInteger:
      Result := AValue.VInteger;
    vtBoolean:
      Result := AValue.VBoolean;
    vtExtended:
      Result := AValue.VExtended^;
    vtPointer:
      Result := TValue.From<Pointer>(AValue.VPointer);
    vtObject:
      Result := AValue.VObject;
    vtClass:
      Result := AValue.VClass;
    vtWideChar:
      Result := string(AValue.VWideChar);
    vtPWideChar:
      Result := string(AValue.VPWideChar);
    vtCurrency:
      Result := AValue.VCurrency^;
    vtVariant:
      Result := TValue.FromVariant(AValue.VVariant^);
    vtInterface:
      Result := TValue.From<IInterface>(IInterface(AValue.VInterface));
    vtWideString:
      Result := string(AValue.VWideString);
    vtInt64:
      Result := AValue.VInt64^;
    vtUnicodeString:
      Result := string(AValue.VUnicodeString);
  end;
end;

function TValueHelper.GetRttiType: TRttiType;
begin
  Result := Context.GetType(TypeInfo);
end;

function TValueHelper.IsBoolean: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Boolean);
end;

function TValueHelper.IsByte: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Byte);
end;

function TValueHelper.IsCardinal: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Cardinal);
{$IFNDEF CPUX64}
  Result := Result or (TypeInfo = System.TypeInfo(NativeUInt));
{$ENDIF}
end;

function TValueHelper.IsCurrency: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Currency);
end;

function TValueHelper.IsDate: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(TDate);
end;

function TValueHelper.IsDateTime: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(TDateTime);
end;

function TValueHelper.IsDouble: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Double);
end;

function TValueHelper.IsFloat: Boolean;
begin
  Result := Kind = tkFloat;
end;

function TValueHelper.IsInstance: Boolean;
begin
  Result := Kind in [tkClass, tkInterface];
end;

function TValueHelper.IsInt64: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Int64);
{$IFDEF CPUX64}
  Result := Result or (TypeInfo = System.TypeInfo(NativeInt));
{$ENDIF}
end;

function TValueHelper.IsInteger: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Integer);
{$IFNDEF CPUX64}
  Result := Result or (TypeInfo = System.TypeInfo(NativeInt));
{$ENDIF}
end;

function TValueHelper.IsInterface: Boolean;
begin
  Result := Assigned(TypeInfo) and (TypeInfo.Kind = tkInterface);
end;

function TValueHelper.IsNumeric: Boolean;
begin
  Result := Kind in [tkInteger, tkChar, tkEnumeration, tkFloat, tkWChar, tkInt64];
end;

function TValueHelper.IsPointer: Boolean;
begin
  Result := Kind = tkPointer;
end;

function TValueHelper.IsShortInt: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(ShortInt);
end;

function TValueHelper.IsSingle: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Single);
end;

function TValueHelper.IsSmallInt: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(SmallInt);
end;

function TValueHelper.IsString: Boolean;
begin
  Result := Kind in [tkChar, tkString, tkWChar, tkLString, tkWString, tkUString];
end;

function TValueHelper.IsTime: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(TTime);
end;

function TValueHelper.IsUInt64: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(UInt64);
{$IFDEF CPUX64}
  Result := Result or (TypeInfo = System.TypeInfo(NativeInt));
{$ENDIF}
end;

function TValueHelper.IsVariant: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Variant);
end;

function TValueHelper.IsWord: Boolean;
begin
  Result := TypeInfo = System.TypeInfo(Word);
end;

function TValueHelper.ToObject: TObject;
begin
  if IsInterface then
    Result := AsInterface as TObject
  else
    Result := AsObject;
end;

class function TValueHelper.ToString(const AValues: array of TValue): string;
begin
  Result := '';

  for var I := Low(AValues) to High(AValues) do
  begin
    if I > Low(AValues) then
      Result := Result + ', ';

    if AValues[I].IsString then
      Result := Result + '''' + TValue.ToString(AValues[I]) + ''''
    else
      Result := Result + TValue.ToString(AValues[I]);
  end;
end;

function TValueHelper.ToVarRec: TVarRec;
begin
  case Kind of
    tkInteger:
    begin
      Result.VType := vtInteger;
      Result.VInteger := AsInteger;
    end;

    tkEnumeration:
    begin
      if IsBoolean then
      begin
        Result.VType := vtBoolean;
        Result.VBoolean := AsBoolean;
      end else
      begin
        Result.VType := vtInteger;
        Result.VInteger := AsInteger;
      end;
    end;

    tkFloat:
    begin
      if IsCurrency then
      begin
        Result.VType := vtCurrency;
        Result.VCurrency := GetReferenceToRawData;
      end else
      begin
        Result.VType := vtExtended;
        Result.VExtended := GetReferenceToRawData;
      end;
    end;

    tkString, tkUString:
    begin
      Result.VType := vtUnicodeString;
      Result.VUnicodeString := Pointer(AsString);
    end;

    tkClass, tkInterface:
    begin
      Result.VType := vtUnicodeString;
      Result.VUnicodeString := Pointer(ToObject.ToString);
    end;
  end;
end;

class function TValueHelper.ToVarRecs(const AValues: array of TValue): TArray<TVarRec>;
begin
  SetLength(Result, Length(AValues));

  for var I := Low(AValues) to High(AValues) do
    Result[I] := AValues[I].ToVarRec;
end;

class function TValueHelper.ToString(const AValue: TValue): string;
begin
  case AValue.Kind of
    tkFloat:
      begin
        if AValue.IsDate then
          Result := DateToStr(AValue.AsDate)
        else if AValue.IsDateTime then
          Result := DateTimeToStr(AValue.AsDateTime)
        else if AValue.IsTime then
          Result := TimeToStr(AValue.AsTime)
        else
          Result := AValue.ToString;
      end;

    tkClass:
      begin
        var LObject := AValue.AsObject;
        Result := Format('%s($%x)', [StripUnitName(LObject.ClassName), NativeInt(LObject)]);
      end;

    tkInterface:
      begin
        var LInterface := AValue.AsInterface;
        var LObject := LInterface as TObject;
        Result := Format('%s($%x) as %s', [StripUnitName(LObject.ClassName), NativeInt(LInterface), StripUnitName(GetTypeName(AValue.TypeInfo))]);
      end
  else
    Result := AValue.ToString;
  end;
end;

function TValueHelper.TryConvert(ATypeInfo: PTypeInfo; out AResult: TValue): Boolean;
begin
  Result := False;

  if ATypeInfo = System.TypeInfo(TValue) then
  begin
    AResult:= Self;
    Exit(True);
  end;

  if Assigned(ATypeInfo) then
  begin
    Result := Conversions[Kind, ATypeInfo.Kind](Self, ATypeInfo, AResult);

    if not Result then
    begin
      case Kind of
        tkRecord: Result := ConvNullable2Any(Self, ATypeInfo, AResult);
      end;

      case ATypeInfo.Kind of
        tkRecord: Result := ConvAny2Nullable(Self, ATypeInfo, AResult);
      end
    end;

    if not Result then
    begin
      Result := TryCast(ATypeInfo, AResult);
    end;
  end;
end;

function TValueHelper.TryConvert<T>(out AResult: TValue): Boolean;
begin
  Result := TryConvert(System.TypeInfo(T), AResult);
end;

type
 // Declare compatible members of TRttiObject in System.Rtti.pas
  TRttiObjectFieldRef = class abstract
  public
    FHandle: Pointer;
    FRttiDataSize: Integer;
    FPackage: Pointer{TRttiPackage};
    FParent: Pointer {TRttiObject};
    FAttributeGetter: Pointer {TFunc<TArray<TCustomAttribute>>};
  end;

  TRttiObjectAccess = class helper for TRttiObject
  public
    procedure Init(Parent: TRttiType; PropInfo: PPropInfoExt);
  end;


procedure TRttiObjectAccess.Init(Parent: TRttiType; PropInfo: PPropInfoExt);
const
  FHANDLE_OFFSET = (SizeOf(Pointer) * 1);
  FPARENT_OFFSET = (SizeOf(Pointer) * 2);
begin
  TRttiObjectFieldRef(Self).FParent := Parent;
  TRttiObjectFieldRef(Self).FHandle := PropInfo;
end;

function TPropInfoExt.NameFld: TTypeInfoFieldAccessor;
begin
  Result.SetData(@NameLength);
end;

function TPropInfoExt.Tail: PPropInfoExt;
begin
  Result := PPropInfoExt(NameFld.Tail);
end;

class constructor TRttiPropertyExtension.Create;
begin
  FRegister := TObjectDictionary<TPair<PTypeInfo, string>, TRttiPropertyExtension>.Create([doOwnsValues]);
  FPatchedClasses := TDictionary<TClass, TClass>.Create;

  TRttiPropertyExtension.InitVirtualMethodTable;
end;

class destructor TRttiPropertyExtension.Destroy;
begin
  for var LClass in FPatchedClasses.Values do
  begin
    var LPointer := PByte(LClass) + vmtSelfPtr;
    FreeMem(LPointer);
  end;

  FPatchedClasses.Free;
  FRegister.Free;
end;

constructor TRttiPropertyExtension.Create(AParent: PTypeInfo; const AName: string; APropertyType: PTypeInfo);
begin
  inherited Create;
  FPropInfo.PropType := Pointer(NativeInt(APropertyType) - SizeOf(PTypeInfo));

  if AName.Length > 255 then
    FPropInfo.NameLength := 255
  else
    FPropInfo.NameLength := AName.Length;

  var LMarshaller: TMarshaller;
  Move(LMarshaller.AsAnsi(AName).ToPointer^, FPropInfo.NameData[0], FPropInfo.NameLength);

  Init(GetRttiType(AParent), @FPropInfo);
  FRegister.Add(TPair<PTypeInfo, string>.Create(AParent, AName), Self);
  PPointer(Self)^ := FPatchedClasses[Self.ClassType];
end;

function TRttiPropertyExtension.DoGetValue(Instance: Pointer): TValue;
begin
  Result := FGetter(Instance);
end;

function TRttiPropertyExtension.DoGetValueStub(Instance: Pointer): TValue;
begin
  Result := DoGetValue(Instance);
end;

procedure TRttiPropertyExtension.DoSetValue(Instance: Pointer; const AValue: TValue);
begin
  FSetter(Instance, AValue);
end;

procedure TRttiPropertyExtension.DoSetValueStub(Instance: Pointer; const AValue: TValue);
begin
  DoSetValue(Instance, AValue);
end;

class function TRttiPropertyExtension.FindByName(AParent: TRttiType; const APropertyName: string): TRttiPropertyExtension;
begin
  for var LPropertyExtension in FRegister.Values do
  begin
    if (LPropertyExtension.Parent = AParent) and SameText(LPropertyExtension.Name, APropertyName) then
    begin
      Result := LPropertyExtension;
      Exit;
    end;
  end;

  if Assigned(AParent.BaseType) then
    Result := FindByName(AParent.BaseType, APropertyName)
  else
    Result := nil
end;

class function TRttiPropertyExtension.FindByName(const AFullPropertyName: string): TRttiPropertyExtension;
begin
  Result := nil;

  var LScope := AFullPropertyName.SubString(0, AFullPropertyName.LastDelimiter('.'));
  var LName := AFullPropertyName.SubString(AFullPropertyName.LastDelimiter('.') + 1);

  for var LProp in FRegister.Values do
  begin
    if (string.Compare(LProp.Name, LName, [coIgnoreCase]) = 0) and LProp.Parent.AsInstance.MetaclassType.QualifiedClassName.EndsWith(LScope, True) then
    begin
      Result := LProp;
      Break;
    end;
  end;
end;

function TRttiPropertyExtension.GetIsReadable: Boolean;
begin
  Result := Assigned(FGetter);
end;

function TRttiPropertyExtension.GetIsReadableStub: Boolean;
begin
  Result := GetIsReadable;
end;

function TRttiPropertyExtension.GetIsWritable: Boolean;
begin
  Result := Assigned(FSetter);
end;

function TRttiPropertyExtension.GetIsWritableStub: Boolean;
begin
  Result := GetIsWritable;
end;

function TRttiPropertyExtension.GetPropInfo: PPropInfo;
begin
  Result := Handle;
end;

function TRttiPropertyExtension.GetPropInfoStub: PPropInfo;
begin
  Result := GetPropInfo;
end;

class procedure TRttiPropertyExtension.InitVirtualMethodTable;
const
  kMaxIndex = 17;  // TRttiInstanceProperty.GetPropInfo
{$POINTERMATH ON}
type
  PVtable = ^Pointer;
{$POINTERMATH OFF}
begin
  var LSize := SizeOf(Pointer) * (1 + kMaxIndex - (vmtSelfPtr div SizeOf(Pointer)));
  var LData := AllocMem(LSize);
  var LPatchedClass := TClass(PByte(LData) - vmtSelfPtr);
  FPatchedClasses.Add(Self, LPatchedClass);
  Move((PByte(Self) + vmtSelfPtr)^, LData^, LSize);

  PVtable(LPatchedClass)[5] := @TRttiPropertyExtension.GetIsReadableStub;
  PVtable(LPatchedClass)[6] := @TRttiPropertyExtension.GetIsWritableStub;
  PVtable(LPatchedClass)[7] := @TRttiPropertyExtension.DoGetValueStub;
  PVtable(LPatchedClass)[8] := @TRttiPropertyExtension.DoSetValueStub;
  PVtable(LPatchedClass)[12] := @TRttiPropertyExtension.GetPropInfoStub;
end;

class function TStrUtils.SplitString(const AStr, ADelimiters: string): TArray<string>;
begin
  Result := nil;

  if AStr <> string.Empty then
  begin
    // Determine the length of the Resulting array
    var LSplitPoints := 0;
    for var I := 0 to AStr.Length - 1 do
      if AStr.IsDelimiter(ADelimiters, I) then Inc(LSplitPoints);

    SetLength(Result, LSplitPoints + 1);

    // Split the string and fill the Resulting array
    var LStartIndex := 0;
    var LCurrentSplit := 0;

    repeat
      var LIndexFound := AStr.IndexOfAny(ADelimiters.ToCharArray, LStartIndex);
      if LIndexFound <> -1 then
      begin
        Result[LCurrentSplit] := AStr.SubString(LStartIndex, LIndexFound - LStartIndex);
        Inc(LCurrentSplit);
        LStartIndex := LIndexFound + 1;
      end;
    until LCurrentSplit = LSplitPoints;

    // Copy the remaining part in case the string does not end in a delimiter
    Result[LSplitPoints] := AStr.SubString(LStartIndex, AStr.Length - LStartIndex + 1);
  end;
end;

class function TListStringUtils.ToArray(const AValues: TList<string>): TArray<string>;
begin
  SetLength(Result,AValues.Count);

  for var I := 0 to AValues.Count - 1 do
    Result[I] := AValues[I];
end;

class function TInterfaceHelper.GetType(const AIntf: IInterface): TRttiInterfaceType;
var
  LImplObj: TObject;
  LGuid: TGUID;
  LIntfType: TRttiInterfaceType;
  LTempIntf: IInterface;
begin
  Result := nil;

  try
    // As far as I know, the cast will fail only when AIntf is obatined from OLE Object
    // Is there any other cases?
    LImplObj := AIntf as TObject;
  except
    // For interfaces obtained from OLE Object
    Result := Context.GetType(TypeInfo(System.IDispatch)) as TRttiInterfaceType;
    Exit;
  end;

  // For interfaces obtained from TRawVirtualClass (e.g. iOS, Android & Mac intf)
  if LImplObj.ClassType.InheritsFrom(TRawVirtualClass) then
  begin
    LGuid := LImplObj.GetField('FIIDs').GetValue(LImplObj).AsType<TArray<TGUID>>[0];
    Result := GetType(LGuid);
  end else
  begin
   // For interfaces obtained from TVirtualInterface
    if LImplObj.ClassType.InheritsFrom(TVirtualInterface) then
    begin
      LGuid := LImplObj.GetField('FIID').GetValue(LImplObj).AsType<TGUID>;
      Result := GetType(LGuid);
    end else
    begin
      // For interfaces obtained from Delphi object. Code taken from Remy Lebeau's answer
      // http://stackoverflow.com/questions/39584234/how-to-obtain-rtti-from-an-interface-reference-in-delphi/
      for LIntfType in (Context.GetType(LImplObj.ClassType) as TRttiInstanceType).GetImplementedInterfaces do
      begin
        if LImplObj.GetInterface(LIntfType.GUID, LTempIntf) and (AIntf = LTempIntf) then
          Exit(LIntfType);
      end;
    end;
  end;
end;

class constructor TInterfaceHelper.Create;
begin
  FInterfaceTypes := TInterfaceTypes.Create;
  FCached := False;
  FCaching := False;
  RefreshCache;
end;

class destructor TInterfaceHelper.Destroy;
begin
  FInterfaceTypes.DisposeOf;
end;

class function TInterfaceHelper.GetQualifiedName(const AIntf: IInterface): string;
var
  LType: TRttiInterfaceType;
begin
  LType := GetType(AIntf);

  if Assigned(LType) then
    Result := LType.QualifiedName
  else
    Result := EmptyStr;
end;

class function TInterfaceHelper.GetMethod(const AIntf: IInterface; const AMethodName: string): TRttiMethod;
begin
  var LType: TRttiInterfaceType := GetType(AIntf);

  if Assigned(LType) then
    Result := LType.GetMethod(AMethodName)
  else
    Result := nil;
end;

class function TInterfaceHelper.GetMethods(const AIntf: IInterface): TArray<TRttiMethod>;
begin
  var LType: TRttiInterfaceType := GetType(AIntf);

  if Assigned(LType) then
    Result := LType.GetMethods
  else
    Result := nil;
end;

class function TInterfaceHelper.GetQualifiedName(const AGuid: TGUID): string;
begin
  var LType: TRttiInterfaceType := GetType(AGuid);

  if Assigned(LType) then
    Result := LType.QualifiedName
  else
    Result := EmptyStr;
end;

class function TInterfaceHelper.GetType(const AGuid: TGUID): TRttiInterfaceType;
begin
  CacheIfNotCachedAndWaitFinish;
  Result := FInterfaceTypes.Items[AGuid];
end;

class function TInterfaceHelper.GetTypeName(const AGuid: TGUID): string;
begin
  var LType: TRttiInterfaceType := GetType(AGuid);

  if Assigned(LType) then
    Result := LType.Name
  else
    Result := EmptyStr;
end;

class function TInterfaceHelper.InvokeMethod(const AIntfInTValue: TValue; const AMethodName: string; const Args: array of TValue): TValue;
var
  LMethod: TRttiMethod;
  LType: TRttiInterfaceType;
begin
  LType := GetType(AIntfInTValue);

  if Assigned(LType) then
    LMethod := LType.GetMethod(AMethodName)
  else
    LMethod := nil;

  if Assigned(LMethod) then
    Result := LMethod.Invoke(AIntfInTValue, Args)
  else
    raise EMethodNotFound.Create(AMethodName);
end;

class function TInterfaceHelper.InvokeMethod(const AIntf: IInterface; const AMethodName: string; const Args: array of TValue): TValue;
begin
  var LMethod := GetMethod(AIntf, AMethodName);

  if not Assigned(LMethod) then
    raise EMethodNotFound.Create(AMethodName);

  Result := LMethod.Invoke(AIntf as TObject, Args);
end;

class function TInterfaceHelper.GetTypeName(const AIntf: IInterface): string;
begin
  var LType: TRttiInterfaceType := GetType(AIntf);

  if Assigned(LType) then
    Result := LType.Name
  else
    Result := EmptyStr;
end;

class procedure TInterfaceHelper.RefreshCache;
begin
  WaitIfCaching;
  FCaching := True;
  FCached := False;

  TThread.CreateAnonymousThread(
    procedure
    begin
      FInterfaceTypes.Clear;

      for var LType in Context.GetTypes do
      begin
        if LType.IsInterface then
        begin
          var LIntfType := (LType as TRttiInterfaceType);

          if TIntfFlag.ifHasGuid in LIntfType.IntfFlags then
            FInterfaceTypes.AddOrSetValue(LIntfType.GUID, LIntfType);
        end;
      end;

      FCaching := False;
      FCached := True;
    end
  ).Start;
end;

class procedure TInterfaceHelper.WaitIfCaching;
begin
  if FCaching then TSpinWait.SpinUntil(
    function: Boolean
    begin
      Result := FCached;
    end
  );
end;

class procedure TInterfaceHelper.CacheIfNotCachedAndWaitFinish;
begin
  if FCached then
    Exit;

  // Need to be protected because FCaching is changed inside. This will block GetType method.
  TMonitor.Enter(FInterfaceTypes);

  if not FCaching then
    RefreshCache;

  TMonitor.Exit(FInterfaceTypes);
  WaitIfCaching;
end;

class function TInterfaceHelper.GetType(const AIntfInTValue: TValue): TRttiInterfaceType;
begin
  var LType := AIntfInTValue.RttiType;

  if LType is TRttiInterfaceType then
    Result := LType as TRttiInterfaceType
  else
    Result := nil;
end;

constructor EMethodNotFound.Create(const AMethodName: string);
begin
  inherited CreateFmt('Method %s not found.', [AMethodName]);
end;


initialization
  Enumerations := TObjectDictionary<PTypeInfo, TStrings>.Create([doOwnsValues]);

finalization
  Enumerations.Free;


end.

