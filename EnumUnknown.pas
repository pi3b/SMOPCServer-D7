unit EnumUnknown;

interface

uses Classes,WinProcs,ComObj,ActiveX;

type
  TS3UnknownEnumerator = class(TComObject, IEnumUnknown)
  private
    nextIndex:Integer;
    theList:TList;
  public
    constructor Create(const inList: TList);
    destructor destroy;override;
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumUnknown): HResult; stdcall;
    procedure AddAnotherList(inList:TList);
  end;

implementation

uses ComServ,OPCDA;

constructor TS3UnknownEnumerator.Create(const inList:TList);
var
 i:integer;
begin
 inherited Create;
 theList:=TList.Create;
 if not Assigned(theList) then
  Exit;
 nextIndex := 0;

 if not Assigned(inList) then Exit;
 for i:=0 to inList.count-1 do
  theList.add(inList[i]);
end;

destructor TS3UnknownEnumerator.destroy;
begin
 theList.Free;
 inherited Destroy;
end;

function TS3UnknownEnumerator.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
 i: integer;
begin
 i:=0;
 if celt < 1 then
  begin
   Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
   Exit;
  end;

 if pceltFetched = nil then
  begin
   Result:=E_INVALIDARG;
   Exit;
  end;

 Result := S_FALSE;
 while (i < celt) do
  begin
   if (nextIndex < theList.Count) then
    begin
     TPointerList(elt)[i]:=theList[nextIndex];
     i:=succ(i);
     nextIndex:=succ(nextIndex);
    end
   else
    begin
     Result:=RPC_X_ENUM_VALUE_OUT_OF_RANGE;
     Break;
    end;
  end;

 pceltFetched^:=i;
 if i = celt then
  Result := S_OK;
end;

function TS3UnknownEnumerator.Skip(celt: Longint): HResult;
begin
 if (nextIndex + celt) <= theList.Count then
  begin
   nextIndex:=nextIndex + celt;
   result:=S_OK;
  end
 else
  begin
   nextIndex:=theList.Count;
   result:=S_FALSE;
  end;
end;

function TS3UnknownEnumerator.Reset: HResult;
begin
 nextIndex:=0;
 result:=S_OK;
end;

function TS3UnknownEnumerator.Clone(out enm: IEnumUnknown): HResult;
begin
 try
  enm:=TS3UnknownEnumerator.Create(theList);
  result:=S_OK;
 except
  result:=E_UNEXPECTED;
 end;
end;

procedure TS3UnknownEnumerator.AddAnotherList(inList:TList);
var
 i:integer;
begin
 if not Assigned(inList) then Exit;
 for i:=0 to inList.count-1 do
  theList.add(inList[i]);
end;

initialization
 TComObjectFactory.Create(ComServer,
                          TS3UnknownEnumerator,
                          IEnumUnknown,
                          'TS3UnknownEnumerator',
                          'SMOPC',
                          ciMultiInstance,
                          tmApartment);
end.
