{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$ENDIF}

function TDA2.AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                  hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                  dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                out pRevisedUpdateRate:DWORD;
                                const riid: TIID;
                                out ppUnk:IUnknown):HResult;stdcall;
var
 s1:string;
 i:longint;
 aGrp:TOPCGroup;
 perDeadband:single;
begin
 result:=S_OK;
 s1:=szName;
 i:=0;

 if (s1 = '') then
  repeat
   s1:=s1 + IntToStr(GetTickCount);
   i:=succ(i);
  until (not IsNameUsedInAnyGroup(s1)) or (i > 9);

 if i > 9 then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 if IsNameUsedInAnyGroup(s1) then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 aGrp:=TOPCGroup.Create(self,grps);
 if aGrp = nil then
  begin
   result:=E_OUTOFMEMORY;
   phServerGroup:=0;
   Exit;
  end;

 grps.Add(aGrp);
 phServerGroup:=GetNewGroupNumber;
 if assigned(pPercentDeadband) then
  perDeadband:=pPercentDeadband^
 else
  perDeadband:=0;

 aGrp.SetUp(s1,bActive,dwRequestedUpdateRate,hClientGroup,0,
            perDeadband,dwLCID,phServerGroup);

 aGrp.ValidateTimeBias(pTimeBias);

 if dwRequestedUpdateRate <> aGrp.requestedUpdateRate then
  result:=OPC_S_UNSUPPORTEDRATE;

 pRevisedUpdateRate:=aGrp.requestedUpdateRate;

 Form1.UpdateGroupCount;
 ppUnk:=aGrp;
end;

function TDA2.GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult; stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.GetGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
var
 i,gNum:integer;
begin
 gNum:=IsGroupNamePresent(grps,szName);     //returns the index to the groups list
 if (addr(ppUnk) = nil) or (gNum = -1)then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 i:=gNum;
// i:=FindIndexViaGrpNumber(grps,gNum);
 if (i = -1) or (i > (grps.count-1) )  then
  begin
   Result:=E_FAIL;              Exit;
  end;
 result:=IUnknown(TOPCGroup(grps[i])).QueryInterface(riid,ppUnk);
end;

function TDA2.GetStatus(out ppServerStatus:POPCSERVERSTATUS):HResult;stdcall;
var
 aFileTime:TFileTime;
begin
 if (addr(ppServerStatus) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 result:=S_OK;
 ppServerStatus:=POPCSERVERSTATUS(CoTaskMemAlloc(sizeof(OPCSERVERSTATUS)));
 if ppServerStatus = nil then
  begin
   ppServerStatus:=nil;         result:=E_OUTOFMEMORY;   Exit;
  end;

 DataTimeToOPCTime(srvStarted,aFileTime);
 CoFiletimeNow(ppServerStatus.ftCurrentTime);
 DataTimeToOPCTime(lastClientUpdate,ppServerStatus.ftLastUpdateTime);

 ppServerStatus.dwServerState:=OPC_STATUS_RUNNING;
 ppServerStatus.dwGroupCount:=GetGroupCount(grps) + GetGroupCount(pubGrps);
 ppServerStatus.dwBandWidth:=100;
 ppServerStatus.wMajorVersion:=1;
 ppServerStatus.wMinorVersion:=2;
 ppServerStatus.wBuildNumber:=5;
 ppServerStatus.szVendorInfo:=StringToLPOLESTR('StateManager');
end;

function TDA2.RemoveGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
var
 i:integer;
 aGrp:TOPCGroup;
begin
 result:=S_OK;
 if hServerGroup < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 i:=FindIndexViaGrpNumber(grps,hServerGroup);
 if i = -1 then            //the client as already freed the group
  Exit;                    //and we deleted it from the server list in the group destroy

 aGrp:=grps[i];
 if (aGrp.RefCount > 2) and not bForce then
  begin
   result:=OPC_S_INUSE;
   Exit;
  end;

 GroupRemovingSelf(grps,hServerGroup);
end;

function TDA2.CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID;
                                         out ppUnk:IUnknown):HResult; stdcall;

 procedure EnumerateForStrings;
 var
  i:integer;
  pvList,pubList:TStringList;
 begin
  pvList:=nil;               pubList:=nil;
  try
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     pvList:=CreateGrpNameList(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     pubList:=CreateGrpNameList(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
      pvList:=CreateGrpNameList(grps);
      pubList:=CreateGrpNameList(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if pvList = nil then
    pvList:=TStringList.Create;
   if pubList <> nil then
    if pubList.count > 0 then
     for i:= 0 to pubList.count-1 do
      pvList.Add(pubList[i]);

   result:=S_OK;
   ppUnk:=TOPCStringsEnumerator.Create(pvList);
  finally
   pvList.Free;
   pubList.Free;
  end;
 end;


 procedure EumerateForUnknown;
 var
  aList:TList;

 procedure AddAGroup(inList:TList);
 var
  i:integer;
  Obj:Pointer;
 begin
  for i:= 0 to inList.count -1 do
   begin
    Obj:=nil;
    IUnknown(TOPCGroup(inList[i])).QueryInterface(IUnknown,Obj);
    if Assigned(Obj) then
     aList.Add(Obj);
   end;
 end;

 begin
  aList:=nil;
  try
   aList:=TList.Create;
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     if Assigned(grps) then
      AddAGroup(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
     if Assigned(grps) then
      AddAGroup(grps);
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if Assigned(aList) and (aList.count > 0) then
    ppUnk:=TS3UnknownEnumerator.Create(aList);
  finally
   aList.Free;
  end;
  result:=S_OK;
 end;

begin
 if not (IsEqualIID(riid,IEnumUnknown) or IsEqualIID(riid,IEnumString)) then
  begin
   result:=E_NOINTERFACE;
   Exit;
  end;

 if IsEqualIID(riid,IEnumString) then
  EnumerateForStrings
 else if IsEqualIID(riid,IEnumUnknown) then
  EumerateForUnknown
 else
  result:=E_FAIL;

end;


