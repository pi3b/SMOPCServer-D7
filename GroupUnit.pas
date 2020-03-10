unit GroupUnit;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses Windows,ActiveX,ComObj,SysUtils,Dialogs,Classes,ServIMPL,OPCDA,axctrls,
     Globals,OpcError,OPCTypes,SMOPC_TLB;

type
  TOPCGroup = class(TTypedComObject,IOPCGroup,IOPCItemMgt,IOPCGroupStateMgt,
                    IOPCPublicGroupStateMgt,IOPCSyncIO,
                    IConnectionPointContainer,IOPCAsyncIO2)
  private
   FIConnectionPoints:TConnectionPoints;
  protected
   property iFIConnectionPoints:TConnectionPoints read FIConnectionPoints
                          write FIConnectionPoints implements IConnectionPointContainer;
//IOPCItemMgt begin
   function AddItems(dwCount:DWORD; pItemArray:POPCITEMDEFARRAY;
                     out ppAddResults:POPCITEMRESULTARRAY;
                     out ppErrors:PResultList):HResult; stdcall;
   function ValidateItems(dwCount:DWORD; pItemArray:POPCITEMDEFARRAY; bBlobUpdate:BOOL; out ppValidationResults:        POPCITEMRESULTARRAY;
                          out ppErrors:PResultList):HResult; stdcall;
   function RemoveItems(dwCount:DWORD; phServer:POPCHANDLEARRAY; out ppErrors:PResultList):HResult; stdcall;
   function SetActiveState(dwCount:DWORD; phServer:POPCHANDLEARRAY; bActive:BOOL; out ppErrors:PResultList):HResult; stdcall;
   function SetClientHandles(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                                            phClient:POPCHANDLEARRAY;
                             out ppErrors:PResultList):HResult; stdcall;
   function SetDatatypes(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                         pRequestedDatatypes:PVarTypeList;
                         out ppErrors:PResultList):HResult; stdcall;
   function CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
//IOPCItemMgt end

//IOPCGroupStateMgt begin
   function GetState(out pUpdateRate:DWORD; out pActive:BOOL; out ppName:POleStr;
                     out pTimeBias:Longint; out pPercentDeadband:Single; out pLCID:TLCID;
                     out phClientGroup:OPCHANDLE; out phServerGroup:OPCHANDLE):HResult;overload;stdcall;
   function SetState(pRequestedUpdateRate:PDWORD; out pRevisedUpdateRate:DWORD; pActive:PBOOL;
                     pTimeBias:PLongint; pPercentDeadband:PSingle; pLCID:PLCID;
                     phClientGroup:POPCHANDLE):HResult; stdcall;
   function SetName(szName:POleStr):HResult;stdcall;
   function CloneGroup(szName:POleStr; const riid: TIID; out ppUnk:IUnknown):HResult;stdcall;
//IOPCGroupStateMgt end

//IOPCPublicGroupStateMgt begin
   function GetState(out pPublic:BOOL):HResult;overload;stdcall;
   function MoveToPublic:HResult;stdcall;
//IOPCPublicGroupStateMgt end

//IOPCSyncIO begin
   function Read(dwSource:OPCDATASOURCE; dwCount:DWORD; phServer:POPCHANDLEARRAY;
                 out ppItemValues:POPCITEMSTATEARRAY; out ppErrors:PResultList):HResult;overload;stdcall;
   function Write(dwCount:DWORD; phServer:POPCHANDLEARRAY; pItemValues:POleVariantArray;
                  out ppErrors:PResultList):HResult;overload;stdcall;
//IOPCSyncIO end

//IOPCAsyncIO2 begin
   function Read(dwCount:DWORD; phServer:POPCHANDLEARRAY; dwTransactionID:DWORD;
                 out pdwCancelID:DWORD; out ppErrors:PResultList):HResult;overload;stdcall;
   function Write(dwCount:DWORD; phServer:POPCHANDLEARRAY; pItemValues:POleVariantArray;
                  dwTransactionID:DWORD; out pdwCancelID:DWORD; out ppErrors:PResultList):HResult;overload;stdcall;
   function Refresh2(dwSource:OPCDATASOURCE; dwTransactionID:DWORD;
                     out pdwCancelID:DWORD):HResult;stdcall;
   function Cancel2(dwCancelID:DWORD):HResult;stdcall;
   function SetEnable(bEnable:BOOL):HResult;stdcall;
   function GetEnable(out pbEnable:BOOL):HResult;stdcall;
//IOPCAsyncIO2 end
  public
   servObj:TDA2;                         //the owner
   tagName:string;                       //the name of this group
   clientHandle:longword;                //the client generates we pass to client
   serverHandle:longword;                //we generate the client will passes to us
   requestedUpdateRate:longword;         //update rate in mills
   lang:longword;                        //lanugage id
   nextUpdate:longword;

   ownList,clItems,asyncList:TList;
   groupActive,groupPublic,onDataChangeEnabled:longbool;
   timeBias:longint;
   percentDeadband:single;
   FOnCallBackConnect:TConnectEvent;
   ClientIUnknown:IUnknown;
   lastMSecUpdate:Comp;
   upStream:TMemoryStream;

   procedure Initialize; override;
   constructor Create(serv:TDA2;oList:TList);
   destructor Destroy;override;
   function ValidateRequestedUpDateRate(dwRequestedUpdateRate:DWORD):DWORD;
   procedure ValidateTimeBias(pTimeBias:PLongint);
   procedure SetUp(szName:string;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                  hClientGroup:OPCHANDLE; pTimeBias:longint; pPercentDeadband:single;
                  dwLCID:DWORD;phServerGroup:longword);
   procedure CallBackOnConnect(const Sink: IUnknown; Connecting: Boolean);
   function GetItemIndexFromServerHandle(servHand:longword;var index:integer):boolean;
   function GetItemIndexFromClientHandle(clHand:longword;var index:integer):boolean;
   function GenerateAsyncCancelID:longword;
   procedure TimeSlice(cTime:TDateTime);
   procedure CloneYourSelf(dGrp:TOPCGroup);
   procedure DoAChangeOccured(aStream:TMemoryStream; cTime:TDateTime);
   procedure AsyncTimeSlice(cTime:TDateTime);
  end;

implementation

uses ComServ,ItemsUnit,AsyncUnit,ItemAttributesOPC,EnumItemAtt,Main;

function IsVariantTypeOK(vType:integer):boolean;
begin
 result:=boolean(vType in [varEmpty..$14]);
end;

//IOPCItemMgt begin
function TOPCGroup.AddItems(dwCount:DWORD; pItemArray:POPCITEMDEFARRAY;
                     out ppAddResults:POPCITEMRESULTARRAY;
                     out ppErrors:PResultList):HResult;stdcall;
var
 i:integer;
 wItem:TOPCItem;
 propID:longword;
 memErr:boolean;
 inItemDef:POPCITEMDEF;

 procedure ClearResultsArray;
 var
  i:integer;
 begin
  for i:= 0 to dwCount -1 do
   begin
    ppAddResults[i].hServer:=0;                 ppAddResults[i].vtCanonicalDataType:=0;
    ppAddResults[i].wReserved:=0;               ppAddResults[i].dwAccessRights:=0;
    ppAddResults[i].dwBlobSize:=0;              ppAddResults[i].pBlob:=nil;
  end;
 end;

begin
 result:=S_OK;
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppAddResults:=POPCITEMRESULTARRAY(CoTaskMemAlloc(dwCount*sizeof(OPCITEMRESULT)));
 memErr:=boolean(ppAddResults = nil);

 if not memErr then
  begin
   ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
   memErr:=boolean(ppErrors = nil);
  end;

 if memErr then
  begin
   if ppAddResults <> nil then  CoTaskMemFree(ppAddResults);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ClearResultsArray;

 for i:= 0 to dwCount -1 do
  begin
   ppErrors[i]:=S_OK;
   inItemDef:=@pItemArray[i];

   if length(inItemDef.szItemID) = 0 then
    begin
     result:=S_FALSE;
     ppErrors[i]:=OPC_E_INVALIDITEMID;
     Continue;
    end;


     propID:=ReturnPropIDFromTagname(inItemDef.szItemID);

     if propID = 0 then
      begin
         if devices.ItemIDAdd(inItemDef.szItemID) then begin

         end else begin
           result:=S_FALSE;
           ppErrors[i]:=OPC_E_UNKNOWNITEMID;
           Continue;
         end;
      end;

   if not IsVariantTypeOK(inItemDef.vtRequestedDataType) then
    begin
     result:=S_FALSE;
     ppErrors[i]:=OPC_E_BADTYPE;
     Continue;
    end;

   wItem:=TOPCItem.Create;
   wItem.servObj:=servObj;
   wItem.serverItemNum:=servObj.GetNewItemNumber;

   wItem.SetActiveState(inItemDef.bActive);
   wItem.SetClientHandle(inItemDef.hClient);
//   wItem.itemIndex:=propID - posItems[low(posItems)].PropID;

   wItem.isWriteAble:=CanPropIDBeWritten(inItemDef.szItemID);
   wItem.canonicalDataType:=ReturnDataTypeFromPropID(inItemDef.szItemID);

   wItem.SetOldValue;

   wItem.SetReqDataType(inItemDef.vtRequestedDataType);
   wItem.strID:=inItemDef.szItemID;
//   wItem.IsSimulate:=IsSimulate(inItemDef.szItemID);
   wItem.pBlob:=inItemDef.pBlob;      
   clItems.Add(wItem);

   ppAddResults[i].hServer:=wItem.serverItemNum;
   ppAddResults[i].vtCanonicalDataType:=wItem.canonicalDataType;

   if wItem.isWriteAble then
    ppAddResults[i].dwAccessRights:=OPC_READABLE or OPC_WRITEABLE
   else
    ppAddResults[i].dwAccessRights:=OPC_READABLE;

   ppAddResults[i].dwBlobSize:=0;
   ppAddResults[i].pBlob:=wItem.pBlob;
  end;
end;

function TOPCGroup.ValidateItems(dwCount:DWORD; pItemArray:POPCITEMDEFARRAY;bBlobUpdate:BOOL;
                   out ppValidationResults:POPCITEMRESULTARRAY;out ppErrors:PResultList):HResult; stdcall;
var
 i:integer;
 memErr:boolean;
 propID:longword;
 inItemDef:POPCITEMDEF;

 procedure ClearResultsArray;
 var
  i:integer;
 begin
  for i:= 0 to dwCount -1 do
   with ppValidationResults[i] do
    begin
     vtCanonicalDataType:=0;    dwAccessRights:=0;
     dwBlobSize:=0;             hServer:=0;
    end;
 end;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppValidationResults:=POPCITEMRESULTARRAY(CoTaskMemAlloc(dwCount*sizeof(OPCITEMRESULT)));
 memErr:=boolean(ppValidationResults = nil);

 if not memErr then
  begin
   ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
   memErr:=boolean(ppErrors = nil);
  end;

 if memErr then
  begin
   if ppValidationResults <> nil then  CoTaskMemFree(ppValidationResults);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 ClearResultsArray;
 for i:= 0 to dwCount -1 do
  begin
   inItemDef:=@pItemArray[i];

   propID:=ReturnPropIDFromTagname(inItemDef.szItemID);  //also cover 0 length
   if propID = 0 then
    begin
     result:=S_FALSE;
     ppErrors[i]:=OPC_E_INVALIDITEMID;
     Continue;
    end;

   ppValidationResults[i].vtCanonicalDataType:=ReturnDataTypeFromPropID(inItemDef.szItemID);
   if CanPropIDBeWritten(inItemDef.szItemID) then
    ppValidationResults[i].dwAccessRights:=OPC_READABLE or OPC_WRITEABLE
   else
    ppValidationResults[i].dwAccessRights:=OPC_READABLE;
   ppValidationResults[i].dwBlobSize:=0;
   ppErrors[i]:=S_OK;
  end;
end;

function TOPCGroup.RemoveItems(dwCount:DWORD; phServer:POPCHANDLEARRAY; out ppErrors:PResultList):HResult; stdcall;
var
 i,x:integer;
 wItem:TOPCItem;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(phServer[i],x) then
   begin
    wItem:=clItems[x];
    clItems.Delete(x);
    TOPCItem(wItem).Free;
    ppErrors[i]:=S_OK;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;


function TOPCGroup.SetActiveState(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                   bActive:BOOL; out ppErrors:PResultList):HResult; stdcall;
var
 i,x:integer;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(phServer[i],x) then
   begin
    TOPCItem(clItems[x]).SetActiveState(bActive);
    ppErrors[i]:=S_OK;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;

function TOPCGroup.SetClientHandles(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                                    phClient:POPCHANDLEARRAY;
                                    out ppErrors:PResultList):HResult; stdcall;
var
 i,x:integer;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(phServer[i],x) then
   begin
    TOPCItem(clItems[x]).SetClientHandle(phClient[i]);
    ppErrors[i]:=S_OK;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;

function TOPCGroup.SetDatatypes(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                                pRequestedDatatypes:PVarTypeList;
                                out ppErrors:PResultList):HResult; stdcall;
var
 i,x:integer;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(phServer[i],x) then
   begin
    if not IsVariantTypeOK(pRequestedDatatypes[i]) then
     ppErrors[i]:=OPC_E_BADTYPE
    else
     begin
      TOPCItem(clItems[x]).SetReqDataType(pRequestedDatatypes[i]);
      ppErrors[i]:=S_OK;
     end;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;

function TOPCGroup.CreateEnumerator(const riid: TIID; out ppUnk: IUnknown): HResult; stdcall;
var
 i:integer;
 aList:TList;
 aAttr:TOPCItemAttributes;
begin
 aList:=nil;
 result:=S_OK;
 if (clItems = nil) or (clItems.count = 0) then
  begin
   result:=S_FALSE;                Exit;
  end;

 try
  aList:=TList.Create;
  if aList = nil then
   begin
    result:=E_OUTOFMEMORY;                Exit;
   end;

  for i:= 0 to clItems.count-1 do
   begin
    aAttr:=TOPCItemAttributes.Create;
    if aAttr = nil then
     begin
      result:=E_OUTOFMEMORY;                Exit;
     end;
    TOPCItem(clItems[i]).FillInOPCItemObject(aAttr);
    aList.Add(aAttr);
   end;

  ppUnk:=TOPCItemAttEnumerator.Create(aList);

 finally
  if (aList <> nil) and (aList.count > 0) then
   for i:= 0 to aList.count-1 do
    TOPCItemAttributes(aList[i]).Free;
  aList.Free;
 end;
end;
//IOPCItemMgt end

//IOPCGroupStateMgt begin
function TOPCGroup.GetState(out pUpdateRate:DWORD; out pActive:BOOL; out ppName:POleStr;
                  out pTimeBias:Longint; out pPercentDeadband:Single; out pLCID:TLCID;
                  out phClientGroup:OPCHANDLE; out phServerGroup:OPCHANDLE):HResult;stdcall;
begin
 pUpdateRate:=requestedUpdateRate;
 pActive:=groupActive;
 ppName:=StringToLPOLESTR(tagName);
 pTimeBias:=timeBias;
 pPercentDeadband:=percentDeadband;            pLCID:=lang;
 phClientGroup:=clientHandle;                  phServerGroup:=serverHandle;
 result:=S_OK;
end;

function TOPCGroup.SetState(pRequestedUpdateRate:PDWORD; out pRevisedUpdateRate:DWORD;
                   pActive:PBOOL; pTimeBias:PLongint; pPercentDeadband:PSingle; pLCID:PLCID;
                   phClientGroup:POPCHANDLE):HResult; stdcall;
begin
 result:=S_OK;
 if assigned(pRequestedUpdateRate) then
  if ValidateRequestedUpDateRate(pRequestedUpdateRate^) <> pRequestedUpdateRate^ then
   result:=OPC_S_UNSUPPORTEDRATE;


 if assigned(pRequestedUpdateRate) then requestedUpdateRate:=pRequestedUpdateRate^;
 if assigned(pActive) then groupActive:=pActive^;
 if assigned(pTimeBias) then timeBias:=pTimeBias^;
 if assigned(pPercentDeadband) then percentDeadband:=pPercentDeadband^;
 if assigned(pLCID) then lang:=pLCID^;
 if assigned(phClientGroup) then clientHandle:=phClientGroup^;

 if (addr(pRevisedUpdateRate) <> nil) then
  pRevisedUpdateRate:=requestedUpdateRate;
end;

function TOPCGroup.SetName(szName:POleStr):HResult;stdcall;
begin
 result:=S_OK;
 if length(szName) = 0 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;
 if (servObj.IsGroupNamePresent(servObj.grps,szName) <> -1) then
  begin
   result:=OPC_E_DUPLICATENAME;
   Exit;
  end;
 tagName:=szName;
end;

function TOPCGroup.CloneGroup(szName:POleStr; const riid: TIID; out ppUnk:IUnknown):HResult;stdcall;
var
 s1:string;
 i:integer;
begin
 if not (IsEqualIID(riid,IID_IOPCGroupStateMgt) or IsEqualIID(riid,IID_IUnknown)) then
  begin
   result:=E_NOINTERFACE;
   Exit;
  end;

 s1:=szName;
 if length(s1) <> 0 then
  if servObj.IsNameUsedInAnyGroup(s1) then
   begin
    result:=OPC_E_DUPLICATENAME;
    Exit;
   end;

 i:=0;
 if servObj.IsNameUsedInAnyGroup(s1) then
  repeat
   s1:=s1+IntToStr(GetTickCount);
   i:=succ(i);
  until (not servObj.IsNameUsedInAnyGroup(s1)) or (i > 9);

 if i > 9 then
  begin
   result:=E_FAIL;
   Exit;
  end;

 ppUnk:=servObj.CloneAGroup(s1,self,result);
 if result <> 0 then
  begin
   ppUnk:=nil;
   Exit;
  end;
 result:=S_OK;
end;
//IOPCGroupStateMgt end

//IOPCPublicGroupStateMgt begin
function TOPCGroup.GetState(out pPublic:BOOL):HResult;stdcall;
begin
 pPublic:=servObj.IsThisGroupPublic(ownList);
 result:=S_OK;
end;

function TOPCGroup.MoveToPublic:HResult;stdcall;
var
 i:integer;
begin
 i:=servObj.IsGroupNamePresent(servObj.pubGrps,tagName);
 if i = -1 then
  begin
   result:=E_FAIL;                    Exit;
  end;
 with servObj do
  begin
   pubGrps.Add(grps[i]);
   GroupRemovingSelf(grps,serverHandle);
   ownList:=pubGrps;
  end;
 result:=S_OK;
end;
//IOPCPublicGroupStateMgt end

//IOPCSyncIO begin
function TOPCGroup.Read(dwSource:OPCDATASOURCE; dwCount:DWORD; phServer:POPCHANDLEARRAY;
                        out ppItemValues:POPCITEMSTATEARRAY; out ppErrors:PResultList):HResult;stdcall;
var
 i,x:integer;
 memErr:boolean;
 ppServer:PDWORDARRAY;

 procedure ClearResultsArray;
 var
  i:integer;
 begin
  for i:= 0 to dwCount -1 do
   begin
    ppItemValues[i].hClient:=0;                 ppItemValues[i].wReserved:=0;
    ppItemValues[i].vDataValue:=0;              ppItemValues[i].wQuality:=0;
  end;
 end;

begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppItemValues:=POPCITEMSTATEARRAY(CoTaskMemAlloc(dwCount * sizeof(OPCITEMSTATE)));
 memErr:=boolean(ppItemValues = nil);

 if not memErr then
  begin
   ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
   memErr:=boolean(ppErrors = nil);
  end;

 if memErr then
  begin
   if ppItemValues <> nil then  CoTaskMemFree(ppItemValues);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 ppServer:=@phServer^;
 ClearResultsArray;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(ppServer[i],x) then
   begin
    TOPCItem(clItems[x]).ReadItemValueStateTime(dwSource,ppItemValues[i]);
    if (dwSource <> OPC_DS_DEVICE) and not groupActive then
     ppItemValues[i].wQuality:=OPC_QUALITY_OUT_OF_SERVICE;
    ppErrors[i]:=S_OK;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;

function TOPCGroup.Write(dwCount:DWORD; phServer:POPCHANDLEARRAY;
                         pItemValues:POleVariantArray; out ppErrors:PResultList):HResult;stdcall;
var
 i,x:integer;
 ppServer:PDWORDARRAY;
begin
 if dwCount < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 result:=S_OK;
 ppServer:=@phServer^;
 for i:= 0 to dwCount -1 do
  if GetItemIndexFromServerHandle(ppServer[i],x) then
   begin
    if not TOPCItem(clItems[x]).isWriteAble then
     begin
      ppErrors[i]:=OPC_E_BADRIGHTS;
      result:=S_FALSE;
     end
    else
     begin
      TOPCItem(clItems[x]).WriteItemValue(pItemValues[i]);
      ppErrors[i]:=S_OK
     end;
   end
  else
   begin
    result:=S_FALSE;
    ppErrors[i]:=OPC_E_INVALIDHANDLE;
   end;
end;
//IOPCSyncIO end

//IOPCAsyncIO2 begin
function TOPCGroup.Read(dwCount:DWORD; phServer: POPCHANDLEARRAY; dwTransactionID:DWORD;
              out pdwCancelID:DWORD; out ppErrors:PResultList):HResult;stdcall;
var
 memErr:boolean;
 i,itemIndex:integer;
 aAsyncObj:TAsyncIO2;
begin
 if ClientIUnknown = nil then
  begin
   result:=CONNECT_E_NOCONNECTION;
   Exit;
  end;

 if (dwCount < 1) then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 ppErrors:=PResultList(CoTaskMemAlloc(dwCount * sizeof(HRESULT)));
 memErr:=boolean(ppErrors = nil);

 if not memErr then
  begin
   aAsyncObj:=TAsyncIO2.Create(self,io2Read,dwTransactionID,
                               dwCount,OPC_DS_DEVICE);
   memErr:=boolean(aAsyncObj = nil);
  end;

 if not memErr then
  try
   GetMem(aAsyncObj.ppServer,dwCount * sizeof(longword));
  except
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 if not memErr then
  begin
   pdwCancelID:=aAsyncObj.cancelID;        //was generated in TAsyncIO2.Create
   aAsyncObj.itemCount:=0;
   for i:= 0 to dwCount-1 do
    begin
     if GetItemIndexFromServerHandle(phServer[i],itemIndex) then
      begin
       aAsyncObj.ppServer[aAsyncObj.itemCount]:=itemIndex;
       aAsyncObj.itemCount:=succ(aAsyncObj.itemCount);
       ppErrors[i]:=S_OK;
      end
     else
      ppErrors[i]:=OPC_E_INVALIDHANDLE;
    end;
  end;

 if aAsyncObj.itemCount < 1 then
  begin
   if Assigned(aAsyncObj) then
    FreeAndNil(aAsyncObj);
   result:=S_FALSE;
  end
 else
  begin
   asyncList.Add(aAsyncObj);
   result:=S_OK;
  end;
end;

function TOPCGroup.Write(dwCount:DWORD; phServer:POPCHANDLEARRAY; pItemValues:POleVariantArray;
               dwTransactionID:DWORD; out pdwCancelID:DWORD; out ppErrors:PResultList):HResult;stdcall;
var
 i:longword;
 itemIndex:integer;
 memErr:boolean;
 aAsyncObj:TAsyncIO2;
begin
 if ClientIUnknown = nil then
  begin
   result:=CONNECT_E_NOCONNECTION;
   Exit;
  end;

 if (dwCount < 1) then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 aAsyncObj:=TAsyncIO2.Create(self,io2Write,dwTransactionID,dwCount,0);
 memErr:=boolean(aAsyncObj = nil);

 if not memErr then
  begin
   ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
   memErr:=boolean(ppErrors = nil);
  end;

 if not memErr then
  try
   GetMem(aAsyncObj.ppServer,dwCount * sizeof(longword));
   GetMem(aAsyncObj.ppValues,dwCount * sizeof(OleVariant));
  except
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 if not memErr then
  begin
   pdwCancelID:=aAsyncObj.cancelID;        //was generated in TAsyncIO2.Create
   aAsyncObj.itemCount:=0;
   for i:= 0 to dwCount-1 do
    begin
     if GetItemIndexFromServerHandle(phServer[i],itemIndex) then
      begin
       aAsyncObj.ppServer[aAsyncObj.itemCount]:=itemIndex;
       aAsyncObj.ppValues[aAsyncObj.itemCount]:=pItemValues[i];
       aAsyncObj.itemCount:=succ(aAsyncObj.itemCount);
       ppErrors[i]:=S_OK;
      end
     else
      ppErrors[i]:=OPC_E_INVALIDHANDLE;
    end;
  end;

 if memErr then
  begin
   if ppErrors <> nil           then    CoTaskMemFree(ppErrors);
   if Assigned(aAsyncObj)       then    FreeAndNil(aAsyncObj);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 if aAsyncObj.itemCount < 1 then
  begin
   if Assigned(aAsyncObj)       then    FreeAndNil(aAsyncObj);
   result:=S_FALSE;
  end
 else
  begin
   asyncList.Add(aAsyncObj);
   result:=S_OK;
  end;
end;

function TOPCGroup.Refresh2(dwSource:OPCDATASOURCE; dwTransactionID:DWORD;
                  out pdwCancelID:DWORD):HResult;stdcall;
var
 i:integer;
 aAsyncObj:TAsyncIO2;
 allInActive:boolean;
begin
 if ClientIUnknown = nil then
  begin
   result:=CONNECT_E_NOCONNECTION;
   Exit;
  end;

 allInActive:=true;
 for i:= 0 to clItems.count-1 do
  if TOPCItem(clItems[i]).GetActiveState then
   begin
    allInActive:=false;
    Break;
   end;

 if (not groupActive) or allInActive then
  begin
   result:=E_FAIL;
   Exit;
  end;

 aAsyncObj:=TAsyncIO2.Create(self,io2Refresh,dwTransactionID,0,dwSource);
 if (aAsyncObj = nil) then
  begin
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 pdwCancelID:=aAsyncObj.cancelID;
 asyncList.Add(aAsyncObj);
 result:=S_OK;
end;

function TOPCGroup.Cancel2(dwCancelID:DWORD):HResult;stdcall;
var
 i:integer;
begin
 result:=E_FAIL;
 if (asyncList = nil) or (asyncList.count = 0) then
  Exit;
 for i:= 0 to asyncList.count-1 do
  if TAsyncIO2(asyncList[i]).cancelID = dwCancelID then
   begin
    result:=S_OK;
    TAsyncIO2(asyncList[i]).isCancelled:=true;
    Break;
   end
end;

function TOPCGroup.SetEnable(bEnable:BOOL):HResult;stdcall;
begin
 if ClientIUnknown = nil then
  begin
   result:=CONNECT_E_NOCONNECTION;
   Exit;
  end;
 onDataChangeEnabled:=bEnable;
 result:=S_OK;
end;

function TOPCGroup.GetEnable(out pbEnable:BOOL):HResult;stdcall;
begin
 if ClientIUnknown = nil then
  begin
   result:=CONNECT_E_NOCONNECTION;
   Exit;
  end;
 pbEnable:=onDataChangeEnabled;
 result:=S_OK;
end;
//IOPCAsyncIO2 end

procedure TOPCGroup.Initialize;
begin
 inherited Initialize;
 FIConnectionPoints.CreateConnectionPoint(IID_IOPCDataCallback,ckMulti,CallBackOnConnect);
end;

constructor TOPCGroup.Create(serv:TDA2;oList:TList);
begin
 FIConnectionPoints:=TConnectionPoints.Create(self);
 Inherited Create;
 servObj:=serv;
 ownList:=oList;
 clItems:=TList.Create;
 asyncList:=TList.Create;
 upStream:=TMemoryStream.Create;
 onDataChangeEnabled:=true;
end;

destructor TOPCGroup.Destroy;
var
 i:integer;
begin
 for i:= 0 to clItems.count-1 do
  TOPCItem(clItems[i]).Free;
 clItems.Free;

 for i:= 0 to asyncList.count-1 do
  TAsyncIO2(asyncList[i]).Free;

 asyncList.Free;
 if Assigned(FIConnectionPoints) then                    FIConnectionPoints.Free;
 if Assigned(upStream) then                              upStream.Free;

 servObj.GroupRemovingSelf(ownList,serverHandle);
end;

function TOPCGroup.ValidateRequestedUpDateRate(dwRequestedUpdateRate:DWORD):DWORD;
begin
 if (dwRequestedUpdateRate < 199) then
  requestedUpdateRate:=1000
 else
  requestedUpdateRate:=dwRequestedUpdateRate;
 nextUpdate:=0;
 result:=requestedUpdateRate;
end;

procedure TOPCGroup.ValidateTimeBias(pTimeBias:PLongint);
var
 timeZoneRec:TTimeZoneInformation;
begin
 if not assigned(pTimeBias) then
  begin
   GetTimeZoneInformation(timeZoneRec);
   timeBias:=timeZoneRec.bias;
  end
 else
 timeBias:=pTimeBias^;
end;


procedure TOPCGroup.SetUp(szName:string;bActive:BOOL;dwRequestedUpdateRate:DWORD;
                hClientGroup:OPCHANDLE; pTimeBias:longint; pPercentDeadband:single;
                dwLCID:DWORD;phServerGroup:longword);
begin
 tagName:=szName;
 groupActive:=bActive;
 ValidateRequestedUpDateRate(dwRequestedUpdateRate);
 clientHandle:=hClientGroup;
 percentDeadband:=pPercentDeadband;
 lang:=dwLCID;
 serverHandle:=phServerGroup;
 groupPublic:=false;
 servObj.lastClientUpdate:=Now;
 onDataChangeEnabled:=true;
 lastMSecUpdate:=TimeStampToMSecs(DateTimeToTimeStamp(Now));
end;

procedure TOPCGroup.CallBackOnConnect(const Sink: IUnknown; Connecting: Boolean);
begin
 if connecting then
  ClientIUnknown:=Sink
 else
  ClientIUnknown:=nil;
end;

function TOPCGroup.GetItemIndexFromServerHandle(servHand:longword;var index:integer):boolean;
var
 i:integer;
begin
 result:=false;
 for i:= 0 to clItems.count-1 do
  if TOPCItem(clItems[i]).serverItemNum = servHand then
   begin
    index:=i;
    result:=true;
    Break;
   end;
end;

function TOPCGroup.GetItemIndexFromClientHandle(clHand:longword;var index:integer):boolean;
var
 i:integer;
begin
 result:=false;
 for i:= 0 to clItems.count-1 do
  if TOPCItem(clItems[i]).clientNum = clHand then
   begin
    index:=i;
    result:=true;
    Break;
   end;
end;

function TOPCGroup.GenerateAsyncCancelID:longword;
const
 cancelIndex:integer = 1;            //Assignable Typed Constants gota lovem
begin
 cancelIndex:=succ(cancelIndex);
 result:=cancelIndex;
end;

procedure TOPCGroup.TimeSlice(cTime:TDateTime);
var
 i:longword;
 curComp:comp;
begin
 if clItems = nil then Exit;
 if (clItems.count = 0) then Exit;

 AsyncTimeSlice(cTime);

 curComp:=TimeStampToMSecs(DateTimeToTimeStamp(cTime));
 if curComp >= (lastMSecUpdate + requestedUpdateRate) then
  begin
   nextUpdate:=0;
   upStream.Clear;
   upStream.Seek(0,soFromBeginning);
   for i:= 0 to clItems.count-1 do
    if TOPCItem(clItems[i]).bActive then begin
//     if
      TOPCItem(clItems[i]).UpdateYourSelf  ;
//     then
      upStream.Write(i,sizeOf(i));
    end;

   if ClientIUnknown <> nil then
    if onDataChangeEnabled and (upStream.position <> 0) then
     DoAChangeOccured(upStream, cTime);
   lastMSecUpdate:=curComp;
  end;
end;

procedure TOPCGroup.CloneYourSelf(dGrp:TOPCGroup);
var
 i:integer;
 wItem:TOPCItem;
begin
 dGrp.ClientIUnknown:=nil;
 dGrp.serverHandle:=servObj.GetNewGroupNumber;
 dGrp.servObj:=servObj;

 dGrp.clientHandle:=clientHandle;
 dGrp.requestedUpdateRate:=requestedUpdateRate;
 dGrp.lang:=lang;
 dGrp.groupActive:=false;
 dGrp.timeBias:=timeBias;
 dGrp.percentDeadband:=percentDeadband;
 dGrp.lastMSecUpdate:=lastMSecUpdate;
 if clItems.count > 0 then
  begin
   for i:=0 to clItems.count-1 do
    begin
     wItem:=TOPCItem.Create;
     TOPCItem(clItems[i]).CopyYourSelf(wItem);
     dGrp.clItems.Add(wItem);
    end;
  end;
end;

procedure TOPCGroup.DoAChangeOccured(aStream:TMemoryStream; cTime:TDateTime);
var
 i:integer;
 aAsyncObj:TAsyncIO2;
begin
 i:=aStream.position div 4;
 aAsyncObj:=nil;
 try
  aAsyncObj:=TAsyncIO2.Create(self,io2Change,0,i,OPC_DS_CACHE);
  aAsyncObj.aStream:=aStream;
  aAsyncObj.HandleThisRequest(cTime);
  finally
   if Assigned(aAsyncObj) then
    FreeAndNil(aAsyncObj);
  end;
end;

procedure TOPCGroup.AsyncTimeSlice(cTime:TDateTime);
var
 i:integer;
 aAsyncObj:TAsyncIO2;
begin
 if (requestedUpdateRate = 0) or
    (clItems.count = 0)       or
    (ClientIUnknown =  nil)   then
  begin
   for i:= asyncList.count-1 downTo 0 do
    TAsyncIO2(asyncList[i]).Free;
   Exit;
  end;

 aAsyncObj:=nil;
 if (asyncList <> nil) and (asyncList.count > 0) then
  begin
   for i:= 0 to asyncList.count - 1 do
    begin
     try
      aAsyncObj:=TAsyncIO2(asyncList[i]);
      aAsyncObj.HandleThisRequest(cTime);
     finally
      FreeAndNil(aAsyncObj);
     end;
    end;
   asyncList.Clear;
  end;
end;

initialization
  TTypedComObjectFactory.Create(ComServer, TOPCGroup, Class_OPCGroup,
    ciMultiInstance, tmApartment);
end.
