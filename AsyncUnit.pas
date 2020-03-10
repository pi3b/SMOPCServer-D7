unit AsyncUnit;

{$IFDEF VER150}
{$WARN UNSAFE_CAST OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses Windows,ActiveX,ComObj,SMOPC_TLB,SysUtils,Dialogs,Classes,ServIMPL,OPCDA,
     axctrls,Globals,GroupUnit,ItemsUnit,OpcError,OPCTypes,Variants;

type
  TAsyncIO2 = class
  public
   grp:TOPCGroup;
   source:integer;
   isCancelled:boolean;
   ppServer:PDWORDARRAY;
   ppValues:POleVariantArray;
   kind,clientTransID,cancelID,itemCount:longword;
   aStream:TMemoryStream;
   constructor Create(aGrp:TOPCGroup;ioKind,cID,count:longword;dwSource:integer);
   destructor Destroy;override;
   procedure HandleRead(cTime:TDateTime);
   procedure HandleWrite(cTime:TDateTime);
   procedure HandleRefresh(cTime:TDateTime);
   procedure HandleChange(aStream:TMemoryStream; cTime:TDateTime);
   procedure HandleThisRequest(cTime:TDateTime);
   function AddItems(phServer:POPCHANDLE):boolean;
   function AddValues(pItemValues:POleVariant):boolean;
  end;

implementation

type
 WORDARRAY = array[0..65535] of WORD;
 PWORDARRAY = ^WORDARRAY;

type
 TFileTimeARRAY = array[0..65535] of TFileTime;
 PTFileTimeARRAY = ^TFileTimeARRAY;

constructor TAsyncIO2.Create(aGrp:TOPCGroup;ioKind,cID,count:longword;dwSource:integer);
begin
 grp:=aGrp;
 cancelID:=grp.GenerateAsyncCancelID;
 kind:=ioKind;
 clientTransID:=cID;
 itemCount:=count;
 ppServer:=nil;
 isCancelled:=false;
 source:=dwSource
end;

destructor TAsyncIO2.Destroy;
begin
 if ppServer <> nil then
  FreeMem(ppServer);
 ppServer:=nil;
 if ppValues <> nil then
  FreeMem(ppValues);
 ppValues:=nil;
end;

procedure TAsyncIO2.HandleRead(cTime:TDateTime);
var
 i:longword;
 Obj:Pointer;
 aItem:TOPCItem;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;
 ppTimeArray:PTFileTimeARRAY;
 masterResult,masterQuality:HRESULT;
begin
 if not Succeeded(grp.ClientIUnknown.QueryInterface(IOPCDataCallback,Obj)) then Exit;
 ppClientItems:=nil;           pVariants:=nil;
 ppErrors:=nil;                ppQualityArray:=nil;
 ppTimeArray:=nil;
 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(longword)));
  if ppClientItems = nil then Exit;

  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount*sizeof(OleVariant)));
  if pVariants = nil then Exit;

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));
  if ppErrors = nil then Exit;

  ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(word)));
  if ppQualityArray = nil then Exit;

  ppTimeArray:=PTFileTimeARRAY(CoTaskMemAlloc(itemCount*sizeof(TFileTime)));
  if ppTimeArray = nil then Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;
  masterQuality:=S_OK;
  for i:= 0 to itemCount-1 do
   begin
    ppTimeArray[i]:=aFileTime;
    pVariants[i]:=VT_EMPTY;
    ppClientItems[i]:=0;
    ppQualityArray[i]:=OPC_QUALITY_BAD;

    aItem:=TOPCItem(TOPCGroup(grp).clItems[ppServer[i]]);

{    if aItem.ReturnAccessRights <> OPC_READABLE then
     begin
      ppErrors[i]:=OPC_E_BADRIGHTS;
      masterResult:=S_FALSE;
      masterQuality:=S_FALSE;
      Continue;
     end;
}

//    aItem.CallBackRead(ppClientItems[i],pVariants[i], ppQualityArray[i]);
     ppClientItems[i]:=aItem.GetClientHandle;
     pVariants[i]:=aItem.ReturnCurrentValue(0);
     ppQualityArray[i]:=aItem.quality;

    if ppQualityArray[i] <> OPC_QUALITY_GOOD then
     masterQuality:=S_FALSE;
    ppErrors[i]:=S_OK;
   end;

  if isCancelled then Exit;

  IOPCDataCallback(Obj).OnReadComplete(clientTransID,
                                       grp.clientHandle,
                                       masterQuality,
                                       masterResult,
                                       itemCount,
                                       @ppClientItems^,
                                       @pVariants^,
                                       @ppQualityArray^,
                                       @ppTimeArray^,
                                       @ppErrors^);

 finally
  if ppClientItems <> nil then  CoTaskMemFree(ppClientItems);
  if pVariants <> nil then      CoTaskMemFree(pVariants);
  if ppErrors <> nil then       CoTaskMemFree(ppErrors);
  if ppQualityArray <> nil then CoTaskMemFree(ppQualityArray);
  if ppTimeArray <> nil then    CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleWrite(cTime:TDateTime);
var
 Obj:Pointer;
 aItem:TOPCItem;
 ppErrors:PResultList;
 i,masterResult:longword;
 ppClientItems:PDWORDARRAY;
begin
 if not Succeeded(grp.ClientIUnknown.QueryInterface(IOPCDataCallback,Obj)) then Exit;
 ppClientItems:=nil;           ppErrors:=nil;
 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(longword)));
  if ppClientItems = nil then Exit;

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));
  if ppErrors = nil then Exit;

  masterResult:=S_OK;
  for i:= 0 to itemCount-1 do
   begin
    if (ppServer[i] > longword((grp.clItems.count-1))) then
     begin
      ppErrors[i]:=OPC_E_INVALIDITEMID;
      masterResult:=S_FALSE;
      Continue;
     end;

    aItem:=TOPCItem(grp.clItems[ppServer[i]]);
    ppClientItems[i]:=aItem.GetClientHandle;
    if not aItem.isWriteAble then
     begin
      ppErrors[i]:=OPC_E_BADRIGHTS;
      masterResult:=S_FALSE;
      Continue;
     end;
    aItem.WriteItemValue(ppValues[i]);
    ppErrors[i]:=S_OK;
   end;

  if isCancelled then Exit;

  IOPCDataCallback(Obj).OnWriteComplete(clientTransID,
                                        grp.clientHandle,
                                        masterResult,
                                        itemCount,
                                        @ppClientItems^,
                                        @ppErrors^);
 finally
  if ppClientItems <> nil then  CoTaskMemFree(ppClientItems);
  if ppErrors <> nil then       CoTaskMemFree(ppErrors);
 end;
end;

procedure TAsyncIO2.HandleRefresh(cTime:TDateTime);
var
 x:integer;
 Obj:Pointer;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 i,masterResult:longword;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;
 ppTimeArray:PTFileTimeARRAY;
begin
 if not Succeeded(grp.ClientIUnknown.QueryInterface(IOPCDataCallback,Obj)) then
  Exit;
 ppClientItems:=nil;           pVariants:=nil;
 ppErrors:=nil;                ppQualityArray:=nil;
 ppTimeArray:=nil;

 try
  itemCount:=0;
  for i:= 0 to grp.clItems.count-1 do
   if TOPCItem(grp.clItems[i]).GetActiveState then
    itemCount:=succ(itemCount);

  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(longword)));
  if ppClientItems = nil then Exit;

  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount*sizeof(OleVariant)));
  if pVariants = nil then Exit;

  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));
  if ppErrors = nil then Exit;

  ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(word)));
  if ppQualityArray = nil then Exit;

  ppTimeArray:=PTFileTimeARRAY(CoTaskMemAlloc(itemCount*sizeof(TFileTime)));
  if ppTimeArray = nil then Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;
  x:=0;
  if (clientTransID <> 0) and (source = OPC_DS_DEVICE) then
   for i:= 0 to grp.clItems.count-1 do
    if TOPCItem(grp.clItems[i]).GetActiveState then
     TOPCItem(grp.clItems[i]).UpdateYourSelf;

  for i:= 0 to grp.clItems.count-1 do
   if TOPCItem(grp.clItems[i]).GetActiveState then
    begin
     ppTimeArray[x]:=aFileTime;
     ppClientItems[x]:=TOPCItem(grp.clItems[i]).GetClientHandle;
     pVariants[x]:=TOPCItem(grp.clItems[i]).ReturnCurrentValue(source);
     ppQualityArray[x]:=TOPCItem(grp.clItems[i]).GetQuality;
     ppErrors[x]:=S_OK;
     x:=succ(x);
    end;

  if isCancelled then Exit;
  IOPCDataCallback(Obj).OnDataChange(clientTransID,
                                     grp.clientHandle,
                                     OPC_QUALITY_GOOD,
                                     masterResult,
                                     itemCount,
                                     @ppClientItems^,
                                     @pVariants^,
                                     @ppQualityArray^,
                                     @ppTimeArray^,
                                     @ppErrors^);

 finally
  if ppClientItems <> nil then  CoTaskMemFree(ppClientItems);
  if pVariants <> nil then      CoTaskMemFree(pVariants);
  if ppErrors <> nil then       CoTaskMemFree(ppErrors);
  if ppQualityArray <> nil then CoTaskMemFree(ppQualityArray);
  if ppTimeArray <> nil then    CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleChange(aStream:TMemoryStream; cTime:TDateTime);
var
 x,k:longword;
 Obj:Pointer;
 aFileTime:TFileTime;
 ppErrors:PResultList;
 i,masterResult:longword;
 ppQualityArray:PWORDARRAY;
 ppClientItems:PDWORDARRAY;
 pVariants:POleVariantArray;  V,V1:olevariant;
 ppTimeArray:PFileTimeARRAY;
 item:TOPCItem;
begin
 if not Succeeded(TOPCGroup(grp).ClientIUnknown.QueryInterface(IOPCDataCallback,Obj)) then
  Exit;

 ppClientItems:=nil;           pVariants:=nil;
 ppErrors:=nil;                ppQualityArray:=nil;
 ppTimeArray:=nil;
 try
  ppClientItems:=PDWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(longword)));
  if ppClientItems = nil then Exit;
  
  pVariants:=POleVariantArray(CoTaskMemAlloc(itemCount*sizeof(OleVariant)));
  if pVariants = nil then Exit;
//  try
//    pVariants[0]:=varEmpty;
//  except
//    on E:Exception do  begin
////      MessageBox(0,PChar(E.Message),'',MB_OK);
//      exit;
//    end;
//  end;
  ppErrors:=PResultList(CoTaskMemAlloc(itemCount*sizeof(HRESULT)));

  if ppErrors = nil then Exit;
   ppQualityArray:=PWORDARRAY(CoTaskMemAlloc(itemCount*sizeof(word)));
  if ppQualityArray = nil then Exit;

  ppTimeArray:=PFileTimeARRAY(CoTaskMemAlloc(itemCount*sizeof(TFileTime)));
  if ppTimeArray = nil then Exit;

  DataTimeToOPCTime(cTime,aFileTime);
  masterResult:=S_OK;

  aStream.Seek(0,soFromBeginning);
  for i:=0 to itemcount-1 do
    begin
      aStream.Read(k,sizeOf(k));
//      if TOPCGroup(grp).clItems.count >= k then
      item:=TOPCItem(TOPCGroup(grp).clItems[k]);
       try
         ppTimeArray[i]:=aFileTime;
//         CallBackRead(ppClientItems[x],pVariants[x], ppQualityArray[x]);

         ppClientItems[i]:=item.GetClientHandle;
//         V:= 'abc';
//         V1:=item.ReturnCurrentValue(0);
//         pVariants[i]:=item.ReturnCurrentValue(0);
         pVariants[i]:=VarAsType(item.ReturnCurrentValue(0),item.vtReqDataType);
         ppQualityArray[i]:=item.quality;
         ppErrors[i]:=S_OK;
         //上面异常后不能回调OnDataChange，否则出错
         if isCancelled then Exit;
         IOPCDataCallback(Obj).OnDataChange(clientTransID,
                                            TOPCGroup(grp).clientHandle,
                                            OPC_QUALITY_GOOD,
                                            masterResult,
                                            itemcount,
                                            @ppClientItems^,
                                            @pVariants^,
                                            @ppQualityArray^,
                                            @ppTimeArray^,
                                            @ppErrors^);

       except
//         ShowMessage('error on handle change :'+item.strID);
         ppTimeArray[i]:=aFileTime;
         ppClientItems[i]:=item.GetClientHandle;
//         pVariants[i]:=0; //不管什么值都出错 why ?
         ppQualityArray[i]:=OPC_QUALITY_OUT_OF_SERVICE;
         ppErrors[i]:=S_FALSE;
       end;
     end;

 finally
  if ppClientItems <> nil then  CoTaskMemFree(ppClientItems);
  if pVariants <> nil then      CoTaskMemFree(pVariants);
  if ppErrors <> nil then       CoTaskMemFree(ppErrors);
  if ppQualityArray <> nil then CoTaskMemFree(ppQualityArray);
  if ppTimeArray <> nil then    CoTaskMemFree(ppTimeArray);
 end;
end;

procedure TAsyncIO2.HandleThisRequest(cTime:TDateTime);
begin
 if isCancelled then Exit;
 case kind of
  io2Read:         HandleRead(cTime);
  io2Write:        HandleWrite(cTime);
  io2Refresh:      HandleRefresh(cTime);
  io2Change:       HandleChange(aStream, cTime)
 end;
end;

function TAsyncIO2.AddItems(phServer:POPCHANDLE):boolean;
var
 i:longword;
begin
 result:=false;
 i:=itemCount*sizeof(longword);
 try
  GetMem(ppServer,i);
 except
  result:=true;
  Exit;
 end;
 Move(phServer^,ppServer^,i);
end;

function TAsyncIO2.AddValues(pItemValues:POleVariant):boolean;
var
 i:longword;
begin
 result:=false;
 i:=itemCount*sizeof(OleVariant);
 try
  GetMem(ppValues,i);
 except
  result:=true;
  Exit;
 end;
 Move(pItemValues^,ppValues^,i);
end;

end.

