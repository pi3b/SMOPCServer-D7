unit ItemPropIMPL;

{$IFDEF VER150}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses Windows,ActiveX,ComObj,SysUtils,Dialogs,Classes,OPCDA,axctrls,
     OpcError,OPCTypes;

type
  TOPCItemProp = class
  public
   function QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                          out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                          out ppvtDataTypes:PVarTypeList):HResult;stdcall;
   function GetItemProperties(szItemID:POleStr;
                              dwCount:DWORD;
                              pdwPropertyIDs:PDWORDARRAY;
                              out ppvData:POleVariantArray;
                              out ppErrors:PResultList):HResult;stdcall;
   function LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                           out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
end;

implementation

uses ServIMPL,GLobals,Main;


function TOPCItemProp.QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                       out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                       out ppvtDataTypes:PVarTypeList):HResult;stdcall;
var
 memErr:boolean;
 propID:longword;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 pdwCount:=1;
 memErr:=false;
 ppPropertyIDs:=PDWORDARRAY(CoTaskMemAlloc(pdwCount*sizeof(DWORD)));
 if ppPropertyIDs = nil then
  memErr:=true;
 if not memErr then
  ppDescriptions:=POleStrList(CoTaskMemAlloc(pdwCount*sizeof(POleStr)));
 if ppDescriptions = nil then
  memErr:=true;
 if not memErr then
  ppvtDataTypes:=PVarTypeList(CoTaskMemAlloc(pdwCount*sizeof(TVarType)));
 if ppvtDataTypes = nil then
  memErr:=true;

 if memErr then
  begin
   if ppPropertyIDs <> nil then  CoTaskMemFree(ppPropertyIDs);
   if ppDescriptions <> nil then  CoTaskMemFree(ppDescriptions);
   if ppvtDataTypes <> nil then  CoTaskMemFree(ppvtDataTypes);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppPropertyIDs[0]:=propID;
 ppDescriptions[0]:=StringToLPOLESTR(szItemID);
 ppvtDataTypes[0]:=ReturnDataTypeFromPropID(szItemID);
 result:=S_OK;
end;

function TOPCItemProp.GetItemProperties(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppvData:POleVariantArray; out ppErrors:PResultList):HResult;stdcall;
var
 data:variant;  pV:Integer; s:string;
 i:integer;
 memErr:boolean;
 propID:longword;
 ppArray:PDWORDARRAY;
begin
 propID:=ReturnPropIDFromTagname(szItemID);

 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppvData:=POleVariantArray(CoTaskMemAlloc(dwCount * sizeof(OleVariant)));
 if ppvData = nil then
  memErr:=true;
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount * sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppvData <> nil then  CoTaskMemFree(ppvData);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppArray:=@pdwPropertyIDs^;
 result:=S_OK;
 for i:= 0 to dwCount-1 do
  begin
   pV:= ppArray[i];
   case pV of
    OPC_PROPERTY_DATATYPE:  data:=ReturnDataTypeFromPropID(szItemID);   //ppArray[i]
    OPC_PROPERTY_VALUE:  begin
          s:=Form1.ReturnItemValue(szItemID);
          data:=s;
        end;
    OPC_PROPERTY_ACCESS_RIGHTS:
     begin
      if CanPropIDBeWritten(szItemID) then
       data:=OPC_READABLE or OPC_WRITEABLE
      else
       data:=OPC_READABLE;
     end;
    //5000 - 5022
//    posItems[low(posItems)].PropID..posItems[high(posItems)].PropID:
    5000..5022:         begin
          s:=Form1.ReturnItemValue(szItemID);
          data:=s;
        end;

    else
     begin
      ppErrors[i]:=OPC_E_INVALID_PID;
      result:=S_FALSE;
      Continue;
     end;
   end;
   try
     ppvData[i]:=data;
     ppErrors[i]:=S_OK;
   except
   end;
  end;

end;

function TOPCItemProp.LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
var
 i:integer;
 propID:longword;
 memErr:boolean;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppszNewItemIDs:=POleStrList(CoTaskMemAlloc(dwCount*sizeof(POleStr)));
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppszNewItemIDs <> nil then  CoTaskMemFree(ppszNewItemIDs);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 for i:= 0 to dwCount-1 do
  begin
   ppszNewItemIDs[i]:=StringToLPOLESTR(szItemID);
   ppErrors[i]:=S_OK;
  end;

 result:=S_OK;
end;

end.
