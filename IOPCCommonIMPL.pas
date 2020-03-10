function TDA2.SetLocaleID(dwLcid:TLCID):HResult;stdcall;
begin
 if (dwLcid = LOCALE_SYSTEM_DEFAULT) or (dwLcid = LOCALE_USER_DEFAULT) then
  begin
   localID:=dwLcid;
   result:=S_OK;
  end
 else
  result:=E_INVALIDARG;
end;

function TDA2.GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
begin
 pdwLcid:=localID;
 result:=S_OK;
end;

function TDA2.QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
begin
 pdwCount:=2;
 pdwLcid:=PLCIDARRAY(CoTaskMemAlloc(pdwCount*sizeof(LCID)));
 if (pdwLcid = nil) then
  begin
   if pdwLcid <> nil then  CoTaskMemFree(pdwLcid);
   result:=E_OUTOFMEMORY;
   Exit;
  end;
 pdwLcid[0]:=LOCALE_SYSTEM_DEFAULT;
 pdwLcid[1]:=LOCALE_USER_DEFAULT;
 result:=S_OK;
end;

function TDA2.GetErrorString(dwError:HResult; out ppString:POleStr):HResult;stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.SetClientName(szName:POleStr):HResult;stdcall;
begin
 if (addr(szName) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 clientName:=szName;
 result:=S_OK;
end;

