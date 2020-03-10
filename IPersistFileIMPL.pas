function TDA2.GetClassID(out classID:TCLSID):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.IsDirty:HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Load(pszFileName:POleStr; dwMode:Longint):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Save(pszFileName:POleStr; fRemember:BOOL):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.SaveCompleted(pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.GetCurFile(out pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;

