function TDA2.GetPublicGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult;stdcall;
begin
 if IsGroupNamePresent(pubGrps,szName) <> -1 then
  begin
   result:=S_OK;
   ppUnk:=self;
  end
 else
 result:=OPC_E_NOTFOUND;
end;

function TDA2.RemovePublicGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
begin
 //do something with the forse if needed
 GroupRemovingSelf(pubGrps,hServerGroup);
 result:=S_OK;
end;
