unit RegDeRegServer;

interface

uses Windows,Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
     Registry,OPCDA,SMOPC_TLB;

procedure RegisterTheServer(name:string);
procedure UnRegisterTheServer(name:string);

implementation

uses ComObj,ComCat;

procedure RegisterTheServer(name:string);
var
 aReg:TRegistry;
 hr:HRESULT;
 myCLSIDString:string;
begin
 myCLSIDString:=GUIDToString(CLASS_DA2);
 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.OpenKey(name,true);
  aReg.WriteString('','SMOPC Data Access Server Version 2.0');
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.OpenKey(name+'\Clsid',true);
  aReg.WriteString('',myCLSIDString);
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

{       this is only needed for Data Access Version 1.0 according
        to the spec.

aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.OpenKey(name+'\OPC',true);
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

}

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.OpenKey('CLSID\'+myCLSIDString,true);
  aReg.WriteString('','SMOPC OPC Data Access 2.0');
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.OpenKey('CLSID\'+myCLSIDString+'\ProgID',true);
  aReg.WriteString('',name);
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 hr:=CreateComponentCategory(CATID_OPCDAServer20,'SMOPC OPC Data Access');
 if hr <> 0 then
  ;
 hr:=RegisterCLSIDInCategory(CLASS_DA2,CATID_OPCDAServer20);
 if hr <> 0 then
  ;
end;

procedure UnRegisterTheServer(name:string);
var
 aReg:TRegistry;
 hr:HRESULT;
 myCLSIDString:string;
begin
 myCLSIDString:=GUIDToString(CLASS_DA2);
 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.DeleteKey('CLSID\'+myCLSIDString+'\ProgID');
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.DeleteKey('CLSID\'+myCLSIDString);
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

{       this is only needed for Data Access Version 1.0 according
        to the spec.

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.DeleteKey(name+'\OPC');
 finally
  aReg.CloseKey;
  aReg.Free;
 end;
}

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.DeleteKey(name+'\Clsid');
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 aReg:=nil;
 try
  aReg:=TRegistry.Create;
  aReg.RootKey:=HKEY_CLASSES_ROOT;
  aReg.DeleteKey(name);
 finally
  aReg.CloseKey;
  aReg.Free;
 end;

 hr:=UnRegisterCLSIDInCategory(CLASS_DA2,CATID_OPCDAServer20);
 if hr <> 0 then
  ;

 hr:=UnCreateComponentCategory(CATID_OPCDAServer20,'SMOPC OPC Data Access');
 if hr <> 0 then
  ;
end;

end.
