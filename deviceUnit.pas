unit deviceUnit;

interface
uses Windows,Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
     RegDeRegServer,AxCtrls,ExtCtrls,IniFiles,SMOPCDevice_TLB,superobject;

type
  TDevices = class(TStringList)
    public
      ini:TIniFile;
      jo:TSuperObject;
      constructor Create();
      destructor Destroy();
      function GetDeviceObject(deviceName:string):TSMOPCDevice;
  end;
//var
//  devices:TDevices;
implementation

{ TDevices }

constructor TDevices.Create;
var i:Integer; sections:TStrings; deviceName:string;  d:TSMOPCDevice;
begin
//  ini:=TIniFile.Create('devices.ini');
//  jo:=TSuperObject.Create();
//  jo.pa
//  sections := TStringList.Create;
//  try
//    ini.ReadSections(sections );
//    for i:=0 to sections.Count-1 do begin
//      deviceName:=sections[i];
//      d:=TSMOPCDevice.Create(nil);
//      d.DeviceType:=ini.ReadString(deviceName,'DeviceType','');
//      Self.AddObject(deviceName,d);
//    end;
//  finally
//    FreeAndNil(sections );
//  end;
end;

destructor TDevices.Destroy;
begin
//  FreeAndNil(ini);
end;

function TDevices.GetDeviceObject(deviceName: string): TSMOPCDevice;
begin
  result:=self.GetObject(self.IndexOf(deviceName)) as TSMOPCDevice;
end;

end.
 