unit Globals;

{$IFDEF VER150}
{$WARN UNSAFE_CODE OFF}
{$ENDIF}

interface

uses
{$IFDEF VER140}
  Variants,
{$ENDIF}
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Windows,Messages,SysUtils,Classes,Graphics,StdCtrls,Forms,Dialogs,Controls,
  ShellAPI,ActiveX,OPCDA,SMOPCDevice_TLB;

type
 itemIDStrings = record
  trunk,branch,leaf:string[255];
 end;

type
  itemProps = record
    PropID: longword;
    tagname: string[64];
    dataType:integer;
    deviceName:string[64];
  end;

const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';

 posItems: array[0..22] of itemProps =
  ((PropID: 5000; tagname: 'Simulate.Complete';                        dataType:VT_BSTR; deviceName:'Simulate'),
   (PropID: 5001; tagname: 'Simulate.Date.Complete';                   dataType:VT_BSTR; deviceName:'Simulate'),
   (PropID: 5002; tagname: 'Simulate.Date.Parts.Day';                  dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5003; tagname: 'Simulate.Date.Parts.Month';                dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5004; tagname: 'Simulate.Date.Parts.Year';                 dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5005; tagname: 'Simulate.Time.Complete';                   dataType:VT_BSTR; deviceName:'Simulate'),
   (PropID: 5006; tagname: 'Simulate.Time.Parts.Hour';                 dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5007; tagname: 'Simulate.Time.Parts.Min';                  dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5008; tagname: 'Simulate.Time.Parts.Seconds';              dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5009; tagname: 'Simulate.Time.Parts.Millseconds';          dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5010; tagname: 'Simulate.Inverted.Date.Complete';         dataType:VT_BSTR; deviceName:'Simulate'),
   (PropID: 5011; tagname: 'Simulate.Inverted.Date.Day';              dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5012; tagname: 'Simulate.Inverted.Date.Month';            dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5013; tagname: 'Simulate.Inverted.Date.Year';             dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5014; tagname: 'Simulate.Inverted.Time.Complete';         dataType:VT_BSTR; deviceName:'Simulate'),
   (PropID: 5015; tagname: 'Simulate.Inverted.Time.Hour';             dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5016; tagname: 'Simulate.Inverted.Time.Min';              dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5017; tagname: 'Simulate.Inverted.Time.Seconds';          dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5018; tagname: 'Simulate.Inverted.Time.Millseconds';      dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5019; tagname: 'Simulate.Test_Tag_1.Actual';              dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5020; tagname: 'Simulate.Test_Tag_1.Inverted';            dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5021; tagname: 'Simulate.Test_Tag_2.Actual';              dataType:VT_UI2; deviceName:'Simulate'),
   (PropID: 5022; tagname: 'Simulate.Test_Tag_2.Inverted';            dataType:VT_UI2; deviceName:'Simulate'));


 io2Read = 1;
 io2Write = 2;
 io2Refresh = 3;
 io2Change      = 4;

var
 itemValues:array[0..22] of word;
 devices:TSMOPCDevice;

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
function ReturnPropIDFromTagname(const tagName:string):longword;
function ReturnTagnameFromPropID(PropID:longword):string;
//function CanPropIDBeWritten(i:longword):boolean;
function CanPropIDBeWritten(const tagName:string):boolean;
//function ReturnDataTypeFromPropID(i:longword):integer;
function ReturnDataTypeFromPropID(const tagName:string):integer;
procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
//function ConvertVariant(cv:variant; reqDataType:TVarType):variant;

function ReturnDeviceNameFromTagname(const TagName:string):string;
function IsSimulate(const TagName:string):boolean;         
function ReturnSimuItemIndex(const tagName:string):Integer;
implementation
                      
function ReturnSimuItemIndex(const tagName:string):Integer;
var i:Integer;
begin
   for i:= low(posItems) to high(posItems) do
    if posItems[i].tagname = tagName then
     begin
      result:=i;
      Exit;
     end;
end;
function ReturnDeviceNameFromTagname(const TagName:string):string;
var p :Integer;
begin
  p:=Pos('.',TagName);
  if p<0 then begin
    result:='';
    exit;
  end;
  result:=Copy(TagName,0,p-1);
end;
function IsSimulate(const TagName:string):boolean;
begin
  result:=ReturnDeviceNameFromTagname(TagName)='Simulate';
end;

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
var
 tempS:string;
 finish:boolean;
 nextloc,strLength: integer;
begin
 {$R-}
 strLength := length(theString);
 finish := false;
 SetLength(tempS,strLength);
 result := tempS;
 nextloc := 1;
 while not finish do
  begin
   if (start < 256) and (theString[start] <> theChar) and
      (theString[start] <> chr(13)) and (start <= strLength) then
    begin
     tempS[nextloc] := theString[start];
     nextloc := succ(nextloc);
     start := succ(start);
    end
   else
    begin
     SetLength(tempS,nextloc-1);      {this sets the length of the string}
     finish:=true;                    {exit the loop}
     result:=tempS;                   {return the value}
    end;
  end;
 {$R+}
end;

function ReturnPropIDFromTagname(const tagName:string):longword;
var
 i:integer;
begin
 result:=0;
   if IsSimulate(tagName) then begin
     for i:= low(posItems) to high(posItems) do
      if posItems[i].tagname = tagName then
       begin
        result:=posItems[i].PropID;
        Exit;
       end;
  end else begin
     result:=devices.ItemIDPropID(tagName);
  end;


end;

function ReturnTagnameFromPropID(PropID:longword):string;
var
 i:integer;
begin
 result:='';
   for i:= low(posItems) to high(posItems) do
    if posItems[i].PropID = PropID then
     begin
      result:=posItems[i].tagname;
      Exit;
     end;
     
     raise Exception.Create('没有实现TagnameFromPropID');
end;

//function CanPropIDBeWritten(i:longword):boolean;
//begin
// i:= i - posItems[low(posItems)].PropID;
// result:=boolean(i in [19..22]);               //the test Test_Tag_X's
//end;

function CanPropIDBeWritten(const tagName:string):boolean;
var i:Integer;
begin
  if(IsSimulate(tagName)) then begin
    i:=ReturnPropIDFromTagname(tagName);
   i:= i - posItems[low(posItems)].PropID;
   result:=boolean(i in [19..22]);               //the test Test_Tag_X's
  end else begin
    result:=devices.ItemIDWriteAble(tagName);
  end;
end;

//function ReturnDataTypeFromPropID(i:longword):integer;
//var
// x:longword;
//begin
// x:= i - posItems[low(posItems)].PropID;
// if (x <= high(posItems)) then
//  result:=posItems[x].dataType
// else
//  result:=VT_UI2;
//end;

function ReturnDataTypeFromPropID(const tagName:string):integer;
var
 x:longword;  i:Integer;
begin
  if IsSimulate(tagName) then begin  
     i:=ReturnPropIDFromTagname(tagName);
     x:= i - posItems[low(posItems)].PropID;
     if (x <= high(posItems)) then
      result:=posItems[x].dataType
     else
      result:=VT_UI2;
  end else begin
    result:=devices.ItemIDDataType(tagName);
  end;
end;
procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
var
 sTime:TSystemTime;
begin
 DateTimeToSystemTime(cTime,sTime);
 SystemTimeToFileTime(sTime,OPCTime);
 LocalFileTimeToFileTime(OPCTime,OPCTime);
end;

//function ConvertVariant(cv:variant; reqDataType:TVarType):variant;
//begin
// try
//  result:=VarAsType(cv,reqDataType);
// except
//  on EVariantError do   result:=DISP_E_TYPEMISMATCH;
// end;
//end;

initialization
  devices:=TSMOPCDevice.Create(nil);
finalization
  FreeAndNil(devices);
end.
