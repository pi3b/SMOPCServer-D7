unit Main;

{$IFDEF VER150}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

interface

uses Windows,Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
     RegDeRegServer,AxCtrls,ExtCtrls,ShutDownRequest, ImgList, ExtCtrlsX,
  Menus;

type
  TForm1 = class(TForm)
    PulseTimer: TTimer;
    DateTimeLbl: TLabel;
    Panel1: TPanel;
    Label2: TLabel;
    ClientConLbl: TLabel;
    GrpCountLbl: TLabel;
    Label1: TLabel;
    trycn1: TTrayIcon;
    il1: TImageList;
    pm1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    btnHide: TButton;
    SDButton: TButton;
    CloseBtn: TButton;
    RegSerBtn: TButton;
    UnRegBtn: TButton;
    btn1About: TButton;
    lbl1: TLabel;
    procedure RegSerBtnClick(Sender: TObject);
    procedure CloseBtnClick(Sender: TObject);
    procedure SDButtonClick(Sender: TObject);
    procedure PulseTimerTimer(Sender: TObject);
    procedure UnRegBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N3Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure btnHideClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure btn1AboutClick(Sender: TObject);
    procedure FormMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  public
   clientsConnected:integer;
   function ReturnItemValue(szItemID:string):variant;     
   procedure UpdateGroupCount;
  end;

var
  Form1: TForm1;


implementation

{$R *.DFM}

uses ServImpl,OPCCOMN,Globals,ActiveX;

const
 serverName = 'SMOPC.DA2';

procedure TForm1.RegSerBtnClick(Sender: TObject);
begin
 RegisterTheServer(serverName);
end;

procedure TForm1.CloseBtnClick(Sender: TObject);
begin
// Close;
 N3Click(nil);
end;

procedure TForm1.SDButtonClick(Sender: TObject);
var
 i:integer;
 Obj: Pointer;
begin
 ShutDownDlg:=nil;
 try
  Application.CreateForm(TShutDownDlg,ShutDownDlg);
  for i:= low(theServers) to high(theServers) do
   begin
    ShutDownDlg.RadioGroup1.controls[i].enabled:=(theServers[i] <> nil);
    if ShutDownDlg.RadioGroup1.controls[i].enabled then
     ShutDownDlg.RadioGroup1.itemIndex:=i;        //select the last one why...
   end;                                           //just so one is selected ;)
  if ShutDownDlg.ShowModal = mrOk then
   begin
    i:=ShutDownDlg.RadioGroup1.itemIndex;
    if theServers[i] <> nil then
     if theServers[i].ClientIUnknown <> nil then
      if Succeeded(theServers[i].ClientIUnknown.QueryInterface(IOPCShutdown,Obj)) then
       IOPCShutdown(Obj).ShutdownRequest('SMOPC IOPCShutdown request.')
      else
       ShowMessage('The client does not support IOPCShutdown.');
   end;
  ShutDownDlg.Release;
 finally
 end;
end;

procedure TForm1.PulseTimerTimer(Sender: TObject);
var
 i:integer;
 cTime:TDateTime;
begin
 try
  PulseTimer.enabled:=false;
  cTime:=Now;

 //0     complete time and date
 //1     complete date
 //2     day
 //3     month
 //4     year
 //5     complete time
 //6     hour
 //7     minute
 //8     second
 //9     millisecond
 //10    complete Date Inverted
 //11    day Inverted
 //12    month Inverted
 //13    year Inverted
 //14    complete time Inverted
 //15    hour Inverted
 //16    minute Inverted
 //17    second Inverted
 //18    millisecond Inverted
 //19    Test_Tag_1
 //20    Test_Tag_1 Inverted
 //21    Test_Tag_2
 //22    Test_Tag_2 Inverted

  DecodeDate(cTime,itemValues[4],itemValues[3],itemValues[2]);
  DecodeTime(cTime,itemValues[6],itemValues[7],itemValues[8],itemValues[9]);
  DateTimeLbl.Caption:=TimeToStr(cTime) + ' ' + DateToStr(cTime);
  for i:= low(theServers) to high(theServers) do
   if theServers[i] <> nil then
    theServers[i].TimeSlice(cTime);
 finally
  PulseTimer.enabled:=true;
 end;
end;

function TForm1.ReturnItemValue(szItemID:string):variant;
var
 s1:string;   ii:Integer;  v:LongInt;
begin
 if(IsSimulate(szItemID)) then begin
     ii:=ReturnSimuItemIndex(szItemID);
     case ii of
      0:       result:=DateTimeLbl.Caption;
      1:
       begin
        s1:=DateToStr(EncodeDate(itemValues[4],itemValues[3],
                                            itemValues[2]));
        result:=s1;
       end;
      5:       result:=TimeToStr(EncodeTime(itemValues[6],itemValues[7],
                                            itemValues[8],itemValues[9]));
      10:      result:=IntToStr(itemValues[13]) + '/' +
                       IntToStr(itemValues[12]) + '/' +
                       IntToStr(itemValues[11]);
      14:      result:=IntToStr(itemValues[15]) + ':' +
                       IntToStr(itemValues[16]) + ':' +
                       IntToStr(itemValues[17]) + ':' +
                       IntToStr(itemValues[18]);
      else     result:=itemValues[ii];
     end;
 end else begin
    result:=devices.GetItemValue(szItemID);
 end;
end;

procedure TForm1.UnRegBtnClick(Sender: TObject);
begin
 UnRegisterTheServer(serverName);
end;

procedure TForm1.UpdateGroupCount;
var
 i,g:integer;
begin
//  if not Visible  then Exit;
 if Application.Terminated then Exit;
 clientsConnected:=0;
 g:=0;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   begin
    clientsConnected:=succ(clientsConnected);
    if Assigned(theServers[i].grps) then
     g:=g + theServers[i].grps.count;
    if Assigned(theServers[i].pubGrps) then
     g:=g + theServers[i].pubGrps.count;
   end;

 Form1.ClientConLbl.caption:=IntToStr(clientsConnected);
 GrpCountLbl.caption:=IntToStr(g);
 SDButton.enabled:=(clientsConnected  > 0);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 s1:string;
begin
 if ParamCount <> 0 then
  begin
   s1:=LowerCase(ParamStr(1));
   if (s1 = 'regserver') or (s1 = 'register') then
    begin
     RegSerBtnClick(self);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end
   else if (s1 = 'unregserver') or (s1 = 'unregister') then
    begin
     UnRegBtnClick(self);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end;
  end;

 UpdateGroupCount;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
 if clientsConnected > 0 then
  CanClose:=(MessageDlg('Clients are connected. Are you sure you want to quit?',
                        mtConfirmation,[mbYes,mbNo],0) =  mrYes)
 else
  CanClose:=true;
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  trycn1.Visible:=False;
  Application.ProcessMessages;
  ExitProcess(0);
end;

procedure TForm1.N2Click(Sender: TObject);
begin
  if MessageBox(0,'是否打开维护界面？','确认',MB_YESNO)= ID_YES then
    Visible:=True;
end;

procedure TForm1.btnHideClick(Sender: TObject);
begin
  Visible:=false;
end;

procedure TForm1.N1Click(Sender: TObject);
begin
  ShowMessage('本程序为完全免费的OPC服务器，支持OPC DA2.0。'#13#10'配合StateManager可以快速对接各种PLC等自定义数据来源。'
  +#13#10'更多信息请访问：https://pi3b.github.io');
end;

procedure TForm1.btn1AboutClick(Sender: TObject);
begin
N1Click(nil);
end;

procedure TForm1.FormMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
ReleaseCapture;
 Perform(WM_SYSCOMMAND, $F017 , 0);
end;

end.
