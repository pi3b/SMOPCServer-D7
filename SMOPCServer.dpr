program SMOPCServer;

uses
  Forms,
  Main in 'Main.pas' {Form1},
  ServIMPL in 'ServIMPL.pas',
  RegDeRegServer in 'RegDeRegServer.pas',
  ComCat in 'comcat.pas',
  Enumstring in 'Enumstring.pas',
  GroupUnit in 'GroupUnit.pas',
  ItemPropIMPL in 'ItemPropIMPL.pas',
  Globals in 'Globals.pas',
  ItemsUnit in 'ItemsUnit.pas',
  ItemAttributesOPC in 'ItemAttributesOPC.pas',
  EnumItemAtt in 'EnumItemAtt.pas',
  AsyncUnit in 'AsyncUnit.pas',
  ShutDownRequest in 'ShutDownRequest.pas' {ShutDownDlg},
  OPCErrorStrings in 'OPCErrorStrings.pas',
  EnumUnknown in 'EnumUnknown.pas',
  OPCCOMN in 'OPCCOMN.pas',
  OPCDA in 'OPCDA.pas',
  OPCerror in 'OPCerror.pas',
  OPCtypes in 'OPCtypes.pas',
  SMOPC_TLB in 'SMOPC_TLB.pas',
  SMOPCServerDevice_TLB in '..\..\App\Delphi7\Imports\SMOPCServerDevice_TLB.pas';

{$R *.TLB}

{$R *.RES}

begin
  Application.Initialize;
  application.ShowMainForm:=False;
  Application.Title := 'SMOPCServer ·þÎñÆ÷';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
