program CalcCodeWork;

uses
  Dialogs,
  MapBasic_INT in '..\MapBasic_INT.pas',
  Common in '..\Common.pas',
  CalcCodePart in 'CalcCodePart.pas';

{$R *.res}

begin
  if (ParamStr(1) = '') OR (ParamStr(2) = '') OR (ParamStr(3) = '') then
    ShowMessage('Это исполняемый файл с параметрами')
  else
    Work(ParamStr(1), ParamStr(2), ParamStr(3), nil);
end.

