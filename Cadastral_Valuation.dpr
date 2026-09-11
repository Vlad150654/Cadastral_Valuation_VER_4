program Cadastral_Valuation;

uses
  Forms,
  CodeNP_Unit in 'CodeNP_Unit.pas' {mainForm},
  Compliance_Unit in 'Compliance_Unit.pas' {ComplianceForm},
  Param_Unit in 'Param_Unit.pas' {ParamForm},
  MapBasic_INT in '..\MapBasic_INT.pas',
  Common in '..\Common.pas';

{$R *.res}

begin
  {$IFDEF DEBUG} ReportMemoryLeaksOnShutdown := True; {$ENDIF}

  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TmainForm, mainForm);
  Application.CreateForm(TComplianceForm, ComplianceForm);
  Application.CreateForm(TParamForm, ParamForm);
  Application.Run;
end.
