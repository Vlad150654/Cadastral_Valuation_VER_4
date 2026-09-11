//https://forum.vingrad.ru/articles/topic-197956/view-all.html
unit Param_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.ExtCtrls,
  IniFiles, Common, Vcl.ComCtrls, Vcl.Menus, Vcl.Buttons, ADODB,
  ComObj, Data.DB;

type
  TParamForm = class(TForm)
    Rules: TPageControl;
    Правила: TTabSheet;
    Excel: TTabSheet;
    Panel1: TPanel;
    rgTitle: TRadioGroup;
    sg: TStringGrid;
    brOk: TBitBtn;
    btCheck: TBitBtn;
    ADOConnection1: TADOConnection;
    rgFlagNP: TRadioGroup;
    btCancel: TBitBtn;
    rgBaseLoader: TRadioGroup;
    rgLocalFactor: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure brOkClick(Sender: TObject);
    procedure btCheckClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RulesChange(Sender: TObject);
    procedure btCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    res : Byte;
    Script: Variant;
    function Examen : Boolean;
    procedure ReadIni;
  public
    { Public declarations }
    ds : TADODataset;
    procedure ReadRule(exelname : String);
    function CalcVRI(rule, category : String; area : Integer) : String;
    function CalcVRI6(rule, category, tipter : String; area : Integer) : String;
    function CalcVRI9(rule, category, tipter : String; area : Integer) : String;
  end;

var
  ParamForm: TParamForm;

implementation


{$R *.dfm}
procedure TParamForm.ReadRule(exelname : String);
var
  a, d, e, f : String;
  r, row : Integer;
  ExcelApp, Workbook, WorkSheet : OleVariant;
begin
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Workbooks.Open(exelname);
  WorkBook := ExcelApp.Workbooks.item[1];
  if FindSheetByName(WorkBook, WorkSheet, 'правила', row, r) then
  try
    sg.RowCount := 1;
    sg.Row := 0;
    for r := 2 to row do
    begin
      d := WorkSheet.Cells[r, 4];
      e := WorkSheet.Cells[r, 5];
      if (d <> '') {and (e <> '')} then
      begin
        a := WorkSheet.Cells[r, 1];
        f := WorkSheet.Cells[r, 6];
        sg.RowCount := sg.RowCount + 1;
        sg.Row := sg.RowCount - 1;
        sg.Cells[0, sg.Row] := d;
        sg.Cells[1, sg.Row] := a;
        sg.Cells[2, sg.Row] := e;
        sg.Cells[3, sg.Row] := f;

        sg.Cells[4, sg.Row] := WorkSheet.Cells[r, 3];
      end;
    end;
    sg.Row := 0;
  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;

    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;

procedure TParamForm.ReadIni;
var
  ini: TMemIniFile;
begin
  ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
  try
    rgTitle.ItemIndex := ini.ReadInteger('Export', 'title', 0);
    rgFlagNP.ItemIndex := ini.ReadInteger('FlagNP', 'flaglocalitycode', 0);
    rgLocalFactor.ItemIndex := ini.ReadInteger('FlagLF', 'flaglocalfactor', 0);
    rgBaseLoader.ItemIndex := ini.ReadInteger('BaseLoader', 'flagloader', 0);
  finally
    ini.Free;
  end;
end;

function TParamForm.Examen : Boolean;
var
  i : Integer;
  s : String;
begin
  Result := true;
  for i := 1 to sg.RowCount-1 do
  begin
//    if sg.Cells[2, i] = '' then
//      s := 'select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels'
//    else
//    if sg.Cells[3, i] = '' then
//      s := format('select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels Where area%s',
//        [sg.Cells[2, i]])
//    else
//      s := format('select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels Where (area%s) and (PARCELCATEGORY%s)',
//        [sg.Cells[2, i], sg.Cells[3, i]]);
//    s := StringReplace(s, ' & ', ' and area', []);

    if sg.Cells[2, i] = '' then
      s := 'select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels'
    else
    if sg.Cells[3, i] = '' then
      s := 'select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels Where '+
        StringReplace(sg.Cells[2, i], 'S', 'area', [rfReplaceAll])
    else
      s := 'select top 1 AREA, PARCELCATEGORY, UTILIZATIONBYDOC from parcels Where '+
        StringReplace(sg.Cells[2, i], 'S', 'area', [rfReplaceAll]) + ' and ' +
        StringReplace(sg.Cells[3, i], 'K', 'PARCELCATEGORY', [rfReplaceAll]);

    ds.Close;
    ds.CommandText := s;
    try
      ds.Open;
      if ds.IsEmpty then
        ShowMessage('Алгоритм в строке ' + InttoStr(i) + ' не выбирает записи'+#10+#10+s);
    except
      Result := false;
      ShowMessage('Ошибка алгоритма в строке ' + InttoStr(i));
      sg.Row := i;
      Winapi.Windows.SetFocus(sg.Handle);
      exit;
    end;
  end;
end;

procedure TParamForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if res = 0 then
    ReadIni;
end;

procedure TParamForm.FormCreate(Sender: TObject);
begin
  ds := TADODataSet.Create(nil);
  ds.CacheSize := 100;
  ds.CursorLocation := clUseServer;
  ds.Connection := ADOConnection1;

  script := CreateOleObject('MSScriptControl.ScriptControl');
  script.Language := 'VBScript';

//  Script := CreateOleObject('ScriptControl');
//  Script.Language := 'JScript';

  sg.ColWidths[0] := 80;
  sg.Cells[0,0] := 'Правило';
  sg.Cells[1,0] := 'Код ВРИ';
  sg.Cells[2,0] := 'Площадь, кв.м';
  sg.Cells[3,0] := 'Категория земель';
  sg.Cells[4,0] := 'Тип территории';

  Rules.ActivePageIndex := 0;
  ReadIni;
end;

procedure TParamForm.FormDestroy(Sender: TObject);
begin
  ds.Close;
  ds.Free;
  script := Unassigned;
end;

procedure TParamForm.FormShow(Sender: TObject);
begin
  res := 0;
end;

procedure TParamForm.btCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TParamForm.brOkClick(Sender: TObject);
var
  ini : TMemIniFile;
begin
  res := 1;

  ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
  try
    ini.WriteInteger('Export', 'title', rgTitle.ItemIndex);
    ini.WriteInteger('FlagNP', 'flaglocalitycode', rgFlagNP.ItemIndex);
    ini.WriteInteger('FlagLF', 'flaglocalfactor', rgLocalFactor.ItemIndex);
    ini.WriteInteger('BaseLoader', 'flagloader', rgBaseLoader.ItemIndex);
  finally
    ini.UpdateFile;
    ini.Free;
  end;

  Close;
end;

procedure TParamForm.btCheckClick(Sender: TObject);
begin
  if Examen then
    ShowMessage('Все хорошо');
end;

procedure TParamForm.RulesChange(Sender: TObject);
begin
  if Rules.ActivePageIndex = 0 then
    btCheck.Show
  else
    btCheck.Hide;
end;

function TParamForm.CalcVRI(rule, category : String; area : Integer) : String;
var
  i : Integer;
  s : String;
begin
  Result := '*';
  for i := 1 to sg.RowCount-1 do
  begin
    if rule = sg.Cells[0, i] then
    begin
      if sg.Cells[3, i] = '' then
        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll])
      else
        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll]) + ' and ' +
          StringReplace(sg.Cells[3, i], 'K', category, [rfReplaceAll]);
//      if sg.Cells[3, i] = '' then
//        s := format('%d%s', [area, sg.Cells[2, i]])
//      else
//        s := format('(%d%s) and (%s%s)', [area, sg.Cells[2, i], category, sg.Cells[3, i]]);
//      s := StringReplace(s, ' & ', format(' and %d', [area]), []);
      if Script.Eval(s) then
      begin
        Result := sg.Cells[1, i];
        exit;
      end;
    end;
  end;
end;

//0 правило 1 код vri 2 площадь 3 категория земель 4 тип территории
function TParamForm.CalcVRI6(rule, category, tipter : String; area : Integer) : String;
var
  i : Integer;
  s : String;
begin
  Result := '*';
  for i := 1 to sg.RowCount-1 do
  begin
    if rule <> sg.Cells[0, i] then
      Continue;

    if CText(sg.Cells[4, i], tipter) then
    begin
      if sg.Cells[3, i] = '' then
        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll])
      else
        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll]) + ' and ' +
          StringReplace(sg.Cells[3, i], 'K', category, [rfReplaceAll]);
      if Script.Eval(s) then
      begin
        Result := sg.Cells[1, i];
        Exit;
      end;
    end;
  end;
end;
//0 правило 1 код vri 2 площадь 3 категория земель 4 тип территории
function TParamForm.CalcVRI9(rule, category, tipter : String; area : Integer) : String;
var
  i : Integer;
  s : String;
begin
  Result := '*';
  for i := 1 to sg.RowCount-1 do
  begin
    if rule <> sg.Cells[0, i] then
      Continue;

    if CText(sg.Cells[4, i], tipter) then
    begin
//      if sg.Cells[3, i] = '' then
//        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll])
//      else
//        s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll]) + ' and ' +
//          StringReplace(sg.Cells[3, i], 'K', category, [rfReplaceAll]);
      s := StringReplace(sg.Cells[2, i], 'S', IntToStr(area), [rfReplaceAll]);
      if Script.Eval(s) then
      begin
        Result := sg.Cells[1, i];
        Exit;
      end;
    end;
  end;
end;

end.
