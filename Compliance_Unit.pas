unit Compliance_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Grids,
  ComObj;

type
  TComplianceForm = class(TForm)
    sg: TStringGrid;
    BitBtn1: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    ExcelApp, Workbook, Sheet, Range, Cell1, Cell2, UsedRange : OleVariant;
  public
    { Public declarations }
    SprName : String;
    procedure LoadSpr;
  end;

var
  ComplianceForm: TComplianceForm;

implementation

{$R *.dfm}

procedure TComplianceForm.BitBtn1Click(Sender: TObject);
var
  i, j : Integer;
  DataParcels : OleVariant;
begin
  Cell1 := Sheet.Cells[1, 1];// Ћева€ верхн€€ €чейка области, в которую будем выводить данные
  Cell2 := Sheet.Cells[sg.RowCount, 3];// ѕрава€ нижн€€ €чейка области, в которую будем выводить данные
  Range := Sheet.Range[Cell1, Cell2];// ќбласть, в которую будем выводить данные

  DataParcels := VarArrayCreate([1, sg.RowCount, 1, 3], varOleStr);
  try
    for i := 1 to sg.RowCount do
      for j := 1 to sg.ColCount do
        DataParcels[i, j] := sg.Cells[j-1, i-1];

  Range.Value := DataParcels;
  finally
    VarClear(DataParcels);
  end;

  Workbook.Save;
  Hide;
end;

procedure TComplianceForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ExcelApp.Workbooks.Close;
  ExcelApp.Quit;
  Range := Unassigned;
  Workbook := Unassigned;
  Sheet := Unassigned;
  ExcelApp := Unassigned;
end;

procedure TComplianceForm.FormShow(Sender: TObject);
begin
  sg.Row := sg.RowCount - 1;
end;

procedure TComplianceForm.LoadSpr;
var
  i, j, k : Integer;
  DataParcels : OleVariant;
begin
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := False;//не показывать предупреждающие сообщени€
  Workbook := ExcelApp.WorkBooks.Open(SprName);//ќткрываем рабочую книгу
  Sheet := WorkBook.WorkSheets.Item[3];
  k := Sheet.Usedrange.Rows.Count;
  Cell1 := Sheet.Cells[1, 1];// Ћева€ верхн€€ €чейка области, в которую будем выводить данные
  Cell2 := Sheet.Cells[k, 3];// ѕрава€ нижн€€ €чейка области, в которую будем выводить данные
  Range := Sheet.Range[Cell1, Cell2];// ќбласть, в которую будем выводить данные

  DataParcels := VarArrayCreate([1, Sheet.Usedrange.Rows.Count, 1, 3], varOleStr);
  DataParcels := Range.Value;
  try
    sg.RowCount := k;
    k := 0;
    for i := 1 to sg.RowCount do
    begin
      for j := 1 to sg.ColCount do
        sg.Cells[j-1, i-1] := DataParcels[i, j];
      if sg.Cells[2, i-1] <> '' then Inc(k);
    end;
  finally
    VarClear(DataParcels);
  end;
  sg.RowCount := k;
end;

end.
