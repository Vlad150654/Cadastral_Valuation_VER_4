unit Valuation_Unit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ComObj, Dialogs,
  ShellApi,
  DB, ADODB, ComCtrls, Grids, StdCtrls,
  ExtCtrls,
  ClipBRD,
  Compliance_Unit,
  Mapbasic_INT,
  CalcCodePart,
  IniFiles,
  GeoTransform,
  System.Zip,
  System.IOUtils,
  System.StrUtils,
  System.RegularExpressions,
  System.RegularExpressionsCore,
  Common, Vcl.Menus, Vcl.Buttons;

const
  strConnect = 'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;'+
    'User ID=zisadmin;Initial Catalog=%s;Data Source=TESTSQL;'+
    'Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'+
    'Workstation ID=G001;Use Encryption for Data=False;Tag with column collation when possible=False';

  StateObjects : Array [0..10] of Integer =(232,258,241,251,274,265,276,278,298,281,296);

  farray : Array [0..11] of String =('ID','CADASTRALNUMBER','CADASTRALBLOCK',
    'DISTRICTNAME','DISTRICTTYPE','CITYNAME','CITYTYPE','LOCALITYNAME',
    'LOCALITYTYPE','LOCALITYCODE','SOURCELOCALITYCODE','FLAGLOCALITYCODE');

Type
  FieldsInfo = Record
    title   : String;//название поля
    str     : Byte;//признак строкового типа
  end;

const
  loc_factors : Array [0..10] of FieldsInfo =((title:'CadastralNumber';str:1),
    (title:'Status_kc';str:1),(title:'Specific_CadastralCost';str:0), (title:'CadastralCost';str:0),
    (title:'CadastralCostDesc';str:1),(title:'id_group';str:0), (title:'loc_VICode';str:1),
    (title:'Segment';str:1), (title:'loc_GroupName';str:1),(title:'loc_SgCode';str:1),
    (title:'Old_Utlbydoc';str:1));

Type
  TValuationForm = class(TForm)
    ADOConnection1: TADOConnection;
    OpenDialog1: TOpenDialog;
    MainMenu1: TMainMenu;
    mnObjects: TMenuItem;
    mnBuildings: TMenuItem;
    mnFlats: TMenuItem;
    mnConstructions: TMenuItem;
    mnUncompleteds: TMenuItem;
    mnParcels: TMenuItem;
    mnSelect: TMenuItem;
    mnCode: TMenuItem;
    mnCodeFromSelection: TMenuItem;
    mnCodeFromBase: TMenuItem;
    Panel1: TPanel;
    lbTables: TListBox;
    lbFields: TListBox;
    Panel4: TPanel;
    CommandString: TEdit;
    cbCPU: TComboBox;
    edSize: TEdit;
    mnBuilds: TMenuItem;
    mnQuery: TMenuItem;
    mnCadNumber: TMenuItem;
    mnRules: TMenuItem;
    mnParams: TMenuItem;
    ADOConnection_test: TADOConnection;
    mnModels: TMenuItem;
    mnBuild: TMenuItem;
    stBar: TStatusBar;
    pg: TProgressBar;
    mnCodeFromFile: TMenuItem;
    cmd: TADOCommand;
    mnVRIFromBase: TMenuItem;
    mnVRIFromFile: TMenuItem;
    mnCheckFloors: TMenuItem;
    mnInfo: TMenuItem;
    mnRulesName: TMenuItem;
    mnCodeFromExcel: TMenuItem;
    mnVRIFromExcel: TMenuItem;
    mnParamBuildings: TMenuItem;
    mnBase: TMenuItem;
    mnFlats_Buildings: TMenuItem;
    N1: TMenuItem;
    ExcelCorrect: TMenuItem;
    mnFlats_Orphan: TMenuItem;
    mnCarParkingSpaces: TMenuItem;
    mnTelda_KVRI_4: TMenuItem;
    mnCompareZU: TMenuItem;
    mnTelda_KVRI_4_Plus: TMenuItem;
    mnTelda_KVRI_2026v5: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure lbTablesClick(Sender: TObject);
    procedure lbFieldsClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormDestroy(Sender: TObject);
    procedure lbTablesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure lbFieldsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ObjectClick(Sender: TObject);
    procedure mnCadNumberClick(Sender: TObject);
    procedure mnParamsClick(Sender: TObject);
    procedure mnModelsClick(Sender: TObject);
    procedure mnBuildClick(Sender: TObject);

    procedure mnCodeFromSelectionClick(Sender: TObject);
    procedure mnCodeFromBaseClick(Sender: TObject);
    procedure mnCodeFromFileClick(Sender: TObject);

    procedure mnVRIFromBaseClick(Sender: TObject);
    procedure mnVRIFromFileClick(Sender: TObject);

    procedure mnQueryClick(Sender: TObject);
    procedure mnCheckFloorsClick(Sender: TObject);
    procedure mnInfoClick(Sender: TObject);
    procedure mnCodeFromExcelClick(Sender: TObject);
    procedure mnVRIFromExcelClick(Sender: TObject);
    procedure mnParamBuildingsClick(Sender: TObject);
    procedure mnFlats_BuildingsClick(Sender: TObject);
    procedure ExcelCorrectClick(Sender: TObject);
    procedure mnFlats_OrphanClick(Sender: TObject);
    procedure mnCarParkingSpacesClick(Sender: TObject);
    procedure mnTelda_KVRI_4Click(Sender: TObject);
    procedure mnCompareZUClick(Sender: TObject);
    procedure mnTelda_KVRI_4_PlusClick(Sender: TObject);
    procedure mnTelda_KVRI_2026Click(Sender: TObject);
    procedure OpenDialog1Close(Sender: TObject);
  private
    { Private declarations }
    reg : TRegEx;
    ds : TADODataset;
    fs : TADODataset;
    cf : TADODataset;
    DataSourcesList : TStringList;//список источников данных для таблицы "источники данных'
    RuTablesList : TStringList;//список всех таблиц базы на русском
    TablesList : TStringList;//список всех таблиц базы на латинском заглавными
    RuFieldsList : TStringList;
    FieldsList : TStringList;
    ElementsList : TStringList;
    ftab, path, basename, basename_ : String;
    row, step, size, cpu : Integer;
    f : TFileStream;
    tmp : String;
    fields4base : StringArray;//Поля из зданий для загрузки в базу
    ResFolder : String;
    DutyMap : String;//Дежурная карта
    CodeCatZem : TStringList;
    log : TStringList;

    procedure Building(lavel : Integer; msg : Boolean);
    procedure AllClick;
    procedure LoadFields;
    procedure CreateTablesList;
    function SetLink : String;
    function ImportDataTAB(msg : Boolean) : Boolean;
    function ImportData : Boolean;
    function SetCommandText : Boolean;
    procedure CreateReport(msg : Boolean);
    procedure ReplaceTitle(tab : String);

    function FindYellowRow(var WorkSheet : OleVariant; var YellowRow : Integer): Boolean;
    function CalcModel(filename : string; canshow : Boolean; param : Byte) : Integer;
    procedure Info(msg : String);
    procedure SelectNewDataBase(Sender: TObject);
    procedure SelectNewSpr(Sender: TObject);
    procedure BaseLoader(tablename : String; msg : Boolean);
    procedure Loc_FactorLoader(excelname : String);
    procedure LoadRules(var list1, list2 : TStringList);
    procedure FreeRules(var list1, list2 : TStringList);
    procedure CalcVri;
    procedure CalcVriXlsx(var exceldata : OleVariant);
  public
    { Public declarations }
    function OnCreateDecompressStream(const InStream: TStream;
      const ZipFile: TZipFile; const Item: TZipHeader; IsEncrypted: Boolean): TStream;
    function LoadDate_KVRI_4(feilename : String; var row2, col2 : Integer; var DataParcels : OleVariant) : Boolean;
    function FindStartRow(var sheet : OleVariant; col : Integer) : Integer;
  end;

var
  mainForm: TValuationForm;

implementation

{$R *.dfm}

uses Param_Unit;

//скачать файл поделив на части через поток
//https://webdelphi.ru/2019/02/kak-v-delphi-skachat-fajl-s-ispolzovaniem-http-client-api/ //tr!!!!
//устанавливаем связи свободного справочника и рабочей таблицы
function TValuationForm.SetLink : String;
var
  i : Integer;
  spr : String;
begin
  Repeat
    for i := 1 to cpu do
    begin
      spr := format('%sСправочник%d.key', [tmp, i]);
      if not FileExists(spr) then
      begin
        Result := ChangeFileExt(spr, '');
        exit;
      end;
    end;
    Sleep(50);
  Until false;
end;

function CalcSizeRowByColor(var WorkSheet : OleVariant;//15773696 49407  65535 5296274
  var row, col, col2 : Integer; rowcolor : Integer) : Boolean;
var
  i, j, r, c : Integer;
begin
  Result := false;
  c := WorkSheet.UsedRange.Columns.Count;
  r := WorkSheet.UsedRange.Rows.Count;
  for i := 1 to r do
    for j := 1 to c do
    begin
       if rowcolor = WorkSheet.Cells[i, j].Interior.Color then
       begin
         row := i;
         col := j;
         col2 := j;
         While rowcolor = WorkSheet.Cells[i, col2].Interior.Color do
           Inc(col2);
         Dec(col2);
         Result := true;
         exit;
       end;
    end;
{Sub ColorTest2()
MsgBox Range("A1").Interior.Color
MsgBox Range("A4:D8").Interior.Color
MsgBox Range("C12:D17").Cells(4).Interior.Color
MsgBox Cells(3, 6).Interior.Color
End Sub}
end;

procedure TValuationForm.mnCompareZUClick(Sender: TObject);
var
  i, j, k, c, row, col, row2, col2, rp, cp : Integer;
  filename, pattern, temppath, source, newzu, oldzu, seg13, s, ss : String;
  ExcelApp, Workbook, WorkSheet, range, DataParcels, data : OleVariant;

  function LoadData(filename : String; var data : OleVariant) : Boolean;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
  begin
    Result := false;
    info('Загрузка ' + ExtractFileName(filename));
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(filename);
      WorkBook := ExcelApp.Workbooks.item[1];
      WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
      FindCadNum(WorkSheet, row, col, row2, col2);//блок данных
      range := WorkSheet.range[format('%s%d:%s%d',
        [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, col2), row2])];
      data := range.value;
      Result := true;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  end;

begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\PATTERNS\TelDa\new_16_04_2024'; {$ENDIF}
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Шаблон сравнения с родителем ЗУ.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Шаблон сравнения с родителем ЗУ.xlsx');

  ResFolder := GetAppFolder + 'RESULT\TelDa\' + AddSlesh(CurrentData);
  if not ForceDirectories(ResFolder) then
    Exc('Не удалось создать каталог' + #10 + ResFolder);

  filename := ResFolder + 'сравнения с родителем ЗУ.xlsx';
  if FileExists(filename) then
    if not DeleteFile(filename) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [filename, #10]));

  OpenDialog1.FileName := '*.xls*';
  OpenDialog1.Title := 'Открытие файла Новые ЗУ';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  newzu := OpenDialog1.FileName;

  OpenDialog1.Title := 'Открытие файла Предыдущие ЗУ';
  if not OpenDialog1.Execute then
    exit;
  oldzu := OpenDialog1.FileName;
  OpenDialog1.Title := 'Открытие файла сегмент13';
  if not OpenDialog1.Execute then
    exit;
  seg13 := OpenDialog1.FileName;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.DisplayAlerts := False;
  try
    pattern := 'PATTERNS\TelDa\Шаблон сравнения с родителем ЗУ.xlsx';
    source := GetAppFolder + pattern;
    temppath := ExtractFilePath(miEval('TempFileName$("")'));
    ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"',
      [source, temppath])), SW_HIDE, true, 0);
    s := temppath + ExtractFileName(source);
    temppath := ChangeFileExt(s, '_.xlsx');
    RenameFile(s, temppath);//новое название должно отличаться от шаблона
    ExcelApp.Workbooks.Open(temppath);//открываем копию
    WorkBook := ExcelApp.Workbooks.item[1];
    if FindSheetByName(WorkBook, WorkSheet, 'Объекты') then
    begin
      LoadData(newzu, data);//загружаем Новые ЗУ
      range := WorkSheet.range[format('%s%d:%s%d',//размер для вставки в шаблоне
        [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, col2), row2])];
      range.value := data;//вставляем в шаблон Новые ЗУ

      FindCadNum(WorkSheet, row, col, rp, cp);//нужно сократить размер колонок до блока формул
      CalcSizeRowByColor(WorkSheet, i, j, cp, 65535);//размер желтого цвета сегмент13
      range := WorkSheet.range[format('%s%d:%s%d',
        [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, cp), rp])];
      DataParcels := range.value;//блок шаблона с новыми данными

      VarClear(data);
      LoadData(oldzu, data);
      c := VarArrayHighBound(data, 2);
      for i := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do//по строкам
      begin
        s := DataParcels[i, 2];
        if (s <> '') and (Pos(';', s) = 0) then
          for j := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
            if s = data[j, 1] then
            begin
              DataParcels[i, c+1] := s;
              for k := 3 to VarArrayHighBound(data, 2) do
                DataParcels[i, c+k-1] := data[j, k];
              break;
            end;
      end;

      VarClear(data);
      LoadData(seg13, data);
      c := cp - VarArrayHighBound(data, 2);
      for i := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do//по строкам
      begin
        s := DataParcels[i, 1];
        for j := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
          if s = data[j, 1] then
          begin
            for k := 3 to VarArrayHighBound(data, 2) do
              DataParcels[i, c+k-2] := data[j, k];
            break;
          end;
      end;

      range.value := DataParcels;//вставляем блок обратно вместе с изменениями

      Info('Копирование фрмул...');
      CalcSizeRowByColor(WorkSheet, row, col, col2, 5296274);//размер зеленого цвета
      s := ExcelNum2Str(WorkSheet, col);
      ss := ExcelNum2Str(WorkSheet, col2);//конец фрпмул
      WorkSheet.Range[format('%s%d:%s%d', [s, row+1, ss, row+1])].Copy;
      WorkSheet.Range[format('%s%d:%s%d', [s, row+2, ss, rp])].PasteSpecial(operation:=xlNone);
    end else
      Exc('Лист "Объекты" не найден в файле шаблона');

    WorkBook.SaveAs(Filename:=filename, FileFormat:= 51, CreateBackup:=False);
    ExcelApp.Visible := true;

  finally
    Info('Готово');
    VarClear(DataParcels);
    VarClear(data);
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
    range := Unassigned;
  end;
end;

procedure TValuationForm.CreateReport(msg : Boolean);
var
  s : String;
  i :Integer;
begin
  pg.Position := 4;
  Info('Создание отчета...');
  Application.ProcessMessages;

  miDo('select * from work1 where LOCALITYCODE="" AND код_НП<>"" into Compare');
  if miEvalInt('TableInfo(Compare, %d)', [TAB_INFO_NROWS]) > 0 then
    miDo('Update Compare Set FLAGLOCALITYCODE = "ГОТОВ_ИЗМЕНЕН_АВТОМАТ"');
  miDo('Close table Compare');

  pg.Position := 5;
  miDo('select * from work1 where LOCALITYCODE<>"" AND код_НП<>"" AND LOCALITYCODE=код_НП into Compare');
  if miEvalInt('TableInfo(Compare, %d)', [TAB_INFO_NROWS]) > 0 then
    miDo('Update Compare Set FLAGLOCALITYCODE = "ГОТОВ_АВТОМАТ"');
  miDo('Close table Compare');

  pg.Position := 6;
  miDo('select * from work1 where LOCALITYCODE<>"" AND код_НП<>"" AND LOCALITYCODE<>код_НП into Compare');
  if miEvalInt('TableInfo(Compare, %d)', [TAB_INFO_NROWS]) > 0 then
    miDo('Update Compare Set FLAGLOCALITYCODE ="ГОТОВ_ИЗМЕНЕН_АВТОМАТ"');
  miDo('Close table Compare');

  pg.Position := 7;
  miDo('Update work1 Set KodMO = Left$(код_НП, 2)');
  miDo('Commit table work1');

  //будем сравнивать код_НП и SOURCELOCALITYCODE
  pg.Position := 8;
  miDo('Select CadastralNumber, код_НП, FLAGLOCALITYCODE, SOURCELOCALITYCODE, '+
    'OLDLOCALITYCODE, KodMO, Проверка from work1 where FLAGLOCALITYCODE<>"" into Compare');//tr!!!  ver3.3

  miDo('Commit table Compare as "%s%s_Compare.tab"', [ResFolder, basename]);
  miDo('Close table Compare');
  miDo('Open table "%s%s_Compare.tab" As Compare', [ResFolder, basename]);

  pg.Position := 9;
  miDo('Alter Table Compare (Rename код_НП LOCALITYCODE)');
  miDo('Select * from Compare where FLAGLOCALITYCODE <> "НЕ ГОТОВ" into Temp noSelect');
  pg.Position := 10;

  if miEvalInt('TableInfo(Temp, %d)', [TAB_INFO_NROWS]) > 0 then
  begin
    Info('Создание файла загрузки...');

    Tab2Excel('Temp', format('%s%s.xlsm', [ResFolder, basename+'_загрузка']), nil);

    if ParamForm.rgBaseLoader.ItemIndex = 0 then
    begin
      OpenDialog1.FileName := format('%s%s.xlsm', [ResFolder, basename+'_загрузка']);
      s := ExtractFileName(OpenDialog1.FileName);
      i := Pos('_загрузка', s);
      i :=  RuTablesList.IndexOf(Copy(s, 1, i-1));
      BaseLoader(TablesList[i], msg);
    end;
  end;
  miDo('Close table Compare');
end;

procedure TValuationForm.mnCodeFromBaseClick(Sender: TObject);
begin
  AllClick;
end;

procedure TValuationForm.Loc_FactorLoader(excelname : String);
var
  col, row, c, r, i, j, k : Integer;
  s1, s2, s3, s, ss, filename, fields, values  : String;
  fds : TFieldDefs;
  fd : TFieldDef;
  ExcelApp, Workbook, WorkSheet : OleVariant;
  sl : TStringList;
begin
//для расчетников заносим в табдицу zisadmin.summ_calc
  AdoConnection1.Execute('DELETE FROM zisadmin.summ_calc');
  fs.Close;
  fs.CommandText := 'select * from zisadmin.summ_calc';
  fs.Open;
//  fs.SaveToFile('D:\Tokyo\Cadastral_Valuation\test\aaa.xml');
  sl := TStringList.Create;
  try
    fds := fs.FieldDefs;
    for i  := 0 to fds.Count-1 do
    begin
      fd := fds.Items[i];
      case fd.DataType of
        ftString, ftBlob, ftMemo, ftFmtMemo, ftBytes, ftVarBytes, ftFixedChar,
        ftFixedWideChar, ftWideMemo, ftByte, ftGuid,
        ftWideString: sl.AddObject(fd.Name, Pointer(fd.Size));//tr!!!
      else
        sl.Add(fd.Name);
      end;
    end;
    sl.SaveToFile('D:\Tokyo\Cadastral_Valuation\test\bbb.txt');

    try
      ExcelApp := CreateOleObject('Excel.Application');
      ExcelApp.Workbooks.Open(OpenDialog1.FileName);
      WorkBook := ExcelApp.Workbooks.item[1];
      WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
      col := WorkSheet.UsedRange.Columns.Count;
      row := WorkSheet.UsedRange.Rows.Count;
      pg.Max := row;
      k := 0;

      for r := 2 to row do
      begin
        if r mod 10 = 0 then
          pg.Position := r;
        s1 := WorkSheet.Cells[1, 1];
        ss := WorkSheet.Cells[r, 1];
        if reg.IsMatch(ss,'66:\d\d:\d\d\d\d\d\d\d:\d{1,}') and CText(s1, 'CadastralNumber') then
        begin
          fields := format('id, %s,', [s1]);
          values := format('%d, ''%s'',', [r, ss]);
          for c := 2 to col do
          begin
            s1 := WorkSheet.Cells[1, c];
            i := sl.IndexOf(s1);//поле по названию
            if i > -1 then
            begin
              ss := WorkSheet.Cells[r, c];
              if ss <> '' then
              begin
                fields := fields + format('%s,', [s1]);

                if sl.Objects[i] <> nil then
                begin
                  j := Integer(sl.Objects[i]);
                  if Length(ss) > j then
                    SetLength(ss, j);
                end;

                if sl.Objects[i] <> nil then
                  values := values + format('''%s'',', [ss])
                else
                  values := values + format('%s,', [ss]);
              end;
            end;
          end;
          s := format('Insert into zisadmin.summ_calc (%s) Values(%s)',
            [Copy(fields, 1, Length(fields)-1), Copy(values, 1, Length(values)-1)]);
          try
            ADOConnection1.BeginTrans;
            ADOConnection1.Execute(s);
            ADOConnection1.CommitTrans;
            Inc(k);
          except
            on e : Exception do
            begin
              ADOConnection1.RollbackTrans;
              Exc(e.Message);
            end;
          end;
        end;

      end;
    finally
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
      ShowMessage(format('занесено %d записей', [k]));
    end;
  finally
    pg.Max := 0;
    sl.Free;
  end;
end;

procedure TValuationForm.BaseLoader(tablename : String; msg : Boolean);
var
  i, j, k : Integer;
  s, ss, sql : String;
  fds : TFieldDefs;
  fd : TFieldDef;
begin
//if pos('test', String(ADOConnection1.ConnectionString)) = 0 then
//exc('не та база');
  Info('Подготовка к обновлению базы '+ tablename);
  BreakCommand := false;
  ADOConnection_test.Close;

  path := ExtractFilePath(OpenDialog1.FileName);

  ADOConnection_test.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;'+
  'Extended Properties="DSN=Excel Files;DBQ=' + OpenDialog1.FileName + ';DefaultDir=' + path +
  ';DriverId=1046;MaxBufferSize=2048;PageTimeout=5;"';

//    ADOConnection_test.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;'+
//      'Data Source=' + OpenDialog1.FileName + ';User ID=Admin;Password=;'+
//      'Extended Properties="Excel 8.0;HDR=YES;HDR=NO;IMEX=1"';// IMEX=1 нужен, чтобы корректно прочитался столбец, если в нем идут вмеперешку ячейки с числами, датами, строками...

  ADOConnection_test.Open;
  try
    fs.Close;
    fs.Connection := ADOConnection_test;
    fs.commandtext := 'select * from [Лист1$]';//Выбираем весь первый лист
    fs.Open;
    fds := fs.FieldDefs;
    try
      fd := fds.Find('SourceLocalityCode');
    except
      fd := nil;
    end;
    if fd <> nil then
    while not fs.eof do
    begin//заменяем номера кодами из справочника
      j := fs.FieldByName('SourceLocalityCode').AsInteger;
      if j <> 0 then
        for i := 0 to DataSourcesList.Count-1 do
        begin
          k := Integer(DataSourcesList.Objects[i]);
          if k = j then
          begin
            fs.Edit;
            fs.FieldByName('SourceLocalityCode').Value := StrToInt(DataSourcesList[i]);
            break;
          end;
        end;
      fs.next;
    end;

    ds.Close;//для определния набора полей делаем выборку
    ds.CommandText := 'Select Top 1 * from ' + tablename;
    ds.Open;
    for i := fds.Count-1 downto 0 do//сопоставим поля Excel и поля в базе
    try
      fd := ds.FieldDefs.Find(fds[i].Name);
    except
      fds[i].Free;//поля отсутствующие в базе удалим
    end;

    sql := '';
    for i := 1 to fds.Count-1 do//создаем запрос, CadastralNumber пропускаем
      sql := sql + format('%s=:%s,', [fds[i].Name, fds[i].Name]);
    sql := format('UPDATE %s SET %s WHERE CadastralNumber=:number',
      [tablename, Copy(sql, 1, Length(sql)-1)]);
    cmd.Parameters.Clear;
    cmd.CommandType := cmdText;
    cmd.Commandtext := sql;
    cmd.Parameters.ParseSQL(sql, True);
    for i := 1 to fds.Count-1 do//назначаем типы полей CadastralNumber пропускаем
    begin
      fd := ds.FieldDefs.Find(fds[i].Name);
      cmd.Parameters[i-1].DataType := fd.DataType;
    end;
    cmd.Parameters.ParamByName('number').DataType := ftString;

    Info('Обновление базы '+ tablename);
    CommandString.SetFocus;
    k := 0;
    pg.Max := fs.RecordCount;
    fs.FindFirst;
    while not fs.eof do
    begin
      Application.ProcessMessages;
      if BreakCommand then exit;
      pg.Position  := pg.Position + 1;
      //последняя запись в  Parameters это CadastralNumber
      cmd.Parameters.ParamByName('number').Value := fs.FieldByName(fds[0].Name).AsString;
      for i := 1 to fds.Count-1 do//CadastralNumber пропускаем
        cmd.Parameters[i-1].Value := fs.FieldByName(fds[i].Name).AsVariant;
      ADOConnection1.BeginTrans;
      try
        cmd.Execute;
        ADOConnection1.CommitTrans;
        Inc(k);
      except
        on e: Exception do
        begin
          ShowMessage(e.Message);
          ADOConnection1.RollbackTrans;
        end;
      end;
      fs.next;
    end;
  finally
    if msg then
      ShowMessage('Обновление завершено' + #10 + format('Обновлено записей %d', [k]));
    fs.Close;
    fs.Connection := ADOConnection1;
    pg.Max := 0;
    ADOConnection_test.Close;
    Info('');
  end;
end;

procedure TValuationForm.Info(msg : String);
begin
  stBar.Panels[1].Text := msg;
  stBar.Repaint;
end;
//Excel. Цвет ячейки
//https://vremya-ne-zhdet.ru/vba-excel/tsvet-yacheyki-zalivka-fon/ //tr!!!!
function TValuationForm.FindYellowRow(var WorkSheet : OleVariant; var YellowRow : Integer): Boolean;
var
  range : OleVariant;
  col : Integer;
begin
  Result := false;
  YellowRow := 1;
  Repeat
    range :=  WorkSheet.Range[format('A%d', [YellowRow])];
    col := range.Interior.Color;
    range := Unassigned;
    if col = 65535 then
    begin
      Result := true;
      exit;
    end;
    Inc(YellowRow);
  Until YellowRow > 10;
end;

function TValuationForm.CalcModel(filename : string; canshow : Boolean; param : Byte) :  Integer;
var
  s, ss, shortname : String;
  i, j, n, row, col, fid, YellowRow : Integer;
  sl : TStringList;
  ExcelApp, Workbook, WorkSheet, range, data : OleVariant;
  complex : Boolean;//признак составного названия поля
  sa : StringArray;
  mc : TMatchCollection;

  function ParceYellowRow : Boolean;
  begin
    Result := false;

    s := WorkSheet.Cells[YellowRow + 3, 1];//проверяем 7 ячейку на наличие второй части запроса: "Select ID,Source_ID,Cadas..."
    if s = '' then
      j := 2//начало и конец запрося в одной строке
    else
      j := 3;//начало в строке YellowRow+2, конец запрося в строке YellowRow+3

    col := WorkSheet.UsedRange.Columns.Count;
    row := WorkSheet.UsedRange.Rows.Count;

    if row > YellowRow + j then//уменьшаем UsedRange до второй строки формул
      WorkSheet.Range[format('A%d:%s%d', [YellowRow+j, ExcelNum2Str(WorkSheet, col), row])].EntireRow.Delete;

    n := 1;//идем строке данных и формул
    range := WorkSheet.Range[format('%s%d', [ExcelNum2Str(WorkSheet, n), YellowRow])]; //65535   желтый
    range.select;
    j := range.Interior.Color;
    While (j = 65535) do//данные желтый цвет   16777215 белый
    begin
      s := WorkSheet.Cells[YellowRow - 2, n];//поле на английском
      if s = '' then
        Exc(shortname + ': Отсутствует название поля в колонке ' + ExcelNum2Str(WorkSheet, n));
      ss := WorkSheet.Cells[YellowRow - 1, n];//таблица
      if complex then
        sl.Add(ss + #35 + s)
      else
        sl.Add(s);

      ss := WorkSheet.Cells[YellowRow, n];//значение из справочника или поле для загрузки в базу
      if ss <> '' then
      begin
        if (param = 0) and ElementsList.Find(s, i) then
          sl.Objects[n-1] := ElementsList.Objects[i];//передаем KLASS_ID
        if param > 0 then
        begin
          SetLength(fields4base, Length(fields4base) + 1);//заносим название поля для загрузки
          fields4base[Length(fields4base)-1] := WorkSheet.Cells[YellowRow-2, n];
        end;
      end;
      Inc(n);
      range := WorkSheet.Range[format('%s%d', [ExcelNum2Str(WorkSheet, n), YellowRow])]; //65535   желтый
      range.select;
      j := range.Interior.Color;
      Result := true;
    end;
    if Result then
    begin
      WorkSheet.Rows[YellowRow].Delete(xlUp);//признак справочника удаляем
      WorkSheet.Rows[YellowRow+1].Delete(xlUp);//строку - шаблон запросов удаляем
      WorkSheet.Rows[YellowRow+1].Delete(xlUp);//строку - шаблон запросов удаляем (может быть 2 строки запроса)

      range :=  WorkSheet.Range[format('A%d:%s%d', [YellowRow, ExcelNum2Str(WorkSheet, n-1), YellowRow])];
      range.Select;
      ExcelApp.Selection.ClearContents;
      ExcelApp.Selection.NumberFormat := '';

      row := WorkSheet.UsedRange.Rows.Count;
      Workbook.Save;
    end;
  end;

begin//CalcModel
  Result := 1;
  BreakCommand := false;
  shortname := ExtractFileName(filename);
  Info(shortname + ' Подготовка запроса...');
  SetLEngth(fields4base, 0);
  sl := TStringList.Create;
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := False;//не показывать предупреждающие сообщения
  try
    ExcelApp.Workbooks.Open(filename);
    WorkBook := ExcelApp.Workbooks.item[1];
    WorkSheet := WorkBook.WorkSheets.item[1];//1 лист
    WorkBook.WorkSheets.item[1].Activate;//обязательно активировать иначе не сработает метод range.select!!!
    Info(shortname + ' Выполнение запроса...');
    if not FindYellowRow(WorkSheet, YellowRow) then
      exit;
    i := YellowRow + 2;//здесь первая строка для запроса
    s := '';
    complex := false;
    Repeat//накапливаем CommandString
      ss := String(WorkSheet.Cells[i, 1]); //число символов в строке 32767
      if (Pos('#CadastralNumber', ss) > 0) or (Pos('#ID', ss) > 0) then
        complex := true;
      s := s + ss;
      Inc(i);
    Until ss = '';

    ds.Close;
    ds.CommandText := s;
    ds.Open;
    if ds.IsEmpty then
      exit;

    Result := 2;
    ds.FindLast;
    pg.Max := ds.RecNo;
    if not ParceYellowRow then
      exit;

    Result := 3;
    ds.FindFirst;
    row := YellowRow;
    Info(shortname + ' Создание таблицы...');
    data := VarArrayCreate([1, pg.Max, 1, col], varOleStr);//размер блока данных
    CommandString.SetFocus;
    while not ds.eof do
    begin
      if (row - 1) mod 10 = 0 then
        pg.Position := pg.Position + 10;
      Application.ProcessMessages;
      if BreakCommand then break;
      for i := 0 to sl.Count-1 do
      begin
        fid := ds.Fields.IndexOf(ds.FieldByName(sl[i]));
        s := ds.Fields[fid].AsString;//пустая строка ддля ds.Fields[fid].IsNull
        //Здания : определение года, этажности, материала стен
        if param > 0 then
        begin
          mc := reg.Matches(s, '\d\d\d\d-\d\d-\d\d', [roIgnoreCase, roMultiLine]);
          if mc.Count > 0 then
          begin
            s := mc.Item[0].Value;
            sa := SplitStr(s, '-');
            s := sa[2] + '.' +sa[1] + '.' + sa[0];
            SetLength(sa, 0);
          end else
          if (param = 2)and (i = 1) then
          begin
            if s = '204001000000' then s := 'Нежилое здание';
            if s = '204002000000' then s := 'Жилой дом';
            if s = '204003000000' then s := 'Многоквартирный дом';
            if s = '204004000000' then s := 'Жилое строение';
          end;
        end;

        if s  <> '' then
        begin
          if sl.Objects[i] <> nil then//справочное значение
          begin
            j := Integer(sl.Objects[i]);
            fs.Close;
            fs.CommandText := format('select * from KLASSELEMENTS where idk=''%s'' AND KLASS_ID=%d', [s, j]);
            fs.Open;
            if not fs.IsEmpty then
              s := #39 + fs.FieldByName('NAME').AsString;
          end;
          case ds.Fields[fid].DataType of//текст
            ftString, ftBlob, ftMemo, ftFmtMemo, ftBytes, ftVarBytes, ftFixedChar,
            ftFixedWideChar, ftWideMemo, ftByte, ftGuid,
            ftWideString: s := #39 + s;
          end;
        end;
        data[ds.RecNo, i+1] := s;
      end;
      ds.next;
      Inc(row);
    end;

    Result := 4;
    ss := ExcelNum2Str(WorkSheet, n-1);
    range := WorkSheet.Range[format('A%d:%s%d', [YellowRow, ss, row-1])];//Область вставки
    range.Value := data;
    Workbook.Save;

    Result := 5;
    Info(shortname + ' Форматирование таблицы...');
    for i := 0 to sl.Count-1 do
    begin
      case ds.FieldByName(sl[i]).DataType of
        ftSingle, ftSmallint, ftInteger, ftWord, ftAutoInc, ftFloat,
        ftExtended, ftCurrency, ftFMTBcd, ftBCD : begin
          ExcelApp.Columns[ExcelNum2Str(WorkSheet, i+1)].Select;
          ExcelApp.Selection.TextToColumns(DataType := XlDelimited);
        end;
      end;
    end;
    Workbook.Save;

    if (col > n-1) and (row-1 > YellowRow) then
    begin
      Result := 6;
      Info(shortname + ' Копирование формул...');
      s := ExcelNum2Str(WorkSheet, n);
      ss := ExcelNum2Str(WorkSheet, col);
      WorkSheet.Range[format('%s%d:%s%d', [s, YellowRow, ss, YellowRow])].Copy;
      WorkSheet.Range[format('%s%d:%s%d', [s, YellowRow+1, s, row-1])].PasteSpecial(operation:=xlNone);
      Workbook.Save;
    end;

    Result := 0;

  finally
    pg.Max := 0;
    sl.Free;
    VarClear(data);
    if Result = 0 then
    begin
      ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
      Workbook.Save;
      if canshow then
        ExcelApp.Visible := true
      else
        ExcelApp.Quit;
    end else
    begin
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
      while FileExists(filename) do
      begin
        Application.ProcessMessages;
        DeleteFile(filename);
      end;
    end;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
    range := Unassigned;
  end;
end;

procedure TValuationForm.mnModelsClick(Sender: TObject);
var
  i, k, d : Integer;
  path, newfile : String;
  sl : TStringList;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER4\Win32\Debug\PATTERNS'; {$ENDIF}
  if OpenDialog1.InitialDir = '' then
    OpenDialog1.InitialDir := GetAppFolder + 'PATTERNS\';
  OpenDialog1.FileName := '*.xls*';
  if OpenDialog1.Execute then
  try
    sl := TStringList.Create;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    path := ExtractFilePath(OpenDialog1.FileName) + CurrentData + '\';
    if not ForceDirectories(path) then
      Exc('Не удалось создать каталог' + #10 + path);

    for i := 0 to OpenDialog1.Files.Count-1 do
    begin
      newfile := ExtractFileName(OpenDialog1.Files[i]);
      if Pos('_', newfile) <> 1 then
        Exc('Файл не является шаблоном'+#10+ExtractFileName(OpenDialog1.Files[i]));
      newfile := path + Copy(newfile, 2, Length(newfile));
      if not CopyFile(PWideChar(OpenDialog1.Files[i]),  PWideChar(newfile), false) then
        Exc('Ошибка копировани файла '+ExtractFileName(OpenDialog1.Files[i])+#10+'Возможно файл открыт.')
      else
        sl.Add(newfile);
    end;

    k := 0;
    for i := 0 to sl.Count-1 do
    begin
      d := CalcModel(sl[i], sl.Count = 1, 0);
      if d = 0 then
        Inc(k);
      if BreakCommand then
        Exc('Расчет прерван');
      if sl.Count = 1 then
        case d of
          0: Exc('Ok');
          1: Exc('Нет выбранных записей');
          2: Exc('Ошика при поиске рабочей области');
          3: Exc('Ошика при создании таблицы');
          4: Exc('Ошика при вставке данных');
          5: Exc('Ошибка форматирование таблицы');
          6: Exc('Ошибка копирования формул');
        end;
    end;
    ShowMessage('Расчет завершен' + #10 + format('Созданных моделей %d', [k]));
    Info('');
  finally
    sl.Free;
  end;
end;

procedure TValuationForm.mnCodeFromSelectionClick(Sender: TObject);
begin
  Building(1, true);
end;

procedure TValuationForm.Building(lavel : Integer; msg : Boolean);
var
  i, k : Integer;
  s, ss, param : String;
begin
  BreakCommand := false;
  if not ForceDirectories(tmp) then
    Exc(format('Не удалолсь создать временный каталог "%s"', [tmp]));

  if (OpenDialog1.FileName = '') or (not CText(ExtractFileExt(OpenDialog1.Filename), '.TAB')) then
    Exc('Не выбрана таблица для обработки');

  Val(edSize.Text, step, i);
  if step = 0 then
    Exc('Не задан размер части.');

  if FileExists(ResFolder + basename + '.xlsm') then
    if not DeleteFile(ResFolder + basename + '.xlsm') then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [basename, #10]));

  if FileExists(ResFolder + basename+'_загрузка' + '.xlsm') then
    if not DeleteFile(ResFolder + basename+'_загрузка' + '.xlsm') then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [basename+'_загрузка', #10]));

  ShellExecute(0, 'Open', 'cmd.exe', '/c del /s /f /q  "c:\temp~files"', nil, SW_HIDE);

  miDo('dim ob as object');
  try
    Info(basename + ': Подготовка');
    miDo('Open table "%sНП_МО_код_02.2020.TAB" As Справочник', [GetAppFolder + 'BOUNDS\']);

    path := ExtractFilePath(OpenDialog1.FileName);
    ftab := OpenDialog1.Filename;
    ss := miEval('TempFileName$("")');
    s := miEval('PathToTableName$("%s")', [ss]);
    if TableExists(ss) then
      miDo('Close table %s', [ss]);

    miDo('Open Table "%s" as FullTable', [ftab]);
    //добавление полей
    miDo('Alter Table "FullTable"(Add код_НП Char(19),алгоритм Integer,Район Integer,'+
      'МО Integer,НП Integer,Источник_НП Integer, KodMO Char(5),'+
//      'OLDLOCALITYCODE Char(50))');
      'OLDLOCALITYCODE Char(50), Проверка Char(1))');//tr!!!!   ver3.3

    size := miEvalInt('TableInfo(FullTable, %d)', [TAB_INFO_NROWS]);//размер полной таблицы
    if size < 10000 then
    begin
      Info(basename + ': Расчет');
      miDo('Close All');
      Work(format('%s', [ftab]), format('%sНП_МО_код_02.2020.TAB', [GetAppFolder+'BOUNDS\']), IntToStr(lavel), pg);
      exit;
    end;

    if size/step < 2 then
      step := Round(size/2 - 1);
    if step = 0 then
      step := 1;
    pg.Max := Round(size/step);
    Val(cbCPU.Text, cpu, i);
    if cpu > pg.Max then
      cpu := pg.Max;

    for i := 1 to cpu do //сперва закинем все справочники по числу CPU
    begin
      miDo('Commit table Справочник As "%sСправочник%d"', [tmp, i]);
      DeleteFile(format('%sСправочник%d.key', [tmp, i]));
    end;
    row := 0;

    for i := 1 to pg.Max do
    begin
      Info(basename + ': Запуск,часть '+Inttostr(i));
      pg.Position := i+1;
      if i = pg.Max then //последняя часть
        miDo('Select * from FullTable where RowId >= %d AND RowID < %d into work noSelect',[row, step*i*2])
      else
        miDo('Select * from FullTable where RowId >= %d AND RowID < %d into work noSelect',[row, step*i]);
      miDo('Commit table work as "%swork%d.tab"', [tmp, i]);
      miDo('Close table work');

      s := SetLink;//находим освободившийся справочник для рабочей таблицы
      param := format('%swork%d %s %d', [tmp, i, s, lavel]);

      f := TFileStream.Create(s + '.key', fmCreate); //создадим пустой ключевой файл - признак
      f.Free;

      if i = pg.Max then//последняя часть
      begin
        Info(basename + ': Расчет, последняя часть');
        Work(format('%swork%d.tab', [tmp, i]), s, IntToStr(lavel), pg)
      end else
        ExShellExecute(GetAppFolder+'CalcCodeWork.exe', param, SW_HIDE, false, 0);

      Inc(row, step);
      Application.ProcessMessages;
      if BreakCommand then exit;
    end;

    miDo('Drop Table FullTable');
    miDo('Close All');
    Info(basename + ': Сборка таблиц');
    pg.Max := Round(size/step);
    pg.Position := 1;

    Repeat//ждем, пока не удалятся все справочники.key (признак завершения части)
      k := 0;
      for i := 1 to pg.Max do
        if FileExists(format('%sСправочник%d.key', [tmp,i])) then
          k := 1;
      Sleep(50);
    Until k = 0;

    miDo('Open table "%s"', [format('%swork1.tab', [tmp])]);
    for i := 2 to pg.Max do
    begin
      pg.Position := i;
      Application.ProcessMessages;
      miDo('Open table "%s"', [format('%swork%d.tab', [tmp, i])]);
      miDo('insert into work1 select * from work%d', [i]);
    end;
    miDo('Commit Table work1 As "%s" TYPE NATIVE Charset "WindowsCyrillic"', [ftab]);

  finally
    if BreakCommand then
      ShowMessage('Обработка прервана')
    else
    try
      pg.Max := 10;
      pg.Position := 1;
      Info(basename + ': Добавление информации из КК');
      miDo('Close All');
      miDo('Open table "%sKK_NP.TAB"', [GetAppFolder + 'BOUNDS\']);
      miDo('Open table "%sKR_NP.TAB"', [GetAppFolder + 'BOUNDS\']);
      miDo('Open table "%s" as work1', [ftab]);
      try//дорабатываем по кадастровым кварталам
        miDo('Select work1.код_НП, work1.алгоритм, work1.Источник_НП, '+
          'KK_NP.Кадастровый_квартал, KK_NP.Код_НП1 from work1, KK_NP '+
          'where work1.CADASTRALBLOCK = KK_NP.Кадастровый_квартал AND work1.алгоритм=0 into No_Algoritm');
      finally
        if miEvalInt('TableInfo(No_Algoritm, %d)', [TAB_INFO_NROWS]) > 0 then
          miDo('Update No_Algoritm Set код_НП=код_НП1, алгоритм=300, Источник_НП=754');
        miDo('Commit Table work1');
      end;
      pg.Position := 2;
      Application.ProcessMessages;
      Info(basename + ': Добавление информации из КР');
      try//дорабатываем по кадастровым районам
        miDo('select * from work1 where Left$(CADASTRALBLOCK, 3) = "66:" into Temp noSelect');
        miDo('Update Temp Set CADASTRALBLOCK = Mid$(CADASTRALBLOCK,4, 2)');
        miDo('Select work1.код_НП, work1.алгоритм, Источник_НП, KR_NP.Кад_район, '+
          'KR_NP.Код_межселенки from KR_NP, work1 where KR_NP.Кад_район='+
          'work1.CADASTRALBLOCK AND work1.алгоритм=0 into No_Algoritm');//связываем таблицы
      finally
        miDo('Rollback Table work1');
        if miEvalInt('TableInfo(No_Algoritm, %d)', [TAB_INFO_NROWS]) > 0 then
          miDo('Update No_Algoritm Set код_НП=Код_межселенки, алгоритм=400, Источник_НП=753');
        miDo('Commit Table work1');
      end;
      pg.Position := 3;
      Application.ProcessMessages;
      miDo('Update work1 Set OLDLOCALITYCODE = LOCALITYCODE');
      miDo('Update work1 Set SOURCELOCALITYCODE = Источник_НП');

      CreateReport(msg);

      if ParamForm.rgTitle.ItemIndex = 0 then
        ReplaceTitle('work1');

      pg.Max := 0;
      if msg then
        ShowMessage('Обработка завершена');
    except
      on e : Exception do ShowMessage(e.Message);
    end;
    miDo('Close All');
    OpenDialog1.FileName := '';
    pg.Max := 0;
    ShellExecute(0,PChar('open'),PChar('cmd'),PChar(format('/c rd /S/Q "%s"',[tmp])),nil,SW_HIDE);//удалене каталога со всем содержимым и подкаталогами
    Info('');
  end;

end;

procedure TValuationForm.ReplaceTitle(tab : String);
var
  i, j, k : Integer;
  sl : TStringList;
  s : String;
  f32 : String[32];
begin
  try
    sl := getFieldsList(tab);
    s := '';
    for i := 0 to sl.Count-1 do
    begin
      k := FieldsList.IndexOf(sl[i]);
      if (k <> -1) AND (RuFieldsList[k] <> FieldsList[k]) then
      begin
        f32 := RuFieldsList[k];
        for j := 1 to Length(f32) do
        if f32[j] in DisabledMiChar then
          f32[j] := '_';
        s := s + format('%s %s,', [sl[i], f32])
      end;
    end;
    mido('Alter Table %s (Rename %s)', [tab, Copy(s, 1, Length(s)-1)]);
  finally
    sl.Free;
  end;
end;

procedure TValuationForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TValuationForm.SelectNewDataBase(Sender: TObject);
var
  ini: TMemIniFile;
begin
  ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
  try
    ini.WriteString('DataBase', 'BaseName', TMenuItem(Sender).Caption);
  finally
    ini.UpdateFile;
    ini.Free;
  end;
  stBar.Panels[0].Text := TMenuItem(Sender).Caption;

  fs.Close;
  ds.Close;
  cf.Close;
  ADOConnection1.Close;
  ADOConnection1.ConnectionString := format(strConnect, [TMenuItem(Sender).Caption]);
  try
    ADOConnection1.Connected := true;
    CreateTablesList;
  except
    RuTablesList.Clear;
    TablesList.Clear;
  end;
end;

procedure TValuationForm.SelectNewSpr(Sender: TObject);
begin
  TMenuItem(Sender).Parent.Caption := TMenuItem(Sender).Caption;
  mnVRIFromBase.Enabled := True;
  mnVRIFromFile.Enabled := True;
  mnVRIFromExcel.Enabled := True;
  ParamForm.ReadRule(GetAppFolder + mnRulesName.Caption + '.xlsx');
end;

procedure TValuationForm.CreateTablesList;
var
  i, k : Integer;
begin
  ds.Close;
  ds.Connection := ADOConnection1;
  fs.Close;
  fs.Connection := ADOConnection1;
  cf.Close;
  cf.Connection := ADOConnection1;

  lbTables.Clear;
  lbFields.Clear;
  RuTablesList.Clear;
  TablesList.Clear;
  fs.CommandText := 'select ID, BASENAME, NAME from RTABLES';//ищем по названию таблицу
  fs.Open;
  while not fs.eof do
  begin
    i := fs.FieldByName('ID').Value;
    for k := Low(StateObjects) to High(StateObjects) do
      if i = StateObjects[k] then
        begin
          RuTablesList.AddObject(fs.FieldByName('NAME').Value, Pointer(i));//название таблицы по русски
          TablesList.AddObject(AnsiUpperCase(fs.FieldByName('BASENAME').Value), Pointer(i));//название таблицы по английски
          break;
        end;
    fs.next;
  end;
  lbTables.Items.Assign(RuTablesList);
end;

procedure TValuationForm.FormCreate(Sender: TObject);
var
  si : TSystemInfo;
  i, j, k : Integer;
  s, id, ver : String;
  sl : TStringList;
  mn : TMenuItem;
  ini : TMemIniFile;
  VerInfo: TVSFixedFileInfo;
begin
  if GetVersionInfo(GetModuleName(HInstance), ver) then
    Caption := Caption + ' ' + ver;

  pstr := ParamStr(1);
  tmp := GetSpecialPath(35) + '\temp~files\';//C:\ProgramData

  log := TStringList.Create;
  mi := CreateOleObject('MapInfo.Application');
  miDo('Set ProgressBars Off');

//  ds := TADODataSet.Create(nil);
//  ds.CacheSize := 100;
//  ds.CursorLocation := clUseServer;
//  ds.Connection := ADOConnection1;
//
//  fs := TADODataSet.Create(nil);
//  fs.CacheSize := 100;
//  fs.CursorLocation := clUseServer;
//  fs.Connection := ADOConnection1;
//
//  cf := TADODataSet.Create(nil);
//  cf.CacheSize := 100;
//  cf.CursorLocation := clUseServer;
//  cf.Connection := ADOConnection1;

  ElementsList := TStringList.Create;
  ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
  try
    DutyMap := ini.ReadString('DutyMap', 'MapName', '');
    stBar.Panels[0].Text := ini.ReadString('DataBase', 'BaseName', 'btiAll');
//    ADOConnection1.Close;
//    ADOConnection1.ConnectionString := format(strConnect, [stBar.Panels[0].Text]);
//    ADOConnection1.Connected := true;

    ElementsList.CommaText := ini.ReadString('Klassificator', 'elements', 'ObjectType');
    for i  := 0 to ElementsList.Count-1 do
    begin
      k := Pos(':', ElementsList[i]);
      if k > 0  then
      begin
        j := StrToInt(Copy(ElementsList[i], k+1, 100));//KLASS_ID
        ElementsList[i] := Copy(ElementsList[i], 1, k-1);
        ElementsList.Objects[i] := Pointer(j);
      end;
    end;
  finally
    ini.Free;
    ElementsList.Sort;
  end;

//  fs.Close;
//  fs.CommandText := 'select name from master.sys.databases WHERE '+
//    'name NOT IN (''master'', ''tempdb'', ''model'', ''msdb'') AND state = 0';
//  fs.Open;
//  while not fs.eof do
//  begin
//    mn := TMenuItem.Create(mnBase);
//    mn.Name := 'mn' + fs.FieldByName('Name').Value;
//    mn.Caption := fs.FieldByName('Name').Value;
//    mn.OnClick := SelectNewDataBase;
//    mnBase.Add(mn);
//    fs.Next;
//  end;
//  fs.Close;

  RuTablesList := TStringList.Create;
  TablesList := TStringList.Create;
  RuFieldsList := TStringList.Create;
  FieldsList := TStringList.Create;

//  CreateTablesList;

  DataSourcesList := TStringList.Create;
//  try//из таблицы ИСТОЧНИКИ ДАННЫХ выбираем связи SourceType и AccessID
//    ds.CommandText := 'select * from DATASOURCES';
//    ds.Open;
//    while not ds.eof do//выбираем данные
//    begin
//      id := ds.FieldByName('id').AsString;
//      i := ds.FieldByName('AccessID').AsInteger;
//      DataSourcesList.AddObject(id, Pointer(i));
//      ds.next;
//    end;
//    DataSourcesList.Sort;
//  finally
//    ds.Close;
//  end;

  GetSystemInfo(si);
  cbCPU.Clear;
  for i := 1 to si.dwNumberOfProcessors do
    cbCPU.Items.Add(IntToStr(i));
  cbCPU.ItemIndex := cbCPU.Items.Count-1;
  ResFolder := GetAppFolder + 'RESULT\';
  ForceDirectories(ResFolder);

  sl := TStringList.Create;
  try
    k := 0;
    FindRecursive(GetAppFolder, '*.xlsx', false, sl);
    for i := 0 to sl.Count-1 do
    begin
      s := AnsiUpperCase(ExtractFileName(sl[i]));
      if Pos('ПРАВИЛА_', s) = 1 then
      begin
        mn := TMenuItem.Create(mnRulesName);
        SetLength(s, Length(s)-5);
        mn.Name := 'mn' + s;
        mn.Caption := s;
        mn.OnClick := SelectNewSpr;
        mnRulesName.Add(mn);
        Inc(k);
      end;
    end;
    if k = 0 then
       Exc('Не найдены справочники правил!');
  finally
    sl.Free;
  end;

end;

procedure TValuationForm.FormDestroy(Sender: TObject);
begin
  if ParamForm.Visible then
    ParamForm.Close;
  miDo('Close All');
  miDo('End MapInfo');
  mi := Unassigned;
  log.Free;
  if ds <> nil then ds.Free;
  if fs <> nil then fs.Free;
  if cf <> nil then cf.Free;
  RuTablesList.Free;
  TablesList.Free;
  RuFieldsList.Free;
  FieldsList.Free;
  DataSourcesList.Free;
  ElementsList.Free;
end;

procedure TValuationForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 27 then
    BreakCommand := true;
end;

//rfields = объединить = fieldmode
procedure TValuationForm.LoadFields;
var
  id : Integer;
begin
  fs.Close;
  id := Integer(TablesList.Objects[lbTables.ItemIndex]);
  fs.CommandText := format('select BASENAME,NAME from RFIELDS WHERE TABLE_ID=%d', [id]);
  fs.Open;
  RuFieldsList.Clear;
  FieldsList.Clear;
  while not fs.eof do
  begin
    RuFieldsList.Add(fs.FieldByName('NAME').Value);//название поля по русски
    FieldsList.Add(AnsiUpperCase(fs.FieldByName('BASENAME').Value));//название поля по английски
    fs.next;
  end;
//  RuFieldsList.SaveToFile(GetAppFolder+'ru'+lbTables.Items[lbTables.ItemIndex]+'.txt');
//  FieldsList.SaveToFile(GetAppFolder+lbTables.Items[lbTables.ItemIndex]+'.txt');
  lbFields.Items.Assign(RuFieldsList);
end;

procedure TValuationForm.lbTablesClick(Sender: TObject);
begin
  if lbTables.ItemIndex > -1 then
  begin
    LoadFields;
    CommandString.Text := format('select * from %s', [TablesList[lbTables.ItemIndex]]);
    mnObjects.Hint := RuTablesList[lbTables.ItemIndex];
  end;
end;

procedure TValuationForm.lbTablesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i : Integer;
begin
  i := lbTables.ItemIndex;
  if Key = 112 then
    lbTables.Items.Assign(RuTablesList);
  if Key = 113 then
    lbTables.Items.Assign(TablesList);
  lbTables.ItemIndex := i;
end;

procedure TValuationForm.lbFieldsClick(Sender: TObject);
var
  i : Integer;
  fields : string;
begin
  if lbTables.ItemIndex <> -1 then
  begin
    if  lbFields.SelCount > 0 then
    begin
      fields := '';
      for i  := 0 to lbFields.Count-1 do
      if lbFields.Selected[i] then
          fields := fields + FieldsList[i] + ',';
      SetLength(fields, Length(fields) - 1);
      CommandString.Text := format('select %s from %s WHERE ', [fields, TablesList[lbTables.ItemIndex]]);
    end else
      CommandString.Text := format('select * from %s WHERE ', [TablesList[lbTables.ItemIndex]]);
  end;
end;

procedure TValuationForm.lbFieldsKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i, k : Integer;
  selar : Array of Integer;
begin
  if Key = 114 then
  begin
    for i := 0 to lbFields.Count-1 do
      lbFields.Selected[i] := false;
    //отмечаем нужные поля
    for k := Low(farray) to High(farray) do
      for i := 0 to lbFields.Count-1 do
        if FieldsList[i] = farray[k] then
          lbFields.Selected[i] := true;
    lbFieldsClick(nil);
    CommandString.Text := CommandString.Text + '(FLAGLOCALITYCODE = ''Не готов'' OR FLAGLOCALITYCODE is Null)';
    exit;
  end;

  SetLength(selar, 0);
  //запомним выборку
  for i := 0 to lbFields.Count-1 do
    if lbFields.Selected[i] then
    begin
      SetLength(selar, Length(selar)+1);
      selar[Length(selar)-1] := i;
    end;

  if Key = 112 then
    lbFields.Items.Assign(RuFieldsList);
  if Key = 113 then
    lbFields.Items.Assign(FieldsList);

  //восстановим выборку
  for i := Low(selar) to High(selar) do
    lbFields.Selected[selar[i]] := true;
  SetLength(selar, 0);
end;
//создание таблицы по выборке (как из текстового файла)
function TValuationForm.ImportDataTAB(msg : Boolean) : Boolean;
var
  s, ss, columnlist, tab : String;
  f32 : String[32];
  i, j, id, k : Integer;
  fld : TStringList;
  fds : TFieldDefs;
  fd : TFieldDef;
begin
  BreakCommand := false;
  Result := false;
  if BreakCommand then Exit;
  if lbTables.ItemIndex = -1 then
    if msg then
      Exc('Не выбрано имя базы')
    else
      Exit;
  if lbFields.SelCount < 1 then
    if msg then
      Exc('Нет выбранных полей таблицы ' + TablesList[lbTables.ItemIndex])
    else
      Exit;

  ds.Close;
  Case ParamForm.rgFlagNP.ItemIndex of
    0 : ds.CommandText := CommandString.Text;
    1 : ds.CommandText := StringReplace(CommandString.Text, 'FLAGLOCALITYCODE = ''Не готов'' OR ', '', []);
    2 : ds.CommandText := StringReplace(CommandString.Text, ' OR FLAGLOCALITYCODE is Null', '',[]);
  end;
  Info('Выборка по запросу');
  ds.Open;

//ds.SaveToFile(format('%s%s.xml', [GetAppFolder, RuTablesList[lbTables.ItemIndex]]));
  if ds.IsEmpty then
    if msg then
      Exc('Нет выбранных записей')
    else
      Exit;
  ds.FindLast;
  pg.Max := ds.RecNo;
  ds.FindFirst;

  fs.Close;
  fs.CommandText := 'select ID, BASENAME, NAME from RTABLES';//ищем по названию таблицу
  fs.Open;
  while not fs.eof do
  begin
    if (not fs.FieldByName('BASENAME').IsNull) and CText(fs.FieldByName('BASENAME').AsString, TablesList[lbTables.ItemIndex]) then
    begin
      id := fs.FieldByName('ID').Value;//запомним ID таблицы и название
      basename := fs.FieldByName('NAME').Value;//название таблицы по русски
      basename_ := AnsiUpperCase(fs.FieldByName('BASENAME').Value);//название таблицы по английски
      fs.Close;
      fs.CommandText := format('select * from RFIELDS WHERE TABLE_ID=%d', [id]);
      fs.Open;
//      fs.SaveToFile(format('%s%s.xml', [GetAppFolder, 'RFIELDS']));
      break;
    end;
    fs.next;
  end;

  fld := TStringList.Create;//поиск выбранных полей базы
  fds := ds.FieldDefs;
  id := -1;
  for i := 0 to fds.Count-1 do
  begin
    fd := fds.Items[i];
    f32 := fd.Name;
    if CText('FLAGLOCALITYCODE', f32)then
      id := i;//признак наличия флага
    case fd.DataType of
      ftString, ftBlob, ftMemo, ftFmtMemo, ftBytes, ftVarBytes, ftFixedChar,
      ftFixedWideChar, ftWideMemo, ftByte, ftGuid,
      ftWideString: fld.AddObject(format('%s Char(%d)', [f32, fd.size]), Pointer(fd.size));
      ftBoolean: fld.Add(f32+' Logical');
      ftFloat, ftSingle, ftExtended, ftCurrency, ftFMTBcd,
      ftBCD: fld.Add(f32+' Float');
      ftDate, ftTimeStamp, ftOraTimeStamp, ftDateTime, ftTime: fld.Add(f32+' Date');
      ftSmallint, ftInteger, ftWord, ftAutoInc, ftLongWord, ftShortint,
      ftLargeInt: fld.Add(f32+' Integer');
      else
        fld.AddObject(format('%s Char(%d)', [f32, fd.size]), Pointer(fd.size));
    end;

  end;

  tab := miEval('PathToTableName$("%s")', [mnObjects.Hint]);//создаем таблицу MI
  if TableExists(tab) then
    miDo('Close table %s', [tab]);
  columnlist := '(';
  for i := 0 to fld.Count-1 do
    columnlist := columnlist + fld[i] + ',';
  columnlist[Length(columnlist)] := ')';

  miDo('Create Table "%s" %s file "%s.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
    [tab, columnlist, ResFolder + mnObjects.Hint]);

  Info(basename + ': Выборка из базы');
  Breakcommand := false;
  pg.Position := 0;
  while not ds.eof do
  begin
    Application.ProcessMessages;
    if BreakCommand then break;
    s := '';
    for i := 0 to fds.Count-1 do
    begin
      if not ds.FieldByName(fds[i].name).IsNull then
      begin
        ss := StringReplace(ds.FieldByName(fds[i].name).AsString, '"', '""', [rfReplaceAll]);
        if fld.Objects[i] <> nil then
          ss := '"' + ss + '"';
      end else
        if fld.Objects[i] <> nil then
          ss := '""'
        else
          ss := '0';

      if CText(fds[i].name, 'SOURCELOCALITYCODE') then
        if DataSourcesList.Find(ss, k) then//заменяем код значение из справоника
          ss := IntToStr(Integer(DataSourcesList.Objects[k]));

      if CText(ss, 'TRUE') then
        ss := '1'
      else
      if CText(ss, 'FALSE') then
        ss := '0';

      s := s + ss + ',';
    end;
    ds.next;
    pg.Position := pg.Position + 1;
    SetLength(s, Length(s) - 1);
    miDo('Insert into %s Values(%s)', [tab, s]);

  end;
  pg.Max := 0;
  fld.Free;

  if not BreakCommand then
  begin
    miDo('Commit table %s', [tab]);
    if id <> -1 then
    begin
      miDo('Select * from %s where FLAGLOCALITYCODE="" into temp noSelect', [tab]);
      miDo('update temp set FLAGLOCALITYCODE="НЕ ГОТОВ"');
      miDo('Update %s Set FLAGLOCALITYCODE = UCase$(FLAGLOCALITYCODE)', [tab]);
      miDo('Commit table %s', [tab]);
    end;

    j := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
    miDo('Close table %s', [tab]);
    OpenDialog1.InitialDir := GetAppFolder;
    OpenDialog1.FileName := (format('%s%s.TAB', [ResFolder, basename]));
    Val(cbCPU.Text, cpu, i);
    edSize.Text := IntToStr(Round(j/cpu+3));
    Application.ProcessMessages;
    Result := true;
  end else
    miDo('Drop table %s', [tab]);

  if msg then
    if BreakCommand then
      ShowMessage('Выборка прервана')
    else
      ShowMessage('Выборка завершена' + #10 + OpenDialog1.FileName);
end;

//для выборки из текстового файла
function TValuationForm.ImportData : Boolean;
var
  tab : String;
  f32 : String[32];
  i : Integer;
  fds : TFieldDefs;
  fd : TFieldDef;
  fld : TStringList;

  procedure Ds2Mif;
  var
    mif, mid : TextFile;
    s, ss, columnlist : String;
    i, j, k : Integer;
  begin
    tab := miEval('PathToTableName$("%s")', [basename]);
    columnlist := 'Version  300' + #10;
    columnlist := columnlist + 'Charset "WindowsCyrillic"' + #10;
    columnlist := columnlist + 'Delimiter ","' + #10;
    columnlist := columnlist + format('Columns %d%s', [fld.Count, #10]);

    for i := 0 to fld.Count-1 do
      columnlist := columnlist + format('%s%s', [fld[i], #10]);
    columnlist := columnlist + format('Data%s', [#10]);

    AssignFile(mif, format('%s.mif', [ResFolder + basename]));
    Rewrite(mif);
    AssignFile(mid, format('%s.mid', [ResFolder + basename]));
    Rewrite(mid);
    Writeln(mif, columnlist);

    pg.Max := ds.recordcount;
    Info(basename + ': Создание таблицы');
    k := 0;
    try
      while not ds.eof do//выбираем данные
      begin
        Inc(k);
        if k mod 10 = 0  then
          pg.Position  := pg.Position + 10;

          s := '';
          for i := Low(farray) to High(farray) do
          begin
            ss := ds.FieldByName(fds[i].name).AsString;
            if ss = '' then
            begin
              if fld.Objects[i] <> nil then
                ss := '""'
              else
                ss := '0';
            end else
              if fld.Objects[i] <> nil then
                ss := StringReplace(ss, '"', '""', [rfReplaceAll]);

//          if not ds.FieldByName(farray[i]).IsNull then
//          begin
//            ss := StringReplace(ds.FieldByName(farray[i]).AsString, '"', '""', [rfReplaceAll]);
//            if fld.Objects[i] <> nil then
//              ss := '"' + ss + '"';
//          end else
//            if fld.Objects[i] <> nil then//объект <> 0 это текстовое поле
//              ss := '""'
//            else
//              ss := '0';

          if CText(farray[i], 'SOURCELOCALITYCODE') then
            if DataSourcesList.Find(ss, j) then//заменяем код значение из справоника
              ss := IntToStr(Integer(DataSourcesList.Objects[j]));

          if CText(ss, 'True') then
            ss := '1'
          else if CText(ss, 'False') then
            ss := '0';

          s := s + ss + ',';
        end;
        Writeln(mif,'none');
        Writeln(mid, Copy(s, 1, Length(s)-1));
        ds.next;
      end;
    finally
      CloseFile(mif);
      CloseFile(mid);
      miDo('Import "%s.mif" Type "MIF" Into "%s.tab" Overwrite', [ResFolder + basename, ResFolder + basename]);
      DeleteFile(format('%s.mif', [ResFolder + basename]));
      DeleteFile(format('%s.mid', [ResFolder + basename]));
      if miEval('TableInfo(%s, %d)', [tab, TAB_INFO_MAPPABLE])='T' then
        miDo('drop map %s', [tab]);//basename модет быть с пробелом, заменяем на tab
    end;
  end;

begin
  Result := false;

  fld := TStringList.Create;//создание списка полей для MI
  try
    fds := ds.FieldDefs;
    for i := 0 to fds.Count-1 do
    begin
      fd := fds.Items[i];
      f32 := fd.Name;
      case fd.DataType of
        ftString, ftBlob, ftMemo, ftFmtMemo, ftBytes, ftVarBytes, ftFixedChar,
        ftFixedWideChar, ftWideMemo, ftByte, ftGuid,
        ftWideString: fld.AddObject(format('%s Char(%d)', [f32, fd.size]), Pointer(fd.size));
        ftBoolean: fld.Add(f32+' Logical');
        ftFloat, ftSingle, ftExtended, ftCurrency, ftFMTBcd,
        ftBCD: fld.Add(f32+' Float');
        ftDate, ftTimeStamp, ftOraTimeStamp, ftDateTime, ftTime: fld.Add(f32+' Date');
        ftSmallint, ftInteger, ftWord, ftAutoInc, ftLongWord, ftShortint,
        ftLargeInt: fld.Add(f32+' Integer');
        else
          fld.AddObject(format('%s Char(%d)', [f32, fd.size]), Pointer(fd.size));
      end;
    end;

    Ds2Mif;

    miDo('Select * from %s where FLAGLOCALITYCODE="" into temp noSelect', [tab]);
    miDo('update temp set FLAGLOCALITYCODE="НЕ ГОТОВ"');//пустые значения определяем как "НЕ ГОТОВ"
    miDo('Update %s Set FLAGLOCALITYCODE = UCase$(FLAGLOCALITYCODE)', [tab]);

    miDo('Commit table %s', [tab]);
    miDo('Close table %s', [tab]);
    OpenDialog1.FileName := (format('%s%s.TAB', [ResFolder, basename]));
    Result := true;
  finally
    fld.Free;
    Info('');
  end;
end;
//повторный вызов этой функции снимает фильтр
procedure Setfilter(var sheet, app  : OleVariant; row : Integer);
begin
  sheet.Rows[row].Select;
  sheet.Range[format('A%d', [row])].Activate;
  app.Selection.AutoFilter;
end;

procedure SetSheetParams(var book, sheet, app : OleVariant;
  sheetname : String; columswith : Array of Integer; title : Integer);
var
  i : Integer;
begin
//  if sheetname <> '' then//для активного листа это делать  ненадо
//  begin
//    i := book.WorkSheets.Count;
//    book.WorkSheets.Item[i].Activate;//активируе последний лист
//    book.WorkSheets.Add(After:=app.ActiveSheet, Count:=1);//добавим лист
//    book.WorkSheets.Item[i + 1].Activate;//активируе последний лист
//    sheet := app.ActiveSheet;
//    sheet.Name := sheetname;
//  end;
  sheet.Rows[title + 1].Select;
  for i := Low(columswith) to High(columswith) do
    sheet.Columns[i+1].ColumnWidth := columswith[i];
  app.ActiveWindow.SplitColumn := 0;
  app.ActiveWindow.SplitRow := title;
  app.ActiveWindow.FreezePanes := True;
end;

procedure CheclTypes(var sheet : OleVariant);
var
  range, data : OleVariant;
  i, j, c : Integer;
begin
  c := sheet.UsedRange.Columns.Count;
  range := sheet.range[format('A1:%s1', [ExcelNum2Str(sheet, c)])];
  data := range.value;
  try
    for j := 1 to c do
    begin
      if SameText(data[1, j], 'String') then
        sheet.Columns[j].NumberFormat := '@' else
      if SameText(data[1, j],'Integer') then
        sheet.Columns[j].NumberFormat := '0' else
      if SameText(data[1, j], 'Data') then
        sheet.Columns[j].NumberFormat := 'm/d/yyyy';
    end;
  finally
    VarClear(data);
  end;
end;

procedure TValuationForm.mnParamBuildingsClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet, range, DataParcels, data, cn : OleVariant;
  filename, pattern, s, ss, temppath, source, buildings, zu : String;
  row, col, row2, col2, zrow, zcol, i, j, k, wall, wallcount, form, linkzu : Integer;
  sa : StringArray;

  function LoadBuildings(var dp : OleVariant; xlsname : String) : Boolean;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
    i, j, cd, rd : Integer;
  begin
    Info('Загрузка данных...');
    Result := false;
    wall := 0;//начало колонок с материалом стен

    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(xlsname);
      WorkBook := ExcelApp.Workbooks.item[1];
      WorkSheet := WorkBook.WorkSheets.item[1];//активный лист 1
      if not FindCadNum(WorkSheet, rd, cd, row, col) then
        Exc('Не найдена колонка КН для связи с ЗУ файле:'+ #10 + xlsname);
      linkzu := cd;//столбец для связи с земельными участками
      range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, col), row])];
      dp := range.value;
      for i := 1 to col do
      begin
        if SameText(String(dp[rd-1, i]), 'материал стен') then
        begin
          if wall = 0 then
            wall := i;
          Inc(wallcount);
        end;
        if SameText(String(dp[rd-1, i]), 'список кадастровых номеров объектов содержащих текущий') then
          linkzu := i;
      end;
      if wallcount = 0 then
        Exc('Не найдена колонка материала стен в файле:'+ #10 + xlsname);
      Result := True;;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  end;

  function LoadZU(var dp : OleVariant; xlsname : String) : Boolean;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
    i, j : Integer;
  begin
    Info('Загрузка земельных участков...');
    Result := false;
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(xlsname);
      WorkBook := ExcelApp.Workbooks.item[1];
      WorkSheet := WorkBook.WorkSheets.item[1];//активный лист 1
      zcol := WorkSheet.UsedRange.Columns.Count;
      zrow := WorkSheet.UsedRange.Rows.Count;
      range := WorkSheet.range[format('B5:%s%d', [ExcelNum2Str(WorkSheet, zcol), zrow])];
      dp := range.value;
      Result := true;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
      if not Result then
        Note('Ошибка при загрузке файла:'+ #10 + xlsname);
    end;
  end;

  procedure IntermediateParameter;//Промежуточный параметр
  var
    data, data2 : OleVariant;
    r, c, k, n, min, max, group : Integer;
    s : String;
    sa : StringArray;
  begin
    info('Промежуточный параметр...');
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;
    k := 0;
    for c := col to col2 do//подсчет числа столбцов с одинаковым названием
      if String(WorkSheet.Cells[row-1, c]) = 'материал стен' then//строка заголовка
      begin
        Inc(k);
        SetLength(sa, k);
        sa[k-1] := ExcelNum2Str(WorkSheet, c);//добавляем название столбца
      end;
    s := sa[Length(sa) - 1];//последний столбец (может быть равен первому)
    range := WorkSheet.range[format('%s%d:%s%d', [sa[0], row-1, s, row2])];
    data := range.value;//данные включают заголовок
    //заполняяем кадастровыми номерами
    data2 := VarArrayCreate([1, VarArrayHighBound(cn, 1), 1, 4], varOleStr);
    for r := VarArrayLowBound(cn,1) to VarArrayHighBound(cn,1) do
      data2[r, 1] := cn[r, 1];
    //заполняяем значениями пропустив заголовок
    for r := VarArrayLowBound(cn, 1)+1 to VarArrayHighBound(cn, 1) do
    begin
      s := data[r, 1];

      for k := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do
        if s = String(DataParcels[k, 1]) then
        begin
          min := DataParcels[k, 2];
          max := DataParcels[k, 2];
        end;

      for c := 2 to Length(sa) do//добавляем данные начиная со 2 столбца
        if String(data[r, c]) <> '' then
        begin
          s := s + ',' + data[r, c];

          for k := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do
            if String(data[r, c]) = String(DataParcels[k, 1]) then
            begin
              group := DataParcels[k, 2];
              if min > group then
                min := group;
              if max < group then
                max := group;
            end;
        end;
      data2[r, 2] := s;
      data2[r, 3] := min;
      data2[r, 4] := max;
    end;
    data2[1, 2] := data[1, 1];//заголовок второго столбца из первого столбца данных
    data2[1, 3] := 'min';
    data2[1, 4] := 'max';

    FindSheetByName(WorkBook, WorkSheet, 'Промежуточный параметр');
    SetSheetParams(WorkBook, WorkSheet, ExcelApp, '', [20, 40], 1);
    range := WorkSheet.range[format('A1:%s%d',
      [ExcelNum2Str(WorkSheet, VarArrayHighBound(data2, 2)), VarArrayHighBound(data2, 1)])];
    range.value := data2;

    VarClear(data);
    VarClear(data2);
    SetLength(sa, 0);
  end;

  procedure WallMaterial;//материал стен
  var
    data, data2 : OleVariant;
    r, c, k : Integer;
    s : String;
    sa : StringArray;
  begin
    info('материал стен...');
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;
    k := 0;
    for c := col to col2 do//подсчет числа столбцов с одинаковым названием
      if String(WorkSheet.Cells[row-1, c]) = 'материал стен' then//строка заголовка
      begin
        Inc(k);
        SetLength(sa, k);
        sa[k-1] := ExcelNum2Str(WorkSheet, c);//добавляем название столбца
      end;
    s := sa[Length(sa) - 1];//последний столбец (может быть равен первому)
    range := WorkSheet.range[format('%s%d:%s%d', [sa[0], row-1, s, row2])];
    data := range.value;//данные включают заголовок

    k := VarArrayHighBound(cn, 1);//считаеы общее количество материалов
    for r := VarArrayLowBound(cn, 1)+1 to VarArrayHighBound(cn, 1) do//заголовок пропускаем
      for c := 2 to Length(sa) do//добавляем данные начиная со 2 столбца
        if String(data[r, c]) <> '' then
          Inc(k);
    //заполняяем кадастровыми номерами повторяя их для повторяющихся материал стен
    data2 := VarArrayCreate([1, k, 1, 5], varOleStr);
    k := 1;
    r := 1;
    While r <= VarArrayHighBound(data2,1) do
    begin
      data2[r, 1] := cn[k, 1];//кадастровый номер
      data2[r, 2] := data[k, 1];//данные первого столбца
      if r > 1 then//первый шаг это заполнение заголовка, пропускам добавления
      for c := 2 to Length(sa) do//добавляем данные начиная со 2 столбца
        if String(data[k, c]) <> '' then
        begin
          Inc(r);
          data2[r, 1] := cn[k, 1];
          data2[r, 2] := data[k, c];
        end;
      Inc(r);
      Inc(k);
    end;
    for r := VarArrayLowBound(data2, 1)+1 to VarArrayHighBound(data2, 1) do
      for k := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do
        if String(data2[r, 2]) = String(DataParcels[k, 1]) then
        begin
          data2[r, 3] := DataParcels[k, 2];
          data2[r, 4] := DataParcels[k, 3];
          data2[r, 5] := DataParcels[k, 4];
        end;
    data2[1, 3] := 'Группа капитальности присвоена';
    data2[1, 4] := 'Возраст';
    data2[1, 5] := 'Возраст2';

    FindSheetByName(WorkBook, WorkSheet, 'Промежуточная табл 1');
    SetSheetParams(WorkBook, WorkSheet, ExcelApp, 'Промежуточная табл 1', [20, 30], 1);
    range := WorkSheet.range[format('A1:%s%d',
      [ExcelNum2Str(WorkSheet, 5), VarArrayHighBound(data2, 1)])];//колонка "Возраст2"
    range.value := data2;

    VarClear(data);
    VarClear(data2);
    SetLength(sa, 0);
  end;

  procedure Table2(colname, sheetname : String; itog : Boolean);//Промежуточная табл X
  var
    data, data2 : OleVariant;
    r, c, k, n : Integer;
    s, kn : String;
    sa : StringArray;
    sl : TStringList;
  begin
    if not FindSheetByName(WorkBook, WorkSheet, sheetname) then
      Exc(format('Лист "%s" не найден.', [sheetname]));

    info(sheetname);
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;
    for c := col to col2 do//ищем колонку с данным названием
      if String(WorkSheet.Cells[row-1, c]) = colname then//строка заголовка
    begin
      range := WorkSheet.range[format('%s%d:%s%d',
        [ExcelNum2Str(WorkSheet, c), row-1, ExcelNum2Str(WorkSheet, c), row2])];
      data := range.value;//данные включают заголовок
      //будем заносить данные в виде длинных строк
      data2 := VarArrayCreate([1, 2, 1, 1], varOleStr);
      data2[2, 1] := cn[1, 1];
      data2[1, 1] := data[1, 1];
      n := 2;
      for r := VarArrayLowBound(cn, 1)+1 to VarArrayHighBound(cn, 1) do
      begin
        kn := cn[r, 1];
        s := data[r, 1];
        sa := SplitStr(s, ';');
        if Length(sa) > 1 then
        for k := Low(sa) to High(sa) do
        begin
          VarArrayRedim(data2, n);
          data2[2, n] := kn;
          data2[1, n] := sa[k];
          Inc(n);
        end;
        SetLength(sa, 0);
      end;
      //переведем строки в столбцы
      VarClear(data);
      data := VarArrayCreate([1, VarArrayHighBound(data2, 2), 1, 2], varOleStr);
      for r := VarArrayLowBound(data2, 2) to VarArrayHighBound(data2, 2) do
      begin
        data[r, 1] := data2[2, r];
        data[r, 2] := data2[1, r];
      end;
    end;

    FindSheetByName(WorkBook, WorkSheet, sheetname);
    WorkSheet.Cells.ClearContents;
    SetSheetParams(WorkBook, WorkSheet, ExcelApp, sheetname, [20, 20], 1);
    range := WorkSheet.range[format('A1:B%d', [VarArrayHighBound(data, 1)])];
    range.value := data;

    Info('подсчет количества КН...');
    if itog then//подсчет количества по кажому КН
    begin
      sl := TStringList.Create;
      sl.Duplicates := dupIgnore;
      sl.Sorted := true;
      for r := VarArrayLowBound(data, 1)+1 to VarArrayHighBound(data, 1) do
        sl.Add(String(data[r, 1]));//уникальные КН

      for r := 0 to sl.Count-1 do
      begin
        k := 0;//подсчет количества уникальных КН
        for n := VarArrayLowBound(data, 1)+1 to VarArrayHighBound(data, 1) do
          if sl[r] = String(data[n, 1]) then
            Inc(k);
        sl.Objects[r] := Pointer(k);
      end;

      if sl.Count > 0 then
      begin
        VarClear(data2);
        data2 := VarArrayCreate([1, VarArrayHighBound(data, 1), 1, 1], varOleStr);
        for r := VarArrayLowBound(data, 1)+1 to VarArrayHighBound(data, 1) do
          if sl.Find(String(data[r, 1]), n) then
            data2[r-1, 1] := Integer(sl.Objects[n])
          else
            data2[r-1, 1] := '0';

        range := WorkSheet.range[format('C2:C%d', [VarArrayHighBound(data, 1)+1])];
        range.value := data2;
      end;
      sl.Free;
    end;

    VarClear(data);
    VarClear(data2);
  end;

begin//mnParamBuildingsClick
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\test\TelDa_XLS'; {$ENDIF}
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Шаблон_Здания_Год_Этажность_Материал.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Шаблон_Здания_Год_Этажность_Материал.xlsx');

  OpenDialog1.FileName := '*.xls*';
  try
    OpenDialog1.Title := 'Открытие файла Зданий';
    if not OpenDialog1.Execute then
      exit;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    buildings := OpenDialog1.FileName;

    OpenDialog1.Title := 'Открытие файла ЗУ';
    if not OpenDialog1.Execute then
      exit;
    zu := OpenDialog1.FileName;

    ResFolder := GetAppFolder + 'RESULT\TelDa\' + AddSlesh(CurrentData);
    if not ForceDirectories(ResFolder) then
      Exc('Не удалось создать каталог' + #10 + ResFolder);
    filename := ExtractFileName(buildings);
    filename := ResFolder + 'Здания_' + ChangeFileExt(filename, '.xlsm');
    if FileExists(filename) then
      if not DeleteFile(filename) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [filename, #10]));

    if LoadZU(data, zu) then//загружаем земельные участки
    try
      if LoadBuildings(DataParcels, buildings) then//загружаем данные в DataParcels
      try
        info('Загрузка шаблона...');
        ExcelApp := CreateOleObject('Excel.Application');
        ExcelApp.DisplayAlerts := False;

        pattern := 'PATTERNS\TelDa\Шаблон_Здания_Год_Этажность_Материал.xlsx';
        source := GetAppFolder + pattern;
        temppath := ExtractFilePath(miEval('TempFileName$("")'));
        ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"',
          [source, temppath])), SW_HIDE, true, 0);
        s := temppath + ExtractFileName(source);
        temppath := ChangeFileExt(s, '_.xlsx');
        RenameFile(s, temppath);//новое название должно отличаться от шаблона
        ExcelApp.Workbooks.Open(temppath);//открываем копию
        WorkBook := ExcelApp.Workbooks.item[1];
        if FindSheetByName(WorkBook, WorkSheet, 'Объекты') then
        begin
          //добавим нужное количество колонок к материалу стен
          if wallcount > 1 then
          begin
            s := ExcelNum2Str(WorkSheet, wall);//первый столбец с матариалом стен
            range := WorkSheet.Range[format('%s:%s', [s, s])];
            range.Select;
            for i := 2 to wallcount do//вставляем дополнительные столбцы
              ExcelApp.Selection.Insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove);
          end;
          form := col + 1 + 2;//начало формул (колонка после данных + информация по ЗУ)
          //вставляем данные,число строк и столбцоы определено в LoadBuildings
          range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, col), row])];
          range.value := DataParcels;
          VarClear(DataParcels);
        end else
          Exc('Лист "Объекты" не найден.');
        if FindSheetByName(WorkBook, WorkSheet, 'СПР МС', row2, col2) then
        begin
          range := WorkSheet.range[format('B2:%s%d', [ExcelNum2Str(WorkSheet, col2), row2])];
          DataParcels := range.value;//загружаем справочник в DataParcels

          FindSheetByName(WorkBook, WorkSheet, 'Объекты');
          FindCadNum(WorkSheet, row, col, row2, col2);
          s := ExcelNum2Str(WorkSheet, col);
          range := WorkSheet.range[format('%s%d:%s%d', [s, row-1, s, row2])];
          cn := range.value;//массив кадастровых номеров с заголовком

          IntermediateParameter;
          WallMaterial;//Промежуточная табл 1
          Table2('список кадастровых номеров машино-мест', 'Промежуточная табл 2', true);
          Table2('список кадастровых номеров объектов содержащих текущий', 'Промежуточная табл 3', false);
          Table2('список кадастровых номеров помещений', 'Промежуточная табл 4', true);

          info('Поиск информации по ЗУ...');
          VarClear(cn);
          VarClear(DataParcels);
          FindSheetByName(WorkBook, WorkSheet, 'Объекты');
          //берем блок данных по размеру от список кадастровых номеров объектов содержащих текущий до формул
          s := ExcelNum2Str(WorkSheet, linkzu);
          ss := ExcelNum2Str(WorkSheet, form-1);
          range := WorkSheet.range[format('%s%d:%s%d', [s, row, ss, row2])];
          DataParcels := range.value;//массив кадастровых номеров с заголовком   380x7
          //добавляем в блок информацию из ЗУ
          for i := VarArrayLowBound(DataParcels, 1) to VarArrayHighBound(DataParcels, 1) do
          begin
            sa := SplitStr(String(DataParcels[i, 1]), ';');
            if Length(sa) = 1 then
            begin
              for j := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
              if data[j, 1] = sa[0] then
              begin
                DataParcels[i, VarArrayHighBound(DataParcels, 2)-1] := data[j, VarArrayHighBound(data, 2)-1];
                DataParcels[i, VarArrayHighBound(DataParcels, 2)] := String(data[j, VarArrayHighBound(data, 2)]);
                break;
              end;
            end;
            SetLength(sa, 0);
          end;
          range.value := DataParcels;//вставляем заполненный блок обратно

          info('Копирование формул...');
          FindSheetByName(WorkBook, WorkSheet, 'Объекты', row2, col2);
          s := ExcelNum2Str(WorkSheet, form);
          ss := ExcelNum2Str(WorkSheet, col2);
          WorkSheet.Range[format('%s%d:%s%d', [s, row, ss, row])].Copy;//копирование форммул
          WorkSheet.Range[format('%s%d:%s%d', [s, row+1, ss, row2])].PasteSpecial(operation:=xlNone);

          for k := form to col2 do
          begin
            if (String(WorkSheet.Cells[1, k]) = '*') and (WorkSheet.Cells[row, k].HasFormula) then
            begin//копируем формулы в память
              s := WorkSheet.Cells[row, k].Formula;
              WorkSheet.Cells[row, k] := s;
              s := ExcelNum2Str(WorkSheet, k);
              ss := ExcelNum2Str(WorkSheet, col2);//конец фрпмул
              WorkSheet.Range[format('%s%d:%s%d', [s, row, ss, row])].Copy;
              WorkSheet.Range[format('%s%d:%s%d', [s, row, ss, row2])].PasteSpecial(operation:=xlNone);
            end;
          end;

          Setfilter(WorkSheet, ExcelApp, row-1);
          SetSheetParams(WorkBook, WorkSheet, ExcelApp, '', [1, 20], row-1);
          WorkBook.SaveAs(Filename:=filename, FileFormat:= 52, CreateBackup:=False);

          info('Готово...');
        end else
          Exc('Лист "СПР МС не найден.');
      finally
        ExcelApp.Visible := true;
        WorkSheet := Unassigned;
        WorkBook := Unassigned;
        ExcelApp := Unassigned;
        VarClear(cn);
      end;
    finally
      VarClear(DataParcels);
      VarClear(data);
    end;
  finally
    OpenDialog1.Title := '';
  end;
end;

procedure TValuationForm.mnFlats_BuildingsClick(Sender: TObject);
var
  sl, log, floormnames : TStringList;
  flats, buildings, filename, pattern, s, s2, temppath, source : String;
  row, col, row2, col2, i, j, k, n, m, floorcount, floor, rb, cb, rb2, cb2, rd, cd : Integer;
  ExcelApp, Workbook, WorkSheet, range, DataParcels, data, data2, datafloor, cn : OleVariant;

  function LoadFlats : Boolean;
  var
    ExcelApp, WorkSheet, range : OleVariant;
    i, j : Integer;
    reg: TRegEx;
    mc : TMatchCollection;
  begin
    Result := false;
    floor := 0;
    floorcount := 0;
    info('Загрузка ' + ExtractFileName(flats));
    VarClear(DataParcels);
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(flats);
      WorkBook := ExcelApp.Workbooks.item[1];
      if not FindSheetByName(WorkBook, WorkSheet, 'Объекты', rd, cd) then//tr!!!
         Exc('лист Объекты в файле помещений не найден!');
//      CheclTypes(WorkSheet, ExcelApp);
      FindCadNum(WorkSheet, row, col, rd, cd);
      range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, cd), rd])];
      DataParcels := range.value;
      for j := 1 to cd do
        if SameText(String(DataParcels[row-1, j]), 'номер этажа')  then
        begin
           if floor = 0 then
             floor := j;
          Inc(floorcount);
        end;
      if floorcount = 0 then
        Exc('Номер этажа в файле ' + ExtractFileName(flats) + ' не найден!');

      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do//по строкам         3
        for j := VarArrayLowBound(DataParcels,2) to VarArrayHighBound(DataParcels,2) do//по столбцам     16
        begin
          DataParcels[i,j] := StringReplace(DataParcels[i,j], 'False', 'ЛОЖЬ', []);
          DataParcels[i,j] := StringReplace(DataParcels[i,j], 'True', 'ИСТИНА', []);
        end;

      Result := true;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      ExcelApp := Unassigned;
    end;
  end;

  function LoadBuildings : Boolean;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
  begin
    Result := false;
    info('Загрузка ' + ExtractFileName(buildings));
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(buildings);
      WorkBook := ExcelApp.Workbooks.item[1];
      if not FindSheetByName(WorkBook, WorkSheet, 'Объекты', rb2, cb2) then
         Exc('лист Объекты в файле помещений не найден!');
//      CheclTypes(WorkSheet, ExcelApp);
      FindCadNum(WorkSheet, rb, cb, rb2, cb2);
      VarClear(data);
      range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, cb2), rb2])];
      data := range.value;
      Result := true;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  end;

  function CalcFloor : Boolean;
  var
    kn : String;
    i, j, n, m : Integer;
  begin
    info('Этажность...');
    if FindSheetByName(WorkBook, WorkSheet, 'Этажность', row2, col2) then
    begin
      //будем заносить данные в виде длинных строк
      VarClear(data2);
      data2 := VarArrayCreate([1, 3, 1, 1], varOleStr);
      data2[1, 1] := DataParcels[row-1, col];//кадастровый номер объекта
      data2[2, 1] := DataParcels[row-1, floor];//номер этажа
      data2[3, 1] := DataParcels[row-1, floor+1];//тип этажа
      n := 2;  m := 0;
      for i := 0 to VarArrayHighBound(DataParcels, 1)-row do//идем по кадастровым номерам
      begin
        k := floor;
        kn := DataParcels[row+i, col];
        VarArrayRedim(data2, n);//первый набор данных
        data2[1, n] := kn;
        data2[2, n] := DataParcels[row+i, k];
        data2[3, n] := DataParcels[row+i, k+1];
        Inc(k, 2);
        Inc(n);
        for j := 2 to floorcount do//оставшиеся столбцы
        begin
          if String(DataParcels[row+i, k]) = '' then
            break;
          VarArrayRedim(data2, n);
          data2[1, n] := kn;
          data2[2, n] := DataParcels[row+i, k];
          data2[3, n] := DataParcels[row+i, k+1];
          Inc(k, 2);
          Inc(n);
        end;
      end;
      //переведем строки в столбцы
      data := VarArrayCreate([1, VarArrayHighBound(data2, 2), 1, 3], varOleStr);
      for j := VarArrayLowBound(data2, 2) to VarArrayHighBound(data2, 2) do
      begin
        data[j, 1] := data2[1, j];//кадастровый номер объекта'
        data[j, 2] := data2[2, j];//номер этажа
        data[j, 3] := data2[3, j];//тип этажа
      end;

      range := WorkSheet.range[format('A1:%s%d',
        [ExcelNum2Str(WorkSheet, VarArrayHighBound(data, 2)), VarArrayHighBound(data, 1)])];
      range.value := data;//вставляем на лист"ЭТАЖНОСТЬ" кадастровый номер объекта, номер этажа, тип этажа

      info('Анализ данных...');
      WorkSheet.Range['D2:F2'].Copy;
      WorkSheet.Range[format('D3:F%d', [VarArrayHighBound(data, 1)])].PasteSpecial(operation:=xlNone);
      //суммируем типы этажей
      range := WorkSheet.range[format('A2:G%d', [VarArrayHighBound(data, 1)])];
      VarClear(data);
      data := range.value;
      i := VarArrayLowBound(data, 1);  //справосник этажноссти
      While i <= VarArrayHighBound(data, 1) do
      begin
        kn := data[i, 1];
        try
          k := data[i, 6];//tr!!!
        except
          k := 0;
          log.Add(data[i, 1] + ': Ошибка определения этажности ' + data[i, 2]);
        end;
        j := i+1;
        While j <= VarArrayHighBound(data, 1) do
        begin
          s := data[j, 1];//kn
          if s <> kn then
            break;
          try
            n := data[j, 6];
          except
            n := 0;
          end;
          Inc(k, n);
          Inc(j);
        end;
        data[i, 7] := k;
        i := j;
      end;

      if log.Count > 0 then exit;//tr!!!

      //Этажность-формируем массив размером VarArrayHighBound(data, 1) в одну колонку
      VarClear(data2);
      data2 := VarArrayCreate([1, VarArrayHighBound(data, 1), 1, 1], varOleStr);
      for i := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
        data2[i, 1] := data[i, 7];
      range := WorkSheet.range[format('G2:G%d', [VarArrayHighBound(data, 1)+1])];//+заголовок
      range.value := data2;//переносим результаты в одну колонку

      FindSheetByName(WorkBook, WorkSheet, 'Результат', row2, col2);

      info('Вставка данных...');
      VarClear(data2);//создаем колонку этажность
      data2 := VarArrayCreate([0, row2-row, 1, 1], varOleStr);
      k := row;//начало КН
      for i := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
      begin
        kn := data[i, 1];
        s := data[i, 7];
        if s <> '' then
        for j := k to row2 do
          if String(DataParcels[j, col]) = kn then
          begin
            data2[j-row, 1] := s;
            Inc(k);
            break;
          end;
      end;

      for i := 1 to floorcount*2 do//удаление колонок "номер этажа +	тип этажа"
        WorkSheet.Columns[floor].Delete(xlToLeft);

      for i := 1 to col2 do//ищем Этажность
        if SameStr(String(WorkSheet.Cells[row-1, i]), 'Этажность') then
        begin
          s := ExcelNum2Str(WorkSheet, i);
          range := WorkSheet.range[format('%s%d:%s%d', [s, row, s, row2])];
          range.Value := data2;//вставляем этажность
          break;
        end;

      VarClear(DataParcels);
      FindCadNum(WorkSheet, row, col, row2, col2);
      range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, col2), row2])];
      DataParcels := range.value;//новые размеры помещений
    end;
  end;

begin//mnFlats_BuildingsClick
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER4\test\TelDa_XLS'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER4\Win32\Debug\RESULT\TelDa'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER4\test\TelDa_XLS\NEW'; {$ENDIF}

  pattern := (Sender as TMenuitem).Hint;
  if not FileExists(GetAppFolder + pattern) then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + pattern);
  {
  pattern=PATTERNS\TelDa\Шаблон_Помещения_Сироты.xlsx
  pattern=PATTERNS\TelDa\Шаблон_Здания_Помещения.xlsx
  }
  ResFolder := GetAppFolder + 'RESULT\TelDa\' + AddSlesh(CurrentData);
  if not ForceDirectories(ResFolder) then
    Exc('Не удалось создать каталог' + #10 + ResFolder);
  filename := ResFolder + 'Здания_'+ (Sender as TMenuitem).Caption  + '.xlsm';

  if FileExists(filename) then
    if not DeleteFile(filename) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [filename, #10]));

  OpenDialog1.FileName := '*.xls*';
  try
    OpenDialog1.Title := 'Открытие файла Зданий';
    if not OpenDialog1.Execute then
      exit;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    buildings := OpenDialog1.FileName;

    OpenDialog1.Title := 'Открытие файла Помещений';
    if not OpenDialog1.Execute then
      exit;
    flats := OpenDialog1.FileName;

    if not LoadFlats then exit;

    info('Загрузка Шаблона...');
    sl := TStringList.Create;//заносим КН зданий в список
    log := TStringList.Create;
    floormnames := TStringList.Create;
    floormnames.Sorted := True;
    ExcelApp := CreateOleObject('Excel.Application');
    try
      source := GetAppFolder + pattern;
      temppath := ExtractFilePath(miEval('TempFileName$("")'));
      ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"',
        [source, temppath])), SW_HIDE, true, 0);
      s := temppath + ExtractFileName(source);
      temppath := ChangeFileExt(s, '_.xlsx');
      RenameFile(s, temppath);//новое название должно отличаться от шаблона
      ExcelApp.Workbooks.Open(temppath);//открываем копию
      WorkBook := ExcelApp.Workbooks.item[1];

      //проверим колонку "номер этажа" на наличие значений на листе "Спр наименований этажей"
      if FindSheetByName(WorkBook, WorkSheet, 'Спр наименований этажей', row2, col2) then
      begin//Заносим в список "Спр наименований этажей"
        VarClear(datafloor);
        range := WorkSheet.range[format('A2:B%d', [row2])];
        datafloor := range.value;
        for i := VarArrayLowBound(datafloor,1) to VarArrayHighBound(datafloor,1) do
          floormnames.AddObject(String(datafloor[i,1]), Pointer(Integer( datafloor[i,2])));
        //проверяем этажность данных с наличием в справочнике
        n := row2;
        While String(WorkSheet.Cells[n, 1]) = '' do//если есть пустые то надо их удалить
          Dec(n);
        rb2 := n;
        k := floor;
        for i := 1 to floorcount do
        begin
          for j := row to rd do//по столбцам
          begin
            s := DataParcels[j, k];
            if (s <> '') and (not floormnames.Find(s, m)) then
            begin
              s2 := '0';
              if not InputQuery('В справочнике этажей отсутствует это наименование',
                '"'+s+'" укажите тип', s2) then
                  exit;//Exc('Выбор прерван');
              if not ((s2 = '0') or (s2 = '1')) then
                Exc('Ошибка в задании типа');
              floormnames.Add(s);
              Inc(n);//добавлям запись в справочник
              WorkSheet.Cells[n, 1] := s;
              WorkSheet.Cells[n, 2] := s2;
            end;
          end;
          Inc(k, 2);
        end;
        if rb2 < n then//были добавления, обновляем исходный справочник
        begin
          Info('Обновление справочника наименований этажей...');
          WorkSheet.Range[format('A1:B%d', [n])].Copy;
          ExcelApp.Workbooks.Open(source);//открываем исходный шаблон
          WorkBook := ExcelApp.Workbooks.item[2];
          FindSheetByName(WorkBook, WorkSheet, 'Спр наименований этажей', i, j);
          WorkSheet.Range['A1'].Select;
          WorkSheet.Paste;
          WorkBook.Save;
          WorkBook.Close;
          WorkBook := ExcelApp.Workbooks.item[1];
        end;
      end else
        Note('Лист "Спр наименований этажей" не найден!' + #10 + 'Проверка этажности невозможна.');

      if FindSheetByName(WorkBook, WorkSheet, 'Результат', row2, col2) then
      begin
        //добавим нужное количество колонок к помещениям
        Info('Вставляем дополнительные столбцы...');
        s := ExcelNum2Str(WorkSheet, floor);//первый столбец
        s2 := ExcelNum2Str(WorkSheet, floor+1);//второй столбец
        range := WorkSheet.Range[format('%s:%s', [s, s2])];
        range.Select;
        for i := 2 to floorcount do//вставляем дополнительные столбцы
          ExcelApp.Selection.Insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove);
        //вставляем целиким все данные помещений в шаблон колонки "номер этажа +	тип этажа" удалить
        range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, cd), rd])];
        range.value := DataParcels;

        CalcFloor;
        if log.Count > 0 then exit;//tr!!!

        info('Поиск в таблице Здания...');
        if not LoadBuildings then exit;//загружаем родителей (зданя)

        //выбираем из зданий КН зданий для быстрого поиска
        for j := rb to rb2 do
          sl.AddObject(data[j, cb], Pointer(j));
        sl.Sort;
        //заполним КН зданий из помещенийдля для связи из колонки "кадастровый номер здания"
        VarClear(cn);
        cn := VarArrayCreate([0, row2-row, 1, 1], varOleStr);//rt!!!
        for j := 0 to VarArrayHighBound(cn, 1) do
          cn[j, 1] := Trim(DataParcels[j+row, col+1]);//начало = 0

        for i := 1 to col2 do//колонки шаблона
        begin
          if SameText(DataParcels[row-2, i], 'buildings') then
            for j := cb to cb2 do//колонки зданий
            begin
              if SameText(DataParcels[row-1, i], data[rb-1, j]) then
              begin
                VarClear(data2);//создаем и заполняем колонку по размеру КН
                data2 := VarArrayCreate([0, VarArrayHighBound(cn, 1), 1, 1], varOleStr);
                for k := 0 to VarArrayHighBound(data2, 1) do
                try
                  if sl.find(cn[k, 1], n) then//ищем КН в зданиях
                  begin
                    n := Integer(sl.Objects[n]);
                    data2[k, 1] := data[n, j] ;
                  end;
                except
                  data2[k, 1] := 'Error';
                end;
                s := ExcelNum2Str(WorkSheet, i);
                range := WorkSheet.range[format('%s%d:%s%d',[s, row, s, VarArrayHighBound(cn, 1)+row])];
                range.value := data2;
              end;
            end;
        end;

        info('Копирование формул...');
        FindSheetByName(WorkBook, WorkSheet, 'Результат', row2, col2);
        FindCadNum(WorkSheet, row, col, row2, col2);
        for k := 1 to col2 do
        begin
          if WorkSheet.Cells[row, k].HasFormula then//копируем формулы в память
          begin
            s := ExcelNum2Str(WorkSheet, k);
            s2 := ExcelNum2Str(WorkSheet, col2);//конец фрпмул
            WorkSheet.Range[format('%s%d:%s%d', [s, row, s2, row])].Copy;
            WorkSheet.Range[format('%s%d:%s%d', [s, row+1, s2, row2])].PasteSpecial(operation:=xlNone);
            break;
          end;
        end;

      end;

      Setfilter(WorkSheet, ExcelApp, row-1);
      SetSheetParams(WorkBook, WorkSheet, ExcelApp, '', [20, 20, 15], row-1);
      WorkBook.SaveAs(Filename:=filename, FileFormat:= 52, CreateBackup:=False);

      info('Готово...');
    finally
      sl.Free;
      filename := ResFolder + 'Здания_'+ (Sender as TMenuitem).Caption  + '.log';
      DeleteFile(filename);

      if floorcount = 0 then
      begin
        ExcelApp.Workbooks.Close;
        ExcelApp.quit;
      end else

      if log.Count > 0 then//tr!!!
      begin
        info('В исходных данных найдены ошибки');
        log.SaveToFile(filename);
        ExShellExecute('Notepad.exe', filename, SW_NORMAL, false, 0);
      end else
        ExcelApp.Visible := true;

      floormnames.Free;
      log.Free;
      VarClear(datafloor);
      VarClear(DataParcels);
      VarClear(data);
      VarClear(data2);
      VarClear(cn);
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  finally
    OpenDialog1.Title := '';
  end;
end;

procedure TValuationForm.mnFlats_OrphanClick(Sender: TObject);
begin
  mnFlats_BuildingsClick(Sender);
end;

procedure TValuationForm.mnCarParkingSpacesClick(Sender: TObject);
var
  car, filename, pattern, temppath, source, s, s2 : String;
  row, col, row2, col2, k : Integer;
  ExcelApp, Workbook, WorkSheet, range, DataParcels : OleVariant;

  function LoadCars : Boolean;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
  begin
    Result := false;
    info('Загрузка ' + ExtractFileName(car));
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Workbooks.Open(car);
      WorkBook := ExcelApp.Workbooks.item[1];
      if not FindSheetByName(WorkBook, WorkSheet, 'Объекты') then
         Exc('лист Объекты в файле не найден!');
      FindCadNum(WorkSheet, row, col, row2, col2);//блок данных
      range := WorkSheet.range[format('%s%d:%s%d',
        [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, col2), row2])];
      DataParcels := range.value;
      Result := true;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  end;

begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\машиноместа'; {$ENDIF}

  pattern := (Sender as TMenuitem).Hint;
  if not FileExists(GetAppFolder + pattern) then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + pattern);

  OpenDialog1.FileName := '*.xls*';
  try
    OpenDialog1.Title := 'Открытие файла Машиномест';
    if not OpenDialog1.Execute then
      exit;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    car := OpenDialog1.FileName;

    ResFolder := GetAppFolder + 'RESULT\TelDa\' + AddSlesh(CurrentData);
    if not ForceDirectories(ResFolder) then
      Exc('Не удалось создать каталог' + #10 + ResFolder);
    filename := ResFolder + 'Машиноместа_' + ExtractFileName(car);

    if FileExists(filename) then
      if not DeleteFile(filename) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [filename, #10]));

    if not LoadCars then exit;

    info('Загрузка Шаблона...');
    ExcelApp := CreateOleObject('Excel.Application');
    try
      source := GetAppFolder + pattern;
      temppath := ExtractFilePath(miEval('TempFileName$("")'));
      ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"',
        [source, temppath])), SW_HIDE, true, 0);
      s := temppath + ExtractFileName(source);
      temppath := ChangeFileExt(s, '_.xlsx');
      RenameFile(s, temppath);//новое название должно отличаться от шаблона
      ExcelApp.Workbooks.Open(temppath);//открываем копию
      WorkBook := ExcelApp.Workbooks.item[1];

      FindSheetByName(WorkBook, WorkSheet, 'Объекты');
      range := WorkSheet.range[format('%s%d:%s%d',
        [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, col2), row2])];
      range.value := DataParcels;

      info('Копирование формул...');
      col := col2 + 1;//начало формул
      col2 := WorkSheet.UsedRange.Columns.Count;
      for k := col to col2 do
      begin
        if WorkSheet.Cells[row, k].HasFormula then//копируем формулы в память
        begin
          s := ExcelNum2Str(WorkSheet, k);
          s2 := ExcelNum2Str(WorkSheet, col2);//конец фрпмул
          WorkSheet.Range[format('%s%d:%s%d', [s, row, s2, row])].Copy;
          WorkSheet.Range[format('%s%d:%s%d', [s, row+1, s2, row2])].PasteSpecial(operation:=xlNone);
          break;
        end;
      end;
      Setfilter(WorkSheet, ExcelApp, row-1);
      SetSheetParams(WorkBook, WorkSheet, ExcelApp, '', [10, 20, 15], row-1);
      WorkBook.SaveAs(Filename:=filename, FileFormat:= 51, CreateBackup:=False);
      ExcelApp.Visible := true;
      info('Готово...');
    finally
      VarClear(DataParcels);
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;
  finally
    OpenDialog1.Title := '';
  end;
end;

function TValuationForm.OnCreateDecompressStream(const InStream: TStream;
  const ZipFile: TZipFile; const Item: TZipHeader;
  IsEncrypted: Boolean): TStream;
begin
  try
    if IsEncrypted then
      ShowMessage('Зашифровано')
    else
      Result := InStream;
  except
    on E: Exception do
      Result := InStream;
  end;
end;

procedure TValuationForm.OpenDialog1Close(Sender: TObject);
begin
  Self.SetFocus;
end;

procedure TValuationForm.ExcelCorrectClick(Sender: TObject);
var
  ZipFile : TZipFile;
  temppath, s, s2 : String;
  mc : TMatchCollection;
  sl, xml : TStringList;
  i, j : Integer;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := GetAppFolder + 'PATTERNS\TelDa\'; {$ENDIF}
  OpenDialog1.FileName := '*.xls*';
  OpenDialog1.Title := 'Открытие файла Зданий';
  if not OpenDialog1.Execute then
    exit;
  temppath := ExtractFilePath(miEval('TempFileName$("")')) + 'temp\';
  ForceDirectories(temppath);
  sl := TStringList.Create;
  xml := TStringList.Create;
  ZipFile := TZipFile.Create;
  ZipFile.OnCreateDecompressStream := OnCreateDecompressStream;
  ZipFile.Open(OpenDialog1.FileName, zmReadWrite);
  try
    ZipFile.ExtractAll(temppath);
    FindRecursive(temppath, '*.xml', true, sl);
    for i := 0 to sl.Count-1 do
    begin
      if Pos('workbook.xml', sl[i]) > 0 then
      begin
        xml.LoadFromFile(sl[i]);
        s := xml.Text;
        s := StringReplace(s, '</definedName>', '</definedName>'+#10, [rfReplaceAll]);
        s2 := '<definedName.+_FilterDatabase.+</definedName>';
        mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
        if mc.count > 1 then
        begin
          for j := mc.count-1 downto 1 do
            s := StringReplace(s, mc.Item[j].Value, '', [rfReplaceAll]);
          xml.Text := s;
          xml.SaveToFile(sl[i]);
          ZipFile.Close;
          s := ExtractFilePath(miEval('TempFileName$("")')) + ExtractFileName(OpenDialog1.FileName);
          ZipFile.ZipDirectoryContents(s, temppath, zcDeflate);
          ZipFile.Close;
          ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',
            [s, OpenDialog1.FileName]), SW_HIDE, true, 0);
          Note('Удалено: ' + mc.Item[1].Value);
          exit;
        end;
      end;
    end;
    Note('Нет изменений');
  finally
    sl.Free;
    xml.Free;
    ZipFile.Free;
  end;
end;

procedure TValuationForm.mnParamsClick(Sender: TObject);
begin
  if mnVRIFromBase.Enabled then
    ParamForm.Show
  else
    ShowMessage('Перед вызовом правил необходимо выбрать справочник.');
end;

procedure TValuationForm.LoadRules(var list1, list2 : TStringList);
var
  i : Integer;
  s, ss, id : AnsiString;
  p : PAnsiChar;
  ExcelApp, Workbook, WorkSheet : OleVariant;
  data, range : OleVariant;

  procedure AddToList(list : TStringList);
  var
    i : Integer;
  begin
    row := WorkSheet.UsedRange.Rows.Count;
    range := WorkSheet.Range[format('A2:D%d', [row])];// Область UndegraundLF1
    data := VarArrayCreate([1, row, 1, 4], varOleStr);
    data := Range.Value;
    for i := VarArrayLowBound(data,1) to VarArrayHighBound(data,1) do
    begin
      id := VarArrayGet(data,[i,1]);//ID
      s := VarArrayGet(data,[i,2]);//ВРИ по документу
      ss := VarArrayGet(data,[i,3]);//код VRI для правила0 / Имя правила для других
      if ss = '' then break;
      ss := ss + '-' + id + '-' + VarArrayGet(data,[i,4]);//для правила7 добавим свой код VRI
      Getmem(p, Length(ss) + 1);
      StrCopy(p, PAnsiChar(ss));
      list.AddObject(Trim(s), @p^);
    end;
    VarClear(data);
    range := Unassigned;
  end;

begin
  list1 := TStringList.Create;//список для Правило0
  list2 := TStringList.Create;//список для Правило1..7
  Info('Открытие таблицы правил');
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Workbooks.Open(GetAppFolder + mnRulesName.Caption + '.xlsx');
  WorkBook := ExcelApp.Workbooks.item[1];

  try
    Info('Загрузка Правила_0'); //ID, название по документу(UtilizationByDoc) и Код ВРИ
    WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
    AddToList(list1);

    Info('Загрузка Правил...');
    WorkSheet := WorkBook.WorkSheets.item[2];//активный 2 лист
    AddToList(list2);
  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.quit;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;

    list1.Sort;
    list2.Sort;
  end;
end;

procedure TValuationForm.FreeRules(var list1, list2 : TStringList);
var
  i : Integer;
  p : PAnsiChar;
begin
  for i := 0 to list1.Count-1 do
  begin
    p := PAnsiChar(list1.Objects[i]);
    Freemem(p);
  end;

  for i := 0 to list2.Count-1 do
  begin
    p := PAnsiChar(list2.Objects[i]);
    Freemem(p);
  end;

  FreeAndNil(list1);
  FreeAndNil(list2)
end;

procedure TValuationForm.CalcVri;
var
  i, j, area : Integer;
  listVRI, list1, list2, listRepeat, listZU : TStringList;
  s, cad, vri, category, bydoc, rule, id : String;
  sa : StringArray;
begin
  if ds.IsEmpty then
    Exc('Нет выбранных записей');
  BreakCommand := false;
  pg.Max := ds.RecordCount;
  listVRI := TStringList.Create;//список для VRI
  listVRI.Add('ID' + #9 + 'CADASTRALNUMBER' + #9 + 'VRICode_NEW' + #9 + 'PraviloVRI');

  listRepeat := TStringList.Create;//список для повторений
  listRepeat.Duplicates := dupIgnore;
  listRepeat.Sorted := True;
  listZU := TStringList.Create;//Список ЗУ
  listZU.Add('список по ЗУ КН' + #9 + 'Код ВРИ по правилу' + #9 + 'Правило' + #9 + 'Номер правила');

  LoadRules(list1, list2);

  Info('Проверка правил...');
  try
    pg.Position := 0;
    while not ds.eof do
    begin
      Application.ProcessMessages;
      if BreakCommand then exit;

      cad := ds.FieldByName('CadastralNumber').AsString;
      bydoc := Trim(StringReplace(ds.FieldByName('UtilizationByDoc').AsString, lf, ' ', [rfReplaceAll]));
      category := ds.FieldByName('ParcelCategory').AsString;
      area := Round(ds.FieldByName('Area').AsFloat);
      vri := '*';
      rule := '-';
      id := '';

      if list1.Find(bydoc, i) then
      begin
        rule := 'Правило0';
        vri := PAnsiChar(list1.Objects[i]);
        sa := SplitStr(vri, '-');
        vri := sa[0];
        id := sa[1];
        SetLength(sa, 0);
      end else

      if list2.Find(bydoc, i) then
      begin
        rule := PAnsiChar(list2.Objects[i]);
        sa := SplitStr(rule, '-');
        rule := sa[0];
        id := sa[1];
        SetLength(sa, 0);
        if CText(rule, 'Правило6') then
          vri := 'этажность'
        else
          vri := ParamForm.CalcVri(Trim(Copy(rule, 8, 10)), category, area);
        if vri = '*' then
          rule := '-';
      end;

      if rule <> '-' then
      begin
        s := id + #9 + rule;
        listRepeat.Add(s);
        if listRepeat.Find(s, i) then
        begin
          j := Integer(listRepeat.Objects[i]);//число повторений
          Inc(j);//увеличиваем на 1
          listRepeat.Objects[i] := Pointer(j);
        end;
      end;

      if vri <> '*' then
      begin
        listZU.Add(cad + #9 + vri + #9 + rule + #9 + id);
        listVRI.Add(id + #9 + cad + #9 + vri + #9 + rule);
      end;

      ds.next;
      pg.Position := pg.Position + 1;
   end;
  finally
    pg.Max := 0;
    Info('');
    listVRI.SaveToFile(path+'VRI.txt');
    listVRI.Free;

    FreeRules(list1, list2);

    listZU.SaveToFile(path+'Список ЗУ.txt');
    listZU.Free;

    listRepeat.Sorted := false;
    for i := 0 to listRepeat.Count-1 do
      listRepeat[i] := listRepeat[i] + #9 + IntToStr(Integer(listRepeat.Objects[i]));
    listRepeat.SaveToFile(path+'Повторения.txt');
    listRepeat.Free;

    ExShellExecute('Notepad.exe', path+'VRI.txt', SW_NORMAL, false, 0);
    ExShellExecute('Notepad.exe', path+'Список ЗУ.txt', SW_NORMAL, false, 0);
    ExShellExecute('Notepad.exe', path+'Повторения.txt', SW_NORMAL, false, 0);

    if BreakCommand then
      Exc('Обработка прервана')
    else
      Note('Обработка завершена');
  end;
end;

procedure TValuationForm.CalcVriXlsx(var exceldata : OleVariant);
var
  i, j, k, col, row, area : Integer;
  d : Double;
  listVRI, list1, list2 : TStringList;
  s, ss, cad, vri, filevri, category, bydoc, landuse, util, utiltxt, rule, id, segmentOfUsage : String;
  ExcelApp, Workbook, WorkSheet, DataParcels, range : OleVariant;
  sa : StringArray;
  ini: TMemIniFile;

  procedure FindVRI(utilizationVRI : String);
  begin
    k := 0;
    vri := '*';
    rule := '-';
    id := '';

    Inc(k);
    if list1.Find(utilizationVRI, i) then
    begin
      rule := 'Правило0';
      vri := PAnsiChar(list1.Objects[i]);
      sa := SplitStr(vri, '-');
      vri := sa[0];
      id := sa[1];
      SetLength(sa, 0);
    end else

    if list2.Find(utilizationVRI, i) then
    begin
      rule := PAnsiChar(list2.Objects[i]);
      sa := SplitStr(rule, '-');
      rule := sa[0];
      id := sa[1];
      SetLength(sa, 0);
      if CText(rule, 'Правило6') then
        vri := 'этажность'
      else
        vri := ParamForm.CalcVri(Trim(Copy(rule, 8, 10)), category, area);
    end;

    if vri <> '*' then
      listVRI.Add(cad + '~' + id + '~' + vri + '~' + rule + '~' + category +
        '~' + bydoc + '~' + IntToStr(area) + '~' + segmentOfUsage)
    else
      listVRI.Add(cad + '~' + '~~~' + category + '~' + bydoc + '~' + IntToStr(area) + '~' + segmentOfUsage);
  end;

begin//CalcVriXlsx
  if VarArrayHighBound(exceldata,1) = 0 then
    Exc('Нет выбранных записей');
  BreakCommand := false;
  pg.Max := VarArrayHighBound(exceldata,1);
  listVRI := TStringList.Create;//Список VRI
  LoadRules(list1, list2);

  CodeCatZem := TStringList.Create;//загружаем справочник категорий земель
  ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
  try
   ini.ReadSectionValues('CodeCatZem', CodeCatZem);
  finally
    ini.Free;
  end;

  Info('Проверка правил...');
  try
    for j := VarArrayLowBound(exceldata,1)+1 to VarArrayHighBound(exceldata,1) do
    begin
      Application.ProcessMessages;
      if BreakCommand then exit;
      cad := exceldata[j, 1];// ds.FieldByName('CadastralNumber').AsString;
      bydoc :=Trim(StringReplace(exceldata[j, 4], lf, ' ', [rfReplaceAll]));//Trim(StringReplace(ds.FieldByName('UtilizationByDoc').AsString, lf, ' ', [rfReplaceAll]));
      landuse := exceldata[j, 5];//ds.FieldByName('UtilizationLandUse').AsString;
      util := exceldata[j, 6];//ds.FieldByName('Utilization').AsString;
      utiltxt := exceldata[j, 7];//ds.FieldByName('UtilizationTxt').AsString;

      category := exceldata[j, 3];//ds.FieldByName('ParcelCategory').AsString;
      category := CodeCatZem.Values[category];

      Val(exceldata[j, 2], d, i);
      area := Round(d);//area := Round(ds.FieldByName('Area').AsFloat);
      segmentOfUsage := exceldata[j, 8];//ds.FieldByName('segmentOfUsage').AsString;
      FindVRI(bydoc);
      pg.Position := pg.Position + 1;
    end;

    filevri := path + 'VRI_' + CurrentData + '.xlsx';
    ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем шаблон
      [GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ.xlsx', path]), SW_HIDE, true, 0);
    DeleteFile(vri);//удаляем старый файл
    RenameFile(path + 'Шаблон для результатов по опр КВРИ.xlsx', filevri);//переименовываем шаблон

    DataParcels := VarArrayCreate([0, listVRI.Count, 0, VarArrayHighBound(exceldata,2)], varOleStr);
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Application.EnableEvents := false;
    ExcelApp.DisplayAlerts := False;
    try
      ExcelApp.Workbooks.Open(filevri);
      WorkBook := ExcelApp.Workbooks.item[1];
      WorkSheet := WorkBook.WorkSheets.item[1];
      for i := 0 to listVRI.Count-1 do
      begin
        sa := SplitStr(listVRI[i], '~');
        for j := 0 to High(sa) do
          DataParcels[i,j] := sa[j];
        SetLength(sa, 0);
      end;
      col := WorkSheet.UsedRange.Columns.Count - 2;//последние 2 колонки - формулы
      range := WorkSheet.Range[format('A2:%s%d',
        [ExcelNum2Str(WorkSheet, col), listVRI.Count+1])];// Область вставки A2:G252
      range.Value := DataParcels;

      s := ExcelNum2Str(WorkSheet, col+1);
      ss := ExcelNum2Str(WorkSheet, col+2);//конец фрпмул
      WorkSheet.Range[format('%s2:%s2', [s, ss])].Copy;
      WorkSheet.Range[format('%s3:%s%d', [s, ss, listVRI.Count+1])].PasteSpecial(operation:=xlNone);

      WorkBook.Save;
      ExcelApp.Visible := true;
    finally
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
///////////////////////////
      Info('Создание шаблона автозапоолнения');
      filevri := path + 'ШАБЛОН_Поля для автозаполнения_Новые' + CurrentData + '.xlsx';
      ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',
        [GetAppFolder + 'PATTERNS\TelDa\ШАБЛОН_Поля для автозаполнения_Новые.xlsx', path]), SW_HIDE, true, 0);
      DeleteFile(filevri);
      RenameFile(path + 'ШАБЛОН_Поля для автозаполнения_Новые.xlsx', filevri);
      ExcelApp := CreateOleObject('Excel.Application');
      ExcelApp.Application.EnableEvents := false;
      ExcelApp.DisplayAlerts := False;
      try
        ExcelApp.Workbooks.Open(filevri);
        WorkBook := ExcelApp.Workbooks.item[1];
        WorkSheet := WorkBook.WorkSheets.item[1];
        range := WorkSheet.Range[format('B4:B%d', [listVRI.Count+3])];// Область вставки
        range.Value := DataParcels;//вставка кадастровых номеров
        col := WorkSheet.UsedRange.Columns.Count;//неизменная часть
        row := WorkSheet.UsedRange.Rows.Count;

        s := ExcelNum2Str(WorkSheet, col);
        WorkSheet.Range[format('C4:%s4', [s])].Copy;
        WorkSheet.Range[format('C5:%s%d', [s, row])].PasteSpecial(operation:=xlNone);

        WorkBook.Save;
        ExcelApp.Visible := true;
      finally
        VarClear(DataParcels);
        WorkSheet := Unassigned;
        WorkBook := Unassigned;
        ExcelApp := Unassigned;
      end;
/////////////////////////
    end;
  finally
    pg.Max := 0;
    Info('');
    listVRI.Free;
    FreeRules(list1, list2);
    CodeCatZem.Free;
  end;
end;

procedure TValuationForm.mnCadNumberClick(Sender: TObject);
var
  s, fields : String;
  sl : TStringList;
  i : Integer;
begin
  if Clipboard.AsText = '' then
    Exc('Буфер обмена пуст');
  if not reg.IsMatch(Clipboard.AsText, '66\:\d\d:\d\d\d\d\d\d\d:\d{1,}') then
    Exc('Буфер обмена не содержит кадастровых номеров');
  if lbTables.ItemIndex = -1 then
    Exc('Не выбрано имя базы');
  if lbFields.SelCount < 1 then
    Exc('Нет выбранных полей таблицы ' + TablesList[lbTables.ItemIndex]);

  sl := TStringList.Create;
  try
    sl.Text := Clipboard.AsText;
    s := '';
    for i := 0 to sl.Count-1 do
    begin
      s := s + 'CADASTRALNUMBER=' +#39 + sl[i] + #39;
      if i < sl.Count-1 then
        s := s + ' OR ';
    end;

    fields := '';
    for i  := 0 to lbFields.Count-1 do
    if lbFields.Selected[i] then
        fields := fields + FieldsList[i] + ',';
    SetLength(fields, Length(fields) - 1);

    CommandString.Text := format('select %s from %s WHERE %s',
      [fields, TablesList[lbTables.ItemIndex], s]);
  finally
    sl.Free;
    BreakCommand := false;
    ImportDataTab(true);
  end;
  ds.Close;
end;

procedure TValuationForm.mnCheckFloorsClick(Sender: TObject);
var
  sl, spr : TStringList;
  i, id, row, k : Integer;
  sprname, s : String;
  DataParcels : OleVariant;
  ExcelApp, Workbook, WorkSheet : OleVariant;
begin
  sl := TStringList.Create;
  FindRecursive(GetAppFolder + 'PATTERNS\Здания\', '*.xls*', False, sl);
  sprname := '';
  for i := 0 to sl.Count-1 do
    if (Pos('этажность', sl[i]) > 0) and (Pos('~$', sl[i]) = 0) then
    begin
      sprname := sl[i];
      break;
    end;
  sl.Free;

  if sprname <> '' then
  begin
    ds.Close;
      ds.CommandText := 'select FLOORS from buildings WHERE (Status_KC=''Не готов'' AND'+
        ' (FlagLocFloors=''Не готов'' OR FlagLocFloors is Null))';
    ds.Open;

    sl := TStringList.Create;
    sl.Sorted := true;
    sl.Duplicates := dupIgnore;
    spr := TStringList.Create;
    spr.Sorted := true;
    spr.Duplicates := dupIgnore;

    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Application.EnableEvents := false;
    ExcelApp.DisplayAlerts := False;//не показывать предупреждающие сообщения
    Workbook := ExcelApp.WorkBooks.Open(sprname);//Открываем рабочую книгу
    WorkSheet := WorkBook.WorkSheets.Item[3];

    try
      Info('Загрузка справочника этажности...');
      row := WorkSheet.UsedRange.Rows.Count;
      DataParcels := VarArrayCreate([1, row], varOleStr);
      try//заполняем справочник
        DataParcels := WorkSheet.UsedRange.Columns[1].Value;
        for i := VarArrayLowBound(DataParcels,1)+1 to VarArrayHighBound(DataParcels,1) do
          if DataParcels[i,1] <> '' then
            spr.Add(String(DataParcels[i,1]));
      finally
        VarClear(DataParcels);
      end;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
      ExcelApp := Unassigned;
    end;

    try
      Info('Загрузка данных...');
      pg.Max := ds.RecordCount;
      while not ds.eof do//выбираем данные и проверяем наличие их в справочнике
      begin
        pg.Position := pg.Position+1;
        s := ds.Fields[0].AsString;
        if not spr.Find(s, k) then
          sl.Add(s);
        ds.next;
      end;
      pg.Max := 0;

      if sl.Count > 0 then
      begin
        Info(format('Новых данных %d', [sl.Count]));
        ComplianceForm := TComplianceForm.Create(nil);
        ComplianceForm.SprName := sprname;
        ComplianceForm.LoadSpr;
        id := ComplianceForm.sg.RowCount;
        ComplianceForm.sg.RowCount := ComplianceForm.sg.RowCount + sl.Count;
        for i := 0 to sl.Count-1 do
          ComplianceForm.sg.Cells[0, id + i] := sl[i];
        ComplianceForm.ShowModal;
        ComplianceForm.Free;
      end else
        ShowMessage('Нет новых данных');
    finally
      sl.Free;
      spr.Free;
      Info('');
    end;
  end else
    ShowMessage('Не найден файл по этажности.');
end;

//кадастровые номера берутся из файла
procedure TValuationForm.mnCodeFromFileClick(Sender: TObject);
var
  sl : TStringList;
  s, sql, fields, command : String;
  i, j : Integer;
begin
  if mnObjects.Hint = '' then
    Exc('Нет выбранных объектов');
  //проверка открытых файлов EXCEL
  for i := 0 to mnObjects.Count-1 do
  begin
    s := mnObjects.Items[i].Caption;
    s := format('%s%s', [ResFolder, s]);
    s := StringReplace(s, ' ', '_', [rfReplaceAll]) + '_загрузка.xlsm';
    if FileExists(s) then//проверяем наличие открыього файла
      if not DeleteFile(s) then
        Exc(format('Файл %s открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));
  end;

  OpenDialog1.FileName := '*.txt';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  s := format(' /c del /s /f /q  "%s"', [tmp]);
  ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//чистим каталог

  BreakCommand := false;
  try
    for i := 0 to mnObjects.Count-1 do
    begin
      Application.ProcessMessages;
      if BreakCommand then exit;

      if mnObjects.Items[i].Caption <> mnObjects.Hint then
        continue;
//      mnObjects.Hint := mnObjects.Items[i].Caption;
      edSize.Text := '5000';
      Info(mnObjects.Items[i].Caption + ': выбор из базы');
      if SetCommandText then//если есть необходимые для выборки поля
      begin
        try
//Временные таблицы SQL Server
//https://www.sqlservertutorial.net/sql-server-basics/sql-server-temporary-tables/ //tr!!!!
          cmd.Execute('DROP TABLE TemporalCadNums');
          cmd.Execute('CREATE TABLE TemporalCadNums(cadnum varchar(35))');
        except
          cmd.CommandText := 'DELETE FROM TemporalCadNums';
          cmd.Execute;
        end;

        sl := TStringList.Create;
        sl.LoadFromFile(OpenDialog1.FileName);
        if (sl.Count > 0) and (not reg.IsMatch(sl[0],'66\:\d\d:\d\d\d\d\d\d\d:\d{1,}')) then
          sl.Delete(0);//заголовок удаляем

        if (sl.Count > 0) and (reg.IsMatch(sl[0],'66\:\d\d:\d\d\d\d\d\d\d:\d{1,}')) then
        try
          cmd.Parameters.Clear();
          cmd.CommandType := cmdText;
          sql := 'Insert into TemporalCadNums (cadnum) values(:code)';
          cmd.Commandtext := sql;
          cmd.Parameters.ParseSQL(sql, True);
          cmd.Parameters.ParamByName('code').DataType := ftString;
          ADOConnection1.BeginTrans;
          pg.Max := sl.Count;
          for j := 0 to sl.Count-1 do
          begin
            cmd.Parameters[0].Value := sl[j];
            cmd.Execute;
          end;

          pg.Max := 0;
          fields := '';//создаем список полей для запроса
          for j := Low(farray) to High(farray)-1 do
            fields := fields + farray[j] +',';
          fields := fields + farray[High(farray)];
          command := format('select %s from %s ', [fields, basename_]);
          ds.Close;
          ds.CommandText := command + 'Where CADASTRALNUMBER=any(Select cadnum from TemporalCadNums)';
          ds.Open;

          if not ds.IsEmpty then//если кадастровые номера имеются в данном объекте
            if ImportData then
              Building(1, false);

        finally
          sl.Free;
          ADOConnection1.RollbackTrans;
          cmd.Execute('DROP TABLE TemporalCadNums');
        end;
      end;
    end;
  finally
    Info('');
    if BreakCommand then
      ShowMessage('Обработка прервана')
    else
      ShowMessage('Обработка завершена');
  end;
end;

procedure TValuationForm.mnCodeFromExcelClick(Sender: TObject);
var
  tab, fields, oldfolder, s : String;
  sl : TStringList;
  i, j, k, row, col, oldtitleiindex : Integer;
  ExcelApp, Workbook, WorkSheet, data, range : OleVariant;

  procedure UpdateExcel;
  begin
    ExcelApp.Workbooks.Close;
    ExcelApp.quit;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;

    ExcelApp.Workbooks.Open(path + basename + '.xlsx');
    WorkBook := ExcelApp.Workbooks.item[1];
    WorkBook.WorkSheets.item[1].Activate;//обязательно активировать иначе не сработает метод range.select!!!
    WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист

    range := WorkSheet.Range[format('B4:E%d', [j+4])];//Область вставки
    range.Value := data;
    Workbook.Save;
  end;

begin
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Объекты.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Объекты.xlsx');
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Классификатор кодов населенных пунктов.TAB') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Классификатор кодов населенных пунктов.TAB');

  OpenDialog1.FileName := '*.xls*';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  fields :='(CADASTRALNUMBER Char(40),CADASTRALBLOCK Char(13),DISTRICTNAME Char(90),'+
    'DISTRICTTYPE Char(5),CITYNAME Char(90),CITYTYPE Char(5),LOCALITYNAME Char(90),'+
    'LOCALITYTYPE Char(5),LOCALITYCODE Char(90),SOURCELOCALITYCODE Integer,FLAGLOCALITYCODE Char(50))';

  sl := TStringList.Create;
  sl.Assign(OpenDialog1.Files);
  oldfolder := ResFolder;
  oldtitleiindex := ParamForm.rgTitle.ItemIndex;//отключим перевод на русский заголовков столбцов
  ParamForm.rgTitle.ItemIndex := 1;

  ResFolder := ExtractFilePath(sl[0]) + CurrentData + '\';
  if not ForceDirectories(ResFolder) then
    Exc('Не удалось создать каталог' + #10 + path);
  ExcelApp := CreateOleObject('Excel.Application');

  for k := 0 to sl.Count-1 do
  begin
    ExcelApp.Workbooks.Open(sl[k]);
    WorkBook := ExcelApp.Workbooks.item[1];
    WorkBook.WorkSheets.item[1].Activate;//обязательно активировать иначе не сработает метод range.select!!!
    WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
    try
      col := WorkSheet.UsedRange.Columns.Count;
      row := WorkSheet.UsedRange.Rows.Count;
      if (col > 3) and (row > 6) then
      begin
        ftab := ResFolder + String(WorkSheet.Cells[6, 2]) + '.TAB';
        tab := miEval('PathToTableName$("%s")', [ftab]);//название с подчеркиванием
        basename := StringReplace(tab, '_', ' ', [rfReplaceAll]);//название с пробелами
        SafeCloseTable(tab);
        miDo('Create Table "%s" %s file "%s" TYPE NATIVE Charset "WindowsCyrillic"',
          [tab, fields, ftab]);

        data := VarArrayCreate([6, row, 3, col], varOleStr);//размер блока данных
        range := WorkSheet.Range[format('C6:%s%d', [ExcelNum2Str(WorkSheet, col), row])]; //J32
        data := Range.Value;
        for i := VarArrayLowBound(data,1) to VarArrayHighBound(data,1) do//по строкам
        begin
          s := '';
          for j := VarArrayLowBound(data,2) to VarArrayHighBound(data,2) do//по столбцам
            s := s + Format('"%s",', [String(VarArrayGet(data,[i,j]))]);
          s := s + '"",0,""';//LOCALITYCODE,SOURCELOCALITYCODE,FLAGLOCALITYCODE
          miDo('Insert into %s Values(%s)', [tab, s]);
        end;
        VarClear(data);
        range := Unassigned;
        miDo('Commit table %s', [tab]);

        miDo('Close table %s', [tab]);
        OpenDialog1.FileName := (format('%s', [ftab]));
        Building(1, false);

        if FileExists(ResFolder + basename + '_Compare.TAB') then
        begin
          path := ResFolder + 'Загрузка\';
          ForceDirectories(path);

          ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"',
            [GetAppFolder + 'PATTERNS\TelDa\Объекты.xlsx', path])), SW_HIDE, true, 0);
          RenameFile(path + 'Объекты.xlsx', path + basename + '.xlsx');

          miDo('Open table "%s" As klass', [GetAppFolder + 'PATTERNS\TelDa\Классификатор кодов населенных пунктов.TAB']);
          miDo('Open table "%s" As result', [ResFolder + basename + '_Compare.TAB']);

          miDo('Add Column result (OLDLOCALITYCODE) '+
            'From klass Set To кодОКТМО Where COL2 = COL1');
          miDo('Update result Set KodMO = "MO" +KodMO');

          j := miEvalInt('TableInfo(result, %d)', [TAB_INFO_NROWS]);
          data := VarArrayCreate([1, i, 1, 4], varOleStr);//размер блока данных
          for i := 1 to j do//по строкам
          begin
            miDo('Fetch rec %d From result', [i]);
            data[i, 1] := miEval('result.CADASTRALNUMBER');
            data[i, 2] := miEval('result.LOCALITYCODE');
            data[i, 3] := miEval('result.OLDLOCALITYCODE');
            data[i, 4] := miEval('result.KODMO');
          end;
          miDo('RollBack table result');
          UpdateExcel;
          VarClear(data);
          range := Unassigned;
          miDo('Close All');
        end;
      end;
    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.quit;
      WorkSheet := Unassigned;
      WorkBook := Unassigned;
    end;
  end;
  ResFolder := oldfolder;
  ParamForm.rgTitle.ItemIndex := oldtitleiindex;
  sl.Free;
  ExcelApp := Unassigned;
  Info('');
  ShowMessage('Обработка завершена');
end;
{------------------------поиск КН в колонке------------------------------------}
function TValuationForm.FindStartRow(var sheet : OleVariant; col : Integer) : Integer;
var
  i, row : Integer;
begin
  Result := 5;
  row := sheet.UsedRange.Rows.Count;
  for i := 1 to row do
    if Pos('66:', sheet.Cells[i, col]) = 1 then
    begin
      Result := i;
      exit;
    end;
end;
//==============================================================================
function TValuationForm.LoadDate_KVRI_4(feilename : String; var row2, col2 : Integer; var DataParcels : OleVariant) : Boolean;
var
  ExcelApp, Workbook, WorkSheet, range : OleVariant;
begin
  Result := false;
  ExcelApp := CreateOleObject('Excel.Application');
  try    //открываем данные и загоняем их в массив
    ExcelApp.Workbooks.Open(feilename);
    WorkBook := ExcelApp.Workbooks.item[1];
    if not FindSheetByName(WorkBook, WorkSheet, 'Объекты', row2, col2) then
      Exc('Лист Объекты не найден');
//    range := WorkSheet.Range[format('A6:%s%d', [ExcelNum2Str(WorkSheet, col2), row2])];
    range := WorkSheet.Range[format('A%d:%s%d', [FindStartRow(WorkSheet, 2),
      ExcelNum2Str(WorkSheet, col2), row2])];
    DataParcels := VarArrayCreate([1, row2, 1, col2], varOleStr); //varVariant
    DataParcels := range.Value;//взяли данные
    Result := true;
  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.quit;
    range := Unassigned;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;
//==============================================================================
procedure TValuationForm.mnTelda_KVRI_4Click(Sender: TObject);
var
  i, k, j, col, row, col2, row2, area, startrow : Integer;
  d : Double;
  listVRI, list1, list2 : TStringList;
  sa : StringArray;
  s, ss, filevri, cad, vri, category, bydoc, rule, id : String;
  ExcelApp, Workbook, WorkSheet, DataParcels, range : OleVariant;
  ini: TMemIniFile;

  procedure FindVRI(utilizationVRI : String);
  begin
    k := 0;
    vri := '*';
    rule := '-';
    Inc(k);

    if list1.Find(utilizationVRI, i) then
    begin
      rule := 'Правило0';
      vri := PAnsiChar(list1.Objects[i]);
      sa := SplitStr(vri, '-');
      vri := sa[0];
      id := sa[1];
      SetLength(sa, 0);
    end else

    if list2.Find(utilizationVRI, i) then
    begin
      rule := PAnsiChar(list2.Objects[i]);
      sa := SplitStr(rule, '-');
      rule := sa[0];
      id := sa[1];
      SetLength(sa, 0);
      if CText(rule, 'Правило6') then
        vri := 'этажность'
      else
        vri := ParamForm.CalcVri(Trim(Copy(rule, 8, 10)), category, area);
    end;

    if vri <> '*' then
      listVRI.Add(vri + '~' + rule + '~' + id)
    else
      listVRI.Add('~~');
  end;

begin//mnTelda_KVRI_4Click
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\РАЗОБРАТЬСЯ'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\ТЕСТ для результатов по опр КВРИ\'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\ТЕСТ для результатов по опр КВРИ_4'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\КВРИ\'; {$ENDIF}

  if not mnVRIFromBase.Enabled then
    Exc('Перед вызовом правил необходимо выбрать справочник.');

  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_4.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_4.xlsx');

  s := mnRulesName.Caption + #10 + 'расчитать по шаблону КВРИ_4' + #10 + 'Подтвердите действие';
  if MessageDlg (s, mtError, mbOKCancel, 0 ) = mrCancel then Exit;

  path := ResFolder + 'VRI\';
  if not ForceDirectories(path) then
    Exc('Не удалось создать каталог' + #10 + path);

  s := path + 'VRI_4_' + CurrentData + '.xlsx';//_4_
  if FileExists(s) then//проверяем наличие открыього файла
    if not DeleteFile(s) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));

  OpenDialog1.FileName := '*.xls*';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  Info('Создание рабочей таблицы...');
  filevri := path + 'VRI_4_' + CurrentData + '.xlsx';
  ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем шаблон
    [GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_4.xlsx', path]), SW_HIDE, true, 0);
  DeleteFile(filevri);//удаляем старый файл
  RenameFile(path + 'Шаблон для результатов по опр КВРИ_4.xlsx', filevri);//переименовываем шаблон

  if LoadDate_KVRI_4(OpenDialog1.FileName, row2, col2, DataParcels) then
  try
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Workbooks.Open(filevri);
    WorkBook := ExcelApp.Workbooks.item[1];
    FindSheetByName(WorkBook, WorkSheet, 'Объекты', row, col);
//    range := WorkSheet.Range[format('A6:%s%d', [ExcelNum2Str(WorkSheet, col2), row2])];
    startrow := FindStartRow(WorkSheet, 2);
    range := WorkSheet.Range[format('A%d:%s%d', [startrow, ExcelNum2Str(WorkSheet, col2), row2])];
    range.Value := DataParcels;//вставили данные

    listVRI := TStringList.Create;//Список VRI
    LoadRules(list1, list2);
    ParamForm.ReadRule(GetAppFolder + mnRulesName.Caption + '.xlsx');

    CodeCatZem := TStringList.Create;//загружаем справочник категорий земель
    ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
    try
     ini.ReadSectionValues('CodeCatZem', CodeCatZem);
    finally
      ini.Free;
    end;

    Info('Проверка правил...');
    pg.Max := VarArrayHighBound(DataParcels,1);
    for j := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
    begin
      pg.Position := pg.Position + 1;
      cad := DataParcels[j, 2];

      category := DataParcels[j, 5];//ds.FieldByName('ParcelCategory').AsString;
      category := CodeCatZem.Values[category];

      bydoc := Trim(StringReplace(DataParcels[j, 7], lf, ' ', [rfReplaceAll]));
      Val(DataParcels[j, 9], d, i);
      area := Round(d);
      FindVRI(bydoc);
    end;
    VarClear(DataParcels);

    DataParcels := VarArrayCreate([0, listVRI.Count-1, 0, 2], varOleStr);
    for i := 0 to listVRI.Count-1 do
    begin
      sa := SplitStr(listVRI[i], '~');
      for k := 0 to High(sa) do
        DataParcels[i,k] := sa[k];
      SetLength(sa, 0);
    end;
    s := ExcelNum2Str(WorkSheet, col2+1);
    ss := ExcelNum2Str(WorkSheet, col2+3);//конец формул
//    range := WorkSheet.Range[format('%s6:%s%d', [s, ss, listVRI.Count+5])];
    range := WorkSheet.Range[format('%s%d:%s%d', [s, startrow, ss, listVRI.Count+5])];
    range.Value := DataParcels;

    s := ExcelNum2Str(WorkSheet, col2+4);
    ss := ExcelNum2Str(WorkSheet, col);//конец формул
//    WorkSheet.Range[format('%s6:%s6', [s, ss])].Copy;
//    WorkSheet.Range[format('%s7:%s%d', [s, ss, row2])].PasteSpecial(operation:=xlNone);
    WorkSheet.Range[format('%s%d:%s%d', [s, startrow, ss, startrow])].Copy;
    WorkSheet.Range[format('%s%d:%s%d', [s, startrow+1, ss, row2])].PasteSpecial(operation:=xlNone);
    WorkBook.Save;
    ExcelApp.Visible := true;
  finally
    pg.Max := 0;
    listVRI.Free;
    FreeRules(list1, list2);
    CodeCatZem.Free;
    VarClear(DataParcels);
    range := Unassigned;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;
//==============================================================================
procedure TValuationForm.mnTelda_KVRI_2026Click(Sender: TObject);
var
  i, k, j, col, row, col2, row2, area, startrow : Integer;
  d : Double;
  listVRI, list1, list2 : TStringList;
  sa : StringArray;
  s, ss, filevri, cad, vri, category, bydoc, code, tipter, rule, id : String;
  ExcelApp, Workbook, WorkSheet, DataParcels, range : OleVariant;
  ini: TMemIniFile;

  procedure FindVRI(utilizationVRI : String);
  begin
    k := 0;
    vri := '*';
    rule := '-';
    id := '';
    Inc(k);

    if list1.Find(utilizationVRI, i) then
    begin
      rule := 'Правило0';
      vri := PAnsiChar(list1.Objects[i]);
      sa := SplitStr(vri, '-');
      vri := sa[0];
      id := sa[1];
      SetLength(sa, 0);
    end else

    if list2.Find(utilizationVRI, i) then
    begin
      rule := PAnsiChar(list2.Objects[i]);
      sa := SplitStr(rule, '-');
      rule := sa[0];
      id := sa[1];
      ss := sa[2];
      SetLength(sa, 0);

      if CText(rule, 'Правило8') then
      begin
        vri := ss;
      end else
      if CText(rule, 'Правило6') or CText(rule, 'Правило7') then
      begin
        vri := ParamForm.CalcVri6(Trim(Copy(rule, 8, 10)), category, tipter, area);
        if CText(vri,  'По правилу столбец D') then
          vri := ss;
      end else
        vri := ParamForm.CalcVri(Trim(Copy(rule, 8, 10)), category, area);
    end;

//    if vri <> '*' then
//      listVRI.Add(cad + '~' + code + '~' + tipter + '~~' + category +
//        '~~' + bydoc + '~~' + IntToStr(area) + '~' + vri + '~' + rule + '~' + id)
//    else
//      listVRI.Add(cad + '~' + code + '~' + tipter + '~~' + category +
//        '~~' + bydoc + '~~' + IntToStr(area) + '~Ручной разбор~' + rule + '~' + id);
    if vri = '*' then
      vri := '~Ручной разбор~';
    listVRI.Add(cad + '~' + code + '~' + tipter + '~~' + category +
      '~~' + bydoc + '~~' + IntToStr(area) + '~' + vri + '~' + rule + '~' + id);
  end;

begin//mnTelda_KVRI_2026Click
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\'; {$ENDIF}
  if not mnVRIFromBase.Enabled then
    Exc('Перед вызовом правил необходимо выбрать справочник.');

  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_2026.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_2026.xlsx');

  s := mnRulesName.Caption + #10 + 'расчитать по шаблону КВРИ_2026' + #10 + 'Подтвердите действие';
  if MessageDlg (s, mtError, mbOKCancel, 0 ) = mrCancel then Exit;

  path := ResFolder + 'VRI\';
  if not ForceDirectories(path) then
    Exc('Не удалось создать каталог' + #10 + path);

  s := path + 'VRI_2026_' + CurrentData + '.xlsx';//_2026_
  if FileExists(s) then//проверяем наличие открыього файла
    if not DeleteFile(s) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));

  OpenDialog1.FileName := '*.xls*';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  Info('Создание рабочей таблицы...');
  filevri := path + 'VRI_2026_' + CurrentData + '.xlsx';
  ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем шаблон
    [GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ_2026.xlsx', path]), SW_HIDE, true, 0);
  DeleteFile(filevri);//удаляем старый файл
  RenameFile(path + 'Шаблон для результатов по опр КВРИ_2026.xlsx', filevri);//переименовываем шаблон

  if LoadDate_KVRI_4(OpenDialog1.FileName, row2, col2, DataParcels) then
  try
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Workbooks.Open(filevri);
    WorkBook := ExcelApp.Workbooks.item[1];
    FindSheetByName(WorkBook, WorkSheet, 'Объекты', row, col);
//    range := WorkSheet.Range[format('A6:%s%d', [ExcelNum2Str(WorkSheet, col2), row2])];
    startrow := FindStartRow(WorkSheet, 2);
    range := WorkSheet.Range[format('A%d:%s%d', [startrow, ExcelNum2Str(WorkSheet, col2), row2])];
    range.Value := DataParcels;//вставили данные

    for i := 2 to col do//конец данных далее идут формулы
      if CText(String(WorkSheet.Cells[5, i]), 'Номер правила') then
        col := i;

    listVRI := TStringList.Create;//Список VRI
    LoadRules(list1, list2);
    ParamForm.ReadRule(GetAppFolder + mnRulesName.Caption + '.xlsx');

    CodeCatZem := TStringList.Create;//загружаем справочник категорий земель
    ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
    try
     ini.ReadSectionValues('CodeCatZem', CodeCatZem);
    finally
      ini.Free;
    end;

    Info('Проверка правил...');
    pg.Max := VarArrayHighBound(DataParcels,1);
    for j := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
    begin
      pg.Position := pg.Position + 1;
      cad := DataParcels[j, 2];
      code := DataParcels[j, 3];
      tipter := DataParcels[j, 4];
      if tipter = '' then
        tipter := 'Отсутствует информация';

      category := DataParcels[j, 6];
      category := CodeCatZem.Values[category];

      bydoc := Trim(StringReplace(DataParcels[j, 8], lf, ' ', [rfReplaceAll]));
      Val(DataParcels[j, 10], d, i);
      area := Round(d);
      FindVRI(bydoc);
    end;

    Info('Заполнение правил');
    VarClear(DataParcels);
    DataParcels := VarArrayCreate([0, listVRI.Count-1, 0, col-1], varOleStr);
    for i := 0 to listVRI.Count-1 do
    begin
      sa := SplitStr(listVRI[i], '~');
      for k := 0 to High(sa) do
        DataParcels[i,k] := sa[k];
      SetLength(sa, 0);
    end;

    Info('Сохранение таблицы...');
//    range := WorkSheet.Range[format('%s6:%s%d', [ExcelNum2Str(WorkSheet, 2),
//      ExcelNum2Str(WorkSheet, col), listVRI.Count+5])];
    range := WorkSheet.Range[format('%s%d:%s%d', [ExcelNum2Str(WorkSheet, 2), startrow,
      ExcelNum2Str(WorkSheet, col), listVRI.Count+5])];
    range.Value := DataParcels;

    col2 := WorkSheet.UsedRange.Columns.Count;
    s := ExcelNum2Str(WorkSheet, col+1);
    ss := ExcelNum2Str(WorkSheet, col2);//конец формул
//    WorkSheet.Range[format('%s6:%s6', [s, ss])].Copy;
//    WorkSheet.Range[format('%s7:%s%d', [s, ss, row2])].PasteSpecial(operation:=xlNone);
    WorkSheet.Range[format('%s%d:%s%d', [s, startrow, ss, startrow])].Copy;
    WorkSheet.Range[format('%s%d:%s%d', [s, startrow+1, ss, row2])].PasteSpecial(operation:=xlNone);
    WorkBook.Save;
    ExcelApp.Visible := true;
  finally
    pg.Max := 0;
    listVRI.Free;
    FreeRules(list1, list2);
    CodeCatZem.Free;
    VarClear(DataParcels);
    range := Unassigned;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;
//==============================================================================
procedure TValuationForm.mnTelda_KVRI_4_PlusClick(Sender: TObject);
var
  i, k, j, col, row, col2, row2, area, startrow : Integer;
  d : Double;
  listVRI, list1, list2 : TStringList;
  sa : StringArray;
  s, ss, filevri, vri, category, bydoc, rule, tipter, id : String;
  ExcelApp, Workbook, WorkSheet, DataParcels, range : OleVariant;
  ini: TMemIniFile;

  procedure FindVRI(utilizationVRI : String);
  begin
    k := 0;
    vri := '*';
    rule := '-';
    id := '';
    Inc(k);

    if list1.Find(utilizationVRI, i) then
    begin
      rule := 'Правило0';
      vri := PAnsiChar(list1.Objects[i]);
      sa := SplitStr(vri, '-');
      vri := sa[0];
      id := sa[1];
      SetLength(sa, 0);
    end else

    if list2.Find(utilizationVRI, i) then
    begin
      rule := PAnsiChar(list2.Objects[i]);
      sa := SplitStr(rule, '-');
      rule := sa[0];
      id := sa[1];
      ss := sa[2];
      SetLength(sa, 0);

      if CText(rule, 'Правило8') then
        vri := ss
      else

      if CText(rule, 'Правило6') or CText(rule, 'Правило7') then
      begin
//        category := DataParcels[j, 7];//для определения категории
//        if category = '' then
//          category := '0'
//        else
//          category := '3002000000';

        vri := ParamForm.CalcVri6(Trim(Copy(rule, 8, 10)), category, tipter, area);
        if CText(vri, 'По правилу столбец D') then
          vri := ss;
      end else

      if CText(rule, 'Правило9') then
      begin
//        category := DataParcels[j, 7];//для определения категории
//        if category = '' then
//          category := '0'
//        else
//          category := '3002000000';

        vri := ParamForm.CalcVri9(Trim(Copy(rule, 8, 10)), category, tipter, area);
        if CText(vri, 'По правилу столбец D') then
          vri := ss;
      end else
        vri := ParamForm.CalcVri(Trim(Copy(rule, 8, 10)), category, area);
    end;

    if vri <> '*' then
      listVRI.Add(vri + '~' + rule + '~' + id)
    else
      listVRI.Add('Ручной разбор~' + rule + '~' + id);
  end;

begin//mnTelda_KVRI_4_PlusClick
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\Большой объем\Новые пакеты'; {$ENDIF}
  if not mnVRIFromBase.Enabled then
    Exc('Перед вызовом правил необходимо выбрать справочник.');

  path := ResFolder + 'VRI\' + CurrentData + '\';
  ForceDirectories(path);
  OpenDialog1.FileName := '*.xls*';
  if not OpenDialog1.Execute then Exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  s := ExtractFileName(OpenDialog1.FileName);
  filevri := path + s;
  if FileExists(filevri) then//проверяем наличие открытого файла
    if not DeleteFile(filevri) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));

  ExcelApp := CreateOleObject('Excel.Application');
  try
    Info('Загрузка данных...');
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    WorkBook := ExcelApp.Workbooks.item[1];
    if not FindSheetByName(WorkBook, WorkSheet, 'Объекты', row2, col2) then
      Exc('Лист Объекты не найден');

    startrow := FindStartRow(WorkSheet, 2);
    range := WorkSheet.Range[format('B%d:%s%d', [startrow, ExcelNum2Str(WorkSheet, col2+3), row2])];
    DataParcels := range.Value;//взяли данные

    listVRI := TStringList.Create;//Список VRI
    LoadRules(list1, list2);
    ParamForm.ReadRule(GetAppFolder + mnRulesName.Caption + '.xlsx');

    CodeCatZem := TStringList.Create;//загружаем справочник категорий земель
    ini := TMemIniFile.Create(GetAppFolder + 'CodeNP.ini');
    try
     ini.ReadSectionValues('CodeCatZem', CodeCatZem);
    finally
      ini.Free;
    end;

    Info('Поиск правил...');
    pg.Max := VarArrayHighBound(DataParcels,1);
    for j := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
    begin
      pg.Position := j;
      tipter := DataParcels[j, 3];//тип территории
      if tipter = '' then
        tipter := 'Отсутствует информация';
      category := DataParcels[j, 4];//категория
      category := CodeCatZem.Values[category];

      bydoc := Trim(StringReplace(DataParcels[j, 5], lf, ' ', [rfReplaceAll]));
      Val(DataParcels[j, 6], d, i);//площадь
      area := Round(d);
      FindVRI(bydoc);
      Application.ProcessMessages;
    end;
    pg.Position := 1;
    pg.Repaint;

    Info('Заполнение правил');
    VarClear(DataParcels);
    DataParcels := VarArrayCreate([0, listVRI.Count-1, 0, 2], varOleStr);
    for i := 0 to listVRI.Count-1 do
    begin
      pg.Position := i + 1;
      sa := SplitStr(listVRI[i], '~');
      for k := 0 to High(sa) do
        DataParcels[i,k] := sa[k];
      SetLength(sa, 0);
    end;
    pg.Max := 0;
    pg.Repaint;

    Info('Сохранение таблицы...');
    range := WorkSheet.Range[format('%s%d:%s%d', [ExcelNum2Str(WorkSheet, col2+1), startrow,
      ExcelNum2Str(WorkSheet, col2+3), listVRI.Count+4])];
    range.Value := DataParcels;
    VarClear(DataParcels);

    WorkSheet.Cells[startrow-1, 9] := 'UtilizationByDoc';
    WorkSheet.Cells[startrow-1, 10] := 'PraviloVRI';
    WorkSheet.Cells[startrow-1, 11] := 'Pravilo Number';
    ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
    WorkBook.SaveAs(Filename:=filevri, FileFormat:= 51, CreateBackup:=False);
    ExcelApp.Visible := true;
  finally
    listVRI.Free;
    FreeRules(list1, list2);
    CodeCatZem.Free;
    range := Unassigned;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
    VarClear(DataParcels);
    Info('Готово');
  end;
end;
//==============================================================================
procedure TValuationForm.mnVRIFromBaseClick(Sender: TObject);
begin
  path := ResFolder + 'VRI_' + CurrentData + '\';
  if not ForceDirectories(path) then
    Exc('Не удалось создать каталог' + #10 + path);
  ds.Close;
  //минимальный набор необходимых полей, по которому можно определить код
  ds.CommandText := 'select CADASTRALNUMBER,AREA,PARCELCATEGORY,UTILIZATIONBYDOC,'+
    'UtilizationLandUse from parcels WHERE status_kc = ''Не готов'' AND '+
    '(flag_vri_all = ''Не готов'' OR flag_vri_all is Null)';

//  ds.CommandText := 'select top 100 CADASTRALNUMBER,AREA,PARCELCATEGORY,UTILIZATIONBYDOC,'+
//    'UtilizationLandUse from parcels';

  ds.Open;

  CalcVri;
end;

procedure TValuationForm.mnVRIFromFileClick(Sender: TObject);
var
  sql, oldfolder : String;
  i : Integer;
  sl : TStringList;
begin
  OpenDialog1.FileName := '*.txt';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  path := ResFolder + 'VRI_' + CurrentData + '\';
  if not ForceDirectories(path) then
    Exc('Не удалось создать каталог' + #10 + path);

  sl := TStringList.Create;
  sl.LoadFromFile(OpenDialog1.FileName);

  if not reg.IsMatch(sl[0], '66:\d\d:\d\d\d\d\d\d\d:\d{1,}') then
    sl.Delete(0);
  try
    cmd.Execute('DROP TABLE TemporalCadNums');
    cmd.Execute('CREATE TABLE TemporalCadNums(cadnum varchar(35))');
  except
    cmd.CommandText := 'DELETE FROM TemporalCadNums';
    cmd.Execute;
  end;

  Info('Выбор данных...');
  try//выбираем данные из базы запросом через временную таблицу с кадастровыми номерами
    cmd.Parameters.Clear();
    cmd.CommandType := cmdText;
    sql := 'Insert into TemporalCadNums (cadnum) values(:code)';
    cmd.Commandtext := sql;
    cmd.Parameters.ParseSQL(sql, True);
    cmd.Parameters.ParamByName('code').DataType := ftString;
    ADOConnection1.BeginTrans;
    for i := 0 to sl.Count-1 do
    begin
      cmd.Parameters[0].Value := sl[i];
      cmd.Execute;
    end;
    ds.Close;
    //минимальный набор необходимых полей, по которому можно определить код
    ds.CommandText := 'select CADASTRALNUMBER,AREA,PARCELCATEGORY,'+
      'UTILIZATIONBYDOC,UtilizationLandUse from parcels Where'+
      ' CADASTRALNUMBER=any(Select cadnum from TemporalCadNums)';
    ds.Open;

    CalcVri;
  finally
    sl.Free;
    ADOConnection1.RollbackTrans;
    cmd.Execute('DROP TABLE TemporalCadNums');
  end;
end;

procedure TValuationForm.mnVRIFromExcelClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet, range : OleVariant;
  data : OleVariant;
  row, col, i : Integer;
  s, ss : String;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\Cadastral_Valuation_VER_4\Win32\Debug\test\TelDa_XLS\ТЕСТ для результатов по опр КВРИ'; {$ENDIF}
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\Шаблон для результатов по опр КВРИ.xlsx');
  if not FileExists(GetAppFolder + 'PATTERNS\TelDa\ШАБЛОН_Поля для автозаполнения_Новые.xlsx') then
    Exc('Не найден шаблон ' + #10 + GetAppFolder + 'PATTERNS\TelDa\ШАБЛОН_Поля для автозаполнения_Новые.xlsx');

  path := ResFolder + 'VRI\';
  if not ForceDirectories(path) then
    Exc('Не удалось создать каталог' + #10 + path);

  s := path + 'VRI_' + CurrentData + '.xlsx';
  if FileExists(s) then
    if not DeleteFile(s) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));

  OpenDialog1.FileName := '*.xls*';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  try
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    WorkBook := ExcelApp.Workbooks.item[1];
    WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
    row := WorkSheet.UsedRange.Rows.Count;
    col := WorkSheet.UsedRange.Columns.Count;
    range := WorkSheet.Range[format('A4:%s%d', [ExcelNum2Str(WorkSheet, col), row])];
    data := VarArrayCreate([1, row-3, 1, col], varOleStr); //varVariant
    data := Range.Value;//взяли данные
    ExcelApp.Workbooks.Close;
    ExcelApp.quit;
    range := Unassigned;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
    //выбираем нужные данные перенося их в левую сторону
    for i := VarArrayLowBound(data,1) to VarArrayHighBound(data,1) do//по строкам
    begin
      data[i,1] := data[i,2];//CADASTRALNUMBER
      data[i,2] := data[i,9];//AREA(areaValue)
      data[i,6] := data[i,7];//Utilization{kindsOfUtilizations}
      data[i,7] := data[i,8];//UtilizationTxt(permittedUseText)
      data[i,8] := {#39 + }data[i,10];//(segmentOfUsage)   Код расчёта вида использования
    end;
    Info('Создание рабочей таблицы...');
    Dec(col, 2);
    //усекаем данные
    VarArrayRedim(data, col);
    data[1,1] := 'CADASTRALNUMBER';
    data[1,2] := 'AREA';
    data[1,3] := 'PARCELCATEGORY';
    data[1,4] := 'UTILIZATIONBYDOC';//parcelsUse bydoc
    data[1,5] := 'UtilizationLandUse';//parcelsUse allowedUse
    data[1,6] := 'Utilization';
    data[1,7] := 'UtilizationTxt';
    data[1,8] := 'segmentOfUsage';

    CalcVriXlsx(data);
  finally
    VarClear(data);
    if BreakCommand then
      ShowMessage('Обработка прервана')
    else
      ShowMessage('Обработка завершена');
  end;
end;

(*
Справочник источников
datasources   -  Источники данных
в таблице Building в полях, имя которых начинается с source - хранится значение поля ID из этого справочника
в файл для загрузки нужно помещать значение поля AccessID из этой таблицы
*)
procedure TValuationForm.mnInfoClick(Sender: TObject);
var
  i, k, d : Integer;
  path, newfile : String;
  sl : TStringList;
begin
  if OpenDialog1.InitialDir = '' then
    OpenDialog1.InitialDir := GetAppFolder + 'PATTERNS\здания\';
  OpenDialog1.FileName := '*.xls*';
  if OpenDialog1.Execute then
  try
    sl := TStringList.Create;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    path := ExtractFilePath(OpenDialog1.FileName) + CurrentData + '\';
    if not ForceDirectories(path) then
      Exc('Не удалось создать каталог' + #10 + path);

    for i := 0 to OpenDialog1.Files.Count-1 do
    begin
      newfile := ExtractFileName(OpenDialog1.Files[i]);
      if Pos('_', newfile) <> 1 then
        Exc('Файл не является шаблоном'+#10+ExtractFileName(OpenDialog1.Files[i]));
      newfile := path + Copy(newfile, 2, Length(newfile));
      if not CopyFile(PWideChar(OpenDialog1.Files[i]),  PWideChar(newfile), false) then
        Exc('Ошибка копировани файла '+ExtractFileName(OpenDialog1.Files[i])+#10+'Возможно файл открыт.')
      else
        sl.Add(newfile);
    end;

    k := 0;
    for i := 0 to sl.Count-1 do
    begin
      if Pos('этажность', OpenDialog1.Files[i]) > 0 then
        d := CalcModel(sl[i], sl.Count = 1, 2)
      else
        d := CalcModel(sl[i], sl.Count = 1, 1);
      if d = 0 then
        Inc(k);
      if BreakCommand then
        Exc('Расчет прерван');
      if sl.Count = 1 then
        case d of
          0: Exc('Ok');
          1: Exc('Нет выбранных записей');
          2: Exc('Ошика при поиске рабочей области');
          3: Exc('Ошика при создании таблицы');
          4: Exc('Ошика при вставке данных');
          5: Exc('Ошибка форматирование таблицы');
          6: Exc('Ошибка копирования формул');
        end;
    end;
    ShowMessage('Расчет завершен' + #10 + format('Созданных моделей %d', [k]));
    Info('');
  finally
    sl.Free;
  end;
end;


procedure TValuationForm.AllClick;
var
  i : Integer;
  s : String;
begin
  for i := 0 to mnObjects.Count-1 do//проверка открытых файлов EXCEL
  begin
    s := mnObjects.Items[i].Caption;
    s := format('%s%s', [ResFolder, s]);
    s := StringReplace(s, ' ', '_', [rfReplaceAll]) + '_загрузка.xlsm';
    if FileExists(s) then
      if not DeleteFile(s) then
        Exc(format('Файл %s открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));
  end;

  BreakCommand := false;
  for i := 0 to mnObjects.Count-1 do
  begin
    if BreakCommand then
      Exc('Обработка прервана');
    edSize.Text := '5000';
    mnObjects.Hint := mnObjects.Items[i].Caption;
    Info(mnObjects.Items[i].Caption);
    Application.ProcessMessages;
    if SetCommandText then
      if ImportDataTAB(false) then
        Building(1, false);
  end;
  ShowMessage('Обработка завершена');
end;

procedure TValuationForm.mnQueryClick(Sender: TObject);
begin
  ImportDataTAB(true);
  ds.Close;
end;
{-----------------------составляем строку запрося для выборки------------------}
function TValuationForm.SetCommandText : Boolean;
var
  j, k : Integer;
  fields : String;
begin
  Result := false;
  if BreakCommand then Exit;
  fs.Close;
  fs.CommandText := 'select ID, BASENAME, NAME from RTABLES';
  fs.Open;
  while not fs.eof do
  begin//ищем по названию выбранную таблицу
    if (not fs.FieldByName('BASENAME').IsNull) and CText(fs.FieldByName('NAME').AsString, mnObjects.Hint) then
    begin
      basename_ := fs.FieldByName('BASENAME').Value;
      basename := fs.FieldByName('NAME').Value;
      for j := 0 to lbTables.Count-1 do
      if CText(TablesList[j], basename_) then
        lbTables.Selected[j] := true
      else
        lbTables.Selected[j] := false;
      LoadFields;

      for k := Low(farray) to High(farray) do
        for j := 0 to lbFields.Count-1 do//отмечаем нужные поля
          if FieldsList[j] = farray[k] then
            lbFields.Selected[j] := true;

      fields := '';//создаем список полей для запроса
      for k  := Low(farray) to High(farray) do
        if FieldsList.IndexOf(farray[k]) <> -1  then
          fields := fields + farray[k] +',';
      SetLength(fields, Length(fields)-1);

      if Pos('FLAGLOCALITYCODE', fields) > 0 then
      begin
        CommandString.Text := format('select %s from %s WHERE (FLAGLOCALITYCODE = ''Не готов'' '+
          'OR FLAGLOCALITYCODE is Null)', [fields, basename_]);
        if mnObjects.Hint = 'Помещения' then
          CommandString.Text := CommandString.Text + ' AND (ORPHAN = 1)';
      end else
      begin
        CommandString.Text := format('select %s from %s WHERE ...', [fields, basename_]);
        Exc('Невозможно задать условие!'+#10+'Нет поля FLAGLOCALITYCODE');
      end;
      Result := true;
      exit;
    end;
    fs.next;
  end;
end;

procedure TValuationForm.ObjectClick(Sender: TObject);
var
  mif, mid : TStreamWriter;
  s, ss : String;
  i, k : Integer;
  reader: TStreamReader;
begin
/////////////////////////////
//mif := TStreamWriter.Create('D:\GDAL\Do_EKB_1.mif');
//mid := TStreamWriter.Create('D:\GDAL\Do_EKB_1.mid');
//mif.WriteLine('Version   300');
//mif.WriteLine('Charset "WindowsCyrillic"');
//mif.WriteLine('Delimiter ","');
//mif.WriteLine('CoordSys Earth Projection 8, 104, "m", 63, 0, 0.9996, 500000, 0 Bounds (-7745844.29597, -9997964.94324) (8745844.29597, 9997964.94324)');
//mif.WriteLine('Columns 1');
//mif.WriteLine('  ID Float');
//mif.WriteLine('Data');
//
//reader := TStreamReader.Create(TFileStream.Create('D:\GDAL\Do_EKB.txt', fmOpenRead));
//
//try
//  k := 0;
//  while not reader.EndOfStream do
//  begin
//    s := reader.ReadLine;
//    Application.ProcessMessages;
//    if Pos('-nan(ind)', s) = 0 then
//    begin
//      i := PosEx(' ', s, 10);
//      ss := Copy(s, i+1, Length(s));
//      s := Copy(s, 1, i-1);
//      mif.WriteLine('Point ' + s);
//      mid.WriteLine(ss);
//      Inc(k);
//      caption := inttostr(k);
//    end;
//    Application.ProcessMessages;
//  end;
//
//finally
//  reader.Free;
//  mif.Free;
//  mid.Free;
//end;
//EXIT;
/////////////////////////////
  BreakCommand := false;
  mnObjects.Hint := (Sender as TMenuItem).Caption;
  SetCommandText;
end;

procedure TValuationForm.mnBuildClick(Sender: TObject);
var
  i, k, r, col, row, currow, model, meter, lastcol : Integer;
  ExcelApp, Workbook, WorkSheet, WBook, WSheet, range : OleVariant;
  data : OleVariant;
  fields : TStringList;
  s, s1, s2, s3, s4, s5, name, tab : String;
  locfactor : Array of String;
  model_checked : TStringList;//списокx моделей для проверка полноты заполнения(Group_Real_Estate_Evaluative_Factor)

  //заполнение заголовка таблицы через запрос по группе моделей
  function PrepareTitle : Boolean;
  var
    n, j, min, max : Integer;
  begin
    meter := 0;

    Result := false;
    for n := 0 to OpenDialog1.Files.Count-1 do
    begin
      tab := ExtractFileName(OpenDialog1.Files[n]);
      Val(tab, min, j);
      if (j <> 5) or (min < 1000) or (min >= 9000) then
      begin
        ShowMessage('Ошибка в имени файла'+#10+OpenDialog1.Files[n]);
        exit;
      end;
    end;
    Result := true;

    i := 0;
    for n := 0 to OpenDialog1.Files.Count-1 do
    begin
      tab := ExtractFileName(OpenDialog1.Files[n]);
      Val(tab, min, j);
      case tab[1] of
        '1' : min := 1000;
        '2' : min := 2000;
        '3' : min := 3000;
        '4' : min := 4000;
        '5' : min := 5000;
        '6' : min := 6000;
        '7' : min := 7000;
        '8' : min := 8000;
      end;
      max := min + 1000;

      ds.Close;
      ds.CommandText := format('select * from Group_Real_Estate_Evaluative_Factor'+
        ' where GROUP_REAL_ESTATE_ID > %d and GROUP_REAL_ESTATE_ID < %d', [min, max]);
      ds.Open; //ds.SaveToFile('D:\Tokyo\DataBase_Calc_Code_NP\PATTERNS\Group_Real_Estate_Evaluative_Factor.xml');

      //накапливаем GROUP_REAL_ESTATE_ID требующих проверки Group_Real_Estate_Evaluative_Factor
      while not ds.eof do
      begin
        s3 := ds.FieldByName('GROUP_REAL_ESTATE_ID').AsString;
        model_checked.Add(s3);
        ds.next;
      end;

      if i = 0 then//при первом проходе выбираем кадастровый номер
      begin
        data := VarArrayCreate([0, 1], varOleStr);//создаем массив для 1 элемента
        data[0] := loc_factors[0].title;//кадастровый номер
        fields.Add(data[0]);
        i := 1;
      end;

      //добавлем уникальные ценовые факторы для каждой модели
      ds.FindFirst;
      while not ds.eof do
      begin
        s3 := 'loc_factor_'+ ds.FieldByName('EVALUATIVE_FACTOR_ID').AsString;
        if fields.IndexOf(s3) = -1 then
        begin
          VarArrayRedim(data, VarArrayHighBound(data, 1) + 1);//увеличиваем размер массива
          data[i] := s3;
          fields.Add(s3);
          Inc(i);
        end;
        ds.next;
      end;
    end;

    //увеличиваем размер массива, добавлем факторы из ЗУ
    VarArrayRedim(data, VarArrayHighBound(data, 1) + High(loc_factors)-1);
    for j := 1 to High(loc_factors) do
    begin
      data[i] := loc_factors[j].title;
      fields.Add(data[i]);
      Inc(i);
    end;

    lastcol := i;//последняя колонка факторов заголовка = "Old_Utlbydoc"
    s := ExcelNum2Str(WorkSheet, i);
    range := WorkSheet.Range[format('A1:%s1', [s])];//Область вставки
    range.Value := data;//строка заголовок
    VarClear(data);
  end;

  //поиск строки и колонки с кадастровым номером
  function FindCadNum_Local : Boolean;
  var
    j : Integer;
  begin
    Result := false;
    col := WSheet.UsedRange.Columns.Count;
    row := WSheet.UsedRange.Rows.Count;
    for j := 1 to col do
    begin
      s1 := WSheet.Cells[1, j];
      s2 := WSheet.Cells[2, j];
      s3 := WSheet.Cells[3, j];
      s4 := WSheet.Cells[4, j];
      s5 := WSheet.Cells[5, j];
      if (s1 = '') and (s2 = '') and (s3 = '') and (s4 = '') and (s5 = '') then
        break;
      r := 0;
      if CText('CadastralNumber', s1) then
        r := 1
      else if CText('CadastralNumber', s2) then
        r := 2
      else if CText('CadastralNumber', s3) then
        r := 3
      else if CText('CadastralNumber', s4) then
        r := 4
      else if CText('CadastralNumber', s5) then
        r := 5;
      if r <> 0 then
      begin
        if reg.IsMatch(String(WSheet.Cells[r+1, j]), '66:\d\d:\d\d\d\d\d\d\d:\d{1,}') then
        begin
          Result := true;
          //удаление пустых кадастровых номеров
          s := WSheet.Cells[row, 1];
          while s = '' do
          begin
            Dec(row);
            s := WSheet.Cells[row, 1];
          end;
          exit;
        end else
        begin
          row := 0;
          r := 0;
          break;
        end;
      end;
    end;
    log.Add(#9+#9+'Не найдена колонка с кадастровым номером');
  end;

  //проверка моделей по названию и запрос ценовых факторов по номеру модели
  function CheckModel : Boolean;
  begin
    Result := false;
    Val(tab, model, k);
    if model > 0 then
    begin
      fs.Close;
      fs.CommandText := format('select * from Group_Real_Estate_Evaluative_Factor where GROUP_REAL_ESTATE_ID=%d', [model]);
      fs.Open;//fs.SaveToFile('D:\Tokyo\DataBase_Calc_Code_NP\PATTERNS\Group_Real_Estate_Evaluative_Factor.xml');

      if model_checked.IndexOf(IntTostr(model)) = -1 then//проверка для этой модели не нужна
      begin
        Result := True;
        exit;
      end;

      if fs.RecordCount = 0 then
        model := 0;
    end;
    if model = 0 then
      log.Add(#9+#9+'Ошибка в названии модели или нет в таблице Group_Real_Estate_Evaluative_Factor (Отношение групп к факторам)')
    else
      Result := true;
  end;

  //проверка правильности названия ценовых факторов и наличе незаполненных ячеек
  procedure Evaluative_Factor_klass;
  var
    sl : TStringList;
    i, id_factor : Integer;
  begin
    if Pos('loc_factor_', s) = 0 then
      exit;
    Val(Copy(s, 12, Length(s)), id_factor, i);
    cf.Close;
    cf.CommandText := format('Select * from Evaluative_Factor_klass where id_factor=%d', [id_factor]);
    cf.Open;//cf.SaveToFile('D:\Tokyo\DataBase_Calc_Code_NP\PATTERNS\Evaluative_Factor_klass.xml');
    if cf.IsEmpty then
      exit;

    sl := TStringList.Create;
    sl.Sorted := true;
    try
//      if VarType(data) and varTypeMask  = varString then
      try
        sl.Add(data);
       except
        for i := VarArrayLowBound(data, 1) to VarArrayHighBound(data, 1) do
        begin
          s2 := VarToStr(data[i, 1]);
          if s2 = '' then
            log.Add(format('Отсутствует %s строка %d', [s, i+1]))
          else
            sl.Add(s2);
        end;
       end;
      for i := 0 to sl.Count-1 do
      begin
        cf.Close;
        cf.CommandText := format('Select * from Evaluative_Factor_klass where id_factor=%d AND QUALITATIVE_VALUE=''%s''', [id_factor, sl[i]]);
        cf.Open;
        if cf.IsEmpty then
          log.Add(format('Ошибка в названии фактора %s : %s', [s, sl[i]]));
      end;
    finally
      sl.Free;
    end;
  end;

  //выборка по столбцам
  procedure Parser;
  var
    c, j : Integer;
  begin
    if FindCadNum_Local then
    begin
      SetLength(locfactor, 0);
      for c := 1 to col do
      begin
        s3 := WSheet.Cells[r, c];//название фактора
        s := AnsiLowerCase(s3);
        if Pos('loc_factor_', s) = 1 then
        begin//ддобавим для проверки в функции Group_Real_Estate_Evaluative_Factor
          SetLength(locfactor, Length(locfactor) + 1);
          locfactor[High(locfactor)] := s;
        end;

        k := fields.IndexOf(s);//номер колонки для вставки
        if k < 0 then//может быть дополнительная колонка типа "loc_factor_'8071"
        begin
          if ParamForm.rgLocalFactor.ItemIndex = 1 then
            continue;//не включать "loc_factor_'8071"
          if Pos('loc_factor_', s) = 0 then
            continue;
          Inc(lastcol);
          WorkSheet.Cells[1, lastcol].Value := s;//добавляем название колонки в заголовок
          fields.Add(s);
          k := fields.Count-1;
        end;

        data := VarArrayCreate([1, row-r], varOleStr);//массив колонок таблицы источника
        s3 := ExcelNum2Str(WSheet, c);//название колонки
        range := WSheet.Range[format('%s%d:%s%d', [s3, r+1, s3, row])];//Область копирования
        data := range.Value;

        s3 := ExcelNum2Str(WorkSheet, k+1);//название колонки
        range := WorkSheet.Range[format('%s%d:%s%d', [s3, currow, s3, currow+row-r-1])];//Область вставки

        for j := Low(loc_factors) to High(loc_factors) do
          if (fields[k] = loc_factors[j].title) and (loc_factors[j].str = 1) then
            range.NumberFormat := '@';//задаем текстовый формат

        range.Value := data;

        Evaluative_Factor_klass;
        VarClear(data);
      end;
      Inc(meter);
    end;
  end;

  //проверка полноты заполнения по моделям
  procedure Group_Real_Estate_Evaluative_Factor;
  var
    j, k : Integer;
  begin
    if Length(locfactor) = 0 then
    begin
      log.Add(format('%s%sПри разборе модели ценовые факторы не найдены', [#9, #9]));
      exit;
    end;

    if not fs.IsEmpty then
    begin
      fs.FindFirst;
      while not fs.eof do
      begin
        s := 'loc_factor_'+ fs.FieldByName('EVALUATIVE_FACTOR_ID').AsString;
        k := 0;
        for j := Low(locfactor) to High(locfactor) do
          if CText(s, locfactor[j]) then
            Inc(k);
        if k = 0 then
          log.Add(format('%s%s%s не найден', [#9, #9, s]));
        fs.next;
      end;
    end;
    SetLength(locfactor, 0);
  end;

  //поиск листа по названию
  function FindList(listname : String): Boolean;
  var
    i : Integer;
  begin
    row := 0;  r := 0;
    Result := false;
    for i := 1 to WBook.Sheets.Count do
      if CText(WBook.Sheets[i].Name, listname) then
      begin
        WSheet := WBook.WorkSheets.Item[listname];
        Result := true;
        exit;
      end;
  end;

begin
  model_checked := TStringList.Create;
  model_checked.Sorted := true;
  model_checked.Duplicates := dupIgnore;
  log.Clear;
  fields := TStringList.Create;
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  Workbook := ExcelApp.WorkBooks.Add;
  WorkSheet := WorkBook.WorkSheets.Item[1];

  try
    OpenDialog1.FileName := '*.xls*';
    if OpenDialog1.Execute then
    begin
      OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
      if Pos('_', ExtractFileName(OpenDialog1.Files[0])) = 1 then
        Exc('Выбран файл шаблонов'+#10+'Необходимо выбрать расчетный файл');

      path := ExtractFilePath(OpenDialog1.FileName) + 'ITOG\';
      ForceDirectories(path);
      name := format('%s%s.xlsm', [path, CurrentData]);
      if FileExists(name) then
        if not DeleteFile(name) then
          Exc('Невозможно удалить файл'+#10+name+#10+'Возможно файл открыт');

      if not PrepareTitle then
        Exc('Ошибка в названии таблиц');

      currow := 2;
      pg.Max := OpenDialog1.Files.Count;
      for i := 0 to OpenDialog1.Files.Count-1 do
      try
        pg.Position := pg.Position + 1;
        WBook := ExcelApp.WorkBooks.Open(OpenDialog1.Files[i]);//Открываем рабочую книгу
        log.Add(OpenDialog1.Files[i]);
        tab := ExtractFileName(OpenDialog1.Files[i]);
        Info(tab);
        WorkSheet.Cells[currow, 1] := tab;
        Inc(Currow);
        if CheckModel then
        begin
          if FindList('ГОТОВ') or FindList('ГОТОВО') or FindList('ПЕРЕСЧЕТ') then
          begin
            log.Add(#9+'лист ' + WSheet.name);
            Parser;
            Group_Real_Estate_Evaluative_Factor;
          end else
            log.Add(#9+#9+'Не найден лист ГОТОВ или ПЕРЕСЧЕТ');

          currow := currow + row-r;// 1048576 строк
          if currow > 1048576 then
            Exc(format('Превышено максимальное количество строк на листе %s currow=%d', [#10, currow]));

          if FindList('ПЕРЕСЧЕТ НЕ ТРЕБУЕТСЯ') then
          begin
            log.Add(#9+'лист ПЕРЕСЧЕТ НЕ ТРЕБУЕТСЯ');
            Parser;
          end else
            log.Add(#9+#9+'Не найден лист ПЕРЕСЧЕТ НЕ ТРЕБУЕТСЯ');

          currow := currow + row-r;// 1048576 строк
          if currow > 1048576 then
            Exc(format('Превышено максимальное количество строк на листе %s currow=%d', [#10, currow]));

        end;
      finally
        WBook.close;
        WBook := Unassigned;
        WSheet := Unassigned;
      end;
      if meter > 0 then
      begin
        ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
        ExcelApp.WindowState := -4137;
        WorkBook.SaveAs(Filename:=name, FileFormat:= 52, CreateBackup:=False);
        ExcelApp.Visible := true;
      end else
        WorkBook.Close;
    end;
  finally
    if path <> '' then
      log.SaveToFile(format('%s%s_log.txt', [path, CurrentData]));
    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    model_checked.Free;
    fields.Free;
    pg.Max := 0;
  end;
end;

end.
