unit CalcCodePart;

interface

uses
  SysUtils,
  OleAuto,
  Variants,
  Dialogs,
  Forms,
  ComCtrls,
  MapBasic_INT,
  classes,
  Common;

  procedure Work(table, spr, param : String; pg : TProgressBar);

implementation

procedure Work(table, spr, param : String;  pg : TProgressBar);
var
  k, row, rn_, mo_, np_ : Integer;
  cadraion, mo, tipmo, np, tipnp, raion, code, address : String;
  needfind : Boolean;

  procedure Updater(algo, loccode : Integer; sel : String);
  var
    cn, cadrai : String;
  begin
    k := miEvalInt('TableInfo(%s, %d)', [sel, TAB_INFO_NROWS]);
    if k = 1 then
    begin
      try
        cadrai := miEval('%s.Кадастровый_район', [sel]);
        cn := miEval('locality.CADASTRALNUMBER');
        if Copy(cn, 1, 5) <> cadrai then
          miDo('Update locality Set Проверка="1" where rowid=%d', [row]);
      except
        cn := '';
      end;

      code := miEval('%s.код_новый', [sel]);//проверяем не получил ли НП новый код
      if code = '' then //нового кода нет, берем текущий
        code := miEval('%s.кодНП_Мо', [sel]);
      needfind := false;
      miDo('Update locality Set код_НП="%s",алгоритм=%d,Район=%d,МО=%d,НП=%d,'+
        'Источник_НП=%d, KodMO="%s" where rowid=%d',
        [code, algo, rn_, mo_, np_, loccode, Copy(code, 1, 2), row]);
    end;
    if Length(sel) = 3 then//это второй уровени выборки (SEL)
      miDo('Close table %s', [sel]);
  end;

begin
  if pg = nil then
    mi := CreateOleObject('MapInfo.Application');

  try
    miDo('Open Table "%s" as locality', [table]);
    miDo('Set Table locality FastEdit On');
    miDo('Set Table locality Undo Off');
    miDo('Open Table "%s" As SPR', [spr]);

    if pg <> nil then
    begin
      pg.Position := 0;
      pg.Max := miEvalInt('TableInfo(locality, %d)', [TAB_INFO_NROWS]);
    end;

    While miEval('EOT(locality)') <> 'T' do
    try
      Application.ProcessMessages;
      if BreakCommand then exit;

      row := miEvalInt('locality.rowid');
      if pg <> nil then
        pg.Position := pg.Position + 1;

      if param = '1' then
      begin
        mo := StringReplace(miEval('locality.CITYNAME'), '"', '', [rfReplaceAll]);
        tipmo := miEval('locality.CITYTYPE');
        np := StringReplace(miEval('locality.LOCALITYNAME'), '"', '', [rfReplaceAll]);
        tipnp := miEval('locality.LOCALITYTYPE');
        raion := StringReplace(miEval('locality.DISTRICTNAME'), '"', '', [rfReplaceAll]);
        cadraion := Copy(miEval('locality.CADASTRALBLOCK'), 1, 5);//выбираем кадастровый район из кадастроваого квартала
      end else
      begin
        mo := ''; tipmo := ''; np := ''; tipnp := ''; raion := ''; cadraion := '';
        address := miEval('locality.Адрес');

        k := Pos(',', address);
        address := Trim(Copy(address, k+1, Length(address)));//адрес без области
        k := Pos(',', address);
        code := Trim(Copy(address, 1, k-1));
        address := Trim(Copy(address, k+1, Length(address)));

        k := Pos(' район', code);
        if k > 0 then
          raion := Trim(Copy(code, 1, k-1));
        k := Pos(' городской округ', code);
        if k > 0 then
          mo := Trim(Copy(code, 1, k-1));

        if (raion = '') and (mo = '') and (code <> '') then
          np := code
        else
          np := address;

        k := Pos(',', np);
        if k > 0 then
          np := Copy(np, 1, k-1);

        k := Pos('пос.', np);
        if k > 0 then
        begin
          tipnp := 'п';
          np := Trim(StringReplace(np, 'пос.', '', [rfReplaceAll]));
        end;
        k := Pos('с.', np);
        if k > 0 then
        begin
          tipnp := 'с';
          np := Trim(StringReplace(np, 'с.', '', [rfReplaceAll]));
        end;
        k := Pos('д.', np);
        if k > 0 then
        begin
          tipnp := 'д';
          np := Trim(StringReplace(np, 'д.', '', [rfReplaceAll]));
        end;
        k := Pos('пгт', np);
        if k > 0 then
        begin
          tipnp := 'пгт';
          np := Trim(StringReplace(np, 'пгт', '', [rfReplaceAll]));
        end;
      end;
////////////////////////////
      rn_ := Integer(raion <> '');
      mo_ := Integer(mo <> '');
      np_ := Integer(np <> '');

      needfind := true;
      if np <> '' then
      begin//выбираем с учетом типа второго населенного пункта
        np := AnsiUpperCase(StringReplace(np, 'ё', 'е', [rfReplaceAll]));
        miDo('Select * from spr Where UCase$(НаименованиеНП_МО)="%s" into SelSpr noSelect', [np]);//совпадение по уникальному названию
        Updater(100, 751, 'SelSpr');
        if needfind and (k > 1) then //найдено несколько одинаковых названий
        begin
          miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
            'префикс="%s" into SEL noSelect', [np, tipnp]);
          Updater(101, 755, 'SEL');
          if needfind then //найдено несколько одинаковых названий одинакового типа
          begin

            if raion <> '' then//НП в составе района
            begin
              miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                'INSTR(1,Наименование_АТЕ,"%s")>0 into SEL noSelect', [np, raion]);
              Updater(102, 752, 'SEL');
              if needfind then //в районе найдено несколько одинаковых названий
              begin
                miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                  'префикс="%s" AND INSTR(1,Наименование_АТЕ,"%s")>0 into SEL noSelect', [np, tipnp, raion]);
                Updater(103, 758, 'SEL');
              end;
            end;

            if needfind and (mo <> '') then//НП в составе МО
            begin
              miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                'INSTR(1,Наименование_МО,"%s")>0 into SEL noSelect', [np, mo]);
              Updater(104, 756, 'SEL');
              if needfind then//в МО несколько одинаковых названий
              begin
                miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND префикс="%s" AND '
                  +'INSTR(1,Наименование_МО,"%s")>0 into SEL noSelect', [np, tipnp, mo]);
                Updater(105, 763, 'SEL');
                if needfind then //вместо названия НП подставляем МО (Шувакиш = Екатеринбург)
                begin
                  miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                    'префикс="%s" into SEL noSelect', [mo, tipmo]);
                  Updater(106, 764, 'SEL');
                end;
              end;
            end;

            if needfind and (cadraion <> '') then //выбираем по вхождению кадастрового номера
            begin
              miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                'Кадастровый_район="%s" into SEL noSelect', [np, cadraion]);
              Updater(107, 757, 'SEL');
              if needfind then//несколько одинаковых названий
              begin
                miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
                  'префикс="%s" AND Кадастровый_район="%s" into SEL noSelect', [np, tipnp, cadraion]);
                Updater(108, 765, 'SEL');
              end;
            end;

          end;
        end;
        miDo('Close table SelSpr');
      end else

      if needfind and (mo <> '') then
      begin
        mo := AnsiUpperCase(StringReplace(mo, 'ё', 'е', [rfReplaceAll]));
        miDo('Select * from spr Where UCase$(НаименованиеНП_МО)="%s" into SelSpr noSelect', [mo]);
        Updater(200, 749, 'SelSpr');
        if needfind and (k > 1) then //найдено несколько одинаковых названий
        begin
          miDo('Select * from SelSpr Where UCase$(НаименованиеНП_МО)="%s" AND '+
            'префикс="%s" into SEL noSelect', [mo, tipmo]);
          Updater(201, 750, 'SEL');
        end;
        miDo('Close table SelSpr');
      end;

      miDo('Fetch Rec %d From locality', [row]);
      MiDo('Fetch Next From locality');
    except
      on e:exception do
      begin
        showmessage(e.Message);
        MiDo('Fetch Next From locality');
      end;
    end;
    miDo('Commit table locality');
  finally
    if pg = nil then
    begin
      miDo('End MapInfo');
      mi := Unassigned;
    end;
    Application.ProcessMessages;
    DeleteFile(ChangeFileExt(spr, '.key'));
  end;
end;
//https://geekon.media/process-antimalware-service-executable    отключить Антивирус
end.
