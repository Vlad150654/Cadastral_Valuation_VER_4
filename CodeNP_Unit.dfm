object mainForm: TmainForm
  Left = 0
  Top = 0
  Caption = #1050#1072#1076#1072#1089#1090#1088#1086#1074#1072#1103' '#1086#1094#1077#1085#1082#1072' '#1086#1073#1098#1077#1082#1090#1086#1074' '#1057#1074#1077#1088#1076#1083#1086#1074#1089#1082#1086#1081' '#1086#1073#1083#1072#1089#1090#1080
  ClientHeight = 173
  ClientWidth = 725
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIForm
  KeyPreview = True
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 725
    Height = 131
    Align = alClient
    TabOrder = 0
    object lbTables: TListBox
      Left = 1
      Top = 1
      Width = 200
      Height = 129
      Align = alLeft
      ItemHeight = 13
      TabOrder = 0
      OnClick = lbTablesClick
      OnKeyDown = lbTablesKeyDown
    end
    object lbFields: TListBox
      Left = 201
      Top = 1
      Width = 523
      Height = 129
      Align = alClient
      Columns = 5
      ItemHeight = 13
      MultiSelect = True
      TabOrder = 1
      OnClick = lbFieldsClick
      OnKeyDown = lbFieldsKeyDown
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 131
    Width = 725
    Height = 42
    Align = alBottom
    TabOrder = 1
    object CommandString: TEdit
      Left = 1
      Top = 1
      Width = 723
      Height = 21
      Align = alTop
      TabOrder = 0
    end
    object cbCPU: TComboBox
      Left = 547
      Top = -5
      Width = 40
      Height = 21
      Style = csDropDownList
      TabOrder = 1
      Visible = False
    end
    object edSize: TEdit
      Left = 680
      Top = 0
      Width = 40
      Height = 21
      TabOrder = 2
      Text = '5000'
      Visible = False
    end
    object stBar: TStatusBar
      Left = 1
      Top = 22
      Width = 392
      Height = 19
      Align = alLeft
      Panels = <
        item
          Width = 90
        end
        item
          Width = 50
        end>
    end
    object pg: TProgressBar
      Left = 393
      Top = 22
      Width = 331
      Height = 19
      Align = alClient
      TabOrder = 4
    end
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;User I' +
      'D=zisadmin;Initial Catalog=btizu2022;Data Source=TESTSQL;Use Pro' +
      'cedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workst' +
      'ation ID=G001;Use Encryption for Data=False;Tag with column coll' +
      'ation when possible=False'
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'SQLOLEDB.1'
    Left = 656
    Top = 24
  end
  object OpenDialog1: TOpenDialog
    OnClose = OpenDialog1Close
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 264
    Top = 32
  end
  object MainMenu1: TMainMenu
    AutoHotkeys = maManual
    Left = 392
    Top = 48
    object mnObjects: TMenuItem
      Caption = #1054#1073#1098#1077#1082#1090#1099
      object mnBuildings: TMenuItem
        Caption = #1047#1076#1072#1085#1080#1103
        OnClick = ObjectClick
      end
      object mnConstructions: TMenuItem
        Caption = #1057#1086#1086#1088#1091#1078#1077#1085#1080#1103
        OnClick = ObjectClick
      end
      object mnUncompleteds: TMenuItem
        Caption = #1054#1073#1098#1077#1082#1090#1099' '#1085#1077#1079#1072#1074#1077#1088#1096#1077#1085#1085#1086#1075#1086' '#1089#1090#1088#1086#1080#1090#1077#1083#1100#1089#1090#1074#1072
        RadioItem = True
        OnClick = ObjectClick
      end
      object mnFlats: TMenuItem
        Caption = #1055#1086#1084#1077#1097#1077#1085#1080#1103
        OnClick = ObjectClick
      end
      object mnParcels: TMenuItem
        Caption = #1047#1077#1084#1077#1083#1100#1085#1099#1077' '#1091#1095#1072#1089#1090#1082#1080
        OnClick = ObjectClick
      end
    end
    object mnSelect: TMenuItem
      Caption = #1042#1099#1073#1086#1088#1082#1072
      object mnQuery: TMenuItem
        Caption = #1055#1086' '#1079#1072#1087#1088#1086#1089#1091' '#1080#1079' '#1082#1086#1084#1072#1085#1076#1099
        OnClick = mnQueryClick
      end
      object mnCadNumber: TMenuItem
        Caption = #1055#1086' '#1082#1072#1076#1072#1089#1090#1088#1086#1074#1099#1084' '#1085#1086#1084#1077#1088#1072#1084' '#1080#1079' '#1073#1091#1092#1077#1088#1072
        OnClick = mnCadNumberClick
      end
    end
    object mnCode: TMenuItem
      Caption = #1050#1086#1076#1099' '#1053#1055
      object mnCodeFromSelection: TMenuItem
        Caption = #1048#1079' '#1074#1099#1073#1086#1088#1082#1080
        OnClick = mnCodeFromSelectionClick
      end
      object mnCodeFromBase: TMenuItem
        Caption = #1048#1079' '#1073#1072#1079#1099' '#1087#1086' '#1074#1089#1077#1084' '#1086#1073#1098#1077#1082#1090#1072#1084
        OnClick = mnCodeFromBaseClick
      end
      object mnCodeFromFile: TMenuItem
        Caption = #1048#1079' '#1092#1072#1081#1083#1072' (*.txt)'
        OnClick = mnCodeFromFileClick
      end
      object mnCodeFromExcel: TMenuItem
        Caption = #1048#1079' '#1058#1077#1083#1044#1072' (*.xlsx)'
        OnClick = mnCodeFromExcelClick
      end
    end
    object mnBuilds: TMenuItem
      Caption = #1047#1076#1072#1085#1080#1103
      object mnCheckFloors: TMenuItem
        Caption = #1055#1088#1086#1074#1077#1088#1082#1072' '#1101#1090#1072#1078#1085#1086#1089#1090#1080
        OnClick = mnCheckFloorsClick
      end
      object mnInfo: TMenuItem
        Caption = #1048#1085#1092#1086#1088#1084#1072#1094#1080#1103
        OnClick = mnInfoClick
      end
    end
    object mnRules: TMenuItem
      Caption = #1055#1088#1072#1074#1080#1083#1072
      object mnRulesName: TMenuItem
        Caption = #1057#1087#1088#1072#1074#1086#1095#1085#1080#1082
      end
      object mnVRIFromBase: TMenuItem
        Caption = #1048#1079' '#1073#1072#1079#1099
        Enabled = False
        OnClick = mnVRIFromBaseClick
      end
      object mnVRIFromFile: TMenuItem
        Caption = #1048#1079' '#1092#1072#1081#1083#1072' (*.txt)'
        Enabled = False
        OnClick = mnVRIFromFileClick
      end
      object mnVRIFromExcel: TMenuItem
        Caption = #1048#1079' '#1058#1077#1083#1044#1072' (*.xlsx)'
        Enabled = False
        OnClick = mnVRIFromExcelClick
      end
    end
    object mnParams: TMenuItem
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099
      OnClick = mnParamsClick
    end
    object mnModels: TMenuItem
      Caption = #1052#1086#1076#1077#1083#1080
      OnClick = mnModelsClick
    end
    object mnBuild: TMenuItem
      Caption = #1057#1073#1086#1088#1082#1072
      OnClick = mnBuildClick
    end
    object mnBase: TMenuItem
      Caption = #1041#1072#1079#1072
    end
    object mnTelDa: TMenuItem
      Caption = 'TelDa'
      object mnParamBuildings: TMenuItem
        Caption = #1047#1076#1072#1085#1080#1103' - '#1086#1087#1088#1077#1076#1077#1083#1077#1085#1080#1077' '#1087#1072#1088#1072#1084#1077#1090#1088#1086#1074
        OnClick = mnParamBuildingsClick
      end
      object mnFlats_Buildings: TMenuItem
        Caption = #1055#1086#1084#1077#1097#1077#1085#1080#1103'-'#1088#1077#1079#1091#1083#1100#1090#1072#1090
        Hint = 'PATTERNS\TelDa\'#1064#1072#1073#1083#1086#1085'_'#1047#1076#1072#1085#1080#1103'_'#1055#1086#1084#1077#1097#1077#1085#1080#1103'.xlsx'
        OnClick = mnFlats_BuildingsClick
      end
      object mnFlats_Orphan: TMenuItem
        Caption = #1055#1086#1084#1077#1097#1077#1085#1080#1103'-'#1089#1080#1088#1086#1090#1099
        Hint = 'PATTERNS\TelDa\'#1064#1072#1073#1083#1086#1085'_'#1055#1086#1084#1077#1097#1077#1085#1080#1103'_'#1057#1080#1088#1086#1090#1099'.xlsx'
        OnClick = mnFlats_OrphanClick
      end
      object mnCarParkingSpaces: TMenuItem
        Caption = #1052#1072#1096#1080#1085#1086#1084#1077#1089#1090#1072
        Hint = 'PATTERNS\TelDa\'#1064#1072#1073#1083#1086#1085' '#1052#1072#1096#1080#1085#1086#1084#1077#1089#1090#1072'.xlsx'
        OnClick = mnCarParkingSpacesClick
      end
      object mnTelda_KVRI_4: TMenuItem
        Caption = #1054#1087#1088#1077#1076#1077#1083#1077#1085#1080#1077' '#1050#1042#1056#1048' 4'
        OnClick = mnTelda_KVRI_4Click
      end
      object mnTelda_KVRI_2026v5: TMenuItem
        Caption = #1054#1087#1088#1077#1076#1077#1083#1077#1085#1080#1077' '#1050#1042#1056#1048' 2026'
        OnClick = mnTelda_KVRI_2026Click
      end
      object mnTelda_KVRI_4_Plus: TMenuItem
        Caption = #1054#1087#1088#1077#1076#1077#1083#1077#1085#1080#1077' '#1050#1042#1056#1048' 4_'#1055#1083#1102#1089
        OnClick = mnTelda_KVRI_4_PlusClick
      end
      object mnCompareZU: TMenuItem
        Caption = #1057#1088#1072#1074#1085#1077#1085#1080#1103' '#1089' '#1088#1086#1076#1080#1090#1077#1083#1077#1084' '#1047#1059
        OnClick = mnCompareZUClick
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object ExcelCorrect: TMenuItem
        Caption = #1048#1089#1087#1088#1072#1074#1083#1077#1085#1080#1077' Excel '#1092#1072#1081#1083#1072
        OnClick = ExcelCorrectClick
      end
    end
  end
  object ADOConnection_test: TADOConnection
    ConnectionString = 'Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False'
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 527
    Top = 40
  end
  object cmd: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 72
    Top = 40
  end
end
