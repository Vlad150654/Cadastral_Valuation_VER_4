object ParamForm: TParamForm
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099
  ClientHeight = 413
  ClientWidth = 537
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Rules: TPageControl
    Left = 0
    Top = 0
    Width = 537
    Height = 374
    ActivePage = Правила
    Align = alClient
    TabOrder = 0
    OnChange = RulesChange
    object Правила: TTabSheet
      Caption = #1055#1088#1072#1074#1080#1083#1072
      object sg: TStringGrid
        Left = 0
        Top = 0
        Width = 529
        Height = 346
        Align = alClient
        DefaultColWidth = 100
        DefaultRowHeight = 20
        RowCount = 6
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goRowSelect, goFixedRowClick]
        TabOrder = 0
        RowHeights = (
          20
          19
          20
          20
          20
          20)
      end
    end
    object Excel: TTabSheet
      Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
      ImageIndex = 1
      object rgTitle: TRadioGroup
        Left = 148
        Top = 278
        Width = 241
        Height = 65
        Caption = #1047#1072#1075#1086#1083#1086#1074#1082#1080' '#1090#1072#1073#1083#1080#1094
        ItemIndex = 0
        Items.Strings = (
          #1055#1077#1088#1077#1074#1086#1076' '#1085#1072' '#1088#1091#1089#1089#1082#1080#1081
          #1053#1077' '#1087#1077#1088#1077#1074#1086#1083#1080#1090#1100)
        TabOrder = 0
      end
      object rgFlagNP: TRadioGroup
        Left = 148
        Top = 25
        Width = 241
        Height = 80
        Caption = #1042#1099#1073#1086#1088' '#1085#1072#1089#1077#1083#1077#1085#1085#1099#1093' '#1087#1091#1085#1082#1090#1086#1074
        ItemIndex = 0
        Items.Strings = (
          #1042#1089#1077' '#1074#1084#1077#1089#1090#1077
          #1053#1086#1074#1099#1077
          #1048#1079#1084#1077#1085#1077#1085#1085#1099#1077)
        TabOrder = 1
      end
      object rgBaseLoader: TRadioGroup
        Left = 148
        Top = 111
        Width = 241
        Height = 65
        Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1086#1074' '#1088#1072#1089#1095#1077#1090#1072
        ItemIndex = 1
        Items.Strings = (
          #1055#1086#1089#1083#1077' '#1088#1072#1089#1095#1077#1090#1072
          #1054#1090#1083#1086#1078#1080#1090#1100' '#1079#1072#1075#1088#1091#1079#1082#1091)
        TabOrder = 2
      end
      object rgLocalFactor: TRadioGroup
        Left = 148
        Top = 192
        Width = 241
        Height = 65
        Caption = #1044#1086#1087#1086#1083#1085#1080#1090#1077#1083#1100#1085#1099#1077' '#1083#1086#1082#1072#1083#1100#1085#1099#1077' '#1092#1072#1082#1090#1086#1088#1099
        ItemIndex = 0
        Items.Strings = (
          #1042#1082#1083#1102#1095#1072#1090#1100' '#1074' '#1089#1073#1086#1088#1082#1091
          #1048#1075#1085#1086#1088#1080#1088#1086#1074#1072#1090#1100)
        TabOrder = 3
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 374
    Width = 537
    Height = 39
    Align = alBottom
    TabOrder = 1
    object brOk: TBitBtn
      Left = 136
      Top = 6
      Width = 75
      Height = 25
      Caption = 'Ok'
      TabOrder = 0
      OnClick = brOkClick
    end
    object btCheck: TBitBtn
      Left = 338
      Top = 6
      Width = 75
      Height = 25
      Caption = #1055#1088#1086#1074#1077#1088#1082#1072
      TabOrder = 1
      WordWrap = True
      OnClick = btCheckClick
    end
    object btCancel: TBitBtn
      Left = 239
      Top = 6
      Width = 75
      Height = 25
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
      OnClick = btCancelClick
    end
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;User I' +
      'D=adievalm;Initial Catalog=btiall;Data Source=TESTSQL;Use Proced' +
      'ure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstati' +
      'on ID=G001;Use Encryption for Data=False;Tag with column collati' +
      'on when possible=False'
    LoginPrompt = False
    Mode = cmRead
    Provider = 'SQLOLEDB.1'
    Left = 173
    Top = 112
  end
end
