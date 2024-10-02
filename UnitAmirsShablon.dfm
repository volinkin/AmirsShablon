object Form1: TForm1
  Left = 97
  Top = 156
  Width = 1361
  Height = 668
  VertScrollBar.Visible = False
  Caption = #1040#1052#1048#1056#1057' '#1064#1072#1073#1083#1086#1085#1099
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -19
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 24
  object Button1: TButton
    Left = 136
    Top = 0
    Width = 97
    Height = 41
    Caption = #1055#1088#1086#1092#1080#1083#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object DateTimePicker1: TDateTimePicker
    Left = 232
    Top = 0
    Width = 81
    Height = 37
    CalAlignment = dtaLeft
    Date = 45326.9332518518
    Format = 'yyyy'
    Time = 45326.9332518518
    DateFormat = dfShort
    DateMode = dmUpDown
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Kind = dtkDate
    ParseInput = False
    ParentFont = False
    TabOrder = 1
    OnChange = DateTimePicker1Change
  end
  object DBNavigator1: TDBNavigator
    Left = 696
    Top = 0
    Width = 220
    Height = 41
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    TabOrder = 2
  end
  object Edit1: TEdit
    Left = 528
    Top = 0
    Width = 169
    Height = 37
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = Edit1Click
    OnEnter = Edit1Enter
    OnKeyPress = Edit1KeyPress
  end
  object Memo1: TMemo
    Left = 0
    Top = 48
    Width = 1265
    Height = 721
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Lines.Strings = (
      'Memo1')
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 4
  end
  object Button2: TButton
    Left = 1096
    Top = 0
    Width = 89
    Height = 41
    Caption = 'WORD'
    TabOrder = 5
    OnClick = Button2Click
  end
  object ComboBox1: TComboBox
    Left = 0
    Top = 0
    Width = 137
    Height = 33
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -21
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 25
    ParentFont = False
    TabOrder = 6
  end
  object Button4: TButton
    Left = 912
    Top = 0
    Width = 97
    Height = 41
    Caption = #1054#1073#1085#1086#1074#1080#1090#1100
    TabOrder = 7
    OnClick = Button4Click
  end
  object ComboBox2: TComboBox
    Left = 312
    Top = 0
    Width = 217
    Height = 32
    ItemHeight = 24
    TabOrder = 8
    Text = 'ComboBox2'
    OnChange = ComboBox2Change
  end
  object Button5: TButton
    Left = 1008
    Top = 0
    Width = 89
    Height = 41
    Caption = #1064#1072#1073#1083#1086#1085
    TabOrder = 9
    OnClick = Button5Click
  end
end
