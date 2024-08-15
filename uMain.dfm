object Form3: TForm3
  Left = 0
  Top = 0
  Caption = 'Form3'
  ClientHeight = 634
  ClientWidth = 954
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object eFileName: TEdit
    Left = 8
    Top = 16
    Width = 857
    Height = 23
    TabOrder = 0
    Text = 'C:\Temp\3.xlsx'
  end
  object Button1: TButton
    Left = 871
    Top = 16
    Width = 75
    Height = 25
    Caption = #48660#47084#50724#44592
    TabOrder = 1
    OnClick = Button1Click
  end
  object StringGrid1: TStringGrid
    Left = 176
    Top = 48
    Width = 777
    Height = 542
    ColCount = 63
    FixedCols = 0
    TabOrder = 2
    ColWidths = (
      122
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64
      64)
  end
  object CheckListBox1: TCheckListBox
    Left = 8
    Top = 48
    Width = 162
    Height = 542
    ItemHeight = 15
    TabOrder = 3
  end
  object btnSelectAll: TButton
    Left = 8
    Top = 592
    Width = 75
    Height = 25
    Caption = #51204#52404' '#49440#53469
    TabOrder = 4
    OnClick = btnSelectAllClick
  end
  object btnUnselectAll: TButton
    Left = 89
    Top = 592
    Width = 75
    Height = 25
    Caption = #51204#52404' '#54644#51228
    TabOrder = 5
    OnClick = btnUnselectAllClick
  end
  object Button2: TButton
    Left = 184
    Top = 600
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 6
    OnClick = Button2Click
  end
end
