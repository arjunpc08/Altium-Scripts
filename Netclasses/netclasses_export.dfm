object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 583
  ClientWidth = 1087
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 16
  object Label2: TLabel
    Left = 416
    Top = 32
    Width = 19
    Height = 16
    Caption = 'List'
  end
  object Button1: TButton
    Left = 72
    Top = 88
    Width = 200
    Height = 48
    Caption = 'Populate List'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 72
    Top = 160
    Width = 200
    Height = 48
    Caption = 'Export to Excel'
    TabOrder = 1
    OnClick = Button2Click
  end
  object ListBox1: TListBox
    Left = 416
    Top = 72
    Width = 224
    Height = 240
    TabOrder = 2
    OnDblClick = ListBox1DblClick
  end
end
