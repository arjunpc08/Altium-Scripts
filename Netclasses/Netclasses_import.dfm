object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 491
  ClientWidth = 1023
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClick = Button1Click
  PixelsPerInch = 120
  TextHeight = 16
  object Button1: TButton
    Left = 64
    Top = 352
    Width = 200
    Height = 48
    Caption = 'Import to Board'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button4: TButton
    Left = 64
    Top = 184
    Width = 200
    Height = 40
    Caption = 'Clear List'
    TabOrder = 1
    OnClick = Button4Click
  end
  object Button3: TButton
    Left = 64
    Top = 104
    Width = 200
    Height = 48
    Caption = 'Import from Excel'
    TabOrder = 2
    OnClick = Button3Click
  end
  object ListBox2: TListBox
    Left = 344
    Top = 72
    Width = 232
    Height = 192
    TabOrder = 3
  end
  object ProgressBar1: TProgressBar
    Left = 304
    Top = 352
    Width = 384
    Height = 48
    TabOrder = 4
  end
end
