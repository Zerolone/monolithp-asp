object Form1: TForm1
  Left = 216
  Top = 123
  Width = 420
  Height = 437
  Caption = 'hackgull'#27983#35272#22120
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Scaled = False
  PixelsPerInch = 96
  TextHeight = 13
  object StatusBar1: TStatusBar
    Left = 0
    Top = 372
    Width = 412
    Height = 19
    Panels = <
      item
        Text = 'SimplePlanel'
        Width = 50
      end>
    SimplePanel = False
  end
  object WebBrowser1: TWebBrowser
    Left = 1
    Top = 32
    Width = 408
    Height = 337
    TabOrder = 1
    ControlData = {
      4C0000002B2A0000D42200000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000E084CB00}
  end
  object CoolBar1: TCoolBar
    Left = 0
    Top = 0
    Width = 412
    Height = 25
    Bands = <
      item
        Control = ComboBox1
        ImageIndex = -1
        MinHeight = 21
        Text = #22320#22336#26639
        Width = 408
      end>
    PopupMenu = PopupMenu1
    object ComboBox1: TComboBox
      Left = 46
      Top = 0
      Width = 358
      Height = 21
      ItemHeight = 13
      TabOrder = 0
      OnKeyPress = ComboBox1KeyPress
    end
  end
  object MainMenu1: TMainMenu
    Left = 312
    Top = 112
    object N1: TMenuItem
      Caption = #25991#20214
      object N2: TMenuItem
        Caption = #25171#24320' (ctrl+0)'
        OnClick = N2Click
      end
      object N3: TMenuItem
        Caption = #20445#23384
      end
      object N4: TMenuItem
        Caption = #21478#23384#20026
      end
      object ctrlp1: TMenuItem
        Caption = #25171#21360' (ctrl+p)'
      end
      object N5: TMenuItem
        Caption = #36864#20986
        OnClick = N5Click
      end
    end
    object N6: TMenuItem
      Caption = #32534#36753
      object N7: TMenuItem
        Caption = #25764#28040
      end
      object N8: TMenuItem
        Caption = #21098#20999
      end
      object N9: TMenuItem
        Caption = #22797#21046
      end
      object N10: TMenuItem
        Caption = #31896#24086
      end
      object N11: TMenuItem
        Caption = #20840#36873
      end
    end
    object N12: TMenuItem
      Caption = #26597#30475
      object N13: TMenuItem
        Caption = #24037#20855#26639
      end
      object N14: TMenuItem
        Caption = #21047#26032
        OnClick = N14Click
      end
    end
    object N15: TMenuItem
      Caption = #24110#21161
      object N16: TMenuItem
        Caption = #20851#20110
      end
      object delphi1: TMenuItem
        Caption = 'delphi'
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #32593#39029#25991#20214
    Left = 176
    Top = 120
  end
  object PopupMenu1: TPopupMenu
    Left = 256
    Top = 120
    object N17: TMenuItem
      Caption = #25991#20214
    end
  end
end
