object Form1: TForm1
  Left = 201
  Top = 128
  Width = 544
  Height = 375
  Caption = 'about:blank'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object WebBrowser1: TWebBrowser
    Left = 0
    Top = 57
    Width = 536
    Height = 272
    Align = alClient
    TabOrder = 0
    OnDownloadBegin = WebBrowser1DownloadBegin
    OnDownloadComplete = WebBrowser1DownloadComplete
    OnNewWindow2 = WebBrowser1NewWindow2
    ControlData = {
      4C000000663700001D1C00000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 0
    Width = 536
    Height = 25
    ButtonHeight = 19
    ButtonWidth = 36
    Caption = 'ToolBar1'
    Flat = True
    List = True
    ShowCaptions = True
    TabOrder = 1
    object ToolButton1: TToolButton
      Left = 0
      Top = 0
      Caption = #21518#36864
      ImageIndex = 0
      OnClick = ToolButton1Click
    end
    object ToolButton2: TToolButton
      Left = 36
      Top = 0
      Caption = #21069#36827
      ImageIndex = 1
      OnClick = ToolButton2Click
    end
    object ToolButton3: TToolButton
      Left = 72
      Top = 0
      Width = 10
      ImageIndex = 5
      ParentShowHint = False
      ShowHint = False
      Style = tbsSeparator
    end
    object ToolButton4: TToolButton
      Left = 82
      Top = 0
      Caption = #20572#27490
      ImageIndex = 2
      OnClick = ToolButton4Click
    end
    object ToolButton5: TToolButton
      Left = 118
      Top = 0
      Caption = #21047#26032
      ImageIndex = 3
      OnClick = ToolButton5Click
    end
    object ToolButton6: TToolButton
      Left = 154
      Top = 0
      Caption = #20027#39029
      ImageIndex = 4
      OnClick = ToolButton6Click
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 25
    Width = 536
    Height = 32
    Align = alTop
    TabOrder = 2
    object Label1: TLabel
      Left = 16
      Top = 8
      Width = 36
      Height = 13
      Caption = #22320#22336#65306
    end
    object Edit1: TEdit
      Left = 56
      Top = 5
      Width = 217
      Height = 21
      TabOrder = 0
      Text = 'http://lshoe/manage'
      OnEnter = BitBtn1Click
    end
    object BitBtn1: TBitBtn
      Left = 280
      Top = 8
      Width = 73
      Height = 21
      Cursor = crIBeam
      Caption = #36716#21040
      Default = True
      TabOrder = 1
      OnClick = BitBtn1Click
      OnEnter = BitBtn1Click
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 329
    Width = 536
    Height = 19
    Panels = <
      item
        Width = 50
      end>
    SimplePanel = False
  end
end
