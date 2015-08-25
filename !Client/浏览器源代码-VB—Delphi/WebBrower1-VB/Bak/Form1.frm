VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Webbrowse"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":1272
   ScaleHeight     =   7110
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   5520
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   2655
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5895
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   360
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   360
      X2              =   120
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   165
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "百度搜索："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   795
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   720
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   840
      X2              =   1080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   2280
      X2              =   1560
      Y1              =   2400
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   1560
      X2              =   2280
      Y1              =   2400
      Y2              =   3120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1440
      Top             =   2280
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   960
      Left            =   1440
      ToolTipText     =   "停止"
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   6600
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   360
      Picture         =   "Form1.frx":126FF
      ToolTipText     =   "刷新"
      Top             =   2280
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   1440
      Picture         =   "Form1.frx":14076
      ToolTipText     =   "前进"
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   240
      Picture         =   "Form1.frx":165E8
      ToolTipText     =   "后退"
      Top             =   1320
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
'转到指定网址
    WebBrowser1.Navigate Combo1.Text
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim existed As Boolean
    If KeyCode = 13 Then       '若按回车键则将浏览器导航到指定网址
        If Left(Combo1.Text, 7) <> "http://" Then
           Combo1.Text = "http://" + Combo1.Text
        End If
        WebBrowser1.Navigate Combo1.Text
        For i = 0 To Combo1.ListCount - 1
           If Combo1.List(i) = Combo1.Text Then
           existed = True
           Exit For
           Else
           existed = False
           End If
        Next
        If Not existed Then
           Combo1.AddItem (Combo1.Text)
        End If
    End If
End Sub


Private Sub Combo2_Click()
Dim url As String
url = "http://www1.baidu.com/baidu?tn=baidu&cl=3&word=" & Combo2.Text
WebBrowser1.Navigate url
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Combo2_Click
  Combo2.AddItem Combo2.Text
End If
End Sub


Private Sub Image1_Click()
On Error Resume Next
WebBrowser1.GoBack       '后退

End Sub

Private Sub Image2_Click()
On Error Resume Next
WebBrowser1.GoForward     '前进
End Sub

Private Sub Image3_Click()
On Error Resume Next
WebBrowser1.Refresh         '停止
End Sub

Private Sub Image4_Click()
On Error Resume Next
WebBrowser1.Stop     '刷新
End Sub

Private Sub Image5_Click()
Me.WindowState = 1
End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub Label3_Click()
On Error GoTo e
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
e:
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Time$
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Dim NewWindow As Form1
    Set NewWindow = New Form1
    NewWindow.Show
    Set ppDisp = NewWindow.WebBrowser1.Object
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Label2.Caption = Text
End Sub
