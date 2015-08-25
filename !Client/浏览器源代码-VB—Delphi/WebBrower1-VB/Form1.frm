VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   Caption         =   "Webbrowse"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   7440
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7440
   Begin VB.CommandButton Command5 
      Caption         =   "停止"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "刷新"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "前进"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "后退"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Text            =   "http://lshoe/manage"
      Top             =   3840
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   5535
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

Private Sub Command1_Click()
On Error GoTo e
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
e:

End Sub

Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.GoBack       '后退

End Sub

Private Sub Command3_Click()
On Error Resume Next
WebBrowser1.GoForward     '前进

End Sub

Private Sub Command4_Click()
On Error Resume Next
WebBrowser1.Refresh         '停止

End Sub

Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.Stop     '刷新

End Sub

Private Sub Image5_Click()
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
