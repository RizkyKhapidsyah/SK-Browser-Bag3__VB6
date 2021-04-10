VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Browser"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   Icon            =   "browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1164
      ButtonWidth     =   953
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open file"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save file"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print Document"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   600
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Stop"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Go Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Go Home"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Forward"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go back"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   10560
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   6975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   10610
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
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "ADDRESS :"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "SaveAs"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnufullscreen 
         Caption         =   "FullScreen"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
End

End Sub

Private Sub Command2_Click()
cd1.CancelError = True
On Error GoTo cancelback
WebBrowser1.GoBack
Exit Sub
cancelback:
Exit Sub

End Sub

Private Sub Command3_Click()
cd1.CancelError = True
On Error GoTo forward
WebBrowser1.GoForward
Exit Sub
forward:
Exit Sub

End Sub

Private Sub Command4_Click()
cd1.CancelError = True
On Error GoTo home
WebBrowser1.GoHome
Exit Sub
home:
Exit Sub

End Sub

Private Sub Command5_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Command6_Click()
WebBrowser1.Stop
End Sub

Private Sub Command7_Click()
WebBrowser1.Refresh
End Sub

Private Sub mnufullscreen_Click()
WebBrowser1.FullScreen = True

End Sub

Private Sub mnuhelp_Click()
cd1.ShowHelp
MsgBox " Till now no help is provided to you. For details of the working principle of this Browser, please see the help provided my Microsoft Corporation in the Internet Explorer 5.0 ", vbOKOnly, "Browser Help"

End Sub

Private Sub mnuopen_Click()
cd1.CancelError = True
On Error GoTo cancelopen

cd1.Filter = "Htm files |*.htm| Html files |*.html| *.*|*.*|"
cd1.FilterIndex = 1
cd1.ShowOpen
WebBrowser1.Navigate cd1.FileName
Exit Sub
cancelopen:
Exit Sub

End Sub


Private Sub mnuprint_Click()
cd1.CancelError = True
On Error GoTo cancelprint
cd1.ShowPrinter
Exit Sub
cancelprint:
Exit Sub

End Sub

Private Sub mnuquit_Click()
Unload Me
End

End Sub


Private Sub mnusave_Click()
cd1.CancelError = True
On Error GoTo save
cd1.Filter = "Html files |*.html| Htm files |*.htm| "
cd1.FilterIndex = 1
cd1.ShowSave
WebBrowser1.Navigate cd1.FileName

Exit Sub
save:
Exit Sub

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
WebBrowser1.Navigate Text1.Text
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case Is = "Open"
mnuopen_Click
Case Is = "Save"
mnusave_Click
Case Is = "Print"
mnuprint_Click

End Select

End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

WebBrowser1.Offline = True
If (InStr(1, URL, "xxx") <> 0) Then
MsgBox " Access denied " & URL
Cancel = True
End If

End Sub



Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress > 0 Then
Label2.Caption = "Downloading....." & Progress & "/" & ProgressMax
Else
Label2.Caption = " Page downloded"
End If

End Sub
