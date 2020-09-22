VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Alert"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Settings"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer freqtester 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   5760
      Top             =   1560
   End
   Begin VB.Timer blocked 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5280
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Log"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H008080FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents IExplore As InternetExplorer
Attribute IExplore.VB_VarHelpID = -1
Private IsActive As Boolean

Private Sub blocked_Timer()
Unload frmBanner
blocked.Enabled = False

End Sub

Private Sub Command1_Click()
Dim LogFIle

X = Right$(App.EXEName, 4)

If Left$(X, 1) = "." Then
   LogFIle = Right$(App.EXEName, Len(App.EXEName) - 4)
Else
   LogFIle = App.EXEName
End If

LogFIle = LogFIle & ".log"
Open LogFIle For Output As #1
   For i = 0 To List1.ListCount
      Print #1, List1.List(i)
    Next i
Close #1

End Sub

Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Command3_Click()
Load frmSettings
frmSettings.Show
End Sub

Private Sub Form_Load()
Set IExplore = New InternetExplorer
If Command$ = "" Then
IExplore.GoHome
Else
   cmds = Right$(Command$, Len(Command$) - 1)
   cmds = Left$(cmds, Len(cmds) - 1)
   IExplore.Navigate cmds
End If

IExplore.Visible = True
Me.Hide
Dim ApPath As String
ApPath = App.Path
If Right$(ApPath, 1) <> "\" Then ApPath = ApPath & "\"

ApPath = ApPath & "settings.inf"
If Dir$(ApPath) <> "" Then

Open ApPath For Binary As #1
Get #1, , RestrictType
Close #1

End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide

End Sub

Private Sub freqtester_Timer()
IsActive = False
freqtester.Enabled = False

End Sub

Private Sub IExplore_NewWindow2(ppDisp As Object, Cancel As Boolean)
Select Case RestrictType
Case 0
    Cancel = True
    Load frmBanner
    frmBanner.lblmessage.Caption = "A POPUP HAS BEEN BLOCKED FROM " + IExplore.LocationName
    frmBanner.Top = 0
    frmBanner.Left = 0
    frmBanner.Width = Screen.Width
    frmBanner.Show
    blocked.Enabled = True
    List1.AddItem IExplore.LocationURL & " was blocked!"
Case 1
    If IsActive = False Then
       IsActive = True
       freqtester.Enabled = True
    Else
       Cancel = True
    Load frmBanner
    frmBanner.lblmessage.Caption = "Popups was recieved to frequently from: " + IExplore.LocationName + " the last 1 was blocked"
    frmBanner.Top = 0
    frmBanner.Left = 0
    frmBanner.Width = Screen.Width
    frmBanner.Show
    blocked.Enabled = True
    List1.AddItem IExplore.LocationURL & " was blocked!"
    End If
End Select


End Sub

Private Sub IExplore_OnQuit()

End
End Sub

Private Sub IExplore_TitleChange(ByVal Text As String)
'You can add coding here. to change the AppTitle to fit the same as the Iexplore title.

Me.Caption = Text

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 PopupMenu frmPopup.mnuRClick, , X, Y
End If

End Sub
