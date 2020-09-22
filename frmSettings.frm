VERSION 5.00
Begin VB.Form frmSettings 
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ok"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   600
      Max             =   2
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblLevelDes 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Block Levels:"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private levels As Collection

Private Sub Command1_Click()
Dim ApPath As String
ApPath = App.Path
If Right$(ApPath, 1) <> "\" Then ApPath = ApPath & "\"

ApPath = ApPath & "settings.inf"
If Dir$(ApPath) <> "" Then Kill ApPath

RestrictType = VScroll1.Value
Open ApPath For Binary As #1
Put #1, , RestrictType
Close #1

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ApPath As String


Set levels = New Collection
levels.Add "Strict Blocks. No Popups allowed"
levels.Add "Rep. blocks. Activates on too frequent popups"
levels.Add "No Blocks. Allows all popups"
lblLevelDes.Caption = VScroll1.Value & " - " & levels.Item(VScroll1.Value + 1)
ApPath = App.Path
If Right$(ApPath, 1) <> "\" Then ApPath = ApPath & "\"

ApPath = ApPath & "settings.inf"
If Dir$(ApPath) <> "" Then

Open ApPath For Binary As #1
Get #1, , RestrictType
Close #1
VScroll1.Value = RestrictType

End If
End Sub

Private Sub VScroll1_Change()
lblLevelDes.Caption = VScroll1.Value & " - " & levels.Item(VScroll1.Value + 1)

End Sub
