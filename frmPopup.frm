VERSION 5.00
Begin VB.Form frmPopup 
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuRClick 
      Caption         =   "&RightClickMenu"
      Begin VB.Menu mnuAPopup 
         Caption         =   "&Allow Popup"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
