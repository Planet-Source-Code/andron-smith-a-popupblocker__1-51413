Attribute VB_Name = "actions"

'This popup blocker is good. But is strict, cause it will not allow any popups
'Now this is to make it a bit less strict.

Public Type User_Settings
    DoPrompt As Boolean 'Hold setting, so that the window will prompt and ask if they want to allow the popup
    AllowedSites As Collection ' A collection of sites, to allow.
    
    Log_Path As String ' A path in which to keep a log file.
End Type
Public RestrictType As Integer







    
