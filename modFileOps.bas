Attribute VB_Name = "modFileOps"
Option Explicit

Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAborted As Boolean
  hNameMaps As Long
  sProgress As String
End Type

Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10

Public Declare Function SHFileOperation Lib "shell32.dll" Alias _
        "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Function GetFileName(Path As String, Ext As Boolean) As String
    
    Dim Cnt As Integer, Cnt2 As Integer
    Cnt = 1
    DoEvents
    Do Until Mid(Path, Len(Path) - Cnt, 1) = "\"
        Cnt = Cnt + 1
    Loop

    If Ext Then
        GetFileName = Mid(Path, Len(Path) - Cnt + 1)
    Else
        Cnt2 = 0
        DoEvents
        Do Until Mid(Path, Len(Path) - Cnt2, 1) = "." Or Cnt2 >= Len(Path)
            Cnt2 = Cnt2 + 1
        Loop
        If Cnt2 >= Len(Path) Then Cnt2 = 0
        GetFileName = Mid(Path, Len(Path) - Cnt + 1, Cnt - Cnt2 - 1)
    End If
    
End Function

Public Function IsRMDriveReady(sDrive As String) As Boolean

  'do a Dir on the drive.
  
   On Error Resume Next
   IsRMDriveReady = Dir(sDrive) <> ""
   On Local Error GoTo 0
   
End Function

