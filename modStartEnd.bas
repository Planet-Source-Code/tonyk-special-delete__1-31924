Attribute VB_Name = "modStartEnd"
Option Explicit

Global mpcSaveCursor As MousePointerConstants

Public Sub Main()
    
    On Error Resume Next
    
    frmMain.Show
    
End Sub

Public Sub EndApp(Optional ByVal blnForce As Boolean = False)
    
    Dim i As Long
    
    On Error Resume Next
    
    For i = Forms.Count - 1 To 0 Step -1
    
        Unload Forms(i) ' Triggers QueryUnload and Form_Unload
         ' If we aren't in blnForce mode and the
         ' unload failed, stop the shutdown.
         Set Forms(i) = Nothing
         
         If Not blnForce Then
         
            If Forms.Count > i Then
               Exit Sub
            End If
            
         End If
         
     Next i
      ' If we are in blnForce mode OR all
      ' forms unloaded, close all files.
     If blnForce Or (Forms.Count = 0) Then Close
      ' If we are in blnForce mode AND all
      ' forms not unloaded, end.
     If blnForce Or (Forms.Count > 0) Then End
     
End Sub

Public Sub SetCursor()
    
    On Error Resume Next
    
    mpcSaveCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass

End Sub

Public Sub ResetCursor()

    On Error Resume Next
    
    Screen.MousePointer = mpcSaveCursor
    
End Sub

