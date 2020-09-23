VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2400
      MultiSelect     =   1  'Simple
      System          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDriveType Lib "kernel32" Alias _
        "GetDriveTypeA" (ByVal nDrive As String) As Long

'Different Drive Types
'0 = "Unknown"
'1 = "No Root On Drive"
'2 = "Removable"
'3 = "Fixed"
'4 = "Network"
'5 = "CD-ROM"
'6 = "RAM Disk"

Dim lngAFlag As Long
Dim strFPath As String
Dim lngFHeight As Long
Dim lngFWidth As Long
Dim blnSuccess As Boolean
Dim strDrive As String
Dim lngDType As Long

Private Sub SendToRecycle(strFile As String, lngAFlag As Long)
    
    Dim SHFileOp As SHFILEOPSTRUCT
    
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = strFile
        .fFlags = lngAFlag Or FOF_NOCONFIRMATION
    End With
    
    On Error GoTo ErrHand
    SHFileOperation SHFileOp
    
    Exit Sub
    
ErrHand:
    MsgBox (Err.Number & "   " & Err.Source & vbCrLf & Err.Description)
    Resume Next
    
End Sub

Private Sub cmdCancel_Click()
    EndApp
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
        
    On Error Resume Next
    
    Dim strMsg As String
    Dim strStyle As String
    Dim strTitle As String
    Dim strResponse As String
    
    strDrive = Drive1.Drive
    lngDType = GetDriveType(strDrive)
    
    If lngDType = 2 Then
ChkRM:
    blnSuccess = IsRMDriveReady(strDrive)
        If blnSuccess Then
            Dir1.Path = strDrive
            File1.Pattern = "*.*"
        Else
            strMsg = "Please Insert Removeable Media Into Drive " & strDrive
            strStyle = vbOKCancel + vbExclamation + vbDefaultButton1
            strTitle = Me.Caption
            strResponse = MsgBox(strMsg, strStyle, strTitle)
            If strResponse = vbCancel Then
                Dir1.Path = App.Path
                Drive1.Drive = "C:"
            Else
                GoTo ChkRM
            End If 'strResponse
        End If 'blnSuccess
    End If 'Drive1.Drive
        
    Dir1.Path = Drive1.Drive
    
End Sub

Public Sub BackupFile(strDFile As String)
    
    On Error Resume Next
    
    Dim strDir As String
    Dim strFName As String
    Dim strSource As String
    Dim strTarget As String
    
    strDir = App.Path & "\Restore"
    
    If Dir$(strDir, vbDirectory) = "" Then
        MkDir (strDir)
    End If
    
    strFName = GetFileName(strDFile, True)
    
    If Right(strFPath, 1) <> "\" Then
        strSource = strFPath & "\" & strFName
    Else
        strSource = strFPath & strFName
    End If
    
    strTarget = strDir & "\" & strFName
    
    If lngDType = 2 Then
        FileCopy strSource, strTarget
        lngAFlag = FOF_ALLOWUNDO
        SendToRecycle strDir & "\" & strFName, lngAFlag
        Kill strSource
    Else
        SendToRecycle strSource, lngAFlag
    End If
    
End Sub

Private Sub cmdDelete_Click()

    On Error Resume Next
    
    Dim strDFile As String
    SetCursor
    
    strFPath = Dir1.Path
    strDFile = File1.List(File1.ListIndex)
    strDFile = strFPath & "\" & strDFile
    lngAFlag = 0&
    
    Call BackupFile(strDFile)
    
    File1.Refresh
    cmdDelete.Enabled = File1.ListIndex > -1
    
    ResetCursor
    
End Sub

Private Sub File1_Click()
    cmdDelete.Enabled = File1.ListIndex > -1
End Sub

Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    Me.Caption = "Removeable Media Delete Utility"
    cmdDelete.Enabled = False
    lngFHeight = Me.Height
    lngFWidth = Me.Width
    
End Sub

Private Sub Form_Resize()
    'The following code prevents resizeing but allows for minimizing
    If Me.Height > 600 And Me.Width > 2400 Then
        Me.Height = lngFHeight
        Me.Width = lngFWidth
    End If
    
End Sub

Private Sub Form_Terminate()
    EndApp
End Sub


