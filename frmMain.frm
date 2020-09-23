VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
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
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2400
      MultiSelect     =   1  'Simple
      System          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
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
Dim strDFile() As String

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

Private Sub Command1_Click()
    Form1.Show
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
    
    If lngDType = 2 Then 'If it's removeable media
ChkRM:
    blnSuccess = IsRMDriveReady(strDrive) 'Check for inserted media in the drive.
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
    
    If Dir$(strDir, vbDirectory) = "" Then 'Check to see if Restore dir is there
        MkDir (strDir)                     'Create it if not
    End If
    
    strFName = GetFileName(strDFile, True)
    ' The following if statement just looks for the backslash
    ' at the end of the source path and adds it if it's not there
    If Right(strFPath, 1) <> "\" Then
        strSource = strFPath & "\" & strFName
    Else
        strSource = strFPath & strFName
    End If
    'Now set the target path to the restore folder
    strTarget = strDir & "\" & strFName
    'Check to see if it media is removeable i.e. floppy drive, zip etc.
    'If removeable then copy to restore folder and send to recycle bin from there
    If lngDType = 2 Then
        FileCopy strSource, strTarget
        SendToRecycle strDir & "\" & strFName, lngAFlag
    Else
        'Not removeable so just send directly to recycle bin from source
        SendToRecycle strSource, lngAFlag
    End If
    
End Sub

Private Sub cmdDelete_Click()

    On Error Resume Next
    
    Dim i As Integer
    Dim c As Integer
    Dim x As Integer
    SetCursor
    
    strFPath = Dir1.Path
    
    For i = 0 To File1.ListCount - 1
        'This looks at all the selected files
        If File1.Selected(i) Then
            c = c + 1
            ReDim Preserve strDFile(1 To c)
            'Again just looking for the backslash and adding it
            'where necessary
            If Right$(strFPath, 1) <> "\" Then
                strDFile(c) = strFPath & "\" & File1.List(i)
            Else
                strDFile(c) = strFPath & File1.List(i)
            End If 'Right$
            lngAFlag = FOF_ALLOWUNDO 'Set delete flag to recycle bin
            Call BackupFile(strDFile(c))
            StatusBar1.Panels(1).Text = "Deleteing " & strDFile(c)
        End If
        
    Next
    
    If c = 0 Then Exit Sub 'nothing selected so exit sub
    'If removeable media then we now can delete from the media since
    'we have sent a copy to the recycle bin.
    If lngDType = 2 Then
        For x = 1 To c
            Kill strDFile(x)
        Next
    End If
    
    File1.Refresh
    cmdDelete.Enabled = File1.ListIndex > -1
    ResetCursor
    StatusBar1.Panels(1).Text = Now
    
End Sub

Private Sub File1_Click()
    cmdDelete.Enabled = File1.ListIndex > -1
End Sub

Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2 'Center the form to the screen
    Me.Caption = "Removeable Media Delete Utility"  'set for title
    cmdDelete.Enabled = False
    lngFHeight = Me.Height
    lngFWidth = Me.Width
    StatusBar1.Panels(1).Text = Now 'Set text of statusbar to todays date
    
End Sub

Private Sub Form_Resize()
    'The following code prevents resizeing but allows for minimizing
    'This will allow resizing below 600 x 2400 unfortunately but most
    'people would not try to resize that small so I think this works
    'ok.
    If Me.Height > 600 And Me.Width > 2400 Then
        Me.Height = lngFHeight
        Me.Width = lngFWidth
    End If
    
End Sub

Private Sub Form_Terminate()
    EndApp
End Sub
