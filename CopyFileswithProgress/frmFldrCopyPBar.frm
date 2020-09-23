VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFldrCopyPBar 
   Caption         =   "Copy (only) Files in a Folder"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox DDir 
      Height          =   1890
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.DirListBox SDir 
      Height          =   1890
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.DriveListBox DDrive 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox SDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   240
      Top             =   3840
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Destination :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Source :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
End
Attribute VB_Name = "frmFldrCopyPBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SourceFolder As String, DestFolder As String

Private Sub DDir_Change()
    DestFolder = Me.DDir.Path
    If Len(DestFolder) <= 3 Then DestFolder = Left$(DestFolder, 2) 'if say DestFolder=f:\ then f:
End Sub

Private Sub DDrive_Change()
    Me.DDir.Path = Me.DDrive.Drive
End Sub

Private Sub SDir_Change()
    SourceFolder = Me.SDir.Path
End Sub

Private Sub SDrive_Change()
    Me.SDir.Path = Me.SDrive.Drive
End Sub

Private Sub cmdcopy_Click()
    Dim AnsYN As String
    If Len(DestFolder) <= 3 Then
        MsgBox "Can't copy to Root directory."
        Exit Sub
    End If
    If SourceFolder = Left$(DestFolder, Len(SourceFolder)) Then
        MsgBox "Please select destination folder which is outside source folder."
        'otherwise it will go in infinite loop as source will go on increasing.
        Exit Sub
    End If
    AnsYN = MsgBox("Copy Folder " & SourceFolder & " to " & DestFolder, vbYesNo + vbQuestion, "Confirm")
    If AnsYN = vbNo Then Exit Sub
    Dim SourceBytes As Long
    Dim myFSO, f
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set f = myFSO.GetFolder(SourceFolder)
    Me.ProgressBar1.Max = myFSO.GetFolder(SourceFolder).Size
    SourceBytes = myFSO.GetFolder(SourceFolder).Size
    Me.ProgressBar1.Value = 0
    Timer1.Enabled = False
    For Each fileP In f.Files
        FileCopy fileP, DestFolder & "\" & Right$(fileP, Len(fileP) - InStrRev(fileP, "\"))
        Label1.Caption = myFSO.GetFolder(DestFolder).Size & " Bytes out of " & SourceBytes & " Bytes copied."
        Me.ProgressBar1.Value = Val(Me.Label1.Caption)
        Timer1.Enabled = True
        Timer1_Timer
        Me.Refresh
    Next
    MsgBox "Copying completed." 'to show that copying is complete.
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

