VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search & Replace"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWhole 
      Caption         =   "Match Whole Word"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "Close"
      Height          =   315
      Left            =   6120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Next"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match Case"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace With"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search For"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position As Long
Private Sub CloseButton_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub FindButton_Click()
    On Error Resume Next
    Dim FindFlags As Integer
    On Error GoTo error_2
    Position = 0
    FindFlags = chkMatch.Value * 4 + chkWhole.Value * 2
    Position = frmKB.docWord.Find(txtSearch.Text, Position + 1, , FindFlags)
    If Position >= 0 Then
        ReplaceButton.Enabled = True
        ReplaceAllButton.Enabled = True
    Else
        MsgBox "MyHelp could not find " & txtSearch.Text, , vbOKOnly + vbExclamation + vbApplicationModal, "Find"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    Err.Clear
    Exit Sub
error_2:
    Err.Clear
End Sub
Private Sub FindNextButton_Click()
    On Error Resume Next
    Dim FindFlags As Long
    On Error GoTo error
    FindFlags = chkMatch.Value * 4 + chkWhole.Value * 2
    Position = frmKB.docWord.Find(txtSearch.Text, Position + 1, , FindFlags)
    If Position > 0 Then
    Else
        MsgBox "MyHelp could not find " & txtSearch.Text, , vbOKOnly + vbExclamation + vbApplicationModal, "Find Next"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    Err.Clear
    Exit Sub
error:
    Err.Clear
End Sub
Private Sub replaceallbutton_Click()
    On Error Resume Next
    Dim FindFlags As Integer
    FindFlags = chkMatch.Value * 4 + chkWhole.Value * 2
    frmKB.docWord.SelText = txtReplace.Text
    Position = frmKB.docWord.Find(txtSearch.Text, Position + 1, , FindFlags)
    While Position > 0
        frmKB.docWord.SelText = txtReplace.Text
        Position = frmKB.docWord.Find(txtSearch.Text, Position + 1, , FindFlags)
        Err.Clear
    Wend
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
    MsgBox "Replacing Completed successfully", vbOKOnly + vbExclamation + vbApplicationModal, "Replace All"
    Err.Clear
End Sub
Private Sub replacebutton_Click()
    On Error Resume Next
    Dim FindFlags As Integer
    frmKB.docWord.SelText = txtReplace.Text
    FindFlags = chkMatch.Value * 4 + chkWhole.Value * 2
    Position = frmKB.docWord.Find(txtSearch.Text, Position + 1, , FindFlags)
    If Position > 0 Then
        frmKB.docWord.SetFocus
    Else
        MsgBox "MyHelp search string could not be found", vbOKOnly + vbExclamation + vbApplicationModal, "Replace"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    Err.Clear
End Sub
