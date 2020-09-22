VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJump 
   Caption         =   "Insert Jump"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10455
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   7680
      Width           =   1335
   End
   Begin VB.OptionButton optNonGreen 
      Caption         =   "Non-Green"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   1215
   End
   Begin VB.OptionButton optNoUnderline 
      Caption         =   "No Underline (not recommennded)"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox txtTopic 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   6840
      Width           =   3855
   End
   Begin VB.TextBox txtWindow 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   7200
      Width           =   8775
   End
   Begin VB.TextBox txtHelpFile 
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   6840
      Width           =   4935
   End
   Begin MSComctlLib.ListView lstTopics 
      Height          =   6255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Context String"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtJump 
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topic / Help File"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Topics"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Word(s)"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmJump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    Dim lItem As ListItem
    Dim jumpTo As String
    If boolIsBlank(txtJump, "selected jump or popup phrase") = True Then Exit Sub
    jumpTo = ""
    If Len(txtTopic.Text) = 0 Then
        If Len(txtHelpFile.Text) = 0 Then
            Set lItem = lstTopics.SelectedItem
            If TypeName(lItem) = "Nothing" Then
                MsgBox "Please select a topic to jump to to complete the process.", vbOKOnly + vbExclamation + vbApplicationModal, "Topic Error"
                Err.Clear
                Exit Sub
            Else
                jumpTo = lItem.SubItems(1)
            End If
        End If
    Else
        If Len(txtHelpFile.Text) <> 0 Then
            jumpTo = txtTopic.Text & "@" & txtHelpFile.Text
        End If
    End If
    If Len(jumpTo) = 0 Then
        MsgBox "Please select a topic to jump to or specify the topic to jump to in another help file.", vbOKOnly + vbExclamation + vbApplicationModal, "Topic Error"
        Err.Clear
        Exit Sub
    End If
    frmKB.docWord.SelRTF = JumpContext(frmKB.docWord.SelText, frmKB.docWord.SelRTF, jumpTo, Me.Tag, optNoUnderline.Value, optNonGreen.Value, txtWindow.Text)
    'frmKB.topicChanged = True
    'frmKB.SaveTopic
    frmKB.JumpsPopsStatus
    Unload Me
    Err.Clear
End Sub
