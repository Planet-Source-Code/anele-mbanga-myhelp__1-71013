VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJumpsPops 
   Caption         =   "Available Jumps & Popups"
   ClientHeight    =   8520
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CheckBox chkCheck 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstTopics 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Check all jumps / popups to delete"
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13785
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Topic Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Topic Context"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmJumpsPops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tChecked As Long
Private Function JumPopupText(ByVal StrLink As String) As String
    On Error Resume Next
    Dim cLine() As String
    cLine = Split(StrLink, "\ulnone\v ")
    StrLink = cLine(0)
    StrLink = Replace$(StrLink, "\uldb ", "")
    JumPopupText = Replace$(StrLink, "\ul ", "")
    Err.Clear
End Function
Private Sub chkCheck_Click()
    On Error Resume Next
    If chkCheck.Value = 1 Then
        LstViewCheckAll Me.lstTopics, True
    Else
        LstViewCheckAll Me.lstTopics, False
    End If
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Sub LoadJumpsPops(jCollection As Collection, pCollection As Collection)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nStr As String
    Dim cLine() As String
    Dim aLine(1 To 4) As String
    Dim nPos As Long
    Dim nOld As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim nTitle As String
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Context"
    lstTopics.ListItems.Clear
    nTot = jCollection.Count
    For nCnt = 1 To nTot
        nStr = jCollection(nCnt)
        nOld = nStr
        nStr = Replace$(nStr, "\uldb ", "")
        nStr = Replace$(nStr, "\v0", "")
        cLine = Split(nStr, "\ulnone\v ")
        nTitle = ""
        tb.Seek "=", cLine(1)
        If tb.NoMatch = False Then nTitle = StringProperCase(tb!Title & "")
        aLine(1) = "Jump"
        aLine(2) = cLine(0)
        aLine(3) = nTitle
        aLine(4) = cLine(1)
        nPos = LstViewUpdate(aLine, lstTopics, "")
        lstTopics.ListItems(nPos).Tag = nOld
        Err.Clear
    Next
    nTot = pCollection.Count
    For nCnt = 1 To nTot
        nStr = pCollection(nCnt)
        nOld = nStr
        nStr = Replace$(nStr, "\ul ", "")
        nStr = Replace$(nStr, "\v0", "")
        cLine = Split(nStr, "\ulnone\v ")
        nTitle = ""
        tb.Seek "=", cLine(1)
        If tb.NoMatch = False Then nTitle = StringProperCase(tb!Title & "")
        aLine(1) = "Popup"
        aLine(2) = cLine(0)
        aLine(3) = nTitle
        aLine(4) = cLine(1)
        nPos = LstViewUpdate(aLine, lstTopics, "")
        lstTopics.ListItems(nPos).Tag = nOld
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    Dim strChecked As String
    strChecked = LstViewCheckedToMV(lstTopics, 1, FM)
    tChecked = MvCount(strChecked, FM)
    If tChecked > 0 Then
        Dim intAns As Integer
        Dim StrLink As String
        Dim StrRTF As String
        Dim nTot As Long
        Dim nCnt As Long
        intAns = MyPrompt("You have opted to delete all the checked popups and jumps, are you sure?", "yn", "q", "Confirm Delete")
        If intAns = vbNo Then Exit Sub
        nTot = lstTopics.ListItems.Count
        For nCnt = 1 To nTot
            If lstTopics.ListItems(nCnt).Checked = False Then GoTo NextTopic
            DoEvents
            StrRTF = lstTopics.ListItems(nCnt).Tag
            StrLink = JumPopupText(StrRTF)
            frmKB.docWord.TextRTF = Replace$(frmKB.docWord.TextRTF, StrRTF, StrLink)
NextTopic:
            Err.Clear
        Next
        'frmKB.topicChanged = True
        'frmKB.SaveTopic
        frmKB.JumpsPopsStatus
    End If
    Unload Me
    Err.Clear
End Sub
Private Sub lstTopics_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If Item.Checked = True Then
        tChecked = tChecked + 1
    Else
        tChecked = tChecked - 1
    End If
    Err.Clear
End Sub
