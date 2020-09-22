VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditTopic 
   Caption         =   "Topic Properties"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7455
   ClipControls    =   0   'False
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
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox fraTopic 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   7215
      TabIndex        =   18
      Top             =   480
      Width           =   7215
      Begin VB.CheckBox chkPopup 
         Caption         =   "Popup Page"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "This topic will be displayed as a popup"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton DeleteFootNote 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6000
         Width           =   255
      End
      Begin VB.CommandButton DeleteMacro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton AddFootNote 
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5640
         Width           =   255
      End
      Begin VB.CommandButton AddMacro 
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox chkBrowseSequence 
         Caption         =   "Browse Sequence (recommended)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton DeleteKeyword 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton AddKeyword 
         Height          =   315
         Left            =   6840
         Picture         =   "frmEditTopic.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   255
      End
      Begin MSComctlLib.ListView lstKeywords 
         Height          =   1695
         Left            =   1560
         TabIndex        =   5
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Keyword(s)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtContextNumber 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtContextString 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
      Begin MSComctlLib.ListView lstMacros 
         Height          =   1695
         Left            =   1560
         TabIndex        =   6
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Keyword(s)"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstFootNotes 
         Height          =   1695
         Left            =   1560
         TabIndex        =   7
         Top             =   5640
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Foot Note"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "A-Keywords"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   5640
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Macro(s)"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "K-Keyword(s)"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Context Number"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Context String"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Label lblTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTopic"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topic"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmEditTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrAdd As String
Private iAnswer As Integer
Private Sub AddFootNote_Click()
    On Error Resume Next
    StrAdd = InputBox("Please type in the a-keywords to add below." & vbCr & "You can separate the a-keywords with a semicolon:", "Add A-Keyword(s)", Replace$(txtTitle.Text, " ", ";", , , vbTextCompare))
    If Len(StrAdd) = 0 Then Exit Sub
    AddKeyWords lstFootNotes, StrAdd
    Err.Clear
End Sub
Private Sub AddKeyword_Click()
    On Error Resume Next
    StrAdd = InputBox("Please type in the k-keywords to add below." & vbCr & "You can separate the k-keywords with a semicolon:", "Add K-Keyword(s)", Replace$(txtTitle.Text, " ", ";", , , vbTextCompare))
    If Len(StrAdd) = 0 Then Exit Sub
    AddKeyWords lstKeywords, StrAdd
    Err.Clear
End Sub
Private Sub AddMacro_Click()
    On Error Resume Next
    StrAdd = InputBox("Please type in the macro to add below.", "Add Macro")
    If Len(StrAdd) = 0 Then Exit Sub
    If LstViewFindItem(lstMacros, StrAdd, search_Text, search_Whole) = 0 Then
        lstMacros.ListItems.Add , , StrAdd
    End If
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    If boolIsBlank(txtTitle, "topic title") = True Then Exit Sub
    If boolIsBlank(txtContextString, "topic context string") = True Then Exit Sub
    If boolIsBlank(txtContextNumber, "topic context number") = True Then Exit Sub
    If lstKeywords.ListItems.Count = 0 Then
        AddKeyWords lstKeywords, txtTitle.Text
    End If
    Dim rRecord(1 To 10) As String
    Dim xLine() As String
    Dim tPos As Long
    rRecord(1) = lblTopic.Caption                                'FullPath
    rRecord(2) = txtTitle.Text                                   'Title
    rRecord(3) = Context_Validate(txtContextString.Text)           'Context
    rRecord(4) = Val(txtContextNumber.Text)                      'Number
    rRecord(5) = chkBrowseSequence.Value                         'Browse
    LstViewRowsToMV lstKeywords, xLine, VM
    rRecord(6) = MvFromArray(xLine, FM)
    LstViewRowsToMV lstMacros, xLine, VM
    rRecord(7) = MvFromArray(xLine, FM)
    LstViewRowsToMV lstFootNotes, xLine, VM
    rRecord(8) = MvFromArray(xLine, FM)
    rRecord(10) = chkPopup.Value
    Dao_WriteRecordArray sProjDb, "Properties", "FullPath", rRecord(1), PropertiesFlds, rRecord
    DAO.DBEngine.Idle
    If chkPopup.Value = 1 Then
        tPos = TreeViewSearchPath(frmKB.treeDms, lblTopic.Caption)
        If tPos > 0 Then
            frmKB.treeDms.Nodes(tPos).Image = "leaf"
            frmKB.treeDms.Nodes(tPos).SelectedImage = "leaf"
            frmKB.treeChanged = True
        End If
    End If
    Unload Me
    Err.Clear
End Sub
Private Sub DeleteFootNote_Click()
    On Error Resume Next
    iAnswer = MsgBox("Are you sure that you want to delete the checked footnote(s).", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Delete")
    If iAnswer = vbNo Then Exit Sub
    LstViewRemoveChecked lstFootNotes, True
    Err.Clear
End Sub
Private Sub DeleteKeyword_Click()
    On Error Resume Next
    iAnswer = MsgBox("Are you sure that you want to delete the checked keyword(s).", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Delete")
    If iAnswer = vbNo Then Exit Sub
    LstViewRemoveChecked lstKeywords, True
    Err.Clear
End Sub
Private Sub DeleteMacro_Click()
    On Error Resume Next
    iAnswer = MsgBox("Are you sure that you want to delete the checked macros(s).", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Delete")
    If iAnswer = vbNo Then Exit Sub
    LstViewRemoveChecked lstMacros, True
    Err.Clear
End Sub
Private Sub lstFootNotes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    StrAdd = LstViewCheckedToMV(lstFootNotes, 1)
    If Len(StrAdd) = 0 Then
        DeleteFootNote.Enabled = False
    Else
        DeleteFootNote.Enabled = True
    End If
    Err.Clear
End Sub
Private Sub lstKeywords_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    StrAdd = LstViewCheckedToMV(lstKeywords, 1)
    If Len(StrAdd) = 0 Then
        DeleteKeyword.Enabled = False
    Else
        DeleteKeyword.Enabled = True
    End If
    Err.Clear
End Sub
Private Sub lstMacros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    StrAdd = LstViewCheckedToMV(lstMacros, 1)
    If Len(StrAdd) = 0 Then
        DeleteMacro.Enabled = False
    Else
        DeleteMacro.Enabled = True
    End If
    Err.Clear
End Sub
Private Sub txtTitle_Change()
    On Error Resume Next
    txtContextString.Text = Context_Validate(txtTitle.Text)
    Err.Clear
End Sub
Public Sub ReadTopicProperties(strPath As String)
    On Error Resume Next
    Dim rRecord() As String
    rRecord = Dao_ReadRecordToArray(sProjDb, "Properties", "FullPath", strPath, PropertiesFlds)
    ReDim Preserve rRecord(10)
    txtTitle.Text = rRecord(2)                   'Title
    txtContextString.Text = rRecord(3)           'Context
    txtContextNumber.Text = rRecord(4)           'Number
    chkBrowseSequence.Value = Val(rRecord(5))    'Browse
    LstViewFromMv lstKeywords, rRecord(6), FM
    LstViewFromMv lstMacros, rRecord(7), FM
    LstViewFromMv lstFootNotes, rRecord(8), FM
    chkPopup.Value = Val(rRecord(10))
    Err.Clear
End Sub
Private Sub AddKeyWords(lstKeywords As ListView, ByVal StrKeywords As String)
    On Error Resume Next
    Dim spKeywords() As String
    Dim spTot As Integer
    Dim spCnt As Integer
    StrKeywords = Replace$(StrKeywords, " ", ";")
    spKeywords = Split(StrKeywords, ";")
    spTot = UBound(spKeywords)
    For spCnt = 0 To spTot
        StrAdd = Trim$(spKeywords(spCnt))
        If Len(StrAdd) = 0 Then GoTo NextKeyWord
        If LstViewFindItem(lstKeywords, StrAdd, search_Text, search_Whole) = 0 Then
            lstKeywords.ListItems.Add , , StrAdd
        End If
NextKeyWord:
        Err.Clear
    Next
    Erase spKeywords
    Err.Clear
End Sub
Public Function LstViewFindItem(lstView As ListView, ByVal StrSearch As String, Optional ByVal SearchWhere As FindWhere = search_Text, Optional SearchItemType As SearchType = search_Whole) As Long
    On Error Resume Next
    Dim itmFound As ListItem
    LstViewFindItem = 0
    Set itmFound = lstView.FindItem(StrSearch, SearchWhere, , SearchItemType)
    If TypeName(itmFound) = "Nothing" Then
        Err.Clear
        Exit Function
    End If
    LstViewFindItem = CLng(itmFound.Index)
    Set itmFound = Nothing
    Err.Clear
End Function
Public Sub Dao_WriteRecordArray(ByVal Dbase As String, ByVal TableName As String, ByVal TableKey As String, ByVal ValuetoSeek As String, FieldsToRead As Variant, FieldsToWrite As Variant, Optional ByVal Overwrite As Boolean = True)
    On Error Resume Next
    If Len(ValuetoSeek) = 0 Then Exit Sub
    Dim adoC As DAO.Database
    Dim adoRs As DAO.Recordset
    Dim spTot As Integer
    Dim spCnt As Integer
    Dim spFld As String
    Set adoC = DAO.OpenDatabase(Dbase)
    Set adoRs = adoC.OpenRecordset(TableName)
    adoRs.Index = TableKey
    adoRs.Seek "=", ValuetoSeek
    spTot = UBound(FieldsToRead)
    Select Case adoRs.NoMatch
    Case True
        adoRs.AddNew
        dbConvertValue adoRs.Fields(TableKey), ValuetoSeek
        For spCnt = 1 To spTot
            spFld = FieldsToRead(spCnt)
            dbConvertValue adoRs.Fields(spFld), FieldsToWrite(spCnt)
            Err.Clear
        Next
        dbConvertValue adoRs.Fields(TableKey), ValuetoSeek
        adoRs.Update
    Case Else
        Select Case Overwrite
        Case False
            adoRs.AddNew
            dbConvertValue adoRs.Fields(TableKey), ValuetoSeek
            For spCnt = 1 To spTot
                spFld = FieldsToRead(spCnt)
                dbConvertValue adoRs.Fields(spFld), FieldsToWrite(spCnt)
                Err.Clear
            Next
            dbConvertValue adoRs.Fields(TableKey), ValuetoSeek
            adoRs.Update
        Case True
            adoRs.Edit
            For spCnt = 1 To spTot
                spFld = FieldsToRead(spCnt)
                dbConvertValue adoRs.Fields(spFld), FieldsToWrite(spCnt)
                Err.Clear
            Next
            dbConvertValue adoRs.Fields(TableKey), ValuetoSeek
            adoRs.Update
        End Select
    End Select
    adoRs.Close
    adoC.Close
    Set adoC = Nothing
    Set adoRs = Nothing
    Err.Clear
End Sub

