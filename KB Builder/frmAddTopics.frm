VERSION 5.00
Begin VB.Form frmAddTopics 
   Caption         =   "Add Topic(s)"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7455
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
   ScaleHeight     =   8520
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox fraNode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   7215
      TabIndex        =   10
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Done"
         Height          =   375
         Left            =   5760
         TabIndex        =   20
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   7920
         Width           =   1335
      End
      Begin VB.PictureBox Frame1 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   0
         ScaleHeight     =   6615
         ScaleWidth      =   7215
         TabIndex        =   11
         Top             =   1200
         Width           =   7215
         Begin VB.ComboBox cboTopicLocation 
            Height          =   315
            ItemData        =   "frmAddTopics.frx":0000
            Left            =   5520
            List            =   "frmAddTopics.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   4440
            Width           =   1575
         End
         Begin VB.CheckBox chkGen 
            Caption         =   "Sequence Generator"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   4800
            Width           =   1935
         End
         Begin VB.OptionButton optLeaf 
            Caption         =   "Page"
            Height          =   255
            Left            =   1800
            TabIndex        =   2
            Top             =   4440
            Width           =   735
         End
         Begin VB.OptionButton optOpenBook 
            Caption         =   "Book"
            Height          =   255
            Left            =   960
            TabIndex        =   1
            Top             =   4440
            Width           =   855
         End
         Begin VB.TextBox txtTitles 
            Height          =   4125
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   0
            Top             =   120
            Width           =   6135
         End
         Begin VB.PictureBox Frame3 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   960
            ScaleHeight     =   1335
            ScaleWidth      =   6135
            TabIndex        =   12
            Top             =   5160
            Width           =   6135
            Begin VB.CommandButton cmdSequence 
               Caption         =   "Apply"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5280
               TabIndex        =   9
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtPrefix 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   5
               Top             =   0
               Width           =   5055
            End
            Begin VB.TextBox txtStart 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   6
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtEnd 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2280
               TabIndex        =   7
               Top             =   480
               Width           =   1215
            End
            Begin VB.CheckBox chkWord 
               Caption         =   "Spell Numbers"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   8
               ToolTipText     =   "Convert numbers to words"
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Topic Prefix"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   18
               Top             =   0
               Width           =   840
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Start/End"
               Enabled         =   0   'False
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   690
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Topic Relationship"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   3840
            TabIndex        =   21
            Top             =   4440
            Width           =   1290
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Titles *"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Image"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   4440
            Width           =   450
         End
         Begin VB.Image MainImage 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   2880
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   255
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Main Topic"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblMain 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "lblMain"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAddTopics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkGen_Click()
    On Error Resume Next
    If chkGen.Value = 0 Then
        txtPrefix.Enabled = False
        txtStart.Enabled = False
        txtEnd.Enabled = False
        chkWord.Enabled = False
        cmdSequence.Enabled = False
        lbl(4).Enabled = False
        lbl(1).Enabled = False
    Else
        txtPrefix.Enabled = True
        txtStart.Enabled = True
        txtEnd.Enabled = True
        chkWord.Enabled = True
        cmdSequence.Enabled = True
        lbl(4).Enabled = True
        lbl(1).Enabled = True
        txtPrefix.SetFocus
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
    Dim xImage As Integer
    xImage = xImage + IIf((optOpenBook.Value = True), 1, 0)
    xImage = xImage + IIf((optLeaf.Value = True), 1, 0)
    If boolIsBlank(txtTitles, "titles") = True Then Exit Sub
    If boolIsBlank(Me.cboTopicLocation, "topic location") = True Then Exit Sub
    If xImage = 0 Then
        MsgBox "Please select the type of image this title will be represented with.", vbOKOnly + vbExclamation + vbApplicationModal, "Image Error"
        optOpenBook.SetFocus
        Err.Clear
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim spTitles() As String
    Dim strTitle As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim strImage As String
    Dim strPath As String
    Dim strPaths As String
    If optOpenBook.Value = True Then strImage = "book"
    If optLeaf.Value = True Then strImage = "leaf"
    strPaths = ""
    spTitles = Split(txtTitles, vbNewLine, , vbTextCompare)
    spTot = UBound(spTitles)
    For spCnt = 0 To spTot
        strTitle = Topic_Validate(Trim$(spTitles(spCnt)))
        If Len(strTitle) = 0 Then GoTo NextTitle
        strPath = TreeView_AddNode(frmKB.treeDms, Tag, strTitle, strImage, Me.cboTopicLocation.Text)
        If Len(strPath) > 0 Then
            strPaths = strPaths & strPath & vbNewLine
        End If
NextTitle:
        Err.Clear
    Next
    SaveTopicsToDb strPaths
    TreeViewSaveToTable frmKB, frmKB.progBar, sProjDb, sProject, frmKB.treeDms
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    Unload Me
    Err.Clear
End Sub
Private Sub cmdSequence_Click()
    On Error Resume Next
    If boolIsBlank(txtPrefix, "topic prefix") = True Then Exit Sub
    If boolIsBlank(txtStart, "starting value") = True Then Exit Sub
    If boolIsBlank(txtEnd, "ending value") = True Then Exit Sub
    If Val(txtEnd.Text) < Val(txtStart.Text) Then
        MsgBox "The ending value cannot be less than the starting value.", vbOKOnly + vbExclamation + vbApplicationModal, "Starting/Ending Error"
        txtStart.SetFocus
        Err.Clear
        Exit Sub
    End If
    If Len(txtTitles.Text) = 0 Then
        txtTitles.Text = SequenceValues
    Else
        txtTitles.Text = txtTitles.Text & vbNewLine & SequenceValues
    End If
    Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FormMoveToNextControl KeyAscii
    Err.Clear
End Sub
Private Sub optLeaf_Click()
    On Error Resume Next
    If optLeaf.Value = True Then
        Set MainImage.Picture = frmKB.imgKB.ListImages(38).Picture
        MainImage.Tag = frmKB.imgKB.ListImages(38).Key
    End If
    Err.Clear
End Sub
Private Sub optOpenBook_Click()
    On Error Resume Next
    If optOpenBook.Value = True Then
        Set MainImage.Picture = frmKB.imgKB.ListImages(1).Picture
        MainImage.Tag = frmKB.imgKB.ListImages(1).Key
    End If
    Err.Clear
End Sub
Private Sub txtTitles_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtTitles
    Err.Clear
End Sub
Private Sub txtPrefix_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtPrefix
    Err.Clear
End Sub
Private Sub txtStart_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtStart
    Err.Clear
End Sub
Private Sub txtEnd_GotFocus()
    On Error Resume Next
    TextBoxHiLite txtEnd
    Err.Clear
End Sub
Private Sub txtTitles_Validate(Cancel As Boolean)
    On Error Resume Next
    txtTitles.Text = StringProperCase(txtTitles.Text)
    Err.Clear
End Sub
Private Sub txtPrefix_Validate(Cancel As Boolean)
    On Error Resume Next
    txtPrefix.Text = StringProperCase(txtPrefix.Text)
    Err.Clear
End Sub
Private Sub txtStart_Validate(Cancel As Boolean)
    On Error Resume Next
    txtStart.Text = StringProperCase(txtStart.Text)
    Err.Clear
End Sub
Private Sub txtEnd_Validate(Cancel As Boolean)
    On Error Resume Next
    txtEnd.Text = StringProperCase(txtEnd.Text)
    Err.Clear
End Sub
Sub FormMoveToNextControl(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
    Case 27   ' escape key
        SendKeys "+{TAB}"
        KeyAscii = 0
        DoEvents
    Case vbKeyReturn          ' catch return key
        Select Case TypeName(Screen.ActiveControl)
        Case "CheckBox", "ComboBox", "MaskEdBox", "OptionButton"
            SendKeys "{TAB}"      ' send tab which changes the element on form
            KeyAscii = 0
            DoEvents
        Case "TextBox"
            If Screen.ActiveControl.MultiLine = False Then
                SendKeys "{TAB}"      ' send tab which changes the element on form
                KeyAscii = 0
                DoEvents
            End If
        End Select
    End Select
    Err.Clear
End Sub
Private Function SequenceValues() As String
    On Error Resume Next
    Dim curTex As String
    Dim newTex As String
    Dim intLoop As Long
    Dim intWord As Long
    Dim intStart As Long
    Dim intEnd As Long
    Dim curNum As String
    Dim curRslt As String
    intWord = chkWord.Value
    intStart = Val(txtStart.Text)
    intEnd = Val(txtEnd.Text)
    curTex = txtPrefix.Text
    curRslt = ""
    For intLoop = intStart To intEnd
        curNum = CStr(intLoop)
        If intWord = 1 Then curNum = StringSpellNumber(curNum)
        newTex = curTex & " " & curNum
        curRslt = curRslt & newTex & vbNewLine
        Err.Clear
    Next
    SequenceValues = curRslt
    Err.Clear
End Function
Private Sub SaveTopicsToDb(strPaths As String)
    On Error Resume Next
    If Len(strPaths) = 0 Then Exit Sub
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    Dim strTitle As String
    Dim lngNew As Long
    spLine = Split(strPaths, vbNewLine, , vbTextCompare)
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Fullpath"
    spTot = UBound(spLine)
    For spCnt = 0 To spTot
        spStr = StringProperCase(Trim$(spLine(spCnt)))
        If Len(spStr) = 0 Then GoTo NextTopic
        strTitle = StringGetFileToken(spStr, "f")
        tb.Seek "=", spStr
        Select Case tb.NoMatch
        Case True
            lngNew = Val(dbNextOpenSequence(db, "Properties", "number"))
            tb.AddNew
            tb!Fullpath = spStr
            tb!Title = strTitle
            tb!Context = Context_Validate(MvFromMv(spStr, 2, , "\"))
            tb!Number = lngNew
            tb!browse = 1
            tb!Keywords = Keywords_Validate(spStr)
            tb!Macros = ""
            tb!Footnotes = ""
            tb!Contents = ""
            tb.Update
        End Select
NextTopic:
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub



