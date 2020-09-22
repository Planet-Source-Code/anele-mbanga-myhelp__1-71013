VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Insert Picture"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox imgPicture 
      AutoRedraw      =   -1  'True
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   7155
      TabIndex        =   8
      Top             =   2280
      Width           =   7215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      ToolTipText     =   "Browse for the picture file"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Path of the picture file"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin VB.CheckBox chkConvert 
         Caption         =   "Convert white pixels (16 Bit Images)"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "The white pixels with be converted to match background"
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton optRight 
         Caption         =   "Right Margin"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Align picture to the right"
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left Margin"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Align picture to the left"
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton optAsText 
         Caption         =   "Align as text character (show in same place as text)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   $"frmPicture.frx":0000
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   600
   End
   Begin VB.Image imgIncrease 
      Height          =   240
      Left            =   600
      MouseIcon       =   "frmPicture.frx":008C
      MousePointer    =   99  'Custom
      Picture         =   "frmPicture.frx":0396
      Stretch         =   -1  'True
      ToolTipText     =   "Increase picture size"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Image imgDecrease 
      Height          =   240
      Left            =   120
      MouseIcon       =   "frmPicture.frx":0920
      MousePointer    =   99  'Custom
      Picture         =   "frmPicture.frx":0C2A
      Stretch         =   -1  'True
      ToolTipText     =   "Decrease picture size"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture File Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1230
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub ClearAll()
    On Error Resume Next
    optAsText.Value = False
    optLeft.Value = False
    optRight.Value = False
    txtPath.Text = ""
    imgPicture.Picture = LoadPicture()
    chkConvert.Value = 0
    Err.Clear
End Sub
Private Sub cmdBrowse_Click()
    On Error GoTo ExitFind
    Dim oW As Long
    Dim oH As Long
    With frmKB.cd1
        .CancelError = True
        .DialogTitle = "Select Picture To Add..."
        .Filter = "Picture Files (*.bmp, *.jpg, *.gif, etc)|*.bmp;*.dib;*.jpeg;*.jpg;*.jpe;*.jfif;*.gif;*.tif;*.tiff;*.png"
        .ShowOpen
        DoEvents
        txtPath.Text = .FileName
        imgPicture.Picture = LoadPicture(.FileName)
        oW = imgPicture.Width
        oH = imgPicture.Height
        If oW > PictureWidth Then oW = PictureWidth
        If oH > PictureHeight Then oH = PictureHeight
        imgPicture.Height = oH
        imgPicture.Width = oW
        imgPicture.PaintPicture imgPicture, 0, 0, oW, oH
        imgPicture.Picture = imgPicture.Image
        imgPicture.Refresh
    End With
ExitFind:
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    Dim intS As Integer
    Dim sLink As String
    Dim sFile As String
    intS = 0
    intS = intS + IIf((Me.optAsText.Value = True), 1, 0)
    intS = intS + IIf((Me.optLeft.Value = True), 1, 0)
    intS = intS + IIf((Me.optRight.Value = True), 1, 0)
    If intS = 0 Then
        MsgBox "Please select a location for the picture.", vbOKOnly + vbExclamation + vbApplicationModal, "Picture Location Error"
        Err.Clear
        Exit Sub
    End If
    If boolIsBlank(txtPath, "picture path") = True Then Exit Sub
    DoEvents
    sFile = sProjPath & "\" & StringGetFileToken(txtPath.Text, "f")
    Call SavePicture(imgPicture.Picture, sFile)
    If optAsText.Value = True Then sLink = "{bmc"
    If optLeft.Value = True Then sLink = "{bml"
    If optRight.Value = True Then sLink = "{bmr"
    If chkConvert.Value = 1 Then sLink = sLink & "t"
    sLink = sLink & " " & sFile & "}"
    frmKB.docWord.SelText = sLink
    'frmKB.topicChanged = True
    'frmKB.SaveTopic
    Unload Me
    Err.Clear
End Sub
Private Sub imgDecrease_Click()
    On Error Resume Next
    ResizePreview txtPath.Text, -100, -100
    Err.Clear
End Sub
Private Sub imgIncrease_Click()
    On Error Resume Next
    ResizePreview txtPath.Text, 100, 100
    Err.Clear
End Sub
Private Sub ResizePreview(sFileName As String, iW As Integer, iH As Integer)
    On Error Resume Next
    imgPicture.Width = imgPicture.Width + iW
    imgPicture.Height = imgPicture.Height + iH
    imgPicture.Picture = LoadPicture(sFileName)
    imgPicture.PaintPicture imgPicture, 0, 0, imgPicture.Width, imgPicture.Height
    imgPicture.Picture = imgPicture.Image
    imgPicture.Refresh
    Err.Clear
End Sub

