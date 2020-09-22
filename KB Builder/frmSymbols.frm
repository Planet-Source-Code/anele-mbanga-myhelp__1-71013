VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Symbols"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
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
   HelpContextID   =   380
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Done"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ListBox lstSymbols 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   380
      Width           =   3855
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub lstSymbols_DblClick()
    On Error Resume Next
    ' Insert selected symbol to rtfText
    frmKB.docWord.SelText = Right$(lstSymbols.Text, 1)
    Err.Clear
End Sub
Private Sub cmdInsert_Click()
    On Error Resume Next
    ' Insert selected symbol to rtfText
    frmKB.docWord.SelText = Right$(lstSymbols.Text, 1)
    Err.Clear
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    'frmKB.topicChanged = True
    'frmKB.SaveTopic
    Unload Me
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim i As Integer
    'Set font name
    For i = 1 To 255
        ' Fills lstSymbols with Symbols
        If i < 10 Then
            lstSymbols.AddItem i & "     -  " & Chr$(i)
        ElseIf i < 100 Then
            lstSymbols.AddItem i & "   -  " & Chr$(i)
        Else
            lstSymbols.AddItem i & " -  " & Chr$(i)
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
