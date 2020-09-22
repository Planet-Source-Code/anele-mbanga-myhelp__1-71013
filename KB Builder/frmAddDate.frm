VERSION 5.00
Begin VB.Form frmAddDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Date/Time"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "frmAddDate.frx":0000
      Left            =   120
      List            =   "frmAddDate.frx":0022
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1320
      X2              =   3840
      Y1              =   220
      Y2              =   220
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1340
      X2              =   3860
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Select a format:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDone_Click()
    On Error Resume Next
    'frmKB.topicChanged = True
    'frmKB.SaveTopic
    Unload Me
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    With frmKB.docWord
        Select Case List1.Text
        Case "1/2/03"
            .SelRTF = Format$(Date, "d/m/yy")
        Case "01/02/03"
            .SelRTF = Format$(Date, "dd/mm/yy")
        Case "01/02/2003"
            .SelRTF = Format$(Date, "dd/mm/yyyy")
        Case "1st February 2003"
            .SelRTF = Format$(Date, "d") & GetSuffix & " " & Format$(Date, "mmmm yyyy")
        Case "Monday"
            .SelRTF = Format$(Date, "dddd")
        Case "Monday 1st"
            .SelRTF = Format$(Date, "dddd ") & Format$(Date, "d") & GetSuffix
        Case "Monday 1st February"
            .SelRTF = Format$(Date, "dddd ") & Format$(Date, "d") & GetSuffix & " " & Format$(Date, "mmmm")
        Case "Monday 1st Febuary 2003"
            .SelRTF = Format$(Date, "dddd ") & Format$(Date, "d") & GetSuffix & " " & Format$(Date, "mmmm yyyy")
        Case "1:30"
            .SelRTF = IIf(Hour(Time) > 12, Hour(Time) - 12, Hour(Time)) & ":" & Format$(Minute(Time), "00")
        Case "13:30"
            .SelRTF = Format$(Time, "hh:mm")
        End Select
    End With
    Err.Clear
End Sub
Private Function GetSuffix() As String
    On Error Resume Next
    Dim Suffix As String
    Select Case Day(Date)
    Case "11", "12", "13"
        Suffix = "th"
    Case Else
        Select Case Right$(Day(Date), 1)
        Case "1"
            Suffix = "st"
        Case "2"
            Suffix = "nd"
        Case "3"
            Suffix = "rd"
        Case Else
            Suffix = "th"
        End Select
    End Select
    GetSuffix = Suffix
    Err.Clear
End Function
Private Sub List1_DblClick()
    On Error Resume Next
    cmdOk_Click
    Err.Clear
End Sub
