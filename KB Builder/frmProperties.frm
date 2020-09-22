VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Properties"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Done"
      Height          =   375
      Left            =   6000
      TabIndex        =   39
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4560
      TabIndex        =   38
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Default"
      Height          =   375
      Left            =   3120
      TabIndex        =   37
      ToolTipText     =   "Update"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Frame Frame6 
         Caption         =   "Font && Color Preview"
         Height          =   2055
         Left            =   2880
         TabIndex        =   12
         Top             =   120
         Width           =   3975
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   " Sample Text"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   " Headline"
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "BackColor"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   1275
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   1275
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Text"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label6 
            Caption         =   "Headline"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Font Color"
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2655
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   1275
            TabIndex        =   16
            Top             =   600
            Width           =   1335
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   1275
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Text"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Headline"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fonts"
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   6735
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   195
            Left            =   1920
            TabIndex        =   33
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   255
            Left            =   6240
            TabIndex        =   32
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "10"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Tahoma"
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "12"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "Tahoma"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Size"
            Height          =   195
            Left            =   3720
            TabIndex        =   30
            Top             =   960
            Width           =   285
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Text"
            Height          =   195
            Left            =   3720
            TabIndex        =   28
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Size"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Headline"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   6480
         TabIndex        =   36
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtCompiler 
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Text            =   "CompilerLocation"
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Made with Knowledge Base Builder"
         Top             =   1200
         Width           =   5655
      End
      Begin VB.ComboBox cmbCompression 
         Height          =   315
         ItemData        =   "frmProperties.frx":0CCA
         Left            =   1200
         List            =   "frmProperties.frx":0CE3
         TabIndex        =   4
         Text            =   "32"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Text            =   "Knowledge Base Builder Project"
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compiler"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Author"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compression"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   300
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkBold_Click()
    On Error Resume Next
    If chkBold.Value = 1 Then
        Text2.FontBold = True
    Else
        Text2.FontBold = False
    End If
    Err.Clear
End Sub
Private Sub cmdOk_Click()
    On Error Resume Next
    Title = txtTitle.Text
    Compression = cmbCompression.Text
    Author = txtAuthor.Text
    HeadlineColor = Picture4.BackColor
    TextColor = Picture5.BackColor
    HeadlineBackColor = Picture6.BackColor
    TextBackColor = Picture7.BackColor
    FontHeadline = Text4.Text
    FontHeadlineSize = Text5.Text
    FontText = Text6.Text
    FontTextSize = Text7.Text
    If chkBold.Value = 1 Then
        FontHeadlineBold = 1
    Else
        FontHeadlineBold = 0
    End If
    CompilerLocation = txtCompiler.Text
    Save_Settings
    Unload Me
    Err.Clear
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    Err.Clear
End Sub
Private Sub cmdReset_Click()
    On Error Resume Next
    HeadlineColor = &H0&
    TextColor = &H0&
    HeadlineBackColor = &HC0FFFF
    TextBackColor = &HC0FFFF
    FontHeadline = "Tahoma"
    FontText = "Tahoma"
    FontHeadlineSize = 12
    FontTextSize = 10
    FontHeadlineBold = 1
    ShowProperties
    Err.Clear
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    Dim tmpString As String
    tmpString = DialogOpen(StringFileFilters, "Help Compiler", , "*.exe")
    If Len(tmpString) > 0 Then txtCompiler.Text = tmpString
    Err.Clear
End Sub
Private Sub Command4_Click()
    On Error Resume Next
    cd1.FontName = Text4.Text
    cd1.flags = cdlCFBoth
    cd1.FontSize = Text5.Text
    If chkBold.Value = 1 Then cd1.FontBold = True
    cd1.ShowFont
    Text4.Text = cd1.FontName
    Text4.FontSize = cd1.FontSize
    Text5.Text = cd1.FontSize
    Text2.Font = cd1.FontName
    Text2.FontSize = cd1.FontSize
    Err.Clear
End Sub
Private Sub Command5_Click()
    On Error Resume Next
    cd1.FontName = Text6.Text
    cd1.flags = cdlCFBoth
    cd1.FontSize = Text7.Text
    cd1.FontBold = False
    cd1.ShowFont
    Text6.Text = cd1.FontName
    Text6.FontSize = cd1.FontSize
    Text7.Text = cd1.FontSize
    Text3.Font = cd1.FontName
    Text3.FontSize = cd1.FontSize
    Err.Clear
End Sub
Private Sub Picture4_Click()
    On Error Resume Next
    Dim NewColor As Long
    cd1.ShowColor
    NewColor = cd1.Color
    If NewColor <> -1 Then
        Picture4.BackColor = NewColor
        Text2.ForeColor = NewColor
    Else
    End If
    Err.Clear
End Sub
Private Sub Picture5_Click()
    On Error Resume Next
    Dim NewColor As Long
    cd1.ShowColor
    NewColor = cd1.Color
    If NewColor <> -1 Then
        Picture5.BackColor = NewColor
        Text3.ForeColor = NewColor
    Else
    End If
    Err.Clear
End Sub
Private Sub Picture6_Click()
    On Error Resume Next
    Dim NewColor As Long
    cd1.ShowColor
    NewColor = cd1.Color
    If NewColor <> -1 Then
        Picture6.BackColor = NewColor
        Text2.BackColor = NewColor
    Else
    End If
    Err.Clear
End Sub
Private Sub Picture7_Click()
    On Error Resume Next
    Dim NewColor As Long
    cd1.ShowColor
    NewColor = cd1.Color
    If NewColor <> -1 Then
        Picture7.BackColor = NewColor
        Text3.BackColor = NewColor
    Else
    End If
    Err.Clear
End Sub
Private Sub TabStrip1_Click()
    On Error Resume Next
    If TabStrip1.Tabs.Item(1).Selected = True Then
        Frame1.Visible = True
        Frame2.Visible = False
        cmdReset.Visible = False
        Frame1.ZOrder 0
    End If
    If TabStrip1.Tabs.Item(2).Selected = True Then
        Frame1.Visible = False
        Frame2.Visible = True
        cmdReset.Visible = True
        Frame2.ZOrder 0
    End If
    Err.Clear
End Sub
Sub ShowProperties()
    On Error Resume Next
    If FontHeadline = "" Then FontHeadline = "Tahoma"
    If FontText = "" Then FontText = "Tahoma"
    If FontHeadlineBold = 0 Then FontHeadlineBold = 1
    If chkBold.Value = 1 Then Text2.FontBold = True
    txtTitle.Text = Title
    txtCompiler.Text = CompilerLocation
    cmbCompression.Text = Compression
    txtAuthor.Text = Author
    Picture4.BackColor = HeadlineColor
    Picture5.BackColor = TextColor
    Picture6.BackColor = HeadlineBackColor
    Picture7.BackColor = TextBackColor
    Text2.ForeColor = Picture4.BackColor
    Text2.Font = FontHeadline
    Text2.FontSize = FontHeadlineSize
    Text2.BackColor = HeadlineBackColor
    Text3.ForeColor = TextColor
    Text3.Font = FontText
    Text3.FontSize = FontTextSize
    Text3.BackColor = TextBackColor
    Text4.Text = FontHeadline
    Text4.FontSize = FontHeadlineSize
    Text5.Text = FontHeadlineSize
    Text6.Text = FontText
    Text6.FontSize = FontTextSize
    Text7.Text = FontTextSize
    chkBold.Value = FontHeadlineBold
    Err.Clear
End Sub
