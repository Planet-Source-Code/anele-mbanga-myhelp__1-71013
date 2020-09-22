VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form HTMLEditor 
   Caption         =   "Nite Editor- HTML Editor"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6945
   HelpContextID   =   470
   Icon            =   "HtmlEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6945
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   480
      WhatsThisHelpID =   470
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"HtmlEditor.frx":038A
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":0404
            Key             =   ""
            Object.Tag             =   "&Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":07A0
            Key             =   ""
            Object.Tag             =   "&Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":0B3C
            Key             =   ""
            Object.Tag             =   "&Exit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":0ED8
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":1274
            Key             =   ""
            Object.Tag             =   "&Preview"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":1610
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":22EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":2688
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":2A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":2DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":315C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":34F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":3C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":3FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":4126
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":4280
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":43DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":4534
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":468E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":47E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":4942
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HtmlEditor.frx":4A9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   470
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   6945
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   75
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   30
         TabIndex        =   1
         Top             =   30
         WhatsThisHelpID =   470
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Description     =   "New"
               Object.ToolTipText     =   "New"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Description     =   "Open"
               Object.ToolTipText     =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Description     =   "Save"
               Object.ToolTipText     =   "Save"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Separator"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Description     =   "Left"
               Object.ToolTipText     =   "Left"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Description     =   "Center"
               Object.ToolTipText     =   "Center"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Description     =   "Right"
               Object.ToolTipText     =   "Right"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Description     =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Description     =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Description     =   "Underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close editor"
               ImageIndex      =   3
            EndProperty
         EndProperty
         MouseIcon       =   "HtmlEditor.frx":4BF6
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      HelpContextID   =   2010
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^G
      End
      Begin VB.Menu line27 
         Caption         =   "-"
         HelpContextID   =   2210
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTables 
      Caption         =   "&Tables"
      HelpContextID   =   2030
      Begin VB.Menu mnuCells 
         Caption         =   "Add the first column"
         Begin VB.Menu mnuColCells 
            Caption         =   "Background"
            Begin VB.Menu mnu1 
               Caption         =   "Black"
            End
            Begin VB.Menu mnu2 
               Caption         =   "Blue"
            End
            Begin VB.Menu mnu3 
               Caption         =   "Blue violet"
            End
            Begin VB.Menu mnu4 
               Caption         =   "Brown"
            End
            Begin VB.Menu mnu5 
               Caption         =   "Cyan"
            End
            Begin VB.Menu mnu6 
               Caption         =   "Dark browm"
            End
            Begin VB.Menu mnu7 
               Caption         =   "Dark green"
            End
            Begin VB.Menu mnu8 
               Caption         =   "Dark blue"
            End
            Begin VB.Menu mnu9 
               Caption         =   "Gold"
            End
            Begin VB.Menu mnu10 
               Caption         =   "Green"
            End
            Begin VB.Menu mnu11 
               Caption         =   "Magenta"
            End
            Begin VB.Menu mnu12 
               Caption         =   "Orange"
            End
            Begin VB.Menu mnu13 
               Caption         =   "Red"
            End
            Begin VB.Menu mnu14 
               Caption         =   "Tan"
            End
            Begin VB.Menu mnu15 
               Caption         =   "White"
            End
            Begin VB.Menu mnu16 
               Caption         =   "Yellow"
            End
         End
      End
      Begin VB.Menu mnuAddCol 
         Caption         =   "Add new column"
         Begin VB.Menu mnuColBac 
            Caption         =   "Backgound"
            Begin VB.Menu mnu1a 
               Caption         =   "Black"
            End
            Begin VB.Menu mnu2a 
               Caption         =   "Blue"
            End
            Begin VB.Menu mnu3a 
               Caption         =   "Blue violet"
            End
            Begin VB.Menu mnu4a 
               Caption         =   "Brown"
            End
            Begin VB.Menu mnu5a 
               Caption         =   "Cyan"
            End
            Begin VB.Menu mnu6a 
               Caption         =   "Dark browm"
            End
            Begin VB.Menu mnu7a 
               Caption         =   "Dark Green"
            End
            Begin VB.Menu mnu8a 
               Caption         =   "Dark blue"
            End
            Begin VB.Menu mnu9a 
               Caption         =   "Gold"
            End
            Begin VB.Menu mnu10a 
               Caption         =   "Green"
            End
            Begin VB.Menu mnu11a 
               Caption         =   "Magenta"
            End
            Begin VB.Menu mnu12a 
               Caption         =   "Orange"
            End
            Begin VB.Menu mnu13a 
               Caption         =   "Red"
            End
            Begin VB.Menu mnu14a 
               Caption         =   "Tan"
            End
            Begin VB.Menu mnu15a 
               Caption         =   "White"
            End
            Begin VB.Menu mnu16a 
               Caption         =   "Yellow"
            End
         End
      End
      Begin VB.Menu mnuAddCH 
         Caption         =   "Add cells"
      End
      Begin VB.Menu line 
         Caption         =   "-"
         HelpContextID   =   2220
      End
      Begin VB.Menu mnuTables1 
         Caption         =   "&Add more Columns"
         Begin VB.Menu mnuCol1 
            Caption         =   "Add One Column"
         End
         Begin VB.Menu mnuCol2 
            Caption         =   "Add Two Columns"
         End
         Begin VB.Menu mnuCol3 
            Caption         =   "Add Three Columns"
         End
         Begin VB.Menu mnuCol4 
            Caption         =   "Add Four Columns"
         End
         Begin VB.Menu mnuCol5 
            Caption         =   "Add more Columns"
         End
         Begin VB.Menu mnuSepCol6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCol7 
            Caption         =   "Add Rows"
         End
         Begin VB.Menu mnuSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "Copy background codes"
         End
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Fo&nts"
      HelpContextID   =   2040
      Begin VB.Menu mnufontname 
         Caption         =   "&Font Name"
         Begin VB.Menu mnuFonts1 
            Caption         =   "Abadi MT Condensed"
         End
         Begin VB.Menu mnuFonts2 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnuFonts3 
            Caption         =   "Arial Black"
         End
         Begin VB.Menu mnuFonts4 
            Caption         =   "Arial Narrow"
         End
         Begin VB.Menu mnuFonts5 
            Caption         =   "Bookman Old Style"
         End
         Begin VB.Menu mnuFonts6 
            Caption         =   "Comic Sans MS"
         End
         Begin VB.Menu mnuFonts7 
            Caption         =   "Courier"
         End
         Begin VB.Menu mnuFonts8 
            Caption         =   "Courier New"
         End
         Begin VB.Menu mnuFonts9 
            Caption         =   "Fixedsys"
         End
         Begin VB.Menu mnuFonts10 
            Caption         =   "Garamond"
         End
         Begin VB.Menu mnuFonts11 
            Caption         =   "Impact"
         End
         Begin VB.Menu mnuFonts12 
            Caption         =   "MS Sans Serif"
         End
         Begin VB.Menu mnuFonts13 
            Caption         =   "MS Serif"
         End
         Begin VB.Menu mnuFonts14 
            Caption         =   "Marlett"
         End
         Begin VB.Menu mnuFonts15 
            Caption         =   "Small Fonts"
         End
         Begin VB.Menu mnuFonts16 
            Caption         =   "Symbol"
         End
         Begin VB.Menu mnuFonts17 
            Caption         =   "System"
         End
         Begin VB.Menu mnuFonts18 
            Caption         =   "Tahoma"
         End
         Begin VB.Menu mnuFonts19 
            Caption         =   "Terminal"
         End
         Begin VB.Menu mnuFonts20 
            Caption         =   "Times New Roman"
         End
         Begin VB.Menu mnuFonts21 
            Caption         =   "Verdana"
         End
         Begin VB.Menu mnuFonts22 
            Caption         =   "Webdings"
         End
         Begin VB.Menu mnuFonts23 
            Caption         =   "Wingdings"
         End
         Begin VB.Menu mnuFonts24 
            Caption         =   "Wingdings 2"
         End
         Begin VB.Menu mnuFonts25 
            Caption         =   "Wingdings 3"
         End
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Size"
         Begin VB.Menu mnuH1 
            Caption         =   "H1"
         End
         Begin VB.Menu mnuH2 
            Caption         =   "H2"
         End
         Begin VB.Menu mnuH3 
            Caption         =   "H3"
         End
         Begin VB.Menu mnuH4 
            Caption         =   "H4"
         End
         Begin VB.Menu mnuH5 
            Caption         =   "H5"
         End
         Begin VB.Menu mnuH6 
            Caption         =   "H6"
         End
      End
      Begin VB.Menu mnucolor 
         Caption         =   "&Color"
         Begin VB.Menu mnuBlack 
            Caption         =   "Black"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuBlueViolet 
            Caption         =   "Blue violet"
         End
         Begin VB.Menu mnuBrown 
            Caption         =   "Brown"
         End
         Begin VB.Menu mnuCyan 
            Caption         =   "Cyan"
         End
         Begin VB.Menu mnuDarkBrown 
            Caption         =   "Dark brown"
         End
         Begin VB.Menu mnuDarkGreen 
            Caption         =   "Dark green"
         End
         Begin VB.Menu mnuDarkPurple 
            Caption         =   "Dark purple"
         End
         Begin VB.Menu mnuGold 
            Caption         =   "Gold"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuMagenta 
            Caption         =   "Magenta"
         End
         Begin VB.Menu mnuOrange 
            Caption         =   "Orange"
         End
         Begin VB.Menu mnuRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuTan 
            Caption         =   "Tan"
         End
         Begin VB.Menu mnuWhite 
            Caption         =   "White"
         End
         Begin VB.Menu mnuYellow 
            Caption         =   "Yellow"
         End
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Styles"
         Begin VB.Menu mnuBlink 
            Caption         =   "Blinking"
         End
         Begin VB.Menu mnuStrikeThrough 
            Caption         =   "Strikethrough"
         End
         Begin VB.Menu mnuTypeWriter 
            Caption         =   "Typewriter"
         End
      End
   End
   Begin VB.Menu mnuOtherss 
      Caption         =   "&Others"
      HelpContextID   =   2050
      Begin VB.Menu mnuUnnumberesLists 
         Caption         =   "Unnumbered Lists"
      End
      Begin VB.Menu mnuNumberedLists 
         Caption         =   "Numbered Lists"
      End
      Begin VB.Menu mnuDefinitionLists 
         Caption         =   "Definition Lists"
      End
      Begin VB.Menu mnuNestedLists 
         Caption         =   "Nested Lists"
      End
      Begin VB.Menu mnuExtendedQuotations 
         Caption         =   "Extended Quotations"
      End
      Begin VB.Menu mnuBreaks 
         Caption         =   "Forced Line Breaks"
      End
      Begin VB.Menu mnuHorRules 
         Caption         =   "Horizontal Rules"
      End
      Begin VB.Menu mnuWhitespace 
         Caption         =   "White space"
      End
   End
End
Attribute VB_Name = "HTMLEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
MDIForm1.Visible = True
End Sub

Private Sub Command2_Click()

End Sub
Private Sub Command3_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()

  
 

RichTextBox1.Text = "<HTML>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & " Web Page </TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>" & vbCrLf & ""
End Sub


Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub
  RichTextBox1.Width = Me.Width - 105
  RichTextBox1.Height = Me.Height - 1005
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me

End Sub

Private Sub mnuH1_Click()
RichTextBox1.SelText = "<h1>" + RichTextBox1.SelText + "</h1>"
End Sub

Private Sub mnunew_Click()
    If MsgBox("Do you want to save your current project?", vbYesNo, "Save") = vbYes Then
        save_Click
        RichTextBox1.Text = "<HTML>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "Web Page</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>" & vbCrLf & ""
    Else
        RichTextBox1.Text = "<HTML>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "Web Page</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>" & vbCrLf & ""
    End If
End Sub



Private Sub mnutoolbar_Click()
If mnutoolbar.Checked = True Then
CoolBar1.Visible = False
mnutoolbar.Checked = False
Else
CoolBar1.Visible = True
mnutoolbar.Checked = True
End If
End Sub

Private Sub open_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    RichTextBox1.Text = alltext$
    Loop
    Close #1
End If

End Sub

Private Sub save_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnunew_Click
        Case "Open"
            open_Click
        Case "Save"
            save_Click
        Case "Close"
            exit_Click
        Case "Left"
            mnuLeft_Click
        Case "Center"
            mnuCenter_Click
        Case "Right"
            mnuRight_Click
        Case "Bold"
            mnuBold_Click
        Case "Italic"
            mnuItalic_Click
        Case "Underline"
            mnuunderline_Click
    End Select
End Sub


Private Sub mnu1_Click()
'richtextbox1.SelText = Chr(13) + Chr(10) + richtextbox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
RichTextBox1.SelText = RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + "Here add new cells for the first column" + Chr(13) + Chr(10) + "Here add the second column" + Chr(13) + Chr(10) + "Here add new cells for the second column, and so on " + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu10_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu10a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu11_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF00FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu11a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF00FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu12_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF7F00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu12a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF7F00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu13_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF0000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu13a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FF0000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu14_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#DB9370>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu14a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#DB9370>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu15_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu15a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu16_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu16a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#FFFF00>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu1a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu2_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu2a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu3_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu3a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu4_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu4a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu5_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu5a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu6_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu6a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu7_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu7a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu8_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu8a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnu9_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub mnu9a_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub mnuBold_Click()
RichTextBox1.SelText = "<b>" + RichTextBox1.SelText + "</b>"
End Sub

Private Sub mnuBreaks_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<BR>"
End Sub

Private Sub mnuCol1_Click()
RichTextBox1.SelText = RichTextBox1.SelText + Chr(13) + Chr(10) + "<P><TABLE BORDER=1>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the first cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Write your text in the second cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "Copy and Paste the following code to add more cells" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol2_Click()
RichTextBox1.SelText = RichTextBox1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol3_Click()
RichTextBox1.SelText = RichTextBox1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol4_Click()
RichTextBox1.SelText = RichTextBox1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol5_Click()
RichTextBox1.SelText = RichTextBox1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "ADD HERE COLUMNS. Select and Paste one of the two lines <P></TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (The cell of the LAST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCol7_Click()
RichTextBox1.SelText = RichTextBox1.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor= >" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the added ROW)" + Chr(13) + Chr(10) + "</TD><TD bgcolor= >" + Chr(13) + Chr(10) + "ADD HERE CELLS. Select and Paste the two lines above: <P>...</TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (Last cell of the added ROW)" + Chr(13) + Chr(10) + "</TD></TR>ADD HERE ROWS"
End Sub

Private Sub mnuCol8_Click()
RichTextBox1.SelText = RichTextBox1.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (4th cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub mnuCopy_Click()
Codes.Visible = True
End Sub

Private Sub mnuDefinitionLists_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<DL>" + Chr(13) + Chr(10) + "<DT> Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "<DT> Second Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "</DL>"
End Sub

Private Sub mnuExtendedQuotations_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<P>Your text" + Chr(13) + Chr(10) + "<BLOCKQUOTE>" + Chr(13) + Chr(10) + "<P> Write your text here to include lengthy quotations in a separate block on the screen" + Chr(13) + Chr(10) + "</P>" + Chr(13) + Chr(10) + "<P> Add more text here if you want</P>" + Chr(13) + Chr(10) + "</BLOCKQUOTE>"
End Sub

Private Sub mnuFonh_Click(Index As Integer)

End Sub

Private Sub mnuFonts1_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Abadi MT Condensed"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts10_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Garamond"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts11_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Impact"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts12_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""MS Sans Serif"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts13_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""MS Serif"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts14_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Marlett"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts15_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Small Fonts"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts16_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Symbol"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts17_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""System"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts18_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Tahoma"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts19_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Terminal"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts2_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts20_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Times New Roman"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts21_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Verdana"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts22_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Webdings"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts23_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts24_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings 2"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts25_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Wingdings 3"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts3_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial Black"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts4_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Arial Narrow"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts5_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Bookman Old Style"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts6_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Comic Sans MS"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts7_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Courier"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts8_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Courier New"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuFonts9_Click()
RichTextBox1.SelText = "<FONT SIZE=""Enter the Size of the Font between -2 and +4""   FACE=""Fixedsys"">" + "Type your Text Here" + "</FONT>" + RichTextBox1.SelText
End Sub

Private Sub mnuH2_Click()
RichTextBox1.SelText = "<h2>" + RichTextBox1.SelText + "</h2>"
End Sub

Private Sub mnuH3_Click()
RichTextBox1.SelText = "<h3>" + RichTextBox1.SelText + "</h3>"
End Sub

Private Sub mnuH4_Click()
RichTextBox1.SelText = "<h4>" + RichTextBox1.SelText + "</h4>"
End Sub

Private Sub mnuH5_Click()
RichTextBox1.SelText = "<h5>" + RichTextBox1.SelText + "</h5>"
End Sub

Private Sub mnuH6_Click()
RichTextBox1.SelText = "<h6>" + RichTextBox1.SelText + "</h6>"
End Sub

Private Sub mnuAddCH_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10)
End Sub


Private Sub mnuHorRules_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<HR SIZE= Enter the desired size    WIDTH=" + "Enter a number %>"
End Sub

Private Sub mnuItalic_Click()
RichTextBox1.SelText = "<i>" + RichTextBox1.SelText + "</i>"
End Sub

Private Sub mnuLink1_Click()
Link1.Show 1
End Sub

Private Sub mnuLink2_Click()
Link2.Show 1
End Sub

Private Sub mnuNestedLists_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "<LI> Second Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub mnuNumberedLists_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<OL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</OL>"
End Sub

Private Sub mnuRight_Click()
RichTextBox1.SelText = "<p align=right>" + RichTextBox1.SelText + "</p>"
End Sub
Private Sub mnuCenter_Click()
RichTextBox1.SelText = "<center>" + RichTextBox1.SelText + "</center>"
End Sub
Private Sub mnuLeft_Click()
RichTextBox1.SelText = "<p align=left>" + RichTextBox1.SelText + "</p>"
End Sub
Private Sub mnuLink_Click()
Form2.Visible = True
End Sub

Private Sub mnuPictureH_Click()
    pichtml.Show
End Sub
Private Sub mnuBlack_Click()
RichTextBox1.SelText = "<FONT COLOR=#000000>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuBlue_Click()
RichTextBox1.SelText = "<FONT COLOR=#0000FF>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuBlueViolet_Click()
RichTextBox1.SelText = "<FONT COLOR=#9F5F9F>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuBrown_Click()
RichTextBox1.SelText = "<FONT COLOR=#A62A2A>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuCyan_Click()
RichTextBox1.SelText = "<FONT COLOR=#00FFFF>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkBrown_Click()
RichTextBox1.SelText = "<FONT COLOR=#5C4033>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkGreen_Click()
RichTextBox1.SelText = "<FONT COLOR=#2F4F2F>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuDarkPurple_Click()
RichTextBox1.SelText = "<FONT COLOR=#871F78>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuGold_Click()
RichTextBox1.SelText = "<FONT COLOR=#CD7F32>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuGreen_Click()
RichTextBox1.SelText = "<FONT COLOR=#00FF00>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuMagenta_Click()
RichTextBox1.SelText = "<FONT COLOR=#FF00FF>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuOrange_Click()
RichTextBox1.SelText = "<FONT COLOR=#FF7F00>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuRed_Click()
RichTextBox1.SelText = "<FONT COLOR=#FF0000>" + RichTextBox1.SelText + "</FONT>"
End Sub

Private Sub mnuTan_Click()
RichTextBox1.SelText = "<FONT COLOR=#DB9370>" + RichTextBox1.SelText + "</FONT>"
End Sub

Private Sub mnuunderline_Click()
RichTextBox1.SelText = "<u>" + RichTextBox1.SelText + "</u>"
End Sub

Private Sub mnuUnnumberesLists_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub mnuWhite_Click()
RichTextBox1.SelText = "<FONT COLOR=#FFFFFF>" + RichTextBox1.SelText + "</FONT>"
End Sub

Private Sub mnuWhitespace_Click()
RichTextBox1.SelText = Chr(13) + Chr(10) + RichTextBox1.SelText + Chr(13) + Chr(10) + "<P>&nbsp;</P>"
End Sub

Private Sub mnuYellow_Click()
RichTextBox1.SelText = "<FONT COLOR=#FFFF00>" + RichTextBox1.SelText + "</FONT>"
End Sub
Private Sub mnuBlink_Click()
RichTextBox1.SelText = "<blink>" + RichTextBox1.SelText + "</blink>"
End Sub
Private Sub mnuBolds_Click()
RichTextBox1.SelText = "<b>" + RichTextBox1.SelText + "</b>"
End Sub
Private Sub mnuItalics_Click()
RichTextBox1.SelText = "<i>" + RichTextBox1.SelText + "</i>"
End Sub
Private Sub mnuStrikeThrough_Click()
RichTextBox1.SelText = "<strike>" + RichTextBox1.SelText + "</strike>"
End Sub
Private Sub mnuTypeWriter_Click()
RichTextBox1.SelText = "<pre>" + RichTextBox1.SelText + "</pre>"
End Sub
Private Sub mnuunderlines_Click()
RichTextBox1.SelText = "<u>" + RichTextBox1.SelText + "</u>"
End Sub
