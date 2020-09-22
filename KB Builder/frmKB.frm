VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmKB 
   Caption         =   "MyHelp"
   ClientHeight    =   8565
   ClientLeft      =   2625
   ClientTop       =   3480
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar progBar 
      Height          =   195
      Left            =   14040
      TabIndex        =   11
      Top             =   4800
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   255
      Left            =   9720
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboTree 
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4080
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboFolders 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList imgKB 
      Left            =   4080
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   47
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":57E2
            Key             =   "book"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5928
            Key             =   "leaf2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5A22
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5B1C
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5C16
            Key             =   "bullet"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5CC4
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5DD6
            Key             =   "center"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5EE8
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":5FFA
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":610C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":621E
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6330
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6442
            Key             =   "left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6554
            Key             =   "new"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6666
            Key             =   "open"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6778
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":688A
            Key             =   "print"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":699C
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6AAE
            Key             =   "save"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6BC0
            Key             =   "spelling"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6CD2
            Key             =   "strike"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6DE4
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":6EF6
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":7008
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":711A
            Key             =   "doubleunderline"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":722C
            Key             =   "project"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":7546
            Key             =   "find"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":7658
            Key             =   "topic"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":7A6A
            Key             =   "book2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":7EBC
            Key             =   "openbook"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":830E
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":88A8
            Key             =   "font"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":E4CA
            Key             =   "image"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":1084C
            Key             =   "format"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":10B66
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":10C78
            Key             =   "picture"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":10D8A
            Key             =   "tools"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":11324
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":11496
            Key             =   "right"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":115A8
            Key             =   "propercase"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":116BA
            Key             =   "selall"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":117CC
            Key             =   "popup"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":118DE
            Key             =   "bullets"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":119F0
            Key             =   "uppercase"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":11E42
            Key             =   "lowercase"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":12294
            Key             =   "wide"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":129E6
            Key             =   "narrow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   8310
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolKB 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   1005
      ButtonWidth     =   1244
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgKB"
      DisabledImageList=   "imgKB"
      HotImageList    =   "imgKB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "new"
            Object.ToolTipText     =   "New project"
            ImageIndex      =   14
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "newfolder"
                  Text            =   "Computer Folder"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "open"
            Object.ToolTipText     =   "Open project"
            ImageIndex      =   15
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
            Object.ToolTipText     =   "Save project"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "deleteproj"
            Object.ToolTipText     =   "Delete project / topics"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compile"
            Key             =   "compile"
            Object.ToolTipText     =   "Compile project"
            ImageIndex      =   26
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hpj"
                  Text            =   "View - Help Project File"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rtf"
                  Text            =   "View - Rich Text File"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mdb"
                  Text            =   "View - Database"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hlp"
                  Text            =   "View - Help File"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "log"
                  Text            =   "View - Log File"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ini"
                  Text            =   "View - Ini Settings"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Topic(s)"
            Key             =   "topicproperties"
            Object.ToolTipText     =   "Add new topic(s)"
            ImageIndex      =   28
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   25
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "addtopic"
                  Text            =   "Add New Topic(s)"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "graphictopics1"
                  Text            =   "Create From Graphics Files"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "filetopics"
                  Text            =   "Create From Files"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "topicsfolder"
                  Text            =   "Create From Folder"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a4"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "jumpto"
                  Text            =   "Jump To"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "popupto"
                  Text            =   "Popup To"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xxd"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "savetopic"
                  Text            =   "Save Topic"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "savealltopics"
                  Text            =   "Save All Topics"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "deletetopic"
                  Text            =   "Delete Topic"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "savertftopic"
                  Text            =   "Topic To RTF"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "savehtmltopic"
                  Text            =   "Topic To HTML"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xd"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "popup"
                  Text            =   "Set Tree Topic As Popup"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dontpopup"
                  Text            =   "Clear Tree Topic Popup"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "browse"
                  Text            =   "Set Tree Topic For Browse"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dontbrowse"
                  Text            =   "Clear Tree Topic For Browse"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "moveup"
                  Text            =   "Move Up"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "movedown"
                  Text            =   "Move Down"
               EndProperty
               BeginProperty ButtonMenu23 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b5"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu24 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "statistics"
                  Text            =   "Statistics"
               EndProperty
               BeginProperty ButtonMenu25 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "jumpspops"
                  Text            =   "Jumps && Popups"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert"
            Key             =   "insert"
            Object.ToolTipText     =   "Insert picture / file"
            ImageIndex      =   36
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "insertfile"
                  Text            =   "File"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "textfile"
                  Text            =   "Text File"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "timedate"
                  Text            =   "Time / Date"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tab"
                  Text            =   "Tab"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "symbol"
                  Text            =   "Symbol"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a7"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "insertgraphic"
                  Text            =   "Picture"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "print"
            Object.ToolTipText     =   "Print topics"
            ImageIndex      =   35
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "printtree"
                  Text            =   "Tree topics"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "printtopic"
                  Text            =   "Topic"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tools"
            Key             =   "tools"
            Object.ToolTipText     =   "Spelling, grammar, thesaurus"
            ImageIndex      =   37
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "spelling"
                  Text            =   "Spelling"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "grammar"
                  Text            =   "Grammar"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a4"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "capturescreen"
                  Text            =   "Capture Entire Screen"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "captureactive"
                  Text            =   "Capture Active Screen"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a7"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "gotoline"
                  Text            =   "Goto Line"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "gotostart"
                  Text            =   "Goto Start"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "gotoend"
                  Text            =   "Goto End"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "help"
            Object.ToolTipText     =   "Help manual"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit the knowledge base builder"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVBar 
      Align           =   3  'Align Left
      BackColor       =   &H00000000&
      Height          =   7740
      Left            =   3855
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7740
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   570
      Width           =   50
   End
   Begin VB.PictureBox picLst 
      Align           =   3  'Align Left
      Height          =   7740
      Left            =   3900
      ScaleHeight     =   7680
      ScaleWidth      =   10545
      TabIndex        =   2
      Top             =   570
      Width           =   10605
      Begin VB.PictureBox picTemp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   13
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.Toolbar toolDocument 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "imgKB"
         DisabledImageList=   "imgKB"
         HotImageList    =   "imgKB"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   31
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut text"
               ImageKey        =   "cut"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy text"
               ImageKey        =   "copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "paste"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sep1"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo"
               ImageKey        =   "undo"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo"
               ImageKey        =   "redo"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sep2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Bold"
               ImageKey        =   "bold"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "Italic"
               ImageKey        =   "italic"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline"
               ImageKey        =   "underline"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyword"
               Object.ToolTipText     =   "Set word as keyword"
               ImageIndex      =   37
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strike"
               ImageKey        =   "strike"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "propercase"
               Object.ToolTipText     =   "Make propercase"
               ImageKey        =   "propercase"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "font"
               Object.ToolTipText     =   "Font properties"
               ImageKey        =   "font"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "color"
               Object.ToolTipText     =   "Color properties"
               ImageKey        =   "picture"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sep3"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "left"
               Object.ToolTipText     =   "Align left"
               ImageKey        =   "left"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               Object.ToolTipText     =   "Align center"
               ImageKey        =   "center"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               Object.ToolTipText     =   "Align right"
               ImageKey        =   "right"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bullets"
               Object.ToolTipText     =   "Bullets"
               ImageKey        =   "bullets"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uppercase"
               Object.ToolTipText     =   "Upper Case"
               ImageKey        =   "uppercase"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "jumpto"
               Object.ToolTipText     =   "Create a jump"
               ImageKey        =   "topic"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "insertfile"
                     Text            =   "File"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "textfile"
                     Text            =   "Text File"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "popupto"
               Object.ToolTipText     =   "Create a popup"
               ImageKey        =   "popup"
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "embed"
               Object.ToolTipText     =   "Create a link to the embedded picture"
               ImageKey        =   "image"
               Style           =   5
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "selall"
               Object.ToolTipText     =   "Select All"
               ImageKey        =   "selall"
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "normal"
               Object.ToolTipText     =   "Normalize topic font and size"
               ImageKey        =   "justify"
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "spelling"
               Object.ToolTipText     =   "Check spelling"
               ImageKey        =   "spelling"
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "grammar"
               Object.ToolTipText     =   "Check grammar"
               ImageKey        =   "narrow"
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Find"
               ImageKey        =   "find"
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox FileT 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmKB.frx":13138
      End
      Begin RichTextLib.RichTextBox docWord 
         Height          =   735
         Left            =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmKB.frx":131B3
      End
   End
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      Height          =   7740
      Left            =   0
      ScaleHeight     =   7680
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   570
      Width           =   3855
      Begin MSComctlLib.TreeView treeDms 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5106
         _Version        =   393217
         Indentation     =   617
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgKB"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.Image imgUser 
         Height          =   240
         Left            =   0
         Picture         =   "frmKB.frx":1322E
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   0
         Picture         =   "frmKB.frx":14070
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFolders 
         Caption         =   "Folders"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bDragging As Boolean
Private iAnswer As Integer
Private jumps As New Collection
Private pops As New Collection
'Stuff for undo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(10000) As String
Public treeChanged As Boolean
Public topicChanged As Boolean
Private myWindows As clsWindows
Private hWndArray() As Long
Private Sub EnableDisableToolDocument(boolEnable As Boolean)
    On Error Resume Next
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim strSel As String
    rsTot = toolDocument.Buttons.Count
    For rsCnt = 1 To rsTot
        toolDocument.Buttons(rsCnt).Enabled = boolEnable
        Err.Clear
    Next
    If Len(docWord.Text) > 0 Then
        toolDocument.Buttons("selall").Enabled = True
        toolDocument.Buttons("normal").Enabled = True
        toolDocument.Buttons("spelling").Enabled = True
        toolDocument.Buttons("find").Enabled = True
        toolDocument.Buttons("grammar").Enabled = True
        toolDocument.Buttons("keyword").Enabled = True
    End If
    If gintIndex <= 0 Then toolDocument.Buttons("undo").Enabled = False
    If gintIndex >= 1 Then toolDocument.Buttons("redo").Enabled = True
    If gintIndex = 10000 Then toolDocument.Buttons("redo").Enabled = False
    toolDocument.Buttons("bold").MixedState = docWord.SelBold
    toolDocument.Buttons("italic").MixedState = docWord.SelItalic
    toolDocument.Buttons("strike").MixedState = docWord.SelStrikeThru
    toolDocument.Buttons("bullets").MixedState = docWord.SelBullet
    toolDocument.Buttons("underline").MixedState = docWord.SelUnderline
    If docWord.SelAlignment = rtfCenter Then
        toolDocument.Buttons("center").MixedState = True
    Else
        toolDocument.Buttons("center").MixedState = False
    End If
    If docWord.SelAlignment = rtfLeft Then
        toolDocument.Buttons("left").MixedState = True
    Else
        toolDocument.Buttons("left").MixedState = False
    End If
    If docWord.SelAlignment = rtfRight Then
        toolDocument.Buttons("right").MixedState = True
    Else
        toolDocument.Buttons("right").MixedState = False
    End If
    If Len(docWord.SelText) > 0 Then
        strSel = UCase$(docWord.SelText)
        If StringAsc(strSel) = StringAsc(docWord.SelText) Then
            toolDocument.Buttons("uppercase").MixedState = True
        Else
            toolDocument.Buttons("uppercase").MixedState = False
        End If
    Else
        toolDocument.Buttons("propercase").Enabled = False
        toolDocument.Buttons("jumpto").Enabled = False
        toolDocument.Buttons("popupto").Enabled = False
        toolDocument.Buttons("cut").Enabled = False
        toolDocument.Buttons("copy").Enabled = False
        toolDocument.Buttons("paste").Enabled = False
        toolDocument.Buttons("bold").Enabled = False
        toolDocument.Buttons("keyword").Enabled = False
        toolDocument.Buttons("italic").Enabled = False
        toolDocument.Buttons("underline").Enabled = False
        toolDocument.Buttons("strike").Enabled = False
        toolDocument.Buttons("font").Enabled = False
        toolDocument.Buttons("color").Enabled = False
        toolDocument.Buttons("left").Enabled = False
        toolDocument.Buttons("center").Enabled = False
        toolDocument.Buttons("right").Enabled = False
        toolDocument.Buttons("bullets").Enabled = False
        toolDocument.Buttons("uppercase").Enabled = False
    End If
    toolDocument.Buttons("embed").Enabled = True
    toolDocument.Buttons("jumpto").Enabled = True
    Err.Clear
End Sub
Private Sub ExactSelText()
    On Error Resume Next
    Dim lStart As Long
    Dim lEnd As Long
    lStart = docWord.SelStart
    lEnd = docWord.SelLength
    Do Until Right$(docWord.SelText, 1) <> " "
        lEnd = lEnd - 1
        docWord.SelStart = lStart
        docWord.SelLength = lEnd
        Err.Clear
    Loop
    Do Until Left$(docWord.SelText, 1) <> " "
        lStart = lStart + 1
        lEnd = lEnd - 1
        docWord.SelStart = lStart
        docWord.SelLength = lEnd
        Err.Clear
    Loop
    sStart = lStart
    sLength = lEnd
    Err.Clear
End Sub
Private Sub docWord_Change()
    On Error Resume Next
    'updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = docWord.TextRTF
        topicChanged = True
    End If
    Err.Clear
End Sub
Private Sub docWord_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyTab Then
        docWord.SelRTF = vbTab
    End If
    Err.Clear
End Sub
'    On Error Resume Next
'    If KeyCode = 13 Then SaveTopic
'    Err.Clear
'End Sub
Private Sub docWord_SelChange()
    On Error Resume Next
    EnableDisableToolDocument True
    topicChanged = True
    Err.Clear
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    EnableDisableToolDocument False
    PutProgressBarInStatusBar StatusBar1, progBar, 4
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    picMain.Width = 5700
    EnableDisableToolDocument False
    docWord.BorderStyle = rtfNoBorder
    picLst.BackColor = &H80000005
    picVBar.BackColor = &H8000000F
    mnuFile.Visible = False
    pPath = ExactPath(App.Path) & "\Projects"
    AddStatusBar StatusBar1, progBar
    MakeDirectory pPath
    UpdateProjectList
    IniPropertiesFlds
    docWord.Font.Name = FontText
    docWord.Font.Size = FontTextSize
    Err.Clear
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Select Case UnloadMode
    Case vbFormControlMenu
        If treeChanged = True Then
            Screen.MousePointer = vbHourglass
            TreeViewSaveToTable Me, progBar, sProjDb, sProject, frmKB.treeDms
            treeChanged = False
            DAO.DBEngine.Idle
            Screen.MousePointer = vbDefault
        End If
        Dim frmOpen As Form
        If mobjWord97.Documents.Count = 0 Then mobjWord97.Quit
        Set mobjWord97 = Nothing
        For Each frmOpen In Forms
            Unload frmOpen
            Err.Clear
        Next
    End Select
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    ResizeStatusBar Me, StatusBar1, progBar
    picLst.Width = Width - picMain.Width - picVBar.Width - 100
    Err.Clear
End Sub
Private Sub mnuFolders_Click(Index As Integer)
    On Error Resume Next
    sProject = mnuFolders(Index).Caption
    sProjPath = pPath & "\" & sProject
    sProjDb = sProjPath & "\" & sProject & ".mdb"
    sProjHTML = sProjPath & "\HTML"
    OpenProject
    Err.Clear
End Sub
Private Sub picVBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    bDragging = True
    picVBar.BackColor = vb3DShadow
    Err.Clear
End Sub
Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If bDragging = True Then
        picMain.Width = x + picMain.Width
        picLst.Width = Width - picMain.Width - picVBar.Width - 100
    End If
    Err.Clear
End Sub
Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    bDragging = False
    picVBar.BackColor = vbButtonFace
    Err.Clear
End Sub
Sub picLst_Resize()
    On Error Resume Next
    toolDocument.Left = 0
    toolDocument.Top = 0
    toolDocument.Width = picLst.Width
    docWord.Left = 60
    docWord.Width = picLst.Width - 150
    docWord.Height = picLst.Height - 150 - toolDocument.Height - 60
    docWord.Top = 60 + toolDocument.Height + 60
    Err.Clear
End Sub
Private Sub picMain_Resize()
    On Error Resume Next
    'Arrange in case width change
    treeDms.Top = 0
    treeDms.Left = 0
    treeDms.Width = picMain.Width - 50
    treeDms.Height = picMain.Height - 50
    Err.Clear
End Sub
Private Sub toolDocument_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim strTemp As String
    Select Case LCase$(Button.Key)
    Case "keyword"
        SaveTopic
        UpdateKeyWords
    Case "embed"
        If docWord.SelRTF = "" Then Exit Sub
        Screen.MousePointer = vbHourglass
        strTemp = StringGetFileToken(sProjHLP, "p") & "\Figure " & StringGetFileToken(iPath, "fo") & ".bmp"
        strTemp = StringNextFile(strTemp, , False)
        'copy picture to clicpboard
        Clipboard.Clear
        VBA.Interaction.SendKeys "^C"
        DoEvents
        picTemp.Picture = Clipboard.GetData()
        picTemp.Refresh
        DoEvents
        VB.SavePicture picTemp.Picture, strTemp
        DoEvents
        If File_Exists(strTemp) = True Then
            docWord.SelText = "{bml " & strTemp & "}"
            docWord.SelFontName = FontText
            docWord.SelFontSize = FontTextSize
            Screen.MousePointer = vbDefault
        Else
            Call MyPrompt("The picture could not be saved. You should retry.", "o", "e", "Save & Link Picture")
        End If
        'topicChanged = True
        'SaveTopic
    Case "propercase"
        If docWord.SelText = "" Then Exit Sub
        docWord.SelRTF = Jump_Propercase(docWord.SelRTF)
        'topicChanged = True
    Case "jumpto"
        If docWord.SelText = "" Then Exit Sub
        RemoveUselessTopics Me, progBar
        ExactSelText
        Dao_ViewSQLNew Me, progBar, sProjDb, "select title,context from properties where title like '" & Left$(docWord.SelText, 1) & "*' order by title;", frmJump.lstTopics, , , , , , False
        frmJump.lstTopics.Checkboxes = False
        frmJump.Caption = "Insert Jump Topic"
        frmJump.txtJump.Text = docWord.SelText
        frmJump.txtHelpFile.Text = ""
        frmJump.txtWindow.Text = "main"
        frmJump.optNoUnderline.Value = False
        frmJump.optNonGreen.Value = False
        frmJump.Tag = "j"
        frmJump.Show
    Case "popupto"
        If docWord.SelText = "" Then Exit Sub
        RemoveUselessTopics Me, progBar
        ExactSelText
        Dao_ViewSQLNew Me, progBar, sProjDb, "select title,context from properties where title like '" & Left$(docWord.SelText, 1) & "*' order by title;", frmJump.lstTopics, , , , , , False
        frmJump.lstTopics.Checkboxes = False
        frmJump.Caption = "Insert Popup Topic"
        frmJump.txtJump.Text = docWord.SelText
        frmJump.txtHelpFile.Text = ""
        frmJump.txtWindow.Text = "main"
        frmJump.optNoUnderline.Value = False
        frmJump.optNonGreen.Value = False
        frmJump.Tag = "p"
        frmJump.Show
    Case "grammar"
        DoEvents
        Word97Do wGrammar, docWord
        'topicChanged = True
    Case "cut"
        Clipboard.SetText docWord.SelText
        docWord.SelText = ""
        'topicChanged = True
    Case "copy"
        Clipboard.SetText docWord.SelText
        'topicChanged = True
    Case "paste"
        docWord.SelText = Clipboard.GetText
        'topicChanged = True
    Case "selall"
        docWord.SelStart = 0
        docWord.SelLength = Len(docWord.Text)
    Case "normal"
        If docWord.SelText = "" Then Exit Sub
        docWord.SelFontName = "Tahoma"
        docWord.SelFontSize = 8
        'topicChanged = True
    Case "undo"
        If gintIndex = 0 Then Exit Sub
        gblnIgnoreChange = True
        gintIndex = gintIndex - 1
        docWord.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
        'topicChanged = True
    Case "redo"
        gblnIgnoreChange = True
        gintIndex = gintIndex + 1
        docWord.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
        'topicChanged = True
    Case "bold"
        docWord.SelBold = Not docWord.SelBold
        'topicChanged = True
    Case "italic"
        docWord.SelItalic = Not docWord.SelItalic
        'topicChanged = True
    Case "underline"
        docWord.SelUnderline = Not docWord.SelUnderline
        'topicChanged = True
    Case "strike"
        docWord.SelStrikeThru = Not docWord.SelStrikeThru
        'topicChanged = True
    Case "font"
        If docWord.SelText = "" Then Exit Sub
        cd1.flags = cdlCFBoth
        cd1.FontBold = docWord.SelBold
        cd1.FontItalic = docWord.SelItalic
        cd1.FontName = docWord.SelFontName
        cd1.FontSize = docWord.SelFontSize
        cd1.ShowFont
        docWord.SelBold = cd1.FontBold
        docWord.SelFontName = cd1.FontName
        docWord.SelFontSize = cd1.FontSize
        docWord.SelItalic = cd1.FontItalic
        'topicChanged = True
    Case "color"
        Dim NewColor As Long
        cd1.ShowColor
        NewColor = cd1.Color
        If NewColor <> -1 Then docWord.SelColor = NewColor
        'topicChanged = True
    Case "left"
        docWord.SelAlignment = rtfLeft
        'topicChanged = True
    Case "center"
        docWord.SelAlignment = rtfCenter
        'topicChanged = True
    Case "right"
        docWord.SelAlignment = rtfRight
        'topicChanged = True
    Case "bullets"
        docWord.SelBullet = Not docWord.SelBullet
        'topicChanged = True
    Case "uppercase"
        strTemp = UCase$(docWord.SelText)
        If StringAsc(strTemp) = StringAsc(docWord.SelText) Then
            ' uppercase, then convert to lowercase
            docWord.SelText = LCase$(docWord.SelText)
        Else
            docWord.SelText = Jump_UpperCase(docWord.SelRTF)
        End If
        'topicChanged = True
    Case "wide"
        docWord.SelText = StrConv(docWord.SelText, vbWide)
        'topicChanged = True
    Case "narrow"
        docWord.SelText = StrConv(docWord.SelText, vbNarrow)
        'topicChanged = True
    Case "spelling"
        DoEvents
        Word97Do wSpelling, docWord
        'topicChanged = True
    Case "find"
        With frmFind
            .txtSearch.Text = docWord.SelText
            .txtReplace.Text = ""
            .chkMatch.Value = 0
            .chkWhole.Value = 0
            .Show
        End With
    End Select
    EnableDisableToolDocument True
    Err.Clear
End Sub
Private Sub toolDocument_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Key = "embed" Then
        RefreshScreens
    End If
    Err.Clear
End Sub
Private Sub toolDocument_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Dim strKey As String
    Dim strTag As String
    strKey = MvField(ButtonMenu.Key, 1, ",")
    strTag = MvField(ButtonMenu.Key, 2, ",")
    Select Case strKey
    Case "screen"
        Dim h As Long
        Dim s As String
        Dim myW As clsWindow
        ' read handle of the window
        h = Val(strTag)
        Set myW = myWindows.ByHandle(CStr(h))
        s = sProjPath & "\" & Replace$(ButtonMenu.Text, "&&", "&") & ".bmp"
        s = NextNewFile(s, False)
        SaveScreen picTemp, s, h
        DoEvents
        Clipboard.Clear
        Clipboard.SetData picTemp.Picture
        docWord.SelStart = Len(docWord.Text)
        docWord.SetFocus
        VBA.Interaction.SendKeys "^V"
        DoEvents
        ApplicationOnTop Me.hWnd
        topicChanged = True
    Case "insertfile"
        InsertFile cd1, docWord, True
        topicChanged = True
    Case "textfile"
        InsertTextFile cd1, docWord, FileT
        topicChanged = True
    End Select
    Err.Clear
End Sub
Private Sub toolKB_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If topicChanged = True Then SaveTopic
    Select Case LCase$(Button.Key)
    Case "help"
        ViewFile App.Path & "\myhelp.hlp"
    Case "new"
askProject:
        sProject = InputBox("Please type in the name of the project to create below:", "New Project", "New Project")
        If Len(sProject) = 0 Then Exit Sub
        Select Case LCase$(sProject)
        Case "new project", "html"
            iAnswer = MsgBox("The project name you have specified " & sProject & " is not acceptable, please enter another name", vbRetryCancel + vbExclamation + vbApplicationModal, "Project Name Error")
            If iAnswer = vbCancel Then Exit Sub
            GoTo askProject
        End Select
        sProjPath = pPath & "\" & sProject
        sProjDb = sProjPath & "\" & sProject & ".mdb"
        sProjHTML = sProjPath & "\HTML"
        If boolDirExists(sProjPath) = False Then
            AddProject
        Else
            iAnswer = MsgBox("A project with this name already exists, do you want to replace it?", vbYesNo + vbQuestion + vbApplicationModal, "Project Exists")
            If iAnswer = vbNo Then Exit Sub
            iAnswer = MsgBox("All the project contents will be erased, are you sure that you want to replace the project?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Replace")
            If iAnswer = vbNo Then Exit Sub
            KillFolderTree sProjPath
            AddProject
        End If
    Case "open"
        PopupMenu mnuFile
    Case "save"
        Screen.MousePointer = vbHourglass
        SaveTopic
        TreeViewSaveToTable Me, progBar, sProjDb, sProject, frmKB.treeDms
        treeChanged = False
        DAO.DBEngine.Idle
        Screen.MousePointer = vbDefault
    Case "compile"
        If File_Exists(CompilerLocation) = False Then
            MsgBox "The Help compiler is not properly installed. The Help compiler should be in this location " & CompilerLocation & "!", vbOKOnly + vbExclamation + vbApplicationModal, "Help Compiler Error"
            Err.Clear
            Exit Sub
        End If
        RemoveUselessTopics Me, progBar
        DAO.DBEngine.Idle
        ResetContents
        'Compile_HTML Me, progBar, FileT
        Compile_CNT Me, progBar
        Compile_Rtf Me, progBar
        Compile_Hpj Me, progBar
        Compile_Project sProjHLP
    Case "exit"
        If treeChanged = True Then
            Screen.MousePointer = vbHourglass
            TreeViewSaveToTable Me, progBar, sProjDb, sProject, frmKB.treeDms
            treeChanged = False
            DAO.DBEngine.Idle
            Screen.MousePointer = vbDefault
        End If
        Dim frmOpen As Form
        If mobjWord97.Documents.Count = 0 Then mobjWord97.Quit
        Set mobjWord97 = Nothing
        Call Dao_DatabaseCompress(sProjDb)
        For Each frmOpen In Forms
            Unload frmOpen
            Err.Clear
        Next
    Case "deleteproj"
        If RecycleFile(Me, sProjPath, , foDelete) = True Then
            UpdateProjectList
            Caption = "MyHelp"
            treeDms.Nodes.Clear
            docWord.Text = ""
        End If
    Case "topicproperties"
        If TypeName(treeDms.SelectedItem) = "Nothing" Then
            MsgBox "Please select a topic to view properties for first.", vbOKOnly + vbExclamation + vbApplicationModal, "Topic Error"
            Err.Clear
            Exit Sub
        End If
        CleanAllControls frmEditTopic
        With frmEditTopic
            .lblTopic.Caption = treeDms.SelectedItem.Fullpath
            .ReadTopicProperties .lblTopic.Caption
            .Tag = treeDms.SelectedItem.Key
            If Len(.txtTitle.Text) = 0 Then
                .txtTitle.Text = treeDms.SelectedItem.Text
                .txtContextString.Text = Replace$(.txtTitle.Text, " ", "_")
            End If
            .DeleteFootNote.Enabled = False
            .DeleteKeyword.Enabled = False
            .DeleteMacro.Enabled = False
            .txtTitle.Locked = True
            .txtContextString.Locked = True
            .Show
        End With
    End Select
    Err.Clear
End Sub
Private Sub UpdateTitle(sProject As String)
    On Error Resume Next
    Caption = "MyHelp: " & sProject
    Err.Clear
End Sub
Private Sub UpdateProjectList()
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim hPos As Long
    cboFolders.Clear
    RecurseFolderToComboBox pPath, cboFolders, False, True
    hPos = LstBoxFindExactItemAPI(cboFolders, "HTML")
    Do Until hPos = -1
        LstBoxRemoveItemAPI cboFolders, "HTML"
        hPos = LstBoxFindExactItemAPI(cboFolders, "HTML")
        Err.Clear
    Loop
    rsTot = toolKB.Buttons("open").ButtonMenus.Count
    For rsCnt = rsTot To 1 Step -1
        toolKB.Buttons("open").ButtonMenus.Remove rsCnt
        Err.Clear
    Next
    rsTot = cboFolders.ListCount - 1
    For rsCnt = 0 To rsTot
        toolKB.Buttons("open").ButtonMenus.Add , "project," & cboFolders.List(rsCnt), cboFolders.List(rsCnt)
        Err.Clear
    Next
    Err.Clear
End Sub
Private Sub AddProject()
    On Error Resume Next
    MakeDirectory sProjPath
    MakeDirectory sProjPath & "\HTML"
    treeDms.Nodes.Clear
    docWord.Text = ""
    TreeView_AddParent treeDms, "ProjectName", sProject
    UpdateTitle sProject
    StatusMessage Me, sProjPath, 1
    UpdateProjectList
    dbCreate sProjDb, True
    TreeViewSaveToTable Me, progBar, sProjDb, sProject, treeDms
    dbCreateTable sProjDb, "Properties", "FullPath,Title,Context,Number,Browse,Keywords,Macros,FootNotes,Contents,Popup,File", "memo,text,memo,long,integer,me,me,me,me,integer,me", "255,255,,,,,,,,,", "1,2,3,4"
    dbCreateTable sProjDb, "DataFiles", "FullPath,Files", "memo,memo", ",", "FullPath"
    sProjContents = Replace$(sProjDb, ".mdb", "contents.dat")
    sProjRTF = Replace$(sProjDb, ".mdb", ".rtf")
    sProjCnt = Replace$(sProjDb, ".mdb", ".cnt")
    sProjHPJ = Replace$(sProjDb, ".mdb", ".hpj")
    sProjHLP = Replace$(sProjDb, ".mdb", ".hlp")
    sProjIni = Replace$(sProjDb, ".mdb", ".ini")
    sProjLog = Replace$(sProjDb, ".mdb", ".log")
    sProjHTML = sProjPath & "\HTML"
    Title = sProject
    Err.Clear
End Sub
Private Sub OpenProject()
    On Error Resume Next
    gintIndex = 0
    treeDms.Nodes.Clear
    docWord.Text = ""
    UpdateTitle sProject
    StatusMessage Me, sProjPath, 1
    TreeViewLoadFromTable Me, progBar, sProjDb, treeDms, sProject
    If treeDms.Nodes.Count = 0 Then
        AddProject
    Else
        treeDms.Nodes.Item(1).Expanded = True
    End If
    sProjContents = Replace$(sProjDb, ".mdb", "contents.dat")
    sProjRTF = Replace$(sProjDb, ".mdb", ".rtf")
    sProjCnt = Replace$(sProjDb, ".mdb", ".cnt")
    sProjHPJ = Replace$(sProjDb, ".mdb", ".hpj")
    sProjHLP = Replace$(sProjDb, ".mdb", ".hlp")
    sProjIni = Replace$(sProjDb, ".mdb", ".ini")
    sProjLog = Replace$(sProjDb, ".mdb", ".log")
    Err.Clear
End Sub
Private Sub ChangeImage(strImage As String)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Screen.MousePointer = vbHourglass
    nTot = treeDms.Nodes.Count
    ProgBarInit Me, progBar, nTot
    For nCnt = 1 To nTot
        Call UpdateProgress(Me, nCnt, progBar, "Changing images to " & strImage)
        If treeDms.Nodes(nCnt).Checked = True Then
            treeDms.Nodes(nCnt).Image = strImage
            treeDms.Nodes(nCnt).SelectedImage = strImage
        End If
        Err.Clear
    Next
    ProgBarClose Me, progBar
    TreeViewSaveToTable Me, progBar, sProjDb, sProject, frmKB.treeDms
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Private Function AskForBitMapFile() As String
    On Error GoTo ExitSub
    With cd1
        .CancelError = True
        .DialogTitle = "Specify Picture File Name"
        .Filter = "Bitmap File (*.bmp)|*.bmp|"
        .ShowSave
        AskForBitMapFile = .FileName
    End With
    Err.Clear
    Exit Function
ExitSub:
    AskForBitMapFile = ""
    Err.Clear
End Function
Private Sub toolKB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Dim strTemp As String
    Dim intAns As Long
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim tNode As Node
    Dim strKey As String
    Dim strSecond As String
    strKey = MvField(ButtonMenu.Key, 1, ",")
    strSecond = MvField(ButtonMenu.Key, 2, ",")
    If topicChanged = True Then SaveTopic
    Select Case LCase$(strKey)
    Case "project"
        sProject = strSecond
        sProjPath = pPath & "\" & sProject
        sProjDb = sProjPath & "\" & sProject & ".mdb"
        sProjHTML = sProjPath & "\HTML"
        OpenProject
    Case "savealltopics"
        intAns = MyPrompt("You have opted to resave all topics available in this project, are you sure?", "yn", "q", "Confirm Resave")
        Screen.MousePointer = vbHourglass
        rsTot = treeDms.Nodes.Count
        ProgBarInit Me, progBar, rsTot, "Saving topics, please wait..."
        For rsCnt = 1 To rsTot
            Set tNode = treeDms.Nodes(rsCnt)
            treeDms_NodeClick tNode
            UpdateProgress Me, rsCnt, progBar, "Saving " & tNode.Fullpath
            SaveTopic
            DoEvents
            Err.Clear
        Next
        ProgBarClose Me, progBar
        RemoveUselessTopics Me, progBar
        Screen.MousePointer = vbDefault
    Case "addtopic"
        If TypeName(treeDms.SelectedItem) = "Nothing" Then
            MsgBox "Please select a parent topic to add the topics to first.", vbOKOnly + vbExclamation + vbApplicationModal, "Parent Topic Error"
            Err.Clear
            Exit Sub
        End If
        CleanAllControls frmAddTopics
        With frmAddTopics
            .cboTopicLocation.AddItem "Child"
            .cboTopicLocation.AddItem "First"
            .cboTopicLocation.AddItem "Last"
            .cboTopicLocation.AddItem "Next"
            .cboTopicLocation.AddItem "Previous"
            .lblMain.Caption = treeDms.SelectedItem.Fullpath
            .Tag = treeDms.SelectedItem.Key
            .chkGen.Value = 0
            .Show
        End With
    Case "normal"
        If docWord.SelText = "" Then Exit Sub
        docWord.SelFontName = "Tahoma"
        docWord.SelFontSize = 8
        'topicChanged = True
    Case "mdb"
        If boolFileExists(sProjDb) = True Then
            Call boolViewFile(sProjDb)
        End If
    Case "rtf"
        If boolFileExists(sProjRTF) = True Then
            Call boolViewFile(sProjRTF)
        End If
    Case "cnt"
        If boolFileExists(sProjCnt) = True Then
            Call boolViewFile(sProjCnt)
        End If
    Case "hpj"
        If boolFileExists(sProjHPJ) = True Then
            Call boolViewFile(sProjHPJ)
        End If
    Case "hlp"
        If boolFileExists(sProjHLP) = True Then
            Call boolViewFile(sProjHLP)
        End If
    Case "ini"
        If boolFileExists(sProjIni) = True Then
            Call boolViewFile(sProjIni)
        End If
    Case "log"
        If boolFileExists(sProjLog) = True Then
            Call boolViewFile(sProjLog)
        End If
    Case "newfolder"
        strTemp = StringBrowseForFolder(Me.hWnd, "Select Folder To Create Project From")
        If Len(strTemp) = 0 Then
            Err.Clear
            Exit Sub
        End If
        sProject = StringGetFileToken(strTemp, "f")
        sProjPath = pPath & "\" & sProject
        sProjDb = sProjPath & "\" & sProject & ".mdb"
        sProjHTML = sProjPath & "\HTML"
        If boolDirExists(sProjPath) = False Then
            AddProject
        Else
            iAnswer = MsgBox("A project with this name already exists, do you want to replace it?", vbYesNo + vbQuestion + vbApplicationModal, "Project Exists")
            If iAnswer = vbNo Then Exit Sub
            iAnswer = MsgBox("All the project contents will be erased, are you sure that you want to replace the project?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Replace")
            If iAnswer = vbNo Then Exit Sub
            KillFolderTree sProjPath
            AddProject
        End If
        ImportComputerFolder strTemp, False
    Case "gotoline"
        Dim lngStart As Long
AskUserLine:
        strTemp = InputBox("Please enter the line number to goto:", "Goto Line", 1)
        If Len(strTemp) = 0 Then
            iAnswer = MsgBox("You have not entered a line number.", vbRetryCancel + vbQuestion + vbApplicationModal, "Goto Line")
            If iAnswer = vbCancel Then Exit Sub
            GoTo AskUserLine
        End If
        lngStart = SendMessage(docWord.hWnd, EM_LINEINDEX, Val(strTemp) - 1, 0&)
        If lngStart = -1 Then
            iAnswer = MsgBox("The line number you have entered is invalid.", vbRetryCancel + vbQuestion + vbApplicationModal, "Goto Line")
            If iAnswer = vbCancel Then Exit Sub
            GoTo AskUserLine
        End If
        docWord.SelStart = lngStart
        docWord.SetFocus
    Case "gotostart"
        docWord.SelStart = 0
        docWord.SetFocus
    Case "gotoend"
        docWord.SelStart = Len(docWord.Text)
        docWord.SetFocus
    Case "capturescreen"
        Set Me.Picture = CaptureScreen()
        strTemp = AskForBitMapFile
        If Len(strTemp) > 0 Then Call SavePicture(Me.Picture, strTemp)
    Case "captureactive"
        Set Me.Picture = CaptureActiveWindow
        strTemp = AskForBitMapFile
        If Len(strTemp) > 0 Then Call SavePicture(Me.Picture, strTemp)
    Case "topicsfolder"
        ImportComputerFolder
    Case "filetopics"
    Case "graphictopics"
    Case "moveup"
        TreeViewMoveNode treeDms, treeDms.SelectedItem, "UP"
        treeDms.SelectedItem.EnsureVisible
        'SaveTopic
        'treeChanged = True
    Case "movedown"
        TreeViewMoveNode treeDms, treeDms.SelectedItem, "DOWN"
        treeDms.SelectedItem.EnsureVisible
        'SaveTopic
        treeChanged = True
    Case "spelling"
        DoEvents
        Word97Do wSpelling, docWord
        'topicChanged = True
    Case "grammar"
        DoEvents
        Word97Do wGrammar, docWord
        'topicChanged = True
    Case "savertftopic"
        DoEvents
        Screen.MousePointer = vbHourglass
        Call boolViewFile(SaveRtfTopic(iPath))
        Screen.MousePointer = vbDefault
    Case "savehtmltopic"
        DoEvents
        Word97Do wSaveHTML, docWord, Replace$(TopicName, " ", "")
        boolViewFile sProjHTML & "\" & Replace$(TopicName, " ", "") & ".html"
    Case "jumpspops"
        Screen.MousePointer = vbHourglass
        JumpsPopsStatus
        Screen.MousePointer = vbDefault
        If jumps.Count > 0 Or pops.Count > 0 Then
            frmJumpsPops.LoadJumpsPops jumps, pops
            frmJumpsPops.chkCheck.Value = 0
            frmJumpsPops.Show 1
        Else
            MsgBox "There are no jumps or popups in this topic.", vbOKOnly + vbInformation + vbApplicationModal, "Jumps & Popups"
        End If
    Case "font"
        If docWord.SelText = "" Then Exit Sub
        cd1.flags = cdlCFBoth
        cd1.ShowFont
        docWord.SelBold = cd1.FontBold
        docWord.SelFontName = cd1.FontName
        docWord.SelFontSize = cd1.FontSize
        docWord.SelItalic = cd1.FontItalic
        'topicChanged = True
    Case "jumpto"
        If docWord.SelText = "" Then Exit Sub
        RemoveUselessTopics Me, progBar
        ExactSelText
        Dao_ViewSQLNew Me, progBar, sProjDb, "select title,context from properties order by number;", frmJump.lstTopics, , , , , , False
        frmJump.lstTopics.Checkboxes = False
        frmJump.Caption = "Insert Jump Topic"
        frmJump.txtJump.Text = docWord.SelText
        frmJump.txtHelpFile.Text = ""
        frmJump.txtWindow.Text = "main"
        frmJump.optNoUnderline.Value = False
        frmJump.optNonGreen.Value = False
        frmJump.Tag = "j"
        frmJump.Show
    Case "popupto"
        If docWord.SelText = "" Then Exit Sub
        RemoveUselessTopics Me, progBar
        ExactSelText
        Dao_ViewSQLNew Me, progBar, sProjDb, "select title,context from properties order by number;", frmJump.lstTopics, , , , , , False
        frmJump.lstTopics.Checkboxes = False
        frmJump.Caption = "Insert Popup Topic"
        frmJump.txtJump.Text = docWord.SelText
        frmJump.txtHelpFile.Text = ""
        frmJump.txtWindow.Text = "main"
        frmJump.optNoUnderline.Value = False
        frmJump.optNonGreen.Value = False
        frmJump.Tag = "p"
        frmJump.Show
    Case "textfile"
        InsertTextFile cd1, docWord, FileT
        'topicChanged = True
    Case "symbol"
        frmSymbols.Show 1
    Case "tab"
        docWord.SelRTF = vbTab
        'SaveTopic
    Case "statistics"
        Statistics
    Case "deletetopic"
        iAnswer = MsgBox("Are you sure that you want to delete the checked topics?" & vbCr & vbCr & "You will not be able to undo your actions, are you sure?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Delete")
        If iAnswer = vbNo Then Exit Sub
        Dim strDel As String
        Screen.MousePointer = vbHourglass
        strDel = TreeViewRemoveChecked(Me, progBar, treeDms, True)
        RemoveTopicsFromDb strDel
        If treeDms.Nodes.Count = 0 Then AddProject
        TreeViewSaveToTable Me, progBar, sProjDb, sProject, treeDms
        DAO.DBEngine.Idle
        If TreeViewTopicPosition(Me, progBar, treeDms, iPath) = 0 Then
            docWord.TextRTF = ""
        End If
        treeChanged = False
        Screen.MousePointer = vbDefault
    Case "katakana"
        docWord.SelText = StrConv(docWord.SelText, vbKatakana)
        'SaveTopic
    Case "hiragana"
        docWord.SelText = StrConv(docWord.SelText, vbHiragana)
        'SaveTopic
    Case "wide"
        docWord.SelText = StrConv(docWord.SelText, vbWide)
        'SaveTopic
    Case "narrow"
        docWord.SelText = StrConv(docWord.SelText, vbNarrow)
        'SaveTopic
    Case "titlecase"
        docWord.SelText = StrConv(docWord.SelText, vbProperCase)
        'SaveTopic
    Case "find"
        With frmFind
            .txtSearch.Text = docWord.SelText
            .txtReplace.Text = ""
            .chkMatch.Value = 0
            .chkWhole.Value = 0
            .Show
        End With
    Case "subscript"
        docWord.SelCharOffset = -55
        'SaveTopic
    Case "noscripting"
        docWord.SelCharOffset = 0
        'SaveTopic
    Case "superscript"
        docWord.SelCharOffset = 55
        'SaveTopic
    Case "special"
        SendKeys ("+{insert}")
    Case "uppercase"
        docWord.SelText = UCase$(docWord.SelText)
        'SaveTopic
    Case "lowercase"
        docWord.SelText = LCase$(docWord.SelText)
        'SaveTopic
    Case "timedate"
        frmAddDate.Show 1
    Case "selall"
        docWord.SelStart = 0
        docWord.SelLength = Len(docWord.Text)
    Case "printtopic"
        PrintTopic
    Case "openbook"
        ChangeImage "book"
    Case "leaf"
        ChangeImage "leaf"
    Case "popup"
        ChangeProperties "popup", 1
    Case "browse"
        ChangeProperties "browse", 1
    Case "dontbrowse"
        ChangeProperties "browse", 0
    Case "dontpopup"
        ChangeProperties "popup", 0
    Case "insertgraphic"
        frmPicture.ClearAll
        frmPicture.Show
    Case "insertfile"
        InsertFile cd1, docWord, True
        'SaveTopic
    Case "cut"
        Clipboard.SetText docWord.SelText
        docWord.SelText = ""
        'SaveTopic
    Case "copy"
        Clipboard.SetText docWord.SelText
        'SaveTopic
    Case "paste"
        docWord.SelText = Clipboard.GetText
        'SaveTopic
    Case "bold"
        docWord.SelBold = Not docWord.SelBold
        'SaveTopic
    Case "italic"
        docWord.SelItalic = Not docWord.SelItalic
        'SaveTopic
    Case "underline"
        docWord.SelUnderline = Not docWord.SelUnderline
        'SaveTopic
    Case "color"
        Dim NewColor As Long
        cd1.ShowColor
        NewColor = cd1.Color
        If NewColor <> -1 Then docWord.SelColor = NewColor
        'SaveTopic
    Case "left"
        docWord.SelAlignment = rtfLeft
        'SaveTopic
    Case "center"
        docWord.SelAlignment = rtfCenter
        'SaveTopic
    Case "right"
        docWord.SelAlignment = rtfRight
        'SaveTopic
    Case "bullets"
        docWord.SelBullet = Not docWord.SelBullet
        'SaveTopic
    Case "strike"
        docWord.SelStrikeThru = Not docWord.SelStrikeThru
        'SaveTopic
    Case "undo"
        If gintIndex = 0 Then Exit Sub
        gblnIgnoreChange = True
        gintIndex = gintIndex - 1
        docWord.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
        'SaveTopic
    Case "redo"
        gblnIgnoreChange = True
        gintIndex = gintIndex + 1
        docWord.TextRTF = gstrStack(gintIndex)
        gblnIgnoreChange = False
        'SaveTopic
    End Select
ExitSub:
    Err.Clear
End Sub
Private Sub RemoveTopicsFromDb(strPaths As String)
    On Error Resume Next
    If Len(strPaths) = 0 Then Exit Sub
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spStr As String
    spLine = Split(strPaths, vbNewLine, , vbTextCompare)
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Fullpath"
    spTot = UBound(spLine)
    For spCnt = 0 To spTot
        spStr = Trim$(spLine(spCnt))
        If Len(spStr) = 0 Then GoTo NextTopic
        tb.Seek "=", spStr
        If tb.NoMatch = False Then tb.Delete
NextTopic:
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub
Private Function RenameTopic(strOldPath As String, ByVal strNewName As String) As Boolean
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim strNewPath As String
    Dim rExists As Boolean
    Dim intAns As Long
    strNewName = StringProperCase(strNewName)
    strNewPath = StringGetFileToken(strOldPath, "P", "\") & "\" & strNewName
    strNewPath = StringProperCase(strNewPath)
    rExists = Dao_RecordExists(sProjDb, "Properties", "FullPath", strNewPath)
    If rExists = True Then
        intAns = MyPrompt("The topic name you have specified already exists in this project." & vbCr & "Do you want to change the topic name on the tree?", "yn", "q", "Confirm Rename")
        If intAns = vbNo Then
            RenameTopic = False
            Err.Clear
            Exit Function
        Else
            RenameTopic = True
            Err.Clear
            Exit Function
        End If
    End If
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Fullpath"
    tb.Seek "=", strOldPath
    If tb.NoMatch = False Then
        tb.Edit
        tb!Title.Value = strNewName
        tb!Context.Value = Context_Validate(MvFromMv(strNewPath, 2, , "\"))
        tb!Keywords.Value = Keywords_Validate(strNewPath)
        tb!Fullpath.Value = strNewPath
        tb.Update
    End If
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    RenameTopic = True
    Err.Clear
End Function
Private Sub RenameChildTopics(strOldPath As String, ByVal strNewName As String)
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim strNewPath As String
    Dim afterPath As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sFullPath As String
    Dim pLen As String
    Dim nFullPath As String
    strNewName = StringProperCase(strNewName)
    strNewPath = StringGetFileToken(strOldPath, "P", "\") & "\" & strNewName
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("select * from Properties where FullPath like '" & strOldPath & "*'")
    tb.MoveLast
    rsTot = tb.RecordCount
    tb.MoveFirst
    Call ProgBarInit(Me, progBar, rsTot)
    pLen = Len(strOldPath)
    For rsCnt = 1 To rsTot
        Call UpdateProgress(Me, rsCnt, progBar, "Updating child topics")
        sFullPath = StringProperCase(tb!Fullpath.Value & "")
        afterPath = Mid$(sFullPath, pLen + 1)
        nFullPath = StringGetFileToken(strOldPath, "P", "\") & "\" & strNewName & afterPath
        tb.Edit
        tb!Fullpath = nFullPath
        tb.Update
        tb.MoveNext
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    ProgBarClose Me, progBar
    Err.Clear
End Sub
Private Sub treeDms_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error Resume Next
    Dim curNode As Node
    Dim intAns As Long
    Dim rExists As Boolean
    Dim strOld As String
    Set curNode = treeDms.SelectedItem
    If TypeName(curNode) = "Nothing" Then Exit Sub
    If topicChanged = True Then SaveTopic
    NewString = StringProperCase(Topic_Validate(NewString))
    If LCase$(curNode.Text) <> LCase$(NewString) Then
        intAns = MyPrompt("You have opted to change this topic, are you sure?" & vbCr & vbCr & "Old Topic Name: " & curNode.Text & vbCr & "New Topic Name: " & NewString, "yn", "q", "Confirm Change")
        If intAns = vbNo Then
            Cancel = True
            Err.Clear
            Exit Sub
        End If
        strOld = curNode.Fullpath
        rExists = RenameTopic(curNode.Fullpath, NewString)
        If rExists = False Then
            Cancel = True
            Err.Clear
            Exit Sub
        End If
        RenameChildTopics strOld, NewString
        treeChanged = True
        Caption = "MyHelp: " & curNode.Fullpath
    End If
    Err.Clear
End Sub
Private Sub treeDms_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If topicChanged = True Then SaveTopic
    If Node.Children > 0 Then
        TreeViewCheckChildren treeDms, Node.Index, Node.Checked
    End If
    Err.Clear
End Sub
Private Sub ChangeProperties(StrFldName As String, Value As Integer)
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim nTot As Long
    Dim nCnt As Long
    Dim nStr As String
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Fullpath"
    nTot = treeDms.Nodes.Count
    Call ProgBarInit(Me, progBar, nTot)
    For nCnt = 1 To nTot
        Call UpdateProgress(Me, nCnt, progBar, "Changing properties")
        nStr = treeDms.Nodes(nCnt).Fullpath
        If treeDms.Nodes(nCnt).Checked = True Then
            tb.Seek "=", nStr
            If tb.NoMatch = False Then
                tb.Edit
                tb.Fields(StrFldName).Value = Value
                tb.Update
            End If
        End If
        Err.Clear
    Next
    ProgBarClose Me, progBar
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub
Private Sub treeDms_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If Node.Index = 1 Then
        docWord.TextRTF = ""
        Err.Clear
        Exit Sub
    End If
    If topicChanged = True Then SaveTopic
    If Node.Image = "project" Then
        docWord.Text = ""
    Else
        Caption = "MyHelp: " & Node.Fullpath
        StatusMessage Me, "Reading topic, please wait..."
        Screen.MousePointer = vbHourglass
        iPath = Node.Fullpath
        TopicName = Node.Text
        ReadContents Node.Fullpath
        StatusMessage Me, "Reading statistics for jumps and popups, please wait..."
        JumpsPopsStatus
        EnableDisableToolDocument False
        Screen.MousePointer = vbDefault
        StatusMessage Me
    End If
    StatusMessage Me, Node.Fullpath, 5
    Err.Clear
End Sub
Sub SaveTopic()
    On Error Resume Next
    StatusMessage Me, "Saving topic, please wait..."
    If Len(iPath) > 0 Then UpdateContents
    EnableDisableToolDocument False
    StatusMessage Me, iPath
    topicChanged = False
    Err.Clear
End Sub
Private Sub PrintTopic()
    On Error GoTo errorhandler
    Dim bcancel As Boolean
    Dim ncopy As Integer
    Dim ncopy_Tot As Integer
    On Error GoTo errorhandler
    bcancel = False
    cd1.flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDNoPageNums Or cdlPDCollate
    cd1.CancelError = True
    cd1.PrinterDefault = True
    cd1.Copies = 1
    cd1.ShowPrinter
    If bcancel = False Then
        PrintRTF docWord, 1440, 1440, 1440, 1440
        ncopy_Tot = cd1.Copies
        For ncopy = 1 To ncopy_Tot
            Err.Clear
        Next
    End If
    Err.Clear
    Exit Sub
errorhandler:
    If Err.Number = cdlCancel Then
        bcancel = True
        Resume Next
    End If
    Err.Clear
End Sub
Private Sub Statistics()
    On Error Resume Next
    Dim words() As String
    Dim lines() As String
    Dim charsexc As Long
    Dim charsinc As Long
    Dim lwords As Long
    Dim llines As Long
    Dim StrMsg As String
    If docWord.SelText = "" Then
        words = Split(docWord.Text, " ")
        lines = Split(docWord.Text, vbCrLf)
        charsexc = Len(docWord.Text) - UBound(words)
        charsinc = Len(docWord.Text)
    Else
        words = Split(docWord.SelText, " ")
        lines = Split(docWord.SelText, vbCrLf)
        charsexc = Len(docWord.SelText) - UBound(words)
        charsinc = Len(docWord.SelText)
    End If
    lwords = UBound(words) + 1
    llines = UBound(lines) + 1
    StrMsg = "Words: " & lwords & vbCr & vbCr & "Characters (exc. spaces): " & charsexc & vbCr & vbCr & "Characters (inc. spaces): " & charsinc & vbCr & vbCr & "Lines: " & llines
    MsgBox StrMsg, vbOKOnly + vbInformation + vbApplicationModal, "Topic Statistics"
    Err.Clear
End Sub
Private Sub UpdateContents()
    On Error Resume Next
    If Len(iPath) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    With tb
        .Index = "fullpath"
        .Seek "=", iPath
        Select Case .NoMatch
        Case False
            .Edit
            If Len(Trim$(docWord.Text)) = 0 Then
                !Contents = ""
            Else
                !Contents = docWord.TextRTF
            End If
            !Keywords = Keywords_Validate(iPath)
            !Context = Context_Validate(MvFromMv(iPath, 2, , "\"))
            .Update
        End Select
    End With
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub


Private Sub UpdateKeyWords()
    On Error Resume Next
    If Len(iPath) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim sKW As String
    Dim tPos As Long
    Dim sTc As String
    
    sTc = docWord.SelText
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    With tb
        .Index = "fullpath"
        .Seek "=", iPath
        Select Case .NoMatch
        Case False
            sKW = !Keywords.Value & ""
            tPos = MvSearch(sKW, sTc, FM)
            If tPos = 0 Then sKW = sKW & FM & sTc
            .Edit
            !Keywords = sKW
            .Update
        End Select
    End With
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub

Private Sub ReadContents(Fullpath As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("select contents from Properties where fullpath = '" & Fullpath & "';")
    tb.MoveLast
    If tb.EOF = False Then
        docWord.TextRTF = tb!Contents.Value & ""
    Else
        docWord.Text = ""
    End If
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Sub JumpsPopsStatus()
    On Error Resume Next
    Set jumps = New Collection
    Set pops = New Collection
    If InStr(1, docWord.TextRTF, "\uldb ", vbTextCompare) > 0 Then
        Set jumps = WordsInBetween(docWord.TextRTF, "\uldb ", "\v0 ")
    End If
    If InStr(1, docWord.TextRTF, "\ul ", vbTextCompare) > 0 Then
        Set pops = WordsInBetween(docWord.TextRTF, "\ul ", "\v0 ")
    End If
    StatusMessage Me, "Jumps: " & jumps.Count, 2
    StatusMessage Me, "Popups: " & pops.Count, 3
    Err.Clear
End Sub
Private Sub ImportComputerFolder(Optional useFolder As String = "", Optional AskForFolder As Boolean = True)
    On Error Resume Next
    Dim StrFolder As String
    If AskForFolder = True Then
        StrFolder = StringBrowseForFolder(Me.hWnd, "Select Folder To Create Topics From")
        If Len(StrFolder) = 0 Then
            Err.Clear
            Exit Sub
        End If
    Else
        StrFolder = useFolder
    End If
    iAnswer = MsgBox("Do you want to retain the structure of the file paths?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Retain")
    Screen.MousePointer = vbHourglass
    LstBoxClearAPI cboTree
    RecurseFolderToComboBox StrFolder, cboTree
    AddTopics sProject, cboTree, iAnswer
    TreeViewSaveToTable Me, progBar, sProjDb, sProject, treeDms
    DAO.DBEngine.Idle
    Screen.MousePointer = vbDefault
    MsgBox cboTree.ListCount & " topics have been added to the project", vbOKOnly + vbInformation + vbApplicationModal, "Topics Added"
    Err.Clear
End Sub
Private Sub AddTopics(StrParent As String, cboTree As ComboBox, bRetain As Integer)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nStr As String
    Dim nPath As String
    Dim nTopic As String
    Dim nParent As String
    Dim oPath As String
    Dim sLink As String
    Dim nPos As Long
    Dim docOle As OLE
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim lngNew As Long
    nTot = cboTree.ListCount
    Call ProgBarInit(Me, progBar, nTot)
    nTot = nTot - 1
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "Fullpath"
    For nCnt = 0 To nTot
        DoEvents
        nPath = cboTree.List(nCnt)
        Call UpdateProgress(Me, nCnt, progBar, "Adding " & nPath)
        nTopic = StringGetFileToken(nPath, "f")
        nTopic = StringProperCase(nTopic)
        nParent = StringGetFileToken(nPath, "p") & "\" & nTopic
        Select Case bRetain
        Case vbYes
            nStr = StrParent & Mid$(nParent, 3)
        Case Else
            nStr = StrParent & "\" & nTopic
        End Select
        nStr = StringProperCase(nStr)
        nPos = TreeViewAddPath(treeDms, nStr, "leaf", "leaf")
        treeDms.Nodes(nPos).EnsureVisible
        If IsFilePicture(nPath) = True Then
            sLink = "{bml " & oPath & "}"
            docWord.Text = sLink
            docWord.SelStart = 0
            docWord.SelLength = Len(docWord.Text)
            docWord.SelAlignment = rtfCenter
        Else
            DoEvents
            docWord.TextRTF = ""
            docWord.OLEObjects.Clear
            Set docOle = docWord.OLEObjects.Add(, , nPath)
            Set docOle = Nothing
        End If
        docWord.Refresh
        tb.Seek "=", nStr
        Select Case tb.NoMatch
        Case True
            lngNew = Val(dbNextOpenSequence(db, "Properties", "number"))
            tb.AddNew
            tb!Fullpath = nStr
            tb!Title = nTopic
            tb!Context = Context_Validate(MvFromMv(nStr, 2, , "\"))
            tb!Number = lngNew
            tb!browse = 1
            tb!Keywords = Keywords_Validate(nStr)
            tb!Macros = ""
            tb!Footnotes = ""
            tb!Contents = docWord.TextRTF
            tb!File = nPath
            tb.Update
        Case False
            tb.Edit
            tb!Contents = docWord.TextRTF
            tb!File = nPath
            tb.Update
        End Select
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    ProgBarClose Me, progBar
    Err.Clear
End Sub
Private Sub ResetContents()
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    nTot = treeDms.Nodes.Count
    Call ProgBarInit(Me, progBar, nTot)
    For nCnt = 2 To nTot
        Call UpdateProgress(Me, nCnt, progBar, "Resetting topic tree structure...")
        If treeDms.Nodes(nCnt).Children = 0 Then
            treeDms.Nodes(nCnt).Image = "leaf"
            treeDms.Nodes(nCnt).SelectedImage = "leaf"
        Else
            treeDms.Nodes(nCnt).Image = "book"
            treeDms.Nodes(nCnt).SelectedImage = "openbook"
        End If
        Err.Clear
    Next
    TreeViewSaveToTable Me, progBar, sProjDb, sProject, treeDms
    ProgBarClose Me, progBar
    Err.Clear
End Sub

Private Sub RefreshScreens()
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    rsTot = toolDocument.Buttons("embed").ButtonMenus.Count
    For rsCnt = rsTot To 1 Step -1
        toolDocument.Buttons("embed").ButtonMenus.Remove rsCnt
        Err.Clear
    Next
    Set myWindows = New clsWindows
    ReDim hWndArray(0)
    Call FindWindowLike(myWindows, hWndArray, 0, "*", "ThunderRT6FormDC")
    Dim myW As clsWindow
    Dim strTitle As String
    Dim strTag As String
    rsTot = myWindows.Count
    Select Case rsTot
    Case 0
    Case Else
        Set myW = myWindows.ByPosition(1)
        strTitle = Replace$(myW.Title, "&", "&&")
        strTag = myW.Handle
        toolDocument.Buttons("embed").ButtonMenus.Add , "screen," & strTag, strTitle
        For rsCnt = 2 To rsTot
            Set myW = myWindows.ByPosition(rsCnt)
            strTitle = Replace$(myW.Title, "&", "&&")
            strTag = myW.Handle
            toolDocument.Buttons("embed").ButtonMenus.Add , "screen," & strTag, strTitle
            Err.Clear
        Next
    End Select
    Err.Clear
End Sub
