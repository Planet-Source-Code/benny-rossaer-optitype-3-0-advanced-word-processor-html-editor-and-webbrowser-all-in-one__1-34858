VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form form1 
   Caption         =   "OptiType 3.0 PSC."
   ClientHeight    =   8595
   ClientLeft      =   600
   ClientTop       =   -6990
   ClientWidth     =   14295
   Icon            =   "OptiWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog savefiledialog 
      Left            =   6780
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   675
      Top             =   6585
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlsavewebpage 
      Left            =   240
      Top             =   7695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   795
      Left            =   1860
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   1785
      ExtentX         =   3149
      ExtentY         =   1402
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.Toolbar webbuttons 
      Align           =   4  'Align Right
      Height          =   7470
      Left            =   13725
      TabIndex        =   23
      Top             =   870
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   13176
      ButtonWidth     =   953
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "optinet"
      DisabledImageList=   "optinet"
      HotImageList    =   "optinet"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vorige"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "volgende"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vernieuwen"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "home"
            Object.ToolTipText     =   "Open start page"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fullscreen"
            Object.ToolTipText     =   "Change webbrowser size"
            ImageIndex      =   6
            Style           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "norefresh"
            Object.ToolTipText     =   "Refresh webpage preview"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "popups"
            Object.ToolTipText     =   "Display popup windows"
            ImageIndex      =   8
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "spepe"
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "segregerg"
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gregherghr"
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "editsite"
            Object.ToolTipText     =   "Webpagina bewerken"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Left            =   12015
      ScaleHeight     =   585
      ScaleWidth      =   600
      TabIndex        =   22
      Top             =   1650
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   45
      TabIndex        =   19
      ToolTipText     =   "Double click on a line to insert the symbol or command."
      Top             =   1305
      Width           =   1770
   End
   Begin VB.ComboBox comboInvoegen 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   18
      ToolTipText     =   "Choose a category here"
      Top             =   960
      Width           =   1770
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   11460
      Top             =   6870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   6333
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":0C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":0FCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":135B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":16E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":1A69
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":1DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":21F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   10515
      Top             =   6930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   6858
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   6333
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":2938
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":2D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":314C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":3569
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":3986
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":3D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":419C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":4532
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":491A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":4D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":50BF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   435
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   767
      ButtonWidth     =   767
      ButtonHeight    =   714
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList5"
      DisabledImageList=   "ImageList5"
      HotImageList    =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vet"
            Description     =   "vet"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cursief"
            Description     =   "cursief"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underl"
            Description     =   "onderl"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "links"
            Description     =   "links"
            Object.ToolTipText     =   "Align left"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "centr"
            Description     =   "centreer"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rechts"
            Description     =   "rechts"
            Object.ToolTipText     =   "Align right"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "font"
            Description     =   "lettertype"
            Object.ToolTipText     =   "Font"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lijst"
            Description     =   "lijst"
            Object.ToolTipText     =   "List"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   7455
         TabIndex        =   28
         Text            =   "Combo5"
         ToolTipText     =   "Word list"
         Top             =   45
         Width           =   1680
      End
      Begin VB.TextBox txtSize 
         Height          =   345
         Left            =   6615
         TabIndex        =   16
         Text            =   "Text1"
         ToolTipText     =   "Font size"
         Top             =   30
         Width           =   390
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         LargeChange     =   4
         Left            =   6990
         Max             =   2
         Min             =   200
         TabIndex        =   15
         Top             =   45
         Value           =   12
         Width           =   315
      End
      Begin VB.ComboBox FontList 
         Height          =   315
         Left            =   4035
         Sorted          =   -1  'True
         TabIndex        =   14
         Text            =   "Combo5"
         ToolTipText     =   "Font"
         Top             =   45
         Width           =   2490
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   767
      ButtonWidth     =   767
      ButtonHeight    =   714
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList4"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nieuw"
            Description     =   "nieuw"
            Object.ToolTipText     =   "New file"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Description     =   "open"
            Object.ToolTipText     =   "Open file"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "opslaan"
            Description     =   "opslaan"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mail"
            Description     =   "mail"
            Object.ToolTipText     =   "Send as e-mail"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "afdrukken"
            Description     =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "voorbeeld"
            Description     =   "preview"
            Object.ToolTipText     =   "Print preview"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Description     =   "htmlpreview"
            Object.ToolTipText     =   "Webpage preview"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "knippen"
            Description     =   "knippen"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kopieren"
            Description     =   "kopieren"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "plakken"
            Description     =   "plakken"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Undo"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Description     =   "undo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "redo"
            Description     =   "redo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         Height          =   285
         Left            =   10590
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Start the search."
         Top             =   90
         Width           =   810
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8505
         TabIndex        =   27
         Text            =   "in"
         Top             =   105
         Width           =   240
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   8790
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Select the current file or an internet search engine."
         Top             =   75
         Width           =   1710
      End
      Begin VB.TextBox txtZoeken 
         Height          =   285
         Left            =   6435
         TabIndex        =   25
         ToolTipText     =   "Type in what you want to search for and cliick Search."
         Top             =   75
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5730
         TabIndex        =   24
         Text            =   "Find:"
         ToolTipText     =   "Search the current file or the internet."
         Top             =   90
         Width           =   765
      End
   End
   Begin VB.Timer X 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11265
      Top             =   4440
   End
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   12015
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   11280
      Top             =   4875
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   8340
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GO"
      Height          =   300
      Left            =   6225
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Browse to this adress"
      Top             =   6585
      Width           =   525
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   2055
      TabIndex        =   8
      Text            =   "Type in a web adress and click Go..."
      Top             =   6570
      Width           =   4170
   End
   Begin RichTextLib.RichTextBox SelToHTML 
      Height          =   735
      Left            =   11910
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"OptiWord.frx":5443
   End
   Begin RichTextLib.RichTextBox SelectionBox 
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"OptiWord.frx":54C5
   End
   Begin RichTextLib.RichTextBox HTMLBox 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"OptiWord.frx":5547
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   9555
      Top             =   6930
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   8640
      Top             =   6915
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.TextBox DocName 
      Height          =   285
      Left            =   7935
      TabIndex        =   1
      Text            =   "File not saved yet!"
      ToolTipText     =   "Document filename"
      Top             =   6450
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   12030
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   12030
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11985
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2775
      Left            =   1935
      TabIndex        =   5
      ToolTipText     =   "Webbrowser"
      Top             =   3690
      Visible         =   0   'False
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox ImageList1 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   11295
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   975
      Visible         =   0   'False
      Width           =   1200
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5400
      Left            =   1905
      TabIndex        =   0
      Top             =   1065
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9525
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      BulletIndent    =   400
      Appearance      =   0
      RightMargin     =   4
      TextRTF         =   $"OptiWord.frx":55C9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList optinet 
      Left            =   12300
      Top             =   6870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   29
      ImageHeight     =   30
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":564D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":5B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":60A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":66C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":6DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":7396
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":7846
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":7D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":817C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":8636
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   29
      ImageHeight     =   30
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":8A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":8FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":94F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":9B11
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OptiWord.frx":A20F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox SaveHTML 
      Height          =   615
      Left            =   4230
      TabIndex        =   31
      Top             =   7470
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"OptiWord.frx":A7E7
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OPEN URL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   690
      Left            =   1185
      TabIndex        =   7
      Top             =   7035
      Width           =   900
   End
   Begin VB.Label lbluitleg 
      BackStyle       =   0  'Transparent
      Height          =   4305
      Left            =   165
      TabIndex        =   21
      Top             =   6240
      Width           =   1785
   End
   Begin VB.Label lblOnderwerp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OptiType 3.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15
      TabIndex        =   20
      Top             =   5850
      Width           =   1755
   End
   Begin VB.Label lblDatum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6855
      TabIndex        =   3
      ToolTipText     =   "Huidige datum"
      Top             =   6570
      Width           =   930
   End
   Begin VB.Label Label2 
      Height          =   9720
      Left            =   15
      TabIndex        =   17
      Top             =   870
      Width           =   1905
   End
   Begin VB.Menu pagina 
      Caption         =   "&Webpage"
      Visible         =   0   'False
      Begin VB.Menu web2editpage 
         Caption         =   "&Edit webpage"
      End
      Begin VB.Menu fezfezfezfezfezf 
         Caption         =   "-"
      End
      Begin VB.Menu web2openpagefromdisk 
         Caption         =   "&Open from disk..."
      End
      Begin VB.Menu ppl 
         Caption         =   "-"
      End
      Begin VB.Menu web2printpage 
         Caption         =   "&Print"
      End
      Begin VB.Menu web2savetodisk 
         Caption         =   "&Save to disk..."
      End
      Begin VB.Menu web2showsource 
         Caption         =   "&View source"
      End
      Begin VB.Menu web2zoekoppagina 
         Caption         =   "&Find on this page..."
      End
      Begin VB.Menu fezopkfezfezfez 
         Caption         =   "-"
      End
      Begin VB.Menu web2close 
         Caption         =   "&Close webbrowser"
      End
   End
   Begin VB.Menu mnuBestand 
      Caption         =   "&File"
      Begin VB.Menu mnuNieuw 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu opensystemfile 
         Caption         =   "&Edit system file"
         Begin VB.Menu openautoexec 
            Caption         =   "Autoexec.bat"
         End
         Begin VB.Menu autoexecnt 
            Caption         =   "Autoexec.nt"
         End
         Begin VB.Menu jfzemjfz 
            Caption         =   "-"
         End
         Begin VB.Menu openconfigsys 
            Caption         =   "Config.sys"
         End
         Begin VB.Menu confignt 
            Caption         =   "Config.nt"
         End
         Begin VB.Menu fiqzef 
            Caption         =   "-"
         End
         Begin VB.Menu openwinini 
            Caption         =   "Win.ini"
         End
         Begin VB.Menu opensystemini 
            Caption         =   "System.ini"
         End
         Begin VB.Menu openprotocolini 
            Caption         =   "Protocol.ini"
         End
         Begin VB.Menu fefze 
            Caption         =   "-"
         End
         Begin VB.Menu openmsdossys 
            Caption         =   "Msdos.sys"
         End
      End
      Begin VB.Menu quicksave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpslaan 
         Caption         =   "Save &as"
         Shortcut        =   {F12}
      End
      Begin VB.Menu saveashtml 
         Caption         =   "Save as &webpage..."
      End
      Begin VB.Menu streepje 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAfdrukken 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu afdrukvoorbeeld 
         Caption         =   "P&rint preview"
      End
      Begin VB.Menu htmlvoorbeeld 
         Caption         =   "Preview as we&bpage"
      End
      Begin VB.Menu batchpreview 
         Caption         =   "Preview as ba&tchfile"
         Enabled         =   0   'False
      End
      Begin VB.Menu UltimateStreepke 
         Caption         =   "-"
      End
      Begin VB.Menu mail 
         Caption         =   "&Send as e-mail..."
      End
      Begin VB.Menu po 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuBewerken 
      Caption         =   "&Edit"
      Begin VB.Menu undoit 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redoit 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu letsallgotothelobby 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKnippen 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuKopieren 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPlakken 
         Caption         =   "&Paste"
      End
      Begin VB.Menu remove 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu streepje3 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select all"
         Shortcut        =   ^A
      End
      Begin VB.Menu nogeenstreepke 
         Caption         =   "-"
      End
      Begin VB.Menu Zoeken 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu vervangen 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu beeld 
      Caption         =   "&View"
      Begin VB.Menu standaardwerkbalk 
         Caption         =   "Standard toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu opmaakbalk 
         Caption         =   "Font toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu spatierules 
         Caption         =   "-"
      End
      Begin VB.Menu htmlpreview 
         Caption         =   "&Preview as webpage"
         Begin VB.Menu htmlview 
            Caption         =   "&Show preview as webpage"
            Checked         =   -1  'True
         End
         Begin VB.Menu htmlprevopt 
            Caption         =   "&Advanced options..."
         End
         Begin VB.Menu autorefresh 
            Caption         =   "&Refresh preview"
            Shortcut        =   {F5}
         End
         Begin VB.Menu restore 
            Caption         =   "&Restore focus to document"
         End
         Begin VB.Menu refreshmanual 
            Caption         =   "&Manually refresh"
         End
      End
      Begin VB.Menu browseroptions 
         Caption         =   "&Webbrowser"
         Begin VB.Menu showpopups 
            Caption         =   "&Display popup windows"
            Checked         =   -1  'True
         End
         Begin VB.Menu browseoffline 
            Caption         =   "&Browse off-line"
            Checked         =   -1  'True
         End
         Begin VB.Menu browstreep 
            Caption         =   "-"
         End
         Begin VB.Menu back 
            Caption         =   "&Back"
         End
         Begin VB.Menu gotonext 
            Caption         =   "&Forward"
         End
         Begin VB.Menu refreshbrowser 
            Caption         =   "R&efresh"
         End
         Begin VB.Menu searchforapage 
            Caption         =   "&Find"
         End
      End
      Begin VB.Menu ohmygodletsgotokennysfuneral 
         Caption         =   "-"
      End
      Begin VB.Menu fullscr 
         Caption         =   "&Full screen"
      End
   End
   Begin VB.Menu mnuOpmaak 
      Caption         =   "&Font"
      Begin VB.Menu mnuVet 
         Caption         =   "&Bold"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCursief 
         Caption         =   "&Italic"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuonderl 
         Caption         =   "&Underlined"
         Checked         =   -1  'True
         Shortcut        =   ^U
      End
      Begin VB.Menu m_syntax 
         Caption         =   "&Mark HTML tags"
         Checked         =   -1  'True
      End
      Begin VB.Menu groovy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuperscript 
         Caption         =   "&Superscript"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSubscript 
         Caption         =   "S&ubscript"
         Checked         =   -1  'True
      End
      Begin VB.Menu groovy2 
         Caption         =   "-"
      End
      Begin VB.Menu choosefont 
         Caption         =   "Choose font"
         Begin VB.Menu fonts 
            Caption         =   "Fonts"
            Index           =   0
         End
      End
      Begin VB.Menu streepke 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLettert 
         Caption         =   "&Font..."
         Shortcut        =   ^D
      End
      Begin VB.Menu achtergrond 
         Caption         =   "&Background..."
      End
      Begin VB.Menu marges 
         Caption         =   "&Margins..."
      End
      Begin VB.Menu ohmygodkennyisdead 
         Caption         =   "-"
      End
      Begin VB.Menu align 
         Caption         =   "&Alignment"
         Begin VB.Menu links 
            Caption         =   "&Left"
            Shortcut        =   ^L
         End
         Begin VB.Menu midden 
            Caption         =   "&Center"
            Shortcut        =   ^E
         End
         Begin VB.Menu rechts 
            Caption         =   "&Right"
            Shortcut        =   ^R
         End
      End
   End
   Begin VB.Menu Invoegen 
      Caption         =   "&Insert"
      Begin VB.Menu datum 
         Caption         =   "&Date"
      End
      Begin VB.Menu tijdinvoegenindocument 
         Caption         =   "&Time"
      End
      Begin VB.Menu auipzsdza 
         Caption         =   "-"
      End
      Begin VB.Menu txt 
         Caption         =   "Text file"
      End
      Begin VB.Menu afbeelding 
         Caption         =   "&Image"
         Begin VB.Menu bestand 
            Caption         =   "From &file..."
         End
      End
      Begin VB.Menu oleinsertfile 
         Caption         =   "&File (Word, Excel, ...)"
      End
      Begin VB.Menu nogmaareenstreepkezeker 
         Caption         =   "-"
      End
      Begin VB.Menu tabelinvoegen 
         Caption         =   "&Table..."
      End
      Begin VB.Menu yeahbabyueafezf 
         Caption         =   "-"
      End
      Begin VB.Menu htmlcode 
         Caption         =   "&HTML"
         Begin VB.Menu htmlimage 
            Caption         =   "&Image"
         End
         Begin VB.Menu horline 
            Caption         =   "&Horizontal ruler"
         End
         Begin VB.Menu link 
            Caption         =   "H&yperlink..."
            Shortcut        =   ^K
         End
         Begin VB.Menu marquee 
            Caption         =   "&Marquee"
         End
         Begin VB.Menu insertHTMLtabel 
            Caption         =   "&Table..."
         End
      End
      Begin VB.Menu batchfile 
         Caption         =   "&Batch-command"
         Enabled         =   0   'False
         Begin VB.Menu cls 
            Caption         =   "Clear screen(CLS)"
         End
         Begin VB.Menu echoit 
            Caption         =   "&Display message (ECHO)..."
         End
         Begin VB.Menu pauzeer 
            Caption         =   "Press any key... (PAUSE)"
         End
         Begin VB.Menu copyafile 
            Caption         =   "Copy a file (COPY)..."
         End
         Begin VB.Menu xcopy 
            Caption         =   "Copy a file (advanced) (XCOPY)..."
         End
         Begin VB.Menu batt 
            Caption         =   "-"
         End
         Begin VB.Menu format 
            Caption         =   "&Format disk (FORMAT)..."
         End
         Begin VB.Menu setdateinbatch 
            Caption         =   "&Set or show date (DATE)"
         End
         Begin VB.Menu settime 
            Caption         =   "&Set or show time (TIME)"
         End
         Begin VB.Menu mem 
            Caption         =   "&Display memory information (MEM)"
         End
         Begin VB.Menu popopoezfe 
            Caption         =   "-"
         End
         Begin VB.Menu insertlabel 
            Caption         =   "&Label"
         End
         Begin VB.Menu gotolabel 
            Caption         =   "&Goto label"
         End
         Begin VB.Menu tobesuxdoesitnot 
            Caption         =   "-"
         End
         Begin VB.Menu ifcommand 
            Caption         =   "&If-command"
         End
         Begin VB.Menu ifexist 
            Caption         =   "If &Exist-command"
         End
      End
   End
   Begin VB.Menu letsdothesel 
      Caption         =   "&Selection"
      Begin VB.Menu wordcountofsel 
         Caption         =   "&Word count..."
      End
      Begin VB.Menu convertitplease 
         Caption         =   "&Convert"
         Begin VB.Menu makeitbig 
            Caption         =   "To &upper case"
         End
         Begin VB.Menu makeitsmall 
            Caption         =   "To &lower case"
         End
      End
      Begin VB.Menu thetruthisoutthere 
         Caption         =   "-"
      End
      Begin VB.Menu saveselection 
         Caption         =   "&Save selection..."
      End
      Begin VB.Menu sendselection 
         Caption         =   "&Send as e-mail..."
      End
   End
   Begin VB.Menu Extra 
      Caption         =   "E&xtra"
      Begin VB.Menu wordlist 
         Caption         =   "Word list"
         Begin VB.Menu makelist 
            Caption         =   "&Create"
         End
         Begin VB.Menu makelistandsave 
            Caption         =   "Create and &save as file..."
         End
      End
      Begin VB.Menu yeahbabytdezfzefze 
         Caption         =   "-"
      End
      Begin VB.Menu wordcountplease 
         Caption         =   "&Word count"
      End
      Begin VB.Menu listwithshortcuts 
         Caption         =   "&List shortcuts"
      End
      Begin VB.Menu overtype 
         Caption         =   "&Enable or disable overtype"
      End
      Begin VB.Menu poppo 
         Caption         =   "-"
      End
      Begin VB.Menu briefwizard 
         Caption         =   "&Letter wizard..."
      End
      Begin VB.Menu checkitoutnow 
         Caption         =   "-"
      End
      Begin VB.Menu aanpassen 
         Caption         =   "&Adjust"
         Begin VB.Menu aanpassenstandaard 
            Caption         =   "&Standard"
         End
         Begin VB.Menu aanpassenopmaak 
            Caption         =   "&Font"
         End
      End
      Begin VB.Menu options 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu Info 
         Caption         =   "&Info..."
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OptiType VERSION 3.0
' LANGUAGE: DUTCH
' PROGRAMMING LANGUAGE: MICROSOFT VISUAL BASIC 6.0 PROFESSIONAL


'Option Explicit


' undo & redo variables:
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Long
Dim gstrStack(1000) As String

Dim undoint As Integer


' other variablen:
Dim allowpopup As Boolean  ' webbrowser
Dim sec As Integer  ' timer
Public bg As String    ' background
Public annuleer As Boolean
Dim autoopslaan As Integer








Private Sub aanpassenopmaak_Click()
Toolbar2.Customize

End Sub

Private Sub aanpassenstandaard_Click()
Toolbar1.Customize

End Sub

Private Sub achtergrond_Click()
On Error Resume Next
    CommonDialog1.ShowColor
    RichTextBox1.BackColor = CommonDialog1.Color
    If language = False Then MsgBox "De achtergrondcolor van een document wordt noch opgeslagen, nog afgedrukt."
    If language = True Then MsgBox "Backgroundcolor won't be saved or printed."
    
End Sub



Private Sub drukaf()
PrintDialog.CancelError = True
On Error GoTo printcancelled
PrintDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If RichTextBox1.SelLength = 0 Then
        PrintDialog.Flags = PrintDialog.Flags + cdlPDAllPages
    Else
        PrintDialog.Flags = PrintDialog.Flags + cdlPDSelection
    End If
    PrintDialog.ShowPrinter
    'Printer.Print ""
    RichTextBox1.SelPrint PrintDialog.hdc

printcancelled:


End Sub

Private Sub Afdrukken_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Display the printer dialog window."

End Sub

Private Sub afdrukvoorbeeld_Click()
 PrintPreview RichTextBox1, 1, 1, 1, 1, 1
End Sub




Private Sub autoexecnt_Click()

On Error GoTo noautont
checktosave
If annuleer = False Then
    RichTextBox1.LoadFile GetWinPath & "\system32\autoexec.nt", 1
    form1.Caption = "OptiType 3.0 - Edit system file:Autoexec.nt"
    DocName.Text = GetWinPath & "\system32\autoexec.nt"
Else
    annuleer = False
End If

Exit Sub

noautont:
MsgBox "The file AUTOEXEC.NT has not been found on your computer.  MThis file only exists if you're running Windows NT, 2000 or XP.", vbCritical, "OptiType"

End Sub

Private Sub autorefresh_Click()
On Error Resume Next

focusherstellen
RichTextBox1.SaveFile "c:\temp.htm", rtfText
    
    
    WebBrowser1.Refresh
   RichTextBox1.SetFocus


End Sub

Private Sub AutoSave_Timer()

On Error GoTo noautosave

If frmOpties.Check2.Value = 1 Then
'MsgBox autoopslaan & " - " & frmOpties.Text1.Text

autoopslaan = autoopslaan + 1

If autoopslaan >= Val(frmOpties.Text1.Text) Then
    autoopslaan = 0
  '  MsgBox "saving"
    
'
If frmOpties.Check3.Value = 1 Then

    a = MsgBox("Het is tijd om uw document op te slaan.  Wilt u dit nu doen?" & vbCrLf & vbCrLf & "(U kunt deze instelling aanpassen via Opties in het menu Extra)", vbYesNo, "Automatisch opslaan")
    If a <> 6 Then Exit Sub
End If
 '
    If Not DocName.Text = "File not saved yet!" Then

    If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
    RichTextBox1.SaveFile DocName.Text, 0
    Else
    RichTextBox1.SaveFile DocName.Text, 1
    End If

    End If
    
End If
End If

Exit Sub

noautosave:



End Sub

Private Sub back_Click()
On Error GoTo cantdoit

WebBrowser1.GoBack
Exit Sub

cantdoit:
MsgBox "Fout bij het inladen van vorige pagina.", vbCritical


End Sub

Private Sub batchpreview_Click()

RichTextBox1.SaveFile App.Path & "\PREVIEW.BAT", rtfText
Shell App.Path & "\PREVIEW.BAT", 1


End Sub

Private Sub bestand_Click()


' Set CancelError is True
    

    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
      CommonDialog2.Flags = cdlOFNHideReadOnly
      CommonDialog2.FileName = ""
    ' Set filters
 '   CommonDialog2.Filter = "Windows bitmap (*.BMP)|*.bmp|Compuserve GIF" & _
 '   "(*.gif)|*.gif|All files|*.*"
    CommonDialog2.Filter = "Windows bitmap (*.BMP)|*.bmp|Compuserve GIF" & _
    "(*.gif)|*.gif|JPEG Filter (*.JPG)|*.jpg|All files|*.*"
  
    ' Specify default filter
    CommonDialog2.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog2.ShowOpen
    
    Picture1.Picture = LoadPicture(CommonDialog2.FileName)
    
    'RichTextBox1.OLEObjects.Add , , CommonDialog2.FileName
        
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    SendMessage RichTextBox1.hWnd, &H302, 0, 0
        
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub

Private Sub binnenkort_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Tot binnenkort"
End Sub

Private Sub briefwizard_Click()

checktosave

If annuleer = False Then frmBrief1.Show vbModeless, Me






End Sub

Private Sub browseoffline_Click()
If browseoffline.Checked = True Then
    browseoffline.Checked = False
    WebBrowser1.Offline = False
Else
    browseoffline.Checked = True
    WebBrowser1.Offline = True
End If

End Sub



Private Sub centreer()

RichTextBox1.RightMargin = RichTextBox1.Width
' ' wordwrap.checked = True
RichTextBox1.SelAlignment = 2

On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub Centreren_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst centreren"

End Sub

Private Sub chkCursief_Click()

If chkCursief = 1 Then RichTextBox1.SelItalic = True: mnuCursief.Checked = True
If chkCursief = 0 Then RichTextBox1.SelItalic = False: mnuCursief.Checked = False


End Sub

Private Sub chkCursief_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst cursief maken"

End Sub

Private Sub chkOnderl_Click()
If chkOnderl = 1 Then RichTextBox1.SelUnderline = True: mnuonderl.Checked = True
If chkOnderl = 0 Then RichTextBox1.SelUnderline = False: mnuonderl.Checked = False


End Sub

Private Sub chkOnderl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst onderlijnen"

End Sub

Private Sub chkVET_Click()

If chkVET = 1 Then RichTextBox1.SelBold = True: mnuVet.Checked = True
If chkVET = 0 Then RichTextBox1.SelBold = False: mnuVet.Checked = False
End Sub

Private Sub chkVET_GotFocus()
If chkVET = 1 Then
    chkVET = 0
    Else
    chkVET = 1
End If
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub chkCursief_GotFocus()
If chkCursief = 1 Then
    chkCursief = 0
    Else
    chkCursief = 1
End If
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub chkOnderl_GotFocus()
If chkOnderl = 1 Then
    chkOnderl = 0
    Else
    chkOnderl = 1
End If
On Error Resume Next: RichTextBox1.SetFocus
End Sub







Private Sub chkVET_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst vet maken"

End Sub



Private Sub classicbg_Click()
form1.Picture = LoadPicture(App.Path & "\achtergr.jpg")

classicbg.Checked = True
millenniumbg.Checked = False
introbg.Checked = False


End Sub

Private Sub clipart_Click()

frmClipart.Show vbModeless, Me

End Sub



Private Sub cls_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "CLS" & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus

End Sub

Private Sub knip()

Clipboard.Clear


On Error Resume Next
Clipboard.SetText RichTextBox1.SelText


RichTextBox1.SelText = ""
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub cmdKnippen_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst op het Windows klembord plaatsen en daarna verwijderen"

End Sub


Private Sub cmdKopieren_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst op het Windows klembord plaatsen"

End Sub

Private Sub plak()

If Clipboard.GetFormat(vbCFText) Then
selectie = Clipboard.GetText
RichTextBox1.SelText = RichTextBox1.SelText & selectie
End If

If Clipboard.GetFormat(vbCFBitmap) Then
    'SendKeys "^V"
    RichTextBox1.OLEObjects.Add , , Clipboard.GetData
    
    
    
End If
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub cmdPlakken_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Tekst invoegen van het Windows klembord"

End Sub

Private Sub redo()

    'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub



Private Sub cmdSendMail_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Uw document verzenden via e-mail"

End Sub

Private Sub undo()

    Toolbar1.Buttons.Item(14).Enabled = True
    
    
    
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub comboInvoegen_Click()

List1.BackColor = vbWhite
If comboInvoegen.ListIndex = 0 Then htmltags
If comboInvoegen.ListIndex = 1 Then dosinvoegen
If comboInvoegen.ListIndex = 2 Then ansi
If comboInvoegen.ListIndex = 3 Then autotekstinvoegen


End Sub

Private Sub comboInvoegen_DropDown()
List1.Clear
List1.BackColor = &H8000000F


End Sub



Private Sub Command3_Click()

Shell App.Path & "\optifile.exe", 1


End Sub

Private Sub Command4_Click()

Shell App.Path & "\optidr.exe", 1


End Sub



Private Sub Combo4_Click()

RichTextBox1.SelText = RichTextBox1.SelText & "<FONT FACE='" & Combo4.Text & "'>"
On Error Resume Next: RichTextBox1.SetFocus

End Sub

Private Sub Command1_Click()

zoeknu



End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Refresh the word list"

End Sub

Private Sub Command2_Click()

If WebBrowser2.Visible = False Then


    
    webbuttons.Buttons.Item(10).Value = tbrUnpressed
      
    form1.htmlview.Checked = True

    If WebBrowser1.Visible = False Then

            form1.RichTextBox1.Height = Fix(form1.RichTextBox1.Height / 2)
            form1.WebBrowser1.Visible = True
            WebBrowser1.Height = RichTextBox1.Height
            htmlview.Checked = True
            frmOpties.Check1.Value = 1
    End If

    showpopups.Enabled = True
    browseoffline.Enabled = True
    back.Enabled = True
    gotonext.Enabled = True
    refreshbrowser.Enabled = True
    searchforapage.Enabled = True

    WebBrowser1.Navigate txtURL.Text

Else

    WebBrowser2.Navigate txtURL.Text

End If



End Sub

Private Sub Command5_Click()

Shell App.Path & "\optinet.exe", 1


End Sub

Private Sub Command7_Click()
Speech1.Speak RichTextBox1.Text

End Sub

Private Sub confignt_Click()
On Error GoTo noconfnt
checktosave
If annuleer = False Then
    RichTextBox1.LoadFile GetWinPath & "\system32\config.nt", 1
    form1.Caption = "OptiType 3.0 - Edit system file:Config.nt"
    DocName.Text = GetWinPath & "\system32\config.nt"
Else
    annuleer = False
End If

Exit Sub

noconfnt:
MsgBox "The file CONFIG.NT has not been found on your computer.  MThis file only exists if you're running Windows NT, 2000 or XP.", vbCritical, "OptiType"

End Sub

Private Sub copyafile_Click()
a = InputBox("File(s) to copy?", "COPY")
b = InputBox("Target folder or files?", "COPY")

RichTextBox1.SelText = RichTextBox1.SelText & "COPY " & a & " " & b & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub datum_Click()

RichTextBox1.SelText = RichTextBox1.SelText & Date$
End Sub

Private Sub DocName_GotFocus()
On Error Resume Next
On Error Resume Next: RichTextBox1.SetFocus

End Sub


Private Sub echoit_Click()

a = InputBox("Type in the message you would like to display.", "ECHO")
RichTextBox1.SelText = RichTextBox1.SelText & "ECHO " & a & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus
End Sub



Private Sub exit_Click()



SendKeys "%{F4}", True


'On Error Resume Next
'Kill "c:\temp.htm"
'Unload Me
'End

End Sub

Private Sub faster_Click()
Speech1.Speed = Speech1.Speed + 5

End Sub

Private Sub FontList_Change()
RichTextBox1.SelFontName = FontList.Text

End Sub

Private Sub fontlist_Click()
RichTextBox1.SelFontName = FontList.Text
On Error Resume Next: RichTextBox1.SetFocus




End Sub





Private Sub FontList_KeyPress(KeyAscii As Integer)

'MsgBox KeyAscii

If KeyAscii = 13 Then   ' enter
    RichTextBox1.SelFontName = FontList.Text
    On Error Resume Next: RichTextBox1.SetFocus
End If
End Sub

Private Sub fonts_Click(Index As Integer)
RichTextBox1.SelFontName = fonts(Index).Caption

End Sub

Private Sub force_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "May The Force Be With You"
End Sub




Private Sub meer_Click()


End Sub

Private Sub Form_Resize()

nieuwformaat


End Sub

Private Sub lblDatum_Click()
MsgBox "You can hide the date with Options in the Extra menu.", vbInformation, "OptiType"
End Sub

Private Sub List1_Click()

If comboInvoegen.ListIndex = 1 Then   ' dos commando
    
    lblOnderwerp.Caption = List1.Text
    
    
    Select Case List1.Text

    Case "ATTRIB"
            lbluitleg.Caption = "Display or chance file properties." & vbCrLf & vbCrLf & "ATTRIB [+R | -R] [+A | -A ] [+S | -S] [+H | -H]   [[drive:][path] filename] [/S] [/D]]" & vbCrLf & vbCrLf & "R: read only" & vbCrLf & "A: archive" & vbCrLf & "S: system" & vbCrLf & "H: hidden" & vbCrLf & "/S: include files in subfolders" & vbCrLf & "/D: include subfolders"
        
        
    Case "BREAK"
                lbluitleg.Caption = "Enable or disable advanced CTRL-checking"
        
    Case "CALL"
            lbluitleg.Caption = "Run another batch file, then return to this one." & vbCrLf & vbCrLf & "CALL [drive:][path]filename.bat [batchparameters]"
                
        
    Case "CD"
            lbluitleg.Caption = "Display or change the current directory" & vbCrLf & vbCrLf & "CHDIR [/D] [drive:][path]" & vbCrLf & "ChDir [..]" & vbCrLf & "CD [/D] [drive:][path]" & vbCrLf & "CD [..]" & vbCrLf & vbCrLf & ".. : upper folder" & vbCrLf & "/D: Change current drive AND path"
        
    Case "CHDIR"
                lbluitleg.Caption = "Display or change the current directory" & vbCrLf & vbCrLf & "CHDIR [/D] [drive:][path]" & vbCrLf & "ChDir [..]" & vbCrLf & "CD [/D] [drive:][path]" & vbCrLf & "CD [..]" & vbCrLf & vbCrLf & ".. : upper folder" & vbCrLf & "/D: Change current drive AND path"
        
    Case "CHKDSK"
            lbluitleg.Caption = "Run CHKDSK (resembles SCANDISk, only WinNT/2000/XP)." & vbCrLf & vbCrLf & "CHKDSK [volume[[path]filename]]] [/F] [/V] [/R] [/X] [/I] [/C] [/L[:size]]" & vbCrLf & vbCrLf & "/F: repair errors" & vbCrLf & "/R: Find and repair damaged clusters" & vbCrLf & "/V: Show filenames"
        
    Case "CLS"
            lbluitleg.Caption = "Clear the screen"
        
    Case "COLOR"
            lbluitleg.Caption = "Change foreground and backgroundcolor.." & vbCrLf & vbCrLf & "0 = black" & vbCrLf & "1 = blue" & vbCrLf & "2 = green" & vbCrLf & "3 = blue" & vbCrLf & "4 = red" & vbCrLf & "5 = purple" & vbCrLf & "6 = yellow" & vbCrLf & "7 = grey" & vbCrLf & "F = white" & vbCrLf & vbCrLf & "Vb: COLOR f1 is blue on white bg."
        
    Case "COPY"
            lbluitleg.Caption = "Copy one or more files to another location." & vbCrLf & vbCrLf & "COPY [/D] [/V] [/N] [/Y | /-Y] [/Z] [/A | /B] bron [/A | /B] [+ source [/A | /B] [+ ...]] [target [/A | /B]]" & vbCrLf & vbCrLf & "/A: textfiles" & vbCrLf & "/B: binair files" & vbCrLf & "/D: decode target file" & vbCrLf & "/V: check file(s)" & vbCrLf & "/N: short filenames" & vbCrLf & "/Y: don't ask to overwrite" & vbCrLf & "/-Y: ask to overwrite"
        
    Case "DATE"
            lbluitleg.Caption = "Display or change the date."
            
    Case "DEL"
            lbluitleg.Caption = "Delete one or more files." & vbCrLf & vbCrLf & "DEL [/P] [/F] [/S] [/Q] [/A[[:]attributes]] names" & vbCrLf & vbCrLf & "/P: ask for confirmation" & vbCrLf & "/F: force read-only files" & vbCrLf & "/S: delete from all subfolders" & vbCrLf & "/Q: don't ask for confirmation" & vbCrLf & "/A: Select by attributes"
            
    Case "DIR"
            lbluitleg.Caption = "View list of files and subfolders in the current folder" & vbCrLf & vbCrLf & "DIR [drive:][path][filename] [/P] [/W] [/D] [/A[[:]attributes]] [/O[[:]order]] [/T[[:]time]] [/S] [/B] [/L] [/N] [/X] [/C]" & vbCrLf & vbCrLf & "/A: show files with certain attributes" & vbCrLf & "/B: text-only" & vbCrLf & "/C: seperate 000's" & vbCrLf & "/-C: don't seperate 000's" & vbCrLf & "/D: sort by column" & vbCrLf & "/L: lower case"
    
    Case "DISKCOPY"
            lbluitleg.Caption = "Copy one disk to another." & vbCrLf & vbCrLf & "DISKCOPY [drive1: [drive2:]] [/V]" & vbCrLf & vbCrLf & "/V: Check copied files for errors."
            
    Case "DOSKEY"
            lbluitleg.Caption = "Create macro's, edit command lines and more." & vbCrLf & vbCrLf & "DOSKEY [/REINSTALL] [/LISTSIZE=size] [/MACROS[:ALL | :EXE-name]] [/HISTORY] [/INSERT | /OVERSTRIKE] [/EXENAME=EXE-name] [/MACROFILE=filename] [macroname=[text]]" & vbCrLf & vbCrLf & "/REINSTALL: load new copy of DOSKEY" & vbCrLf & "/HISTORY: display all saved commands"
            
    Case "DISKCOMP"
            lbluitleg.Caption = "Compare the contents of two floppy disks." & vbCrLf & vbCrLf & "DISKCOMP [drive1: [drive2:]]"
            
    Case "ECHO"
            lbluitleg.Caption = "Display a message, or enable or disable the ECHO command." & vbCrLf & vbCrLf & "ECHO [ON | OFF]" & vbCrLf & "ECHO [message]" & vbCrLf & vbCrLf & "Type @ before a command to hide the command itself.  Example: @ECHO OFF."
            
    Case "ERASE"
            lbluitleg.Caption = "Delete one or more files." & vbCrLf & vbCrLf & "ERASE [/P] [/F] [/S] [/Q] [/A[[:]attributes]] names" & vbCrLf & vbCrLf & "/P: ask for confirmation" & vbCrLf & "/F: force read-only files" & vbCrLf & "/S: delete from all subfolders" & vbCrLf & "/Q: don't ask for confirmation" & vbCrLf & "/A: Select files with certain attributes"
            
      Case "EXIT"
        lbluitleg.Caption = "Closes the command prompt. Use EXIT /B to close only the batch file."
            
    Case "FIND"
        lbluitleg.Caption = "Search for a string in one or more files." & vbCrLf & vbCrLf & "FIND [/V] [/C] [/N] [/I] [/OFF[LINE]] 'string' [[drive:][path]filename[ ...]]" & vbCrLf & vbCrLf & "/V: display all lines WITHOUT the search string" & vbCrLf & "/C: display only the number of lines found" & vbCrLf & "/N: display line numbers" & vbCrLf & "/I: not case sensitive"
        
    Case "FOR"
        lbluitleg.Caption = "Create a loop." & vbCrLf & vbCrLf & "FOR %variable IN (set) DO command [parameters]" & vbCrLf & vbCrLf & "set = one or more files." & vbCrLf & vbCrLf & "Use FOR /? for more information."
        
    Case "FORMAT"
        lbluitleg.Caption = "Format a drive." & vbCrLf & vbCrLf & "FORMAT volume: [/FS:filesystem] [/V:volumename] [/Q] [/A:size] [/C]" & vbCrLf & vbCrLf & "/FS: FAT, FAT32 or NTFS" & vbCrLf & "/Q: Quick format" & vbCrLf & "/C: Use compression (NTFS)" & vbCrLf & "/X: Detach volume first" & vbCrLf & "/F: size" & vbCrLf & "/N: number of sectors."
            
    Case "GOTO"
        lbluitleg.Caption = "Jump to a label in a batch file." & vbCrLf & vbCrLf & "GOTO label" & vbCrLf & vbCrLf & "Example: " & vbCrLf & vbCrLf & "GOTO start" & vbCrLf & vbCrLf & ":start"
            
    Case "IF"
        lbluitleg.Caption = "Run an optional command." & vbCrLf & vbCrLf & "IF [NOT] string1==string2 command" & vbCrLf & vbCrLf & "IF [NOT] ERRORLEVEL number command" & vbCrLf & vbCrLf & "IF [NOT] EXIST filename command" & vbCrLf & vbCrLf & "Type IF /? for more information."
            
    Case "LABEL"
        lbluitleg.Caption = "Create, edit or change a drive's volume name." & vbCrLf & vbCrLf & "LABEL [drive:][name]" & vbCrLf
        
    Case "MD"
        lbluitleg.Caption = "Create a new (sub)folder." & vbCrLf & vbCrLf & "MKDIR [drive:]path" & vbCrLf & "MD [drive:]path"
        
     Case "MODE"
        lbluitleg.Caption = "Configure system devices." & vbCrLf & vbCrLf & "Serial port: " & vbCrLf & " MODE COMm[:] [BAUD=b] [PARITY=p] [DATA=d] [STOP=s] [to=on|off] [xon=on|off] [odsr=on|off] [octs=on|off] [dtr=on|off|hs] [rts=on|off|hs|tg] [idsr=on|off]" & vbCrLf & "Print ports:" & vbCrLf & "MODE LPTn[:]=COMm[:]"
            
    Case "MOVE"
        lbluitleg.Caption = "Move files" & vbCrLf & vbCrLf & "MOVE [/Y | /-Y] [drive:][path]filename[,...] target" & vbCrLf & vbCrLf & "/Y: don't ask for confirmation" & vbCrLf & "/-Y: ask for confirmation before overwriting"
                        
    Case "PATH"
        lbluitleg.Caption = "Display or set a search path for executable files." & vbCrLf & vbCrLf & "PATH [[drive:]path[;...][;%PATH%]" & vbCrLf & vbCrLf & "%PATH%: add the previous path."
        
    Case "PAUSE"
        lbluitleg.Caption = "Pauses the execution of a batch file and displays the message 'Press any key to continue'."
        
    Case "PROMPT"
        lbluitleg.Caption = "Changes the command prompt." & vbCrLf & vbCrLf & "PROMPT [text]" & vbCrLf & vbCrLf & "$A: &" & vbCrLf & "$C: <" & vbCrLf & "$D: date" & vbCrLf & "$N: drive" & vbCrLf & "$P: drive and path" & vbCrLf & "$T: time"
            
    Case "RD"
        lbluitleg.Caption = "Delete a folder." & vbCrLf & vbCrLf & "RMDIR [/S] [/Q] [drive:]path" & vbCrLf & vbCrLf & "/S: Include subfolders" & vbCrLf & "/Q: Don't ask for confirmation."
        
    Case "Rem"
        lbluitleg.Caption = "Add comments: everything after the word REM at the beginning of a line will be ignored."
    
    Case "REN"
        lbluitleg.Caption = "Rename one or more files." & vbCrLf & vbCrLf & "RENAME [drive:][path]filename1 filename2" & vbCrLf & vbCrLf & "filename1 and filename2 must be on the same drive."
    
    Case "RENAME"
        lbluitleg.Caption = "Rename one or more files." & vbCrLf & vbCrLf & "RENAME [drive:][path]filename1 filename2" & vbCrLf & vbCrLf & "filename1 and filename2 must be on the same drive."
    
    Case "REPLACE"
        lbluitleg.Caption = "Replace files." & vbCrLf & vbCrLf & "REPLACE [drive:][path1]filename [drive2:][path2] [/A] [/P] [/R] [/W]" & vbCrLf & vbCrLf & "/A: add new files to target folder" & vbCrLf & "/P: ask for confirmation" & vbCrLf & "/R: replace read-only files."
        
    Case "RMDIR"
        lbluitleg.Caption = "Delete a folder." & vbCrLf & vbCrLf & "RMDIR [/S] [/Q] [drive:]path" & vbCrLf & vbCrLf & "/S: Delete the folder itself" & vbCrLf & "/Q: Don't ask for confirmation."
        
    Case "SET"
        lbluitleg.Caption = "Display, set or edit environment variables." & vbCrLf & vbCrLf & "SET [variable=[string]]" & vbCrLf & vbCrLf & "SET without options displays the current environment." & vbCrLf & "Type SET /? for more information/"
        
    Case "START"
        lbluitleg.Caption = "Open a seperate window to execute a program." & vbCrLf & vbCrLf & "START ['program'] [/Dpath] [/I] [/MIN] [/MAX] [/SEPARATE | /SHARED] [/LOW | /NORMAL | /HIGH | /REALTIME | /ABOVENORMAL | /BELOWNORMAL] [/WAIT] [/B] [command/program] [parameters]"
        
    Case "SUBST"
        lbluitleg.Caption = "Attach a driveletter to a path." & vbCrLf & vbCrLf & "SUBST [drive1: [drive2:]path]" & vbCrLf & vbCrLf & "SUBST drive: /D = delete drive"
        
    Case "TIME"
        lbluitleg.Caption = "Display or change the system time."
        
    Case "TREE"
        lbluitleg.Caption = "Shows a graphical view of the folder's structure" & vbCrLf & vbCrLf & "TREE [drive:][path] [/F] [/A]" & vbCrLf & vbCrLf & "/F: show filenames" & vbCrLf & "/A: Use ASCII characters"
        
    Case "TYPE"
        lbluitleg.Caption = "Display the contents of a textfile." & vbCrLf & vbCrLf & "TYPE [drive:][path]filename"
    
    Case "VER"
        lbluitleg.Caption = "Displays the Windows version."
        
    Case "VOL"
        lbluitleg.Caption = "Display a drive's volume information." & vbCrLf & vbCrLf & "VOL [drive:]"
        
    Case "XCOPY"
        lbluitleg.Caption = "Copy files or folder structures." & vbCrLf & vbCrLf & "XCOPY source [target] [/A | /M] [/D[:date]] [/P] [/S [/E]] [/V] [/W] [/C] [/I] [/Q] [/F] [/L] [/G] [/H] [/R] [/T] [/U] [/K] [/N] [/O] [/X] [/Y] [/-Y] [/Z] [/EXCLUDE: bestand1[+bestand2] [+bestand3]...]" & vbCrLf & vbCrLf & "Type XCOPY /? for more information"
        
    End Select
    



'form1.List1.AddItem "VER"
'form1.List1.AddItem "VOL"
'form1.List1.AddItem "XCOPY"
    
    
    
ElseIf comboInvoegen.ListIndex = 0 Then   ' html tag

Select Case List1.Text

    Case "HTML-page header"
    
        lblOnderwerp.Caption = "HTML-Page"
    
        lbluitleg.Caption = "Insert the structure for a new webpage and insert HTML, HEAD and BODY tags."
    

    Case "Line"
        lblOnderwerp.Caption = "Tag: <BR>"
        lbluitleg.Caption = "Add a new line."

    Case "Paragraph"
        lblOnderwerp.Caption = "Tag: <P>"
        lbluitleg.Caption = "Add a new paragraph."

    Case "Image"
        lblOnderwerp.Caption = "Tag: <IMG SRC=...>"
        lbluitleg.Caption = "Insert a picture"

    Case "Hyperlink"
        lblOnderwerp.Caption = "Tag: <A HREF=...>"
        lbluitleg.Caption = "Insert a hyperlink"
        
        
    Case "Horiz.rule"
        lblOnderwerp.Caption = "Tag: <HR>"
        lbluitleg.Caption = "Insert a horizontal ruler"
        
    Case "Applet"
        lblOnderwerp.Caption = "Tag: <APPLET>"
        lbluitleg.Caption = "Insert a Java applet"
    
    Case "Quote"
        lblOnderwerp.Caption = "Tag: <CITE>"
        lbluitleg.Caption = "Insert a quote"
                
    Case "Numbered list"
        lblOnderwerp.Caption = "Tag: <OL>"
        lbluitleg.Caption = "Insert a numbered list"
        
    Case "List"
        lblOnderwerp.Caption = "Tag: <UL>"
        lbluitleg.Caption = "Insert a list"
        
    Case "Comment"
        lblOnderwerp.Caption = "Tag: <!--..-->"
        lbluitleg.Caption = "Insert comments that aren't visible on the webpage"
                
    Case "Bold"
        lblOnderwerp.Caption = "Tag: <B>"
        lbluitleg.Caption = "Make text bold."
                
    Case "Italic"
        lblOnderwerp.Caption = "Tag: <I>"
        lbluitleg.Caption = "Make text italic."
    
    Case "Underline"
        lblOnderwerp.Caption = "Tag: <U>"
        lbluitleg.Caption = "Underline text."

    Case "Subscript"
        lblOnderwerp.Caption = "Tag: <SUB>"
        lbluitleg.Caption = "Text in subscript"
        
    Case "Superscript"
        lblOnderwerp.Caption = "Tag: <SUP>"
        lbluitleg.Caption = "Text in superscript"
        
    Case "Center"
        lblOnderwerp.Caption = "Tag: <CENTER>"
        lbluitleg.Caption = "Center object"
        
    Case "Font size 1"
        lblOnderwerp.Caption = "<FONT SIZE=1>"
        lbluitleg.Caption = "Font size 1."
        
    Case "Font size 2"
        lblOnderwerp.Caption = "<FONT SIZE=2>"
        lbluitleg.Caption = "Set font size 2."
        
    Case "Font size 3"
        lblOnderwerp.Caption = "<FONT SIZE=3>"
        lbluitleg.Caption = "Set font size 3."
        
    Case "Font size 4"
        lblOnderwerp.Caption = "<FONT SIZE=4>"
        lbluitleg.Caption = "Set font size 4."
        
    Case "Font size 5"
        lblOnderwerp.Caption = "<FONT SIZE=5>"
        lbluitleg.Caption = "Set font size 5."
        
    Case "Font size 6"
        lblOnderwerp.Caption = "<FONT SIZE=6>"
        lbluitleg.Caption = "Set font size 6."
        
    Case "Font size 7"
        lblOnderwerp.Caption = "<FONT SIZE=7>"
        lbluitleg.Caption = "Set font size 7."
        
    Case "color: black"
        lblOnderwerp.Caption = "color: black"
        lbluitleg.Caption = "HTML-code for black"
        

    Case "Color: white"
        lblOnderwerp.Caption = "Color: white"
        lbluitleg.Caption = "HTML-code for white"
        
    Case "Color: red"
        lblOnderwerp.Caption = "Color: red"
        lbluitleg.Caption = "HTML-code for red"
        
    Case "Color: yellow"
        lblOnderwerp.Caption = "Color: yellow"
        lbluitleg.Caption = "HTML-code for yellow"
        
    Case "Color: blue"
        lblOnderwerp.Caption = "Color: blue"
        lbluitleg.Caption = "HTML-code for blue"
        
    Case "Color: grey"
        lblOnderwerp.Caption = "Color: grey"
        lbluitleg.Caption = "HTML-code for grey"
        
    Case "Color: silver"
        lblOnderwerp.Caption = "Color: silver"
        lbluitleg.Caption = "HTML-code for silver"
        
    Case "Color: green"
        lblOnderwerp.Caption = "Color: green"
        lbluitleg.Caption = "HTML-code for green"
        
    Case "Table"
        lblOnderwerp.Caption = "Tag: <TABLE ...>"
        lbluitleg.Caption = "Insert HTML code for a table"
    
        
    

    
End Select





    

End If
On Error Resume Next
RichTextBox1.SetFocus


End Sub

Private Sub List1_DblClick()



If comboInvoegen.ListIndex = 1 Then   ' dos commando
    RichTextBox1.SelText = RichTextBox1.SelText & List1.Text & " "

ElseIf comboInvoegen.ListIndex = 3 Then    ' autotekst
    RichTextBox1.SelText = RichTextBox1.SelText & List1.Text & " "
    

ElseIf comboInvoegen.ListIndex = 0 Then   ' html tag

Select Case List1.Text

    Case "HTML-page header"
    
    Close
    Open App.Path & "\user.txt" For Input As #1
    Input #1, userdude
    Close
    
        RichTextBox1.SelText = "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE> </TITLE>" & vbCrLf & "<META NAME='Generator' Content='OptiType 3.0'>" & vbCrLf & "<META NAME='Author' Content='" & userdude & "'>" & vbCrLf & "<META NAME='Keywords' Content=' '" & vbCrLf & "<META NAME='Description' Content=' '>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY BGCOLOR='#FFFFFF' TEXT='#000000' LINK='#FF0000' VLINK='#800000' ALINK='#FF00FF'>" & vbCrLf & vbCrLf & RichTextBox1.SelText & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"



    Case "Line"
        RichTextBox1.SelText = RichTextBox1.SelText & "<BR>"

    Case "Paragraph"
        RichTextBox1.SelText = "<P>" & RichTextBox1.SelText & "</P>"

    Case "Image"
        RichTextBox1.SelText = RichTextBox1.SelText & "<IMG SRC='' ALT='' BORDER=0>"

    Case "Hyperlink"
        RichTextBox1.SelText = "<A HREF=''>" & RichTextBox1.SelText & "</A>"
        
    Case "Horiz.rule"
        RichTextBox1.SelText = RichTextBox1.SelText & "<HR>"
        
    Case "Applet"
        RichTextBox1.SelText = "<APLET CODE='' LANGUAGE=JAVASCRIPT>" & RichTextBox1.SelText & "</APPLET>"
    
    Case "Quote"
        RichTextBox1.SelText = "<CITE>" & RichTextBox1.SelText & "</CITE"
        
    Case "Numbered list"
        RichTextBox1.SelText = "<OL>" & vbCrLf & "<LI>" & RichTextBox1.SelText & vbCrLf & "</OL>   "
        
    Case "List"
        RichTextBox1.SelText = "<UL>" & vbCrLf & "<LI>" & RichTextBox1.SelText & vbCrLf & "</UL>   "

    Case "Comment"
        RichTextBox1.SelText = "<!--" & RichTextBox1.SelText & "-->"
        
    Case "Bold"
        RichTextBox1.SelText = "<B>" & RichTextBox1.SelText & "</B>"
        
    Case "Italic"
        RichTextBox1.SelText = "<I>" & RichTextBox1.SelText & "</I>"
    
    Case "Underline"
        RichTextBox1.SelText = "<U>" & RichTextBox1.SelText & "</U>"

    Case "Subscript"
        RichTextBox1.SelText = "<SUB>" & RichTextBox1.SelText & "</SUB>"
        
    Case "Superscript"
        RichTextBox1.SelText = "<SUP>" & RichTextBox1.SelText & "</SUP>"
        
    Case "Center"
        RichTextBox1.SelText = "<CENTER>" & RichTextBox1.SelText & "</CENTER>"
        
    Case "Font size 1"
        RichTextBox1.SelText = "<FONT SIZE='1'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 2"
        RichTextBox1.SelText = "<FONT SIZE='2'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 3"
        RichTextBox1.SelText = "<FONT SIZE='3'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 4"
        RichTextBox1.SelText = "<FONT SIZE='4'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 5"
        RichTextBox1.SelText = "<FONT SIZE='5'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 6"
        RichTextBox1.SelText = "<FONT SIZE='6'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Font size 7"
        RichTextBox1.SelText = "<FONT SIZE='7'>" & RichTextBox1.SelText & "</FONT>"
        
    Case "Color: black"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#000000'"

    Case "Color: white"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#FFFFFF'"
        
    Case "Color: red"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#FF0000'"
        
    Case "Color: yellow"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#FFFF00'"
        
    Case "Color: blue"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#0000FF'"
        
    Case "Color: grey"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#808080'"
        
    Case "Color: silver"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#C0C0C0'"
        
    Case "Color: green"
        RichTextBox1.SelText = RichTextBox1.SelText & "'#008000'"
        
    Case "Table"
        htmltabel = "<TABLE ALIGN='left' BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='100%'>"
        htmltabel = htmltabel & vbCrLf & "<TR ALIGN='left' VALIGN='middle'>" & vbCrLf
        htmltabel = htmltabel & "<TH></TH>" & vbCrLf & "<TH></TH>" & vbCrLf
        htmltabel = htmltabel & "<TR ALIGN='left' VALIGN='middle'>" & vbCrLf
        htmltabel = htmltabel & "<TD>   </TD>" & vbCrLf & "<TD>    </TD>" & vbCrLf & "</TABLE>"
        RichTextBox1.SelText = RichTextBox1.SelText & htmltabel
    
    

    
End Select




ElseIf comboInvoegen.ListIndex = 2 Then   ' ansi
    RichTextBox1.SelText = RichTextBox1.SelText & Right$(List1.Text, 1)
    

End If
On Error Resume Next
RichTextBox1.SetFocus

End Sub

Private Sub makelist_Click()


On Error GoTo getmeouttahere
Combo5.Clear
a = CountWords(1, RichTextBox1, True, Combo5)
Combo5.ListIndex = 0

getmeouttahere:
Exit Sub

End Sub

Private Sub makelistandsave_Click()



On Error GoTo getmeouttahere
Combo5.Clear
a = CountWords(1, RichTextBox1, True, Combo5)
Combo5.ListIndex = 0

' ------------------------- code voor opslaan als:
On Error GoTo ErrHandler
    
   ' Set CancelError is True
    form1.CommonDialog2.CancelError = True
    Rem On Error GoTo ErrHandler
    ' Set flags
    form1.CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    form1.CommonDialog2.Filter = "Text files(*.txt)|*.txt|All files|*.*"
    ' Specify default filter
     form1.CommonDialog2.FilterIndex = 1
    form1.CommonDialog2.DialogTitle = "Save word list as..."
    form1.CommonDialog2.FileName = "word list.txt"
    ' Display the Open dialog box
    form1.CommonDialog2.ShowSave
    ' Display name of selected file

    Open CommonDialog2.FileName For Output As #6

        
    For i = 0 To Combo5.ListCount
        Print #6, Combo5.Text
        Combo5.ListIndex = Combo5.ListIndex + 1
    Next i
    
    Close #6
 
    Combo5.ListIndex = 0
    
    On Error Resume Next
    form1.RichTextBox1.SetFocus
        
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub







getmeouttahere:
Exit Sub


End Sub

Private Sub mnuhelpabout_Click()

contenthelp

End Sub



Private Sub openautoexec_Click()

On Error GoTo noautoexec

checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    DocName.Text = "c:\autoexec.bat"
    RichTextBox1.LoadFile "c:\autoexec.bat", 1
    form1.Caption = "OptiType 3.0 - Edit system file: Autoexec.bat"
Else
    annuleer = False
End If

Exit Sub

noautoexec:
MsgBox "The file C:\AUTOEXEC.BAT has not been found on your computer.", vbCritical, "OptiType"


End Sub

Private Sub openconfigsys_Click()
On Error GoTo noconfig
checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    DocName.Text = "c:\config.sys"
    RichTextBox1.LoadFile "c:\config.sys", 1
    form1.Caption = "OptiType 3.0 - Edit system file:Config.sys"
Else
    annuleer = False
End If

Exit Sub

noconfig:
MsgBox "The file C:\CONFIG.SYS has not been found on your computer.", vbCritical, "OptiType"

End Sub

Private Sub openmsdossys_Click()
On Error GoTo nomsdos
checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    RichTextBox1.LoadFile "c:\msdos.sys", 1
    form1.Caption = "OptiType 3.0 - Edit system file: Msdos.sys"
    DocName.Text = "c:\msdos.sys"
Else
    annuleer = False
End If

Exit Sub

nomsdos:
MsgBox "The file C:\MSDOS.SYS has not been found on your computer.", vbCritical, "OptiType"

End Sub

Private Sub openprotocolini_Click()
On Error GoTo noprotocol
checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    RichTextBox1.LoadFile GetWinPath & "\protocol.ini", 1
    form1.Caption = "OptiType 3.0 - Edit system file: Protocol.ini"
    DocName.Text = GetWinPath & "\protocol.ini"
Else
    annuleer = False
End If

Exit Sub

noprotocol:
MsgBox "Het bestand PROTOCOL.INI has not been found on your computer.", vbCritical, "OptiType"
End Sub

Private Sub opensystemini_Click()
On Error GoTo nosystem
checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    RichTextBox1.LoadFile GetWinPath & "\system.ini", 1
    form1.Caption = "OptiType 3.0 - Edit system file: System.ini"
    DocName.Text = GetWinPath & "\system.ini"
Else
    annuleer = False
End If

Exit Sub

nosystem:
MsgBox "Het bestand SYSTEM.INI has not been found on your computer.", vbCritical, "OptiType"

End Sub

Private Sub openwinini_Click()
On Error GoTo nowin
checktosave
If annuleer = False Then
RichTextBox1.SelBullet = False
    RichTextBox1.LoadFile GetWinPath & "\win.ini", 1
    form1.Caption = "OptiType 3.0 - Edit system file: Win.ini"
    DocName.Text = GetWinPath & "\win.ini"
Else
    annuleer = False
End If

Exit Sub

nowin:
MsgBox "Het bestand WIN.INI has not been found on your computer.", vbCritical, "OptiType"
End Sub

Private Sub opmaakbalk_Click()

If opmaakbalk.Checked = True Then
    opmaakbalk.Checked = False
    Toolbar2.Visible = False
Else
    opmaakbalk.Checked = True
    Toolbar2.Visible = True
End If



End Sub

Private Sub optiknoppen_Click()
If optiknoppen.Checked = True Then
    optiknoppen.Checked = False
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
Else
    optiknoppen.Checked = True
    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True
End If


End Sub

Private Sub options_Click()

frmOpties.Show

End Sub

Private Sub pauzeer_click()

RichTextBox1.Text = RichTextBox1.Text & "PAUSE" & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus



End Sub


Private Sub Form_Activate()
On Error Resume Next: RichTextBox1.SetFocus

End Sub

Private Sub Form_GotFocus()
On Error Resume Next: RichTextBox1.SetFocus

End Sub

Private Sub Form_Load()
 'gLngMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)
' remove the comment on the line above to disable the IE right-click menu.


comboInvoegen.Clear

comboInvoegen.AddItem "HTML Tags"
comboInvoegen.AddItem "DOS commands"
comboInvoegen.AddItem "Symbols"
comboInvoegen.AddItem "Autotext"

comboInvoegen.ListIndex = 0


Combo1.Clear

 Combo1.AddItem "This document"
 Combo1.AddItem "Internet: Google"
 Combo1.AddItem "Internet: Altavista"
 Combo1.AddItem "Internet: Yahoo"
 Combo1.AddItem "Internet: Webcrawler"
  Combo1.AddItem "Internet: Excite"
 Combo1.AddItem "Internet: Download.com"

Combo1.ListIndex = 0

annuleer = False





Label2.Height = Me.Height


RichTextBox1.Left = 2400
RichTextBox1.Width = form1.Width - 2820
Label2.Width = Me.Width - RichTextBox1.Width - 420

comboInvoegen.Width = Label2.Width - 100
List1.Width = Label2.Width - 100

'RichTextBox1.Top = toolbar2.top + toolbar2.height
RichTextBox1.Top = Toolbar2.Top + Toolbar2.Height

RichTextBox1.Height = form1.Height - 2640



Label1.Top = RichTextBox1.Top + RichTextBox1.Height + 80
txtURL.Top = RichTextBox1.Top + RichTextBox1.Height + 80
Command2.Top = RichTextBox1.Top + RichTextBox1.Height + 80

lblDatum.Top = RichTextBox1.Top + RichTextBox1.Height + 80
DocName.Top = RichTextBox1.Top + RichTextBox1.Height + 80

' -------------------------------------------
WebBrowser1.Width = RichTextBox1.Width

'WebBrowser1.Top = RichTextBox1.Top + (RichTextBox1.Height / 2)
WebBrowser1.Top = RichTextBox1.Top + (RichTextBox1.Height)
WebBrowser1.Height = RichTextBox1.Height / 2


' -----------------------------------------

bron = 3
tijd = 0




Dim datumweergeven, statusbalkweergeven, automatischopslaan, automatischopslaanbevestiging As Integer

setoptions

MousePointer = 11
StatusBar1.SimpleText = "Welcome to OptiType!"

    allowpopup = True
    showpopups.Checked = True


webbuttons.Buttons.Item(10).Value = tbrPressed
showpopups.Checked = True
browseoffline.Checked = False

showpopups.Enabled = False
browseoffline.Enabled = False
back.Enabled = False
gotonext.Enabled = False
refreshbrowser.Enabled = False
searchforapage.Enabled = False


Open "c:\temp.htm" For Output As #1
Print #1, " "
Close


'RichTextBox1.Height = Fix(5415 / 2)

verversing = 3

WebBrowser1.Navigate "c:\temp.htm"

Dim i As Integer
bg = "WHITE"

    

    

RichTextBox1.Text = ""

mnuSuperscript.Checked = False
mnuSubscript.Checked = False



FontList.Text = RichTextBox1.SelFontName

m_syntax.Checked = False


htmlsyntax = False

WindowState = vbMaximized



mnuVet.Checked = False
mnuCursief.Checked = False
mnuonderl.Checked = False
' wordwrap.Checked = False




'RichTextBox1.RightMargin = RichTextBox1.Width
'RichTextBox1.RightMargin = 0

' wordwrap.Checked = True



setfont

frmSplash.Hide



'RichTextBox1.RightMargin = 0
Combo5.Text = "Word list"

MousePointer = 0

If Command$ = "" And shownew = 1 Then
        frmNieuw.Show vbModal
Else
    leegdocumentmaken
End If
        






If UCase$(Right$(Command$, 3)) = "BMP" Or UCase$(Right$(Command$, 3)) = "GIF" Or UCase$(Right$(Command$, 3)) = "JPG" Then
    On Error Resume Next
    form1.htmlview.Checked = False
    form1.quicksave.Enabled = True
    form1.mnuOpslaan.Enabled = True
    leegdocumentmaken
    form1.RichTextBox1.SetFocus
    RichTextBox1.OLEObjects.Add , , Command$
    Exit Sub
End If


If Len(Command$) > 0 Then
    
    On Error GoTo cantloadfile
    If UCase$(Right$(Command$, 3)) = "RTF" Then
    RichTextBox1.LoadFile Command$, 0
    Else
    RichTextBox1.LoadFile Command$, 1
    End If
    
    DocName.Text = Command$
    
            
End If

'form1.RichTextBox1.SelRightIndent = 1600
    fonts(0).Caption = Screen.fonts(0)
    For i = 1 To Screen.FontCount - 1
        Load fonts(i)
       fonts(0).Caption = Screen.fonts(i)
       FontList.AddItem Screen.fonts(i)
  
          
    Next




Exit Sub

cantloadfile:

MsgBox "The file specified in the command line could not be loaded.", vbCritical, "Can't find or open file"



End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        form1.Enabled = False
              PopupMenu beeld
         form1.Enabled = True
        
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

StatusBar1.SimpleText = "OptiType 3.0  Benny Rossaer 1999 - 2002"


End Sub




Private Sub Form_Unload(Cancel As Integer)



If Len(RichTextBox1.Text) > 1 Then
    
    If DocName.Text = "File not saved yet!" Then
        a = MsgBox("Your document hasn't been saved yet.  Would you like to save it now?", vbYesNoCancel, "Save file")
        If a = 6 Then OpslaanAls
        If a = 7 Then GoTo quitprogram
        If a = 2 Then Cancel = -1
        
    Else
        a = MsgBox("Do you want to save your changes to this document?", vbYesNoCancel, "Save changes")
        If a = 6 Then
                If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
                RichTextBox1.SaveFile DocName.Text, 0
                Else
                RichTextBox1.SaveFile DocName.Text, 1
                End If
        End If
        If a = 7 Then GoTo quitprogram
        If a = 2 Then Cancel = -1
    End If

End If


quitprogram:
On Error Resume Next
If Cancel <> -1 Then
Kill "c:\temp.htm"

Unload Me
End

Else
    Exit Sub
End If




End Sub

Private Sub format_Click()

frmFormatDisk.Show


End Sub

Private Sub fullscr_Click()


If WindowState = vbNormal Then
    WindowState = vbMaximized
    nieuwformaat
Else
    WindowState = vbNormal
    nieuwformaat
End If


End Sub

Private Sub gauw_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Tot gauw"
End Sub

Private Sub geachte_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Geachte"
End Sub

Private Sub geachteheer_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Geachte heer"
End Sub

Private Sub geachtemevrouw_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Geachte mevrouw"
End Sub

Private Sub gotolabel_Click()
a = InputBox("Type in the name of the label you would like to jump to.", "LABEL")
form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & "GOTO " & a & vbCrLf
On Error Resume Next: form1.RichTextBox1.SetFocus
End Sub

Private Sub gotonext_Click()
On Error GoTo whoops

WebBrowser1.GoForward
Exit Sub

whoops:
' = " | Fout bij het inladen van volgende pagina."

End Sub

Private Sub groeten_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Hartelijke groeten"
End Sub

Private Sub groetjes_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Groetjes"
End Sub

Private Sub hoogachtend_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Hoogachtend"
End Sub

Private Sub horline_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "<HR>"

End Sub






Private Sub HTMLPrevKnop_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Display a web preview in a webbrowser."

End Sub

Private Sub htmlprevopt_Click()
ViewHTMLOpties.Show vbModeless, Me

End Sub

Private Sub htmlview_Click()

If htmlview.Checked = True And WebBrowser1.Visible = True Then
    RichTextBox1.Height = form1.Height - 2640
    'form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False
    htmlview.Checked = False

    form1.Caption = "OptiType 3.0 - New text file"
ElseIf htmlview.Checked = False And WebBrowser1.Visible = False Then
    form1.RichTextBox1.Height = Fix(form1.RichTextBox1.Height / 2)
    form1.WebBrowser1.Visible = True
    WebBrowser1.Height = RichTextBox1.Height
    htmlview.Checked = True

    frmOpties.Check1.Value = 1
    form1.Caption = "OptiType 3.0 - New website"
End If

End Sub

Private Sub htmlvoorbeeld_Click()



showHTMLPreview



End Sub

Private Sub htmlimage_Click()
frmAfbeelding.Show

End Sub

Private Sub ifcommand_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "IF "
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub ifexist_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "IF EXIST "
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub info_Click()
frmAbout.Show vbModeless, Me
End Sub

Private Sub insertHTMLtabel_Click()
frmHTMLtabel.Show

End Sub

Private Sub Font_Change()

On Error Resume Next

RichTextBox1.SelFontName = Font.Text

End Sub

Private Sub insertlabel_Click()

a = InputBox("Typ de naam voor het label...", "LABEL")
form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & ":" & a & vbCrLf
On Error Resume Next: form1.RichTextBox1.SetFocus
End Sub

Private Sub introbg_Click()
form1.Picture = LoadPicture(App.Path & "\newbg.jpg")
classicbg.Checked = False
millenniumbg.Checked = False
introbg.Checked = True

End Sub





Private Sub slower_Click()
Speech1.Speed = Speech1.Speed - 5

End Sub

Private Sub start_Click()

Speech1.Age (6)





Speech1.Speak RichTextBox1.Text
End Sub

Private Sub stop_Click()
Speech1.AudioReset


End Sub

Private Sub printpreview_Click()

End Sub



Private Sub standaardwerkbalk_Click()

If standaardwerkbalk.Checked = True Then

    standaardwerkbalk.Checked = False
    Toolbar1.Visible = False
    
Else
    standaardwerkbalk.Checked = True
    Toolbar1.Visible = True
End If





End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
MsgBox "You can hide the statusbar by going to Options in the Extra menu.", vbInformation, "OptiType"

' = " | Statusbalk verwijderen"

End Sub



Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub tabelinvoegen_Click()
frmTabel.Show vbModeless, Me
End Sub


Private Sub tijdinvoegenindocument_Click()
RichTextBox1.SelText = RichTextBox1.SelText & Time$
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' = " | Met bestanden werken"
Select Case Button.Key
    
    Case Is = "voorbeeld"
        PrintPreview RichTextBox1, 1, 1, 1, 1, 1
         ' = " | Afdrukken"
    
    Case Is = "nieuw"           ' Open file.
        Nieuwbestand
        
    Case Is = "open"              ' Save file.
        openbestand
        
        
    Case Is = "opslaan"
        bewaren
        
    Case Is = "mail"
        frmEmail.RichTextBox1.TextRTF = RichTextBox1.TextRTF
        ' = " | Documenten e-mailen"
        frmEmail.Show
        
    Case Is = "afdrukken"
        drukaf
        ' = " | Afdrukken"
        
    Case Is = "preview"
        showHTMLPreview
        ' = " | Websites ontwerpen"
        
    Case Is = "knippen"
        knip
        ' = " | Knippen en plakken"
        
    Case Is = "kopieren"
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.SelText
        On Error Resume Next: RichTextBox1.SetFocus
        ' = " | Knippen en plakken"
        
    Case Is = "plakken"
        plak
        ' = " | Knippen en plakken"
        
    Case Is = "undo"
        undo
        ' = " | Ongedaan maken en herhalen"
        
    Case Is = "redo"
        redo
        ' = " | Ongedaan maken en herhalen"
        
        
    
        
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' = " | Opmaakkenmerken"
Select Case Button.Key
    Case Is = "vet"
        
        
        
        If Toolbar2.Buttons.Item(1).Value = tbrPressed Then RichTextBox1.SelBold = True: mnuVet.Checked = True: Exit Sub
        If Toolbar2.Buttons.Item(1).Value = tbrUnpressed Then RichTextBox1.SelBold = False: mnuVet.Checked = False
                
        
                
    Case Is = "cursief"
        If Toolbar2.Buttons.Item(2).Value = tbrPressed Then RichTextBox1.SelItalic = True: mnuCursief.Checked = True
        If Toolbar2.Buttons.Item(2).Value = tbrUnpressed Then RichTextBox1.SelItalic = False: mnuCursief.Checked = False
        
        
    Case Is = "underl"
         If Toolbar2.Buttons.Item(3).Value = tbrPressed Then RichTextBox1.SelUnderline = True: mnuonderl.Checked = True
        If Toolbar2.Buttons.Item(3).Value = tbrUnpressed Then RichTextBox1.SelUnderline = False: mnuonderl.Checked = False
    
    Case Is = "links"
        centrLeft
        
    Case Is = "centr"
        centreer
        
    Case Is = "rechts"
        centrrechts
        
    Case Is = "font"
        WijzigLettertype
        
    Case Is = "lijst"
      '  Opsomming
      If Toolbar2.Buttons.Item(11).Value = tbrPressed Then RichTextBox1.SelBullet = True: mnuonderl.Checked = True
    If Toolbar2.Buttons.Item(11).Value = tbrUnpressed Then RichTextBox1.SelBullet = False: mnuonderl.Checked = False
        
    
    End Select
    


End Sub

Private Sub txtSize_GotFocus()
' = " | Lettertypen"
End Sub

Private Sub txtsize_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Hier ziet u de Font size."

End Sub

Private Sub Lettertype_Click()
   WijzigLettertype
   
   
End Sub
Private Sub WijzigLettertype()
On Error Resume Next
'Finds 7 Properties
CommonDialog1.FontItalic = RichTextBox1.SelItalic
CommonDialog1.FontBold = RichTextBox1.SelBold
CommonDialog1.FontName = RichTextBox1.SelFontName
CommonDialog1.FontSize = RichTextBox1.SelFontSize
CommonDialog1.FontStrikethru = RichTextBox1.SelStrikeThru
CommonDialog1.FontUnderline = RichTextBox1.SelUnderline
CommonDialog1.Color = RichTextBox1.SelColor
'Sets Flags and Shows FontSelect
CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects Or 262144
CommonDialog1.ShowFont
'Returns 7 Properties
'AdjustButtons
RichTextBox1.SelFontName = CommonDialog1.FontName
'Font.Caption = FName
txtSize.Text = CommonDialog1.FontSize
VScroll1.Value = CommonDialog1.FontSize

RichTextBox1.SelColor = CommonDialog1.Color


RichTextBox1.SelBold = CommonDialog1.FontBold
RichTextBox1.SelItalic = CommonDialog1.FontItalic
RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru

On Error Resume Next: RichTextBox1.SetFocus

'i(0) = cd1.FontBold
'i(1) = cd1.FontItalic
'i(2) = cd1.FontUnderline
'i(3) = cd1.FontStrikethru'

End Sub




Private Sub Lettertype_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "View the font dialog window."

End Sub

Private Sub liefs_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Liefs"
End Sub

Private Sub liefsteouders_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Liefste ouders"
End Sub

Private Sub link_Click()

frmHyperlink.Show


End Sub

Private Sub links_Click()
  RichTextBox1.RightMargin = RichTextBox1.Width
        ' ' wordwrap.checked = True
        RichTextBox1.SelAlignment = 0
        On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub centrLeft()
RichTextBox1.RightMargin = RichTextBox1.Width
' ' wordwrap.checked = True
RichTextBox1.SelAlignment = 0
On Error Resume Next: RichTextBox1.SetFocus

End Sub



Private Sub LinksUitlijnen_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Align left"

End Sub

Private Sub listwithshortcuts_Click()

Dim a, b, c, d As String

a = "CTRL-N: New" & vbCrLf & "CTRL-O: Open" & vbCrLf & "CTRL-S: Save" & vbCrLf & "F12 - Save as" & vbCrLf & "CTRL-P: Print" & vbCrLf
b = "CTRL-X: Cut" & vbCrLf & "CTRL-C: Copy" & vbCrLf & "CTRL-V: Paste" & vbCrLf & "CTRL-A: Select All" & vbCrLf & "CTRL-F: Find" & vbCrLf & "CTRL-H: Replace" & vbCrLf & vbCrLf & "CTRL-L: Align left" & vbCrLf & "CTRL-E: Center" & vbCrLf & "CTRL-R: Align right" & vbCrLf
c = "CTRL-Z: Undo" & vbCrLf & "CTRL-B: Bold" & vbCrLf & "CTRL-I: Italic" & vbCrLf & "CTRL-U: Underline" & vbCrLf
d = "F1: Help"

MsgBox a & vbCrLf & b & vbCrLf & c & vbCrLf & d




End Sub

Private Sub m_syntax_Click()

If m_syntax.Checked = True Then
    m_syntax.Checked = False
    htmlsyntax = False
    Else
    m_syntax.Checked = True
    htmlsyntax = True
End If
    

End Sub

Private Sub mail_Click()
StatusBar1.SimpleText = "With this function, you can e-mail your document to someone.  A compatible e-mailclient, like Outlook or Outlook Express, should be installed on your computer."

ebron = 1
frmEmail.RichTextBox1.TextRTF = RichTextBox1.TextRTF
frmEmail.Show vbModeless, Me



End Sub

Private Sub makeitbig_Click()

RichTextBox1.SelText = UCase$(RichTextBox1.SelText)


End Sub

Private Sub makeitsmall_Click()
RichTextBox1.SelText = LCase(RichTextBox1.SelText)

End Sub

Private Sub marges_Click()
frmMarges.Show vbModeless, Me

End Sub

Private Sub marquee_Click()
Dim a, krant As String

a = InputBox("Type in the text you would like to display in the marquee...", "HTML: Marquee")

krant = "<MARQUEE>" & a & "</MARQUEE>"


RichTextBox1.SelText = RichTextBox1.SelText & krant

On Error Resume Next: RichTextBox1.SetFocus



End Sub

Private Sub mem_Click()
form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & "MEM" & vbCrLf
On Error Resume Next: form1.RichTextBox1.SetFocus
End Sub

Private Sub Mevrouw_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Mevrouw"
End Sub

Private Sub midden_Click()
       RichTextBox1.RightMargin = RichTextBox1.Width
        ' ' wordwrap.checked = True
        RichTextBox1.SelAlignment = 2
        On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub Mijnheer_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Mijnheer"

End Sub

Private Sub millenniumbg_Click()
form1.Picture = LoadPicture(App.Path & "\bg.jpg")
classicbg.Checked = False
millenniumbg.Checked = True
introbg.Checked = False

End Sub

Private Sub mnuAfdrukken_Click()


PrintDialog.CancelError = True
On Error GoTo cantprintit

PrintDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If RichTextBox1.SelLength = 0 Then
        PrintDialog.Flags = PrintDialog.Flags + cdlPDAllPages
    Else
        PrintDialog.Flags = PrintDialog.Flags + cdlPDSelection
    End If
    PrintDialog.ShowPrinter
    'Printer.Print ""
    RichTextBox1.SelPrint PrintDialog.hdc

cantprintit:



End Sub

Private Sub mnuCursief_Click()
If Toolbar2.Buttons.Item(2).Value = tbrUnpressed Then
    Toolbar2.Buttons.Item(2).Value = tbrPressed
    RichTextBox1.SelItalic = True
    mnuCursief.Checked = True
Else
    Toolbar2.Buttons.Item(2).Value = tbrUnpressed
    RichTextBox1.SelItalic = False
    mnuCursief.Checked = False
End If


End Sub

Private Sub mnuKnippen_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""

End Sub

Private Sub mnuKopieren_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText

End Sub

Private Sub mnuLettert_Click()
WijzigLettertype
End Sub

Private Sub mnuNieuw_Click()
checktosave

If annuleer = False Then
    DocName.Text = "File not saved yet!"
    RichTextBox1.Text = ""
    RichTextBox1.SelBullet = False
    setfont
    frmNieuw.Show
End If
End Sub

Private Sub mnuonderl_Click()
If Toolbar2.Buttons.Item(3).Value = tbrUnpressed Then
    Toolbar2.Buttons.Item(3).Value = tbrPressed
    RichTextBox1.SelUnderline = True
    mnuonderl.Checked = True
Else
    Toolbar2.Buttons.Item(3).Value = tbrUnpressed
    RichTextBox1.SelUnderline = False
    mnuonderl.Checked = False
End If

End Sub

Private Sub mnuOpenen_Click()
checktosave

If annuleer = False Then
    Openen
Else
   annuleer = False
End If

End Sub

Private Sub mnuOpslaan_Click()
OpslaanAls

End Sub

Private Sub mnuPlakken_Click()


If Clipboard.GetFormat(vbCFText) Then
selectie = Clipboard.GetText
RichTextBox1.SelText = RichTextBox1.SelText & selectie
End If

If Clipboard.GetFormat(vbCFBitmap) Then
    SendKeys "^V"
End If


End Sub

Private Sub mnuSubscript_Click()

   If mnuSubscript.Checked = False Then
            RichTextBox1.SelCharOffset = -35: mnuSubscript.Checked = True: mnuSuperscript.Checked = False
    Else
      RichTextBox1.SelCharOffset = 0: mnuSubscript.Checked = False: mnuSuperscript.Checked = False
End If
   
On Error Resume Next: RichTextBox1.SetFocus


End Sub

Private Sub mnuSuperscript_Click()
   If mnuSuperscript.Checked = False Then
            RichTextBox1.SelCharOffset = 60: mnuSuperscript.Checked = True: mnuSubscript.Checked = False
    Else
      RichTextBox1.SelCharOffset = 0: mnuSuperscript.Checked = False: mnuSubscript.Checked = False
End If
   
On Error Resume Next: RichTextBox1.SetFocus

End Sub

Private Sub mnuVet_Click()

If Toolbar2.Buttons.Item(1).Value = tbrUnpressed Then
    Toolbar2.Buttons.Item(1).Value = tbrPressed
    RichTextBox1.SelBold = True
    mnuVet.Checked = True
Else
    Toolbar2.Buttons.Item(1).Value = tbrUnpressed
    RichTextBox1.SelBold = False
    mnuVet.Checked = False
End If


End Sub

Private Sub Nieuwbestand()

If Len(RichTextBox1.Text) > 1 Then
    
    If DocName.Text = "File not saved yet!" Then
   a = MsgBox("Your document hasn't been saved yet.  Would you like to save it now?", vbYesNoCancel, "Save file")
        
        If a = 6 Then OpslaanAls
        If a = 2 Then Exit Sub
    Else
        a = MsgBox("Would you like to save changes to your file?", vbYesNoCancel, "Save changes")
        
        If a = 6 Then
                If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
                RichTextBox1.SaveFile DocName.Text, 0
                Else
                RichTextBox1.SaveFile DocName.Text, 1
                End If
        End If
        If a = 2 Then Exit Sub
    End If

End If


RichTextBox1.Text = ""
DocName.Text = "File not saved yet!"
setfont
frmNieuw.Show


End Sub

Private Sub Openen()

   ' Set CancelError is True
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog2.Filter = "OptiType Rich Text (*.rtf)|*.rtf|Text files " & _
    "(*.txt)|*.txt|Webpage (*.htm, *.html)|*.ht*|Batch-file(*.bat)|*.bat|All files|*.*"
    ' Specify default filter
    CommonDialog2.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog2.ShowOpen
    ' Display name of selected file

    
   form1.Caption = "OptiType 3.0 - " & CommonDialog2.FileName
   
   
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "RTF" Then
    RichTextBox1.LoadFile CommonDialog2.FileName, 0
    RichTextBox1.SelBullet = False
    Else
    RichTextBox1.LoadFile CommonDialog2.FileName, 1
    RichTextBox1.SelBullet = False
    End If
    
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "BAT" Then
        batchpreview.Enabled = True
        form1.batchfile.Enabled = True
    Else
        batchpreview.Enabled = False
        form1.batchfile.Enabled = False
    End If
        
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "HTM" Or UCase$(Right$(CommonDialog2.FileName, 3)) = "TML" Then
    nieuwhtmlmakenzonderwizard
    End If
    
    
        
    DocName.Text = CommonDialog2.FileName
    
    form1.AutoSave.Enabled = True
    
            
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub NIEUW_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Creates a new document, webpage, or batch file."

End Sub

Private Sub oleinsertfile_Click()
' Set CancelError is True
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog2.Filter = "Microsoft Word (*.DOC)|*.doc|Microsoft Excel " & _
    "(*.xls)|*.xls|Windows 3.1 Write|*.wri|BMP image|*.bmp|GIF image|*.gif|JPG image|*.jpg|All files|*.*"
    ' Specify default filter
    CommonDialog2.FilterIndex = 7
    ' Display the Open dialog box
    CommonDialog2.ShowOpen
    
    RichTextBox1.OLEObjects.Add , , CommonDialog2.FileName
        
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub openbestand()

If Len(RichTextBox1.Text) > 1 Then
    
    If DocName.Text = "File not saved yet!" Then
        a = MsgBox("This document hasn't been saved yet.  Would you like to save it now?", vbYesNoCancel, "Save file")
        
        If a = 6 Then OpslaanAls
        If a = 2 Then Exit Sub
    Else
        a = MsgBox("Would you like to save changes to your document?", vbYesNoCancel, "Save changes")
        
        If a = 6 Then
                If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
                RichTextBox1.SaveFile DocName.Text, 0
                Else
                RichTextBox1.SaveFile DocName.Text, 1
                End If
        End If
        If a = 2 Then Exit Sub
End If

End If


Openen


End Sub



Private Sub OPEN_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Open een tekstbestand of RTF-bestand van schijf."

End Sub

Private Sub bewaren()

On Error GoTo cantsave
If DocName.Text = "File not saved yet!" Then

    OpslaanAls

Else

    If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
    RichTextBox1.SaveFile DocName.Text, 0
    Else
    RichTextBox1.SaveFile DocName.Text, 1
    End If

End If

Exit Sub

cantsave:


contenthelp

End Sub

Private Sub OPSLAAN_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Save your document as RTF, TXT, HTML or BAT."

End Sub

Private Sub Opsomming()

On Error Resume Next: RichTextBox1.SetFocus


If Toolbar2.Buttons.Item(11).Value = tbrPressed Then
RichTextBox1.SelBullet = True
Else
RichTextBox1.SelBullet = False
End If


End Sub

Private Sub Opsomming_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Create a list"

End Sub

Private Sub overtype_Click()

' = " | Overtype modus"


On Error Resume Next: RichTextBox1.SetFocus
SendKeys "{INSERT}"

End Sub

Private Sub quicksave_Click()

On Error GoTo cantsave

If DocName.Text = "File not saved yet!" Then

    OpslaanAls

Else

    If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
    RichTextBox1.SaveFile DocName.Text, 0
    Else
    RichTextBox1.SaveFile DocName.Text, 1
    End If

End If

Exit Sub

cantsave:


contenthelp

End Sub

Private Sub rechts_Click()
        RichTextBox1.RightMargin = RichTextBox1.Width
        ' ' wordwrap.checked = True
        RichTextBox1.SelAlignment = 1
        On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub centrrechts()
RichTextBox1.RightMargin = RichTextBox1.Width
' ' wordwrap.checked = True
RichTextBox1.SelAlignment = 1

On Error Resume Next: RichTextBox1.SetFocus
End Sub



Private Sub RECHTSUITLIJNEN_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
StatusBar1.SimpleText = "Align right"

End Sub

Private Sub redoit_Click()
   'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub refreshbrowser_Click()
On Error Resume Next
WebBrowser1.Refresh
End Sub

Private Sub refreshmanual_Click()
    
    On Error Resume Next
    RichTextBox1.SaveFile "c:\temp.htm", rtfText
    
    
    WebBrowser1.Refresh
    
    sec = 0
End Sub

Private Sub remove_Click()
RichTextBox1.SelText = ""
End Sub

Private Sub restore_Click()

focusherstellen

End Sub

Private Sub RichTextBox1_Change()
On Error Resume Next
If form1.Visible = True Then
X.Enabled = False
End If
    
undoint = undoint + 1

If undoint = 5 Then
    'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        On Error Resume Next: gstrStack(gintIndex) = RichTextBox1.TextRTF
    End If
    
    undoint = 0
End If

If form1.Visible = True Then X.Enabled = True




    
End Sub

Private Sub RichTextBox1_keydown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
    
'If KeyCode = 66 Then If CtrlDown Then MakeItBold
'If KeyCode = 73 Then If CtrlDown Then MakeItItalic
'If KeyCode = 85 Then If CtrlDown Then MakeItUnderline



If KeyCode = vbKeyTab Then
    RichTextBox1.SelText = RichTextBox1.SelText & Space$(8)
    On Error Resume Next: RichTextBox1.SetFocus
    
End If

    
    
End Sub


Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
On Error GoTo nopress



    Static InTag As Boolean
    Static InQuote As Boolean


    If htmlsyntax = True Then
        If KeyAscii = Asc("<") Then KeyAscii = 0: MakeColor: InTag = True
        If KeyAscii = Asc(">") Then RichTextBox1.SelColor = vbBlack: InTag = False


        If InTag = True Then


            If KeyAscii = Asc("""") Then
                RichTextBox1.SelColor = vbMagenta


                If InQuote = True Then
                    InQuote = False
                Else
                    InQuote = True
                End If

            Else


                If InQuote = False Then
                    If KeyAscii = Asc(" ") Then RichTextBox1.SelColor = vbRed
                End If

            End If

        End If

    End If

nopress:

End Sub



Private Sub richtextbox1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        RichTextBox1.Enabled = False
              PopupMenu mnuBewerken
         RichTextBox1.Enabled = True
        
    End If
End Sub


Private Sub RichTextBox1_SelChange()


If RichTextBox1.SelBullet = False Then
    Toolbar2.Buttons.Item(11).Value = tbrUnpressed
Else
    Toolbar2.Buttons.Item(11).Value = tbrPressed
End If
    


If RichTextBox1.SelCharOffset = 0 Then
    mnuSuperscript.Checked = False
    mnuSubscript.Checked = False
End If

If RichTextBox1.SelCharOffset > 0 Then
    mnuSuperscript.Checked = True
    mnuSubscript.Checked = False
End If

If RichTextBox1.SelCharOffset < 0 Then
    mnuSuperscript.Checked = False
    mnuSubscript.Checked = True
End If

If RichTextBox1.SelBold = True Then
       
   Toolbar2.Buttons.Item(1).Value = tbrPressed
           
  Else
   Toolbar2.Buttons.Item(1).Value = tbrUnpressed
End If
    
    If RichTextBox1.SelItalic = True Then
    Toolbar2.Buttons.Item(2).Value = tbrPressed
    Else
    Toolbar2.Buttons.Item(2).Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelUnderline = True Then
    Toolbar2.Buttons.Item(3).Value = tbrPressed
    Else
    Toolbar2.Buttons.Item(3).Value = tbrUnpressed
    End If
  
  'If RichTextBox1.SelFontSize <> Null Then
On Error Resume Next
txtSize.Text = RichTextBox1.SelFontSize
  
  


  
FontList.Text = RichTextBox1.SelFontName
Font.FontName = RichTextBox1.SelFontName


End Sub

Private Sub saveashtml_Click()
' = " | Websites ontwerpen"


' Set CancelError is True
    savefiledialog.CancelError = False
    On Error GoTo ErrHandler
    ' Set flags
    savefiledialog.Flags = cdlOFNHideReadOnly
    ' Set filters
    savefiledialog.Filter = "Webpage (*.htm)|*.htm|Text files" & _
    " (*.txt)|*.txt|All files|*.*"
    ' Specify default filter
    savefiledialog.FilterIndex = 1
    ' Display the Open dialog box
    savefiledialog.ShowSave
    ' Display name of selected file
 
 
 form1.RichTextBox1.SaveFile savefiledialog.FileName, 1
 
 
     
    Exit Sub
        
ErrHandler:
    'User pressed the Cancel button
    ' MsgBox "Error occured."
    Unload Me
    
    
    Exit Sub






End Sub

Private Sub saveselection_Click()


SelectionBox.Text = RichTextBox1.SelText


   ' Set CancelError is True
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog2.Filter = "OptiType RTF (*.RTF)|*.rtf|Text files" & _
    " (*.txt)|*.txt|All files|*.*"
    ' Specify default filter
    CommonDialog2.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog2.ShowSave
    ' Display name of selected file

   
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "RTF" Then
    SelectionBox.SaveFile CommonDialog2.FileName, 0
    Else
    SelectionBox.SaveFile CommonDialog2.FileName, 1
    End If
    
         
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub



End Sub

Private Sub searchforapage_Click()
On Error Resume Next
WebBrowser1.GoSearch

End Sub

Private Sub selectall_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub



Private Sub sendselection_Click()

ebron = 5
frmEmail.Show vbModeless, Me

End Sub

Private Sub setdateinbatch_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "DATE" & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub settime_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "TIME" & vbCrLf
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub showpopups_Click()
If showpopups.Checked = True Then
    allowpopup = False
    showpopups.Checked = False
    webbuttons.Buttons.Item(11).Value = tbrUnpressed
Else
    allowpopup = True
    showpopups.Checked = True
    webbuttons.Buttons.Item(11).Value = tbrPressed
End If
    


End Sub








Private Sub txt_Click()
   
   
   ' Set CancelError is True
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog2.Filter = "Text files|*.txt|Get text from any file" & _
    "(*.*)|*.*"
    ' Specify default filter
    CommonDialog2.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog2.ShowOpen
    ' Display name of selected file

 
 
    Open CommonDialog2.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, regel
    RichTextBox1.SelText = RichTextBox1.SelText & vbCrLf & regel
    Loop
    Close
 'RichTextBox1.SelStart
     
    
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub

Private Sub Undo_Click()
On Error Resume Next: RichTextBox1.SetFocus

SendKeys "^{Z}"




End Sub

Private Sub txtSize_Change()
On Error Resume Next
' = " | Lettertypen"
RichTextBox1.SelFontSize = Val(txtSize.Text)
On Error Resume Next: RichTextBox1.SetFocus
End Sub

Private Sub txtURL_GotFocus()
' = " | Surfen op Internet in OptiType "
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)



If KeyAscii = vbKeyReturn Then

    If WebBrowser2.Visible = False Then

    showpopups.Enabled = True
    browseoffline.Enabled = True
    back.Enabled = True
    gotonext.Enabled = True
    refreshbrowser.Enabled = True
    searchforapage.Enabled = True

    webbuttons.Buttons.Item(10).Value = tbrUnpressed
 
    form1.htmlview.Checked = True

    If WebBrowser1.Visible = False Then
        form1.RichTextBox1.Height = Fix(form1.RichTextBox1.Height / 2)
        form1.WebBrowser1.Visible = True
        WebBrowser1.Height = RichTextBox1.Height
        htmlview.Checked = True
        frmOpties.Check1.Value = 1
    End If

    WebBrowser1.Navigate txtURL.Text

    Else

    WebBrowser2.Navigate txtURL.Text

End If

End If


End Sub

Private Sub txtZoeken_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then zoeknu


End Sub

Private Sub undoit_Click()
    

    
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub



Private Sub vervangen_Click()
frmReplace.Show vbModeless, Me


End Sub

Private Sub voorbaat_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Bij voorbaat dank"
End Sub

Private Sub vrgroet_Click()
RichTextBox1.SelText = RichTextBox1.SelText & "Met vriendelijke groet"
End Sub

Private Sub VScroll1_Change()
' = " | Lettertypen"
txtSize.Text = Val(VScroll1.Value)
' = " |  Lettertypen"
End Sub

Private Sub VScroll1_GotFocus()
' = " | Lettertypen"
End Sub

Private Sub web2close_Click()
hidewebbrowser2

End Sub

Private Sub web2editpage_Click()
SaveHTML.Text = getsourcecode(form1.WebBrowser2.LocationURL)
SaveHTML.SaveFile App.Path & "\page.txt", 1
nieuwhtmlmakenzonderwizard
form1.RichTextBox1.LoadFile App.Path & "\page.txt", 1
focusherstellen

End Sub

Private Sub web2openpagefromdisk_Click()



   ' Set CancelError is True
    cdlsavewebpage.CancelError = True
    On Error GoTo ErrHandler
    cdlsavewebpage.DialogTitle = "Open web page from disk..."
    ' Set flags
    cdlsavewebpage.Flags = cdlOFNHideReadOnly
    ' Set filters
    cdlsavewebpage.Filter = "HTML-page|*.htm;*.html|All files|*.*"
    ' Specify default filter
    cdlsavewebpage.FilterIndex = 1
    ' Display the Open dialog box
    cdlsavewebpage.ShowOpen
    ' Display name of selected file

   
    WebBrowser2.Navigate cdlsavewebpage.FileName
    
       
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub web2printpage_Click()
WebBrowser2.SetFocus
SendKeys "^p"
End Sub

Private Sub web2savetodisk_Click()
SaveHTML.Text = getsourcecode(form1.WebBrowser2.LocationURL)



   ' Set CancelError is True
    cdlsavewebpage.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    cdlsavewebpage.DialogTitle = "Save web page..."
    cdlsavewebpage.Flags = cdlOFNHideReadOnly
    ' Set filters
    cdlsavewebpage.Filter = "Webpage (*.HTM)|*.htm|All files|*.*"
    ' Specify default filter
    cdlsavewebpage.FilterIndex = 1
    ' Display the Open dialog box
    cdlsavewebpage.ShowSave
    ' Display name of selected file

  
    SaveHTML.SaveFile cdlsavewebpage.FileName, 1
       
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub web2showsource_Click()
frmViewHTML.Show vbModeless, Me

End Sub

Private Sub web2zoekoppagina_Click()
WebBrowser2.SetFocus
SendKeys "^f"
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
txtURL.Text = URL
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
txtURL.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_GotFocus()
On Error Resume Next
'RichTextBox1.SetFocus

End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If allowpopup = True Then
    Cancel = False
    DoEvents
    
ElseIf allowpopup = False Then
    Cancel = True
End If
End Sub



Private Sub WebBrowser1_TitleChange(ByVal Text As String)
form1.Caption = "OptiType 3.0 - " & Text
End Sub

Private Sub WebBrowser2_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
txtURL.Text = URL
End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
txtURL.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser2_NewWindow2(ppDisp As Object, Cancel As Boolean)
If allowpopup = True Then
    Cancel = False
    DoEvents
    
ElseIf allowpopup = False Then
    Cancel = True
End If
End Sub



Private Sub WebBrowser2_TitleChange(ByVal Text As String)
form1.Caption = "OptiType 3.0 - " & Text
End Sub

Private Sub webbuttons_ButtonClick(ByVal Button As MSComctlLib.Button)
 'On Error Resume Next

 
 Select Case Button.Key
    
    Case Is = "vorige"
        On Error Resume Next
        
        If WebBrowser2.Visible = True Then WebBrowser2.GoBack Else WebBrowser1.GoBack
        
    
    Case Is = "volgende"           ' Open file.
            On Error Resume Next
            If WebBrowser2.Visible = True Then WebBrowser2.GoForward Else WebBrowser1.GoForward
            

        
    Case Is = "vernieuwen"              ' Save file.
            On Error Resume Next
            If WebBrowser2.Visible = True Then WebBrowser2.Refresh Else WebBrowser1.Refresh

        
    Case Is = "home"              ' Save file.
            If WebBrowser2.Visible = True Then WebBrowser2.GoHome Else WebBrowser1.GoHome
        
    Case Is = "stop"              ' Save file.
            On Error Resume Next
            WebBrowser1.Stop
            WebBrowser2.Stop
        
    
        
    Case Is = "fullscreen"      ' browser volledig scherm
            If WebBrowser2.Visible = False Then
    
            webbuttons.Buttons.Item(15).Visible = True
            pagina.Visible = True
            showwebbrowser2
                
                
            ElseIf WebBrowser2.Visible = True Then
            pagina.Visible = False
            
            webbuttons.Buttons.Item(15).Visible = False
         
            hidewebbrowser2
            
            
            End If
                
        
    Case Is = "norefresh"
            focusherstellen
             RichTextBox1.SaveFile "c:\temp.htm", rtfText
    
    
         WebBrowser1.Refresh
           On Error Resume Next: RichTextBox1.SetFocus
    


    Case Is = "popups"
       
If showpopups.Checked = True Then
    allowpopup = False
    showpopups.Checked = False
    webbuttons.Buttons.Item(11).Value = tbrUnpressed
Else
    allowpopup = True
    showpopups.Checked = True
    webbuttons.Buttons.Item(11).Value = tbrPressed
End If
    
    Case Is = "editsite"
        SaveHTML.Text = getsourcecode(form1.WebBrowser2.LocationURL)
        SaveHTML.SaveFile App.Path & "\page.txt", 1
        nieuwhtmlmakenzonderwizard
        form1.RichTextBox1.LoadFile App.Path & "\page.txt", 1
        focusherstellen
    


End Select



End Sub

Private Sub wordcountofsel_Click()

MsgBox "Number of words in selection: " & CountWords(2, RichTextBox1, False), vbInformation, "Wourd count"


End Sub

Private Sub wordcountplease_Click()

MsgBox "Number of words in document: " & CountWords(1, RichTextBox1, False), vbInformation, "Word count"



End Sub

Private Sub wordwrap_Click()

If WordWrap.Checked = False Then
    RichTextBox1.RightMargin = RichTextBox1.Width
    ' wordwrap.Checked = True
Else
    RichTextBox1.RightMargin = 0
    ' wordwrap.Checked = False
End If

End Sub

Private Sub X_Timer()

X.Enabled = False


End Sub

Private Sub X2_Timer()


X2.Interval = 10000

End Sub

Private Sub xcopy_Click()
frmXCopy.Show

End Sub



Private Sub Zoeken_Click()
frmFind.Option2.Value = True
frmFind.txtzoeknaar.Text = ""
frmFind.Show vbModeless, Me






End Sub

Sub MakeColor()

    RichTextBox1.SelText = "<"
    RichTextBox1.SelColor = vbBlue
End Sub

Private Sub setfont()
RichTextBox1.SelFontName = "Times New Roman"
RichTextBox1.SelFontSize = "12"
RichTextBox1.SelBold = False
RichTextBox1.SelItalic = False
RichTextBox1.SelUnderline = False
RichTextBox1.SelBullet = False
RichTextBox1.SelColor = QBColor(0)
RichTextBox1.SelStrikeThru = False

FontList.Text = "Times New Roman"
txtSize.Text = "12"

RichTextBox1.RightMargin = RichTextBox1.Width
' ' wordwrap.checked = True
RichTextBox1.SelAlignment = 0
'form1.on error resume next: richtextbox1.setfocus


RichTextBox1.SelIndent = 0
RichTextBox1.SelRightIndent = 0
RichTextBox1.SelHangingIndent = 0


lblDatum.Caption = Date



End Sub


Public Function CountWords(countmodus As Integer, rtBox As RichTextBox, BuildList As Boolean, Optional TargetList, Optional SizeLimit = 0) As Long

    'Keith Gardner (kgard@mhonline.net) - November 1999
    'Required Globals(Just cut and past the following into General De
    '     clarations.)
    'Public WordList() As String
    'Public WordCount As Long
    'Public UnqWordCount As Long
    'Counters
    Dim X As Long, y As Long 'Loop counters
    'Flags
    Dim AddNow As Boolean 'Flag to indicate If it is time to add word
    Dim FoundIt As Boolean 'Flag to indicate If word was found in list
    'Temporary storage
    Dim ThisText As String 'Holds full text of the rich text control
    Dim ThisWord As String 'Holds current word as it is built
    Dim ThisChar As Integer 'Holds ASCII value of current character
    Dim PrevChar As Integer 'Holds ASCII value of previous character
    
    
    If countmodus = 1 Then ThisText = Trim(rtBox.Text)
    If countmodus = 2 Then ThisText = Trim(rtBox.SelText)
    


    If ThisText = "" Then
        WordCount = -1
    Else
        WordCount = 0
        UnqWordCount = 0


        If BuildList Then
            ReDim wordlist(2, 1)
        End If


        PrevChar = 0


        For X = 1 To Len(ThisText)
            ThisChar = Asc(Mid(ThisText, X, 1))


            Select Case ThisChar
                Case 13 'Line feed


                If PrevChar <> 10 Then
                    WordCount = WordCount + 1
                    AddNow = True
                End If


                Case 32 'Space
                WordCount = WordCount + 1
                AddNow = True
                Case 10, 33, 34, 39, 40, 41, 63 'Ignore LF, "!", """, "'", "(", ")", "?"
                Case 44, 46 'Ignore "," or "." unless it's in a number
                If PrevChar >= 48 And PrevChar <= 57 Then ThisWord = ThisWord & Chr(ThisChar)
                Case Else 'ThisChar not a delimiter
                ThisWord = ThisWord & Chr(ThisChar)


                If X = Len(ThisText) - 1 Then
                    AddNow = True 'Add last word in list
                End If


            End Select


        'Building WordList?


        If BuildList And AddNow Then
            'Look for the word in the list


            For y = 1 To UnqWordCount


                If ThisWord = wordlist(1, y) Then 'Found it!
                    FoundIt = True
                    wordlist(2, y) = wordlist(2, y) + 1
                End If


                If FoundIt Then Exit For
            Next




            If Not FoundIt Then
                UnqWordCount = UnqWordCount + 1
                ReDim Preserve wordlist(2, UnqWordCount)
                wordlist(1, y) = ThisWord
                wordlist(2, y) = 1
            End If


            FoundIt = False 'Reset flag
            ThisWord = ""
        End If


        AddNow = False
        PrevChar = ThisChar
    Next


    WordCount = WordCount + 1 'Add one to the word count


    If Not IsMissing(TargetList) Then


        With TargetList


            For X = 1 To UnqWordCount


                If Len(wordlist(1, X)) > SizeLimit Then
                    .AddItem wordlist(1, X) & " - (" & wordlist(2, X) & ")"
                End If


            Next


        End With


    End If


End If


CountWords = WordCount
End Function



Sub checktosave()
annuleer = False

If Len(RichTextBox1.Text) > 1 Then
    
    If DocName.Text = "File not saved yet!" Then
        a = MsgBox("This file hasn't been saved yet.  Would you like to save it now?", vbYesNoCancel, "Save file")
        If a = 6 Then annuleer = False: OpslaanAls
        If a = 2 Then annuleer = True: Exit Sub
    Else
        a = MsgBox("Would you like to save changes to your file?", vbYesNoCancel, "Save changes")
        If a = 6 Then
                annuleer = False
                If UCase$(Right$(DocName.Text, 3)) = "RTF" Then
                RichTextBox1.SaveFile DocName.Text, 0
                Else
                RichTextBox1.SaveFile DocName.Text, 1
                End If
        End If
        If a = 2 Then annuleer = True: Exit Sub
    End If

End If
End Sub



Private Sub zoeknu()

Select Case Combo1.ListIndex

    Case 0   ' zoeken in huidig document
    
        form1.RichTextBox1.SelStart = 0
        If form1.RichTextBox1.SelStart = 0 And form1.RichTextBox1.SelLength = 0 Then
        lStart = InStr(form1.RichTextBox1.Text, txtZoeken.Text)
        Else
        lStart = InStr(form1.RichTextBox1.SelStart + 2, form1.RichTextBox1.Text, txtZoeken.Text)
        End If
    
        If lStart > 0 Then
        form1.RichTextBox1.SelStart = lStart - 1
        form1.RichTextBox1.SelLength = Len(txtZoeken.Text)
        Else
        MsgBox txtZoeken.Text & " hasn't been found in this file.", vbCritical, "String not found"
  
        End If

    Case 1  ' Google
            s = "http://www.google.com/search?q=" & txtZoeken.Text & "&hl=nl&lr="
    Case 2  ' altavista
            s = "http://www.altavista.com/sites/search/web?q=" & txtZoeken.Text & "&pg=q&kl=XX"
    Case 3  ' yahoo
           s = "http://search.yahoo.com/bin/search?p=" & txtZoeken.Text
    Case 4   ' webcrawler
            s = "http://dpxml.webcrawler.com/info.wbcrwl/dog/webresults.htm?&qkw=" & txtZoeken.Text
  
    Case 5   ' excite
            s = "http://search.excite.com/search.gw?searchType=Concept&search=" & txtZoeken.Text & "&category=default&mode=relevance&showqbe=1&display=html3,hb"
    Case 6  ' download.com
            s = "http://download.cnet.com/search/redirector/1,10207,0-0,00.html?qt=" & txtZoeken.Text & "&tg=dl-10001&cn=&tt=srch"


End Select

If Combo1.ListIndex > 0 Then

' webbrowser weergeven

If form1.WebBrowser1.Visible = False Then
 form1.RichTextBox1.Height = Fix(form1.RichTextBox1.Height / 2)
    form1.WebBrowser1.Visible = True
    WebBrowser1.Height = RichTextBox1.Height
    htmlview.Checked = True

    frmOpties.Check1.Value = 1
End If
    
    form1.WebBrowser1.Navigate s
    If WebBrowser2.Visible = True Then WebBrowser2.Navigate s


End If




End Sub
