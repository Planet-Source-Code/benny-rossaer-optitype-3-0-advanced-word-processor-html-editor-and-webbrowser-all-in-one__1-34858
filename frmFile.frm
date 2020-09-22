VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmFile 
   Caption         =   "Form2"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   8310
   Begin VB.CheckBox chkVET 
      DisabledPicture =   "frmFile.frx":0000
      DownPicture     =   "frmFile.frx":0F42
      Height          =   375
      Left            =   3360
      Picture         =   "frmFile.frx":1F7C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Vet (CTRL+B)"
      Top             =   560
      Width           =   375
   End
   Begin VB.CheckBox chkCursief 
      DisabledPicture =   "frmFile.frx":2EBE
      DownPicture     =   "frmFile.frx":3DF2
      Height          =   375
      Left            =   3720
      Picture         =   "frmFile.frx":4E31
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Cursief (CTRL+I)"
      Top             =   560
      Width           =   375
   End
   Begin VB.CheckBox chkOnderl 
      DisabledPicture =   "frmFile.frx":5D65
      DownPicture     =   "frmFile.frx":6E24
      Height          =   375
      Left            =   4080
      Picture         =   "frmFile.frx":800A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Onderlijnd (CTRL+U)"
      Top             =   560
      Width           =   375
   End
   Begin VB.CommandButton Lettertype 
      Height          =   375
      Left            =   5760
      Picture         =   "frmFile.frx":90C9
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Lettertype wijzigen"
      Top             =   560
      Width           =   375
   End
   Begin VB.CommandButton NIEUW 
      Height          =   375
      Left            =   600
      Picture         =   "frmFile.frx":AAFF
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nieuw OptiWORD RTF bestand maken"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton OPEN 
      Height          =   375
      Left            =   960
      Picture         =   "frmFile.frx":BC1F
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "RTF of TXT bestand openen"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton OPSLAAN 
      Height          =   375
      Left            =   1320
      Picture         =   "frmFile.frx":D136
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Bestand opslaan"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Afdrukken 
      Height          =   375
      Left            =   1800
      Picture         =   "frmFile.frx":E60D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Adrukken"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Centreren 
      Height          =   375
      Left            =   4920
      Picture         =   "frmFile.frx":E8B9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   560
      Width           =   375
   End
   Begin VB.CommandButton LinksUitlijnen 
      Height          =   375
      Left            =   4560
      Picture         =   "frmFile.frx":FA98
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   560
      Width           =   375
   End
   Begin VB.CommandButton RECHTSUITLIJNEN 
      Height          =   375
      Left            =   5280
      Picture         =   "frmFile.frx":10DCD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   560
      Width           =   375
   End
   Begin VB.CheckBox Opsomming 
      DisabledPicture =   "frmFile.frx":11F79
      DownPicture     =   "frmFile.frx":121F5
      Height          =   375
      Left            =   6240
      Picture         =   "frmFile.frx":1355D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   560
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   315
      LargeChange     =   4
      Left            =   2760
      Max             =   2
      Min             =   100
      TabIndex        =   2
      Top             =   600
      Value           =   10
      Width           =   315
   End
   Begin VB.TextBox lblFONT 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox DocName 
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Text            =   "Document nog niet bewaard!"
      Top             =   6600
      Width           =   4215
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   0
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   0
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   600
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Typ hier uw tekst in."
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9551
      _Version        =   327680
      ScrollBars      =   3
      BulletIndent    =   400
      Appearance      =   0
      TextRTF         =   $"frmFile.frx":137D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      ToolTipText     =   "Text Size"
      Top             =   600
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   3360
      Top             =   560
      Width           =   1095
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
