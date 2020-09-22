VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmSendmail 
   BackColor       =   &H00808000&
   Caption         =   "Nieuw e-mailbericht - OptiMail"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog printdialog 
      Left            =   3960
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog invoegen 
      Left            =   4800
      Top             =   4305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4905
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog bijlage 
      Left            =   5565
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bijlage..."
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3735
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   327680
      ScrollBars      =   3
      TextRTF         =   $"frmSendmail.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Annuleer"
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Top             =   2040
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verstuur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   600
      Width           =   4950
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   120
      Width           =   4950
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Benny Rossaer, 2000"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Left            =   4905
      TabIndex        =   10
      Top             =   5295
      Width           =   1770
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bericht:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Onderwerp:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mailadres ontvanger:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   165
      TabIndex        =   0
      Top             =   45
      Width           =   1815
   End
   Begin VB.Menu file 
      Caption         =   "&Bestand"
      Begin VB.Menu new 
         Caption         =   "&Nieuw"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Openen"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Opslaan"
         Shortcut        =   ^S
      End
      Begin VB.Menu dezdezdezdez 
         Caption         =   "-"
      End
      Begin VB.Menu printit 
         Caption         =   "&Afdrukken"
         Shortcut        =   ^P
      End
      Begin VB.Menu itsoutthere 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "&Afsluiten"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Bewerken"
      Begin VB.Menu cut 
         Caption         =   "&Knippen"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "K&opiëren"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Plakken"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete 
         Caption         =   "&Verwijderen"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu fdez 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Alles selecteren"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Invoegen"
      Begin VB.Menu txt 
         Caption         =   "&Tekstbestand"
      End
      Begin VB.Menu insertattach 
         Caption         =   "&Bijlage"
      End
      Begin VB.Menu tze 
         Caption         =   "-"
      End
      Begin VB.Menu invoegendatum 
         Caption         =   "&Datum"
      End
      Begin VB.Menu invoegentijd 
         Caption         =   "&Tijd"
      End
   End
   Begin VB.Menu extr 
      Caption         =   "&Extra"
      Begin VB.Menu countwordsordie 
         Caption         =   "&Woorden tellen..."
      End
   End
   Begin VB.Menu h 
      Caption         =   "&Help"
      Begin VB.Menu info 
         Caption         =   "&Info"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmSendmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub countwordsordie_Click()

End Sub

Private Sub delete_Click()

End Sub

Private Sub Form_Load()

End Sub
