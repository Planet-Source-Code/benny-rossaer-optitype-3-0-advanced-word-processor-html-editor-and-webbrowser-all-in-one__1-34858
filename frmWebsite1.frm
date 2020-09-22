VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmWebsite1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OptiType Wizard - Create new webpage"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmWebsite1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   7605
      TabIndex        =   26
      Top             =   3210
      Width           =   1200
   End
   Begin VB.TextBox txtbeschrijving 
      Height          =   360
      Left            =   3360
      TabIndex        =   8
      Top             =   6315
      Width           =   5280
   End
   Begin VB.TextBox txtkeywords 
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Top             =   5685
      Width           =   5280
   End
   Begin VB.TextBox txtauteur 
      Height          =   345
      Left            =   3360
      TabIndex        =   6
      Top             =   5085
      Width           =   5280
   End
   Begin MSComDlg.CommonDialog SaveFileDialog 
      Left            =   7200
      Top             =   3930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox savehtml 
      Height          =   855
      Left            =   7830
      TabIndex        =   16
      Top             =   3885
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      _Version        =   393217
      TextRTF         =   $"frmWebsite1.frx":08CA
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3315
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4395
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3315
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3795
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   5985
      TabIndex        =   9
      Top             =   7050
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3330
      TabIndex        =   3
      Top             =   3225
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3330
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2625
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3330
      TabIndex        =   1
      Top             =   2040
      Width           =   5280
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "The three settings on the left configure the site's META tags and helps search engines to find your page.  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   -15
      TabIndex        =   25
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      Height          =   810
      Left            =   -15
      TabIndex        =   23
      Top             =   6900
      Width           =   9075
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   270
      Left            =   315
      TabIndex        =   22
      Top             =   6345
      Width           =   2880
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Keywords:"
      Height          =   345
      Left            =   315
      TabIndex        =   21
      Top             =   5715
      Width           =   2880
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author's name:"
      Height          =   345
      Left            =   315
      TabIndex        =   20
      Top             =   5100
      Width           =   2880
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Create new webpage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   180
      TabIndex        =   18
      Top             =   195
      Width           =   7035
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9075
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Left margin:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   270
      TabIndex        =   15
      Top             =   4365
      Width           =   2880
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Upper margin:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   270
      TabIndex        =   14
      Top             =   3855
      Width           =   2880
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Filename for picture on background:"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   285
      TabIndex        =   13
      Top             =   3180
      Width           =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Background color:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   285
      TabIndex        =   12
      Top             =   2640
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Web page title (appears on browser window):"
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   285
      TabIndex        =   11
      Top             =   1995
      Width           =   2880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Here you can change the settings of your new web page by adjusting the options below.  All options are optional."
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   255
      TabIndex        =   0
      Top             =   1170
      Width           =   6495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Height          =   750
      Left            =   0
      TabIndex        =   19
      Top             =   975
      Width           =   9075
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
      Height          =   1995
      Left            =   90
      TabIndex        =   24
      Top             =   4920
      Width           =   9075
   End
End
Attribute VB_Name = "frmWebsite1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

htmlheader = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 3.2 Final//EN'>"
htmlheader = htmlheader & vbCrLf & "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & Text1.Text & "</TITLE>" & vbCrLf
htmlheader = htmlheader & "<META NAME='Generator' CONTENT='OptiType - (C) Benny Rossaer - E-mail: benny.rossaer@pandora.be'>"
htmlheader = htmlheader & vbCrLf & "<META NAME='Author' CONTENT='" & txtauteur.Text & "'>" & vbCrLf
htmlheader = htmlheader & "<META NAME='Keywords' CONTENT='" & txtkeywords.Text & "'>" & vbCrLf
htmlheader = htmlheader & "<META NAME='Description' CONTENT='" & txtbeschrijving.Text & "'>" & vbCrLf
htmlheader = htmlheader & "</HEAD>" & vbCrLf & "<BODY "

If Combo1.ListIndex = 0 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=WHITE"
If Combo1.ListIndex = 1 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=BLACK"
If Combo1.ListIndex = 2 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=RED"
If Combo1.ListIndex = 3 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=YELLOW"
If Combo1.ListIndex = 4 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=GREEN"
If Combo1.ListIndex = 5 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=BLUE"
If Combo1.ListIndex = 6 Then htmlheader = htmlheader & "TOPMARGIN=" & Combo2.ListIndex & " LEFTMARGIN=" & Combo3.ListIndex & " BGCOLOR=BRAUN"

If Len(Text2.Text) > 0 Then htmlheader = htmlheader & " BACKGROUND='" & Text2.Text & "'"


htmlheader = htmlheader & ">" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"

form1.RichTextBox1.Text = htmlheader
Unload Me



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

SaveFileDialog.DialogTitle = "Selecteer een achtergrondafbeelding..."
SaveFileDialog.CancelError = False
SaveFileDialog.Filter = "JPEG-afbeelding (*.jpg)|*.jpg|GIF-afbeelding (*.gif)|*.gif|Bitmap-afbeelding (*.bmp)|*.bmp|Alle bestanden (*.*)|*.*"
SaveFileDialog.FilterIndex = 0

SaveFileDialog.ShowOpen

Text2.Text = SaveFileDialog.FileName

End Sub

Private Sub Form_Load()

Open App.Path & "\user.txt" For Input As #1
Dim maker As String
Line Input #1, maker
Close

txtauteur.Text = maker

Combo1.Clear
Combo2.Clear
Combo3.Clear


For i = 0 To 10
Combo2.AddItem i
Combo3.AddItem i
Next

Combo2.ListIndex = 2
Combo3.ListIndex = 2



Combo1.AddItem "Wit"
Combo1.AddItem "Zwart"
Combo1.AddItem "Rood"
Combo1.AddItem "Geel"
Combo1.AddItem "Groen"
Combo1.AddItem "Blauw"
Combo1.AddItem "Bruin"

Combo1.ListIndex = 0
Combo2.ListIndex = 2
Combo3.ListIndex = 2





End Sub

