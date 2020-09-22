VERSION 5.00
Begin VB.Form frmBrief3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create letter"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "Brief3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDatum 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtOnsKenmerk 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtUwKenmerk 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtBericht 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "This is the final step.  You may click Finish to create your letter."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "OptiType has filled in the date according to the current date of your system.  You can change this if needed."
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Our reference"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your reference"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your message"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   5160
      X2              =   240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You can now fill in the properties of your letter.  These will be used to comply to the BIN standards."
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmBrief3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
leegdocumentmaken



    form1.Caption = "OptiType 3.0 - New letter"
autotekstinvoegen

'form1.' wordwrap.checked = False
'Dim kop As String
'Dim gegevens As String

kop = frmBrief1.txtAfzender.Text & vbCrLf & vbCrLf & frmBrief2.txtBestemmeling.Text & vbCrLf & vbCrLf & vbCrLf

gegevens1 = "your message                  your reference                         our reference                  date" & vbCrLf
gegevens2 = gegevens2 & frmBrief3.txtBericht.Text & Space$(32 - Len(frmBrief3.txtBericht.Text))
gegevens2 = gegevens2 & frmBrief3.txtUwKenmerk.Text & Space$(35 - Len(frmBrief3.txtUwKenmerk.Text))
gegevens2 = gegevens2 & frmBrief3.txtOnsKenmerk.Text & Space$(29 - Len(frmBrief3.txtOnsKenmerk.Text))
gegevens2 = gegevens2 & frmBrief3.txtDatum.Text
gegevens2 = gegevens2

form1.RichTextBox1.Text = ""
form1.RichTextBox1.Text = vbCrLf & vbCrLf & kop
form1.RichTextBox1.Text = form1.RichTextBox1.Text & gegevens1
form1.RichTextBox1.Text = form1.RichTextBox1.Text & gegevens2 & vbCrLf & vbCrLf

form1.RichTextBox1.Text = form1.RichTextBox1.Text & frmBrief2.txtOnderwerp.Text & vbCrLf & vbCrLf




aanspr = frmBrief2.Combo1.Text



form1.RichTextBox1.Text = form1.RichTextBox1.Text & aanspr & vbCrLf & vbCrLf & "[TYPE IN THE CONTENTS OF YOUR LETTER]" & vbCrLf & vbCrLf


einde = frmBrief2.Combo2.Text

form1.RichTextBox1.Text = form1.RichTextBox1.Text & einde & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

form1.RichTextBox1.Text = form1.RichTextBox1.Text & frmBrief1.txtNaam & vbCrLf & frmBrief1.txtFunctie.Text



' EIGENSCHAPPENREGEL EN ONDERWERP IN VET!!! ---------------
X = form1.RichTextBox1.Find(frmBrief2.txtOnderwerp.Text)
form1.RichTextBox1.SelBold = True
    form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
a = form1.RichTextBox1.Find("uw bericht van")
form1.RichTextBox1.SelBold = True
    form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
b = form1.RichTextBox1.Find("uw kenmerk")
form1.RichTextBox1.SelBold = True
form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
c = form1.RichTextBox1.Find("ons kenmerk")
form1.RichTextBox1.SelBold = True
form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
d = form1.RichTextBox1.Find("datum")
form1.RichTextBox1.SelBold = True
form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
' ----------------------------------------------------------


' MARGES INSTELLEN -----------------------------------------
    form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = Len(form1.RichTextBox1.Text)
Dim marge As Integer
form1.RichTextBox1.SelIndent = 3 * 400
    form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = 0
' EINDE MARGES INSTELLEN ---------------------





MsgBox "OptiType has created the basic layout of your letter.  You may now do the following:" & vbCrLf & vbCrLf & "1) Erase the text [TYPE IN THE CONTENTS OF YOUR LETTER] and start typing your letter here/" & vbCrLf & vbCrLf & "2) Check the entire letter and make adjustments if needed. Save the file and print it using the File menu." & vbCrLf & vbCrLf & "3) At the bottom of the document, above your name, there are 6 empty lines.  This is where you can put your signature.", vbInformation, "Create Letter Wizard"


On Error Resume Next

form1.Show
form1.RichTextBox1.SetFocus

 form1.RichTextBox1.Height = form1.Height - 2640
    'form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False
    form1.htmlview.Checked = False

Unload frmBrief1
Unload frmBrief2
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
frmBrief3.Hide
frmBrief2.Show vbModal


End Sub

Private Sub Form_Load()
txtDatum.Text = Year(Date) & "-" & Month(Date) & "-" & Day(Date)






End Sub
