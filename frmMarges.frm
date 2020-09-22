VERSION 5.00
Begin VB.Form frmMarges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust margins"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmMarges.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ToepassenOp 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Inspringen 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox RechterMarge 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox LinkerMarge 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Apply to:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "cm"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New paragraph at:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Right margin:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Left margin:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMarges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If ToepassenOp.ListIndex = 1 Then
    form1.RichTextBox1.SelStart = 0
    form1.RichTextBox1.SelLength = Len(form1.RichTextBox1.Text)
End If

Dim marge As Integer



marge = CInt(LinkerMarge.Text)

form1.RichTextBox1.SelIndent = marge * 400

marge = CInt(RechterMarge.Text)

If marge > 0 Then
    ' automatische terugloop inschakelen als er een rechtermarge
    ' is opgegeven.
    form1.RichTextBox1.RightMargin = form1.RichTextBox1.Width
   ' form1.' wordwrap.checked = True
End If

form1.RichTextBox1.SelRightIndent = marge * 400

marge = CInt(Inspringen.Text)

form1.RichTextBox1.SelHangingIndent = marge * 400




Unload Me

End Sub

Private Sub Form_Load()

ToepassenOp.AddItem "Selected text only"
ToepassenOp.AddItem "Complete document"

ToepassenOp.ListIndex = 0


LinkerMarge.Text = form1.RichTextBox1.SelIndent
RechterMarge.Text = form1.RichTextBox1.SelRightIndent
Inspringen.Text = form1.RichTextBox1.SelHangingIndent

End Sub
