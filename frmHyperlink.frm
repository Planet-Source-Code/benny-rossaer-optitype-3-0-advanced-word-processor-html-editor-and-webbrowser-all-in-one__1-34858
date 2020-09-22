VERSION 5.00
Begin VB.Form frmHyperlink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add hyperlink to web page"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmHyperlink.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2625
      TabIndex        =   5
      Top             =   2970
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox txtTekst 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Optional: type in the target frame below."
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in the text that will represent the link (for example ""click here"")."
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in the URL you want the hyperlink to refer to:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4455
   End
End
Attribute VB_Name = "frmHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txtTarget.Text = UCase$(txtTarget.Text)

Dim hyperlink As String

If Not Len(txtTarget.Text) > 0 Then
    hyperlink = "<A HREF='" & txtFileName.Text & "'>" & txtTekst.Text & "</A>"
Else
    hyperlink = "<A HREF='" & txtFileName.Text & "' TARGET='" & txtTarget.Text & "'>" & txtTekst.Text & "</A>"
End If

If bron = 1 Then Kolommen.RichTextBox1.SelText = Kolommen.RichTextBox1.SelText & hyperlink
If bron = 2 Then Kolommen.RichTextBox2.SelText = Kolommen.RichTextBox2.SelText & hyperlink
If bron = 3 Then form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & hyperlink


Unload Me




End Sub

Private Sub Command2_Click()
Unload Me

End Sub

