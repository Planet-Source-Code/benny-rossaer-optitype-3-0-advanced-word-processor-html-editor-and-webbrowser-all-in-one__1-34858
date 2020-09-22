VERSION 5.00
Begin VB.Form frmVervangen 
   Caption         =   "Tekst vervangen"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Vervangen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Vervangen_Cancel 
      Caption         =   "Annuleren"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Vervangen_OK 
      Caption         =   "OK"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Vervangen_Door 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Vervangen_Zoektekst 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vervangen door:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Typ de tekst in die u wilt zoeken:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmVervangen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Vervangen_Cancel_Click()
Unload Me

End Sub

Private Sub Vervangen_OK_Click()

Unload Me

a = form1.RichTextBox1.Find(Vervangen_Zoektekst)
If a = -1 Then form1.mnuhelp.Caption = " | Zoekstring niet gevonden in document."


form1.RichTextBox1.SelText = Vervangen_Door






End Sub
