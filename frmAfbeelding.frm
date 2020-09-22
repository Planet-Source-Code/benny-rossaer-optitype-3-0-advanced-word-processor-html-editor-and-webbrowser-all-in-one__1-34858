VERSION 5.00
Begin VB.Form frmAfbeelding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML: Insert picture"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmAfbeelding.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert picture"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtHyperlink 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alignment:"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
      Begin VB.OptionButton Option3 
         Caption         =   "Right"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Center"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
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
      Caption         =   "Border size:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If you want to add a hyperlink to this image, please type in the URL here:"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   10
      Top             =   1095
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in the filename of the picture you want to insert:"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   3210
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   660
   End
End
Attribute VB_Name = "frmAfbeelding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim afb As String

If Len(txtHyperlink.Text) > 0 Then

    afb = "<A HREF='" & txtHyperlink.Text & "'><IMG SRC='" & txtFileName.Text & "' "
    If Option1.Value = True Then afb = afb & "ALIGN=LEFT "
    If Option2.Value = True Then afb = afb & "ALIGN=CENTER "
    If Option3.Value = True Then afb = afb & "ALIGN=RIGHT "
    
    afb = afb & "BORDER=" & Str$(Combo1.ListIndex + 1) & "></A>"
    
Else


    afb = "<IMG SRC='" & txtFileName.Text & "' "
    If Option1.Value = True Then afb = afb & "ALIGN=LEFT "
    If Option2.Value = True Then afb = afb & "ALIGN=CENTER "
    If Option3.Value = True Then afb = afb & "ALIGN=RIGHT "
    
    afb = afb & "BORDER=" & Str$(Combo1.ListIndex + 1) & ">"
End If

If bron = 1 Then Kolommen.RichTextBox1.SelText = Kolommen.RichTextBox1.SelText & afb
If bron = 2 Then Kolommen.RichTextBox2.SelText = Kolommen.RichTextBox2.SelText & afb
If bron = 3 Then form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & afb


Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Option1.Value = True

For i = 1 To 10
    Combo1.AddItem i
Next

Combo1.ListIndex = 0


End Sub
