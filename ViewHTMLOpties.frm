VERSION 5.00
Begin VB.Form ViewHTMLOpties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Preview options"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   ControlBox      =   0   'False
   ForeColor       =   &H80000000&
   Icon            =   "ViewHTMLOpties.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Web page preview size:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "ViewHTMLOpties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

If form1.WebBrowser1.Visible = True Then

If Combo1.ListIndex = 0 Then  ' 25 %
form1.RichTextBox1.Height = ((form1.Height - 2640) * (3 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (1 / 4))
nieuwformaat

End If

If Combo1.ListIndex = 1 Then
    form1.RichTextBox1.Height = (form1.Height - 2640) / 2
    form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
    form1.WebBrowser1.Height = (form1.Height - 2640) / 2
    nieuwformaat
End If
    
If Combo1.ListIndex = 2 Then  ' 75 %
form1.RichTextBox1.Height = ((form1.Height - 2640) * (1 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (3 / 4))
nieuwformaat
    End If
    
    End If
    
verversing = Val(Combo2.Text)


Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Activate()

If form1.WebBrowser1.Visible = False Then
    Label1.Enabled = False
    Combo1.Enabled = False
Else
    Label1.Enabled = True
    Combo1.Enabled = True
End If


End Sub

Private Sub Form_GotFocus()
If form1.WebBrowser1.Visible = False Then
    Label1.Enabled = False
    Combo1.Enabled = False
Else
    Label1.Enabled = True
    Combo1.Enabled = True
End If

End Sub

Private Sub Form_Load()

If form1.WebBrowser1.Visible = False Then
    Label1.Enabled = False
    Combo1.Enabled = False
Else
    Label1.Enabled = True
    Combo1.Enabled = True
End If


Combo1.AddItem "25%"
Combo1.AddItem "50%"
Combo1.AddItem "75%"

'If form1.RichTextBox1.Height = 4059 Then Combo1.ListIndex = 0
'If form1.RichTextBox1.Height = Fix(5415 / 2) Then Combo1.ListIndex = 1
'If form1.RichTextBox1.Height = 1354 Then Combo1.ListIndex = 2

Combo1.ListIndex = 1



End Sub

