VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find.."
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Search..."
      Height          =   1170
      Left            =   1290
      TabIndex        =   6
      Top             =   855
      Width           =   3540
      Begin VB.OptionButton Option2 
         Caption         =   "Start searching from the cursor's position"
         Height          =   420
         Left            =   270
         TabIndex        =   3
         Top             =   645
         Value           =   -1  'True
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In the entire file"
         Height          =   360
         Left            =   285
         TabIndex        =   2
         Top             =   270
         Width           =   2805
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   400
      Left            =   3765
      TabIndex        =   5
      Top             =   2145
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find next..."
      Height          =   400
      Left            =   1920
      TabIndex        =   4
      Top             =   2145
      Width           =   1725
   End
   Begin VB.TextBox txtzoeknaar 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   300
      Width           =   3540
   End
   Begin VB.Label Label1 
      Caption         =   "Find this:"
      Height          =   360
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   1275
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


 If form1.RichTextBox1.SelStart = 0 And form1.RichTextBox1.SelLength = 0 Then
  lStart = InStr(form1.RichTextBox1.Text, txtzoeknaar.Text)
 Else
  lStart = InStr(form1.RichTextBox1.SelStart + 2, form1.RichTextBox1.Text, txtzoeknaar.Text)
 End If

 'lStart = InStr(form1.RichTextBox1.Text, txtzoeknaar.Text)

 If lStart > 0 Then
  form1.RichTextBox1.SelStart = lStart - 1
  form1.RichTextBox1.SelLength = Len(txtzoeknaar.Text)
 Else
  MsgBox txtzoeknaar.Text & " hasn't been found.", vbCritical, "String not found"
  
 End If
 
 
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Option1_Click()
form1.RichTextBox1.SelStart = 0
End Sub
