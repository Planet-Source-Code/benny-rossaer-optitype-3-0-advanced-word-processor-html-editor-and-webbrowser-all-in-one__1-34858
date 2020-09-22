VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and replace"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   Icon            =   "frmReplace.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Replace"
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   2610
      Width           =   1350
   End
   Begin VB.TextBox txtvervangdoor 
      Height          =   315
      Left            =   1515
      TabIndex        =   2
      Top             =   735
      Width           =   3315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search..."
      Height          =   1170
      Left            =   1305
      TabIndex        =   8
      Top             =   1320
      Width           =   3540
      Begin VB.OptionButton Option2 
         Caption         =   "Start searching from the cursor's position"
         Height          =   420
         Left            =   285
         TabIndex        =   4
         Top             =   645
         Value           =   -1  'True
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Caption         =   "The entire file"
         Height          =   360
         Left            =   300
         TabIndex        =   3
         Top             =   270
         Width           =   2805
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   400
      Left            =   3780
      TabIndex        =   7
      Top             =   2610
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find next..."
      Height          =   400
      Left            =   375
      TabIndex        =   5
      Top             =   2610
      Width           =   1725
   End
   Begin VB.TextBox txtzoeknaar 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   300
      Width           =   3360
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with:"
      Height          =   480
      Left            =   210
      TabIndex        =   9
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Find:"
      Height          =   360
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   1275
   End
End
Attribute VB_Name = "frmReplace"
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
  MsgBox txtzoeknaar.Text & " has not been found in this file.", vbCritical, "String not found"
  
 End If
 
 
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
If form1.RichTextBox1.SelText = txtzoeknaar.Text Then
  form1.RichTextBox1.SelText = txtvervangdoor.Text
 End If
 If form1.RichTextBox1.SelStart = 0 And form1.RichTextBox1.SelLength = 0 Then
  lStart = InStr(form1.RichTextBox1.Text, txtzoeknaar.Text)
 Else
  lStart = InStr(form1.RichTextBox1.SelStart + 2, form1.RichTextBox1.Text, txtzoeknaar.Text)
 End If

 If lStart > 0 Then
  form1.RichTextBox1.SelStart = lStart - 1
  form1.RichTextBox1.SelLength = Len(txtzoeknaar.Text)
 Else
 MsgBox txtzoeknaar.Text & " has not been found in this file.", vbCritical, "String not found"
 End If
End Sub

Private Sub Option1_Click()
form1.RichTextBox1.SelStart = 0
End Sub
