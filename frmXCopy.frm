VERSION 5.00
Begin VB.Form frmXCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy file(s) with XCOPY"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmXCopy.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert command"
      Height          =   375
      Left            =   375
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Overwrite automatically if file already exists"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include subfolders"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination (folder or file(s)):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source file(s):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmXCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

opdracht = "XCOPY " & Text1.Text & " " & Text2.Text
If Check1.Value = 1 Then opdracht = opdracht & " /S"
If Check2.Value = 1 Then opdracht = opdracht & " /Y"

form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & opdracht & vbCrLf
frmXCopy.Hide
form1.Show

On Error Resume Next: form1.RichTextBox1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
form1.Show

End Sub

