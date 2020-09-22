VERSION 5.00
Begin VB.Form frmFormatDisk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format a disk (batch command)"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmFormatDisk.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Make disk bootable "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Quick format"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2190
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert command"
      Height          =   375
      Left            =   390
      TabIndex        =   3
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the drive you would like to format from the batch file:"
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
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmFormatDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
opdracht = "FORMAT " & Left$(Drive1.Drive, 2)


If Check1.Value = 1 Then opdracht = opdracht & " /Q"
If Check2.Value = 1 Then opdracht = opdracht & " /S"

form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & opdracht & vbCrLf

form1.Show

On Error Resume Next: form1.RichTextBox1.SetFocus
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

