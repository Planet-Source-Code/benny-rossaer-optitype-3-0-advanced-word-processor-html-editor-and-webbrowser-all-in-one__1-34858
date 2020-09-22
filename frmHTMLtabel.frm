VERSION 5.00
Begin VB.Form frmHTMLtabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add table to webpage"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmHTMLtabel.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alignment"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "Right"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Center"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Border size="
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numer of columns:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of rows:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmHTMLtabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Dim tabel As String


If Option1.Value = True Then tabel = "<DIV ALIGN=LEFT>"
If Option2.Value = True Then tabel = "<DIV ALIGN=CENTER>"
If Option3.Value = True Then tabel = "<DIV ALIGN=RIGHT>"

tabel = tabel & "<TABLE WIDTH=100% BORDER=" & Combo3.ListIndex + 1 & ">"

' <TR> = rijen
' <TD> = kolommen


For i = 1 To Combo1.ListIndex + 1  ' voor rijen

    tabel = tabel & "<TR>"
    
        For j = 1 To Combo2.ListIndex + 1
            tabel = tabel & "   " & "<td width=50%>&nbsp;</td>"
            Next j

    tabel = tabel & "</TR>"
    
Next

tabel = tabel & "</TABLE>" & "</DIV>"

If bron = 1 Then Kolommen.RichTextBox1.SelText = Kolommen.RichTextBox1.SelText & tabel
If bron = 2 Then Kolommen.RichTextBox2.SelText = Kolommen.RichTextBox2.SelText & tabel
If bron = 3 Then form1.RichTextBox1.SelText = form1.RichTextBox1.SelText & tabel

Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

For i = 1 To 30
    Combo1.AddItem i
    Combo2.AddItem i
Next
Combo1.ListIndex = 0
Combo2.ListIndex = 0

For i = 1 To 10
    Combo3.AddItem i
Next
Combo3.ListIndex = 0

Option1.Value = True


End Sub
