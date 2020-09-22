VERSION 5.00
Begin VB.Form frmTabel 
   Caption         =   "Insert table..."
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmTabel.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox kolomgrootte 
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   750
      Width           =   690
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1425
      Width           =   1335
   End
   Begin VB.CommandButton cmdTabelInvoegen 
      Caption         =   "Insert table"
      Height          =   375
      Left            =   1230
      TabIndex        =   4
      Top             =   1425
      Width           =   1455
   End
   Begin VB.ComboBox Kolommen 
      Height          =   315
      Left            =   3645
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.ComboBox Rijen 
      Height          =   315
      ItemData        =   "frmTabel.frx":08CA
      Left            =   1080
      List            =   "frmTabel.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Column width in spaces:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Columns:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rows:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmTabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdTabelInvoegen_Click()

'MsgBox Rijen.ListIndex
'MsgBox Kolommen.ListIndex

Dim tabel_kolom As String


If Kolommen.ListIndex = 0 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 1 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 2 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 3 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 4 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 5 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"
If Kolommen.ListIndex = 6 Then tabel_kolom = "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|" & Space$(kolomgrootte.ListIndex + 6) & "|"

Dim rij As String




For i = 1 To Rijen.ListIndex + 1


form1.RichTextBox1.SelText = vbCrLf & form1.RichTextBox1.SelText & tabel_kolom

Next

Unload Me


End Sub

Private Sub Form_Load()



For i = 1 To 30
Rijen.AddItem i
Next

Rijen.ListIndex = 4

For i = 1 To 7
Kolommen.AddItem i
Next

Kolommen.ListIndex = 1



For i = 5 To 50
kolomgrootte.AddItem i
Next

kolomgrootte.ListIndex = 15


End Sub


