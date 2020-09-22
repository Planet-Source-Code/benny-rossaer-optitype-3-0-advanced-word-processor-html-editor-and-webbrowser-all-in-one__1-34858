VERSION 5.00
Begin VB.Form frmAddContact 
   Caption         =   "Create new contact..."
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   Icon            =   "frmAddContact.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   3945
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add contact"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1455
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1455
      TabIndex        =   7
      Top             =   3030
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2565
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2025
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   255
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Adress:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   1605
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cell phone:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2070
      Width           =   1065
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1125
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3045
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3525
      Width           =   855
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Close

Open App.Path & "\list.o2k" For Append As #1

Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
Print #1, Text5.Text
Print #1, Text6.Text
Print #1, Text7.Text
Print #1, Text8.Text


Close

getbook


MsgBox Text1.Text & " has been added to your adress book.", vbInformation, "OptiType"
Unload Me


End Sub

Private Sub Command2_Click()

Unload Me
End Sub

