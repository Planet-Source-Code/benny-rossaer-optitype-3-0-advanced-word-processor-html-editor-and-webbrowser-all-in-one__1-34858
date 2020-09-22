VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmBrief2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Letter"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   FillColor       =   &H00FFC0C0&
   Icon            =   "Brief2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   375
      Left            =   2505
      TabIndex        =   5
      Top             =   5625
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4800
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox txtOnderwerp 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   6015
   End
   Begin RichTextLib.RichTextBox txtBestemmeling 
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   1185
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Brief2.frx":08CA
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5295
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5625
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4) Select a phrase to end the letter:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3) Type in a greeting."
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2) Type in a subject for your letter."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1) Type in the data (name, adress, city) of the person you would like to adress this letter to."
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   6480
      X2              =   360
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You can now adjust several more settings.  You can still change these later."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmBrief2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBrief2.Hide
frmBrief3.Show vbModal


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
frmBrief2.Hide
frmBrief1.Show vbModal

End Sub

Private Sub Form_Load()



Combo1.AddItem "Dear sir,"
Combo1.AddItem "Dear parents,"

Combo1.ListIndex = 0

Combo2.AddItem "With kind regards,"
Combo2.AddItem "Regards,"
Combo2.AddItem "See you soon,"
Combo2.AddItem "Love,"


Combo2.ListIndex = 0




End Sub


