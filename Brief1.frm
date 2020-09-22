VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmBrief1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Letter"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "Brief1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFunctie 
      Height          =   285
      Left            =   3135
      TabIndex        =   3
      Top             =   4335
      Width           =   2175
   End
   Begin VB.TextBox txtNaam 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox txtAfzender 
      Height          =   2205
      Left            =   1215
      TabIndex        =   1
      Top             =   1410
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3889
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Brief1.frx":08CA
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Your function:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name to close letter:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   5280
      X2              =   240
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please type in your personal data (name, adress, city) in the text box below."
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   345
      TabIndex        =   6
      Top             =   705
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to OptiType's Create Letter wizard.  This wizard helps you to create a new letter, and does most of the work for you."
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmBrief1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBrief1.Hide
frmBrief2.Show vbModal


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

