VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   15
   ClientTop       =   1380
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   360
      Top             =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Benny Rossaer, 1999 - 2002."
      Height          =   540
      Left            =   2730
      TabIndex        =   2
      Top             =   2505
      Width           =   2370
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to OptiType."
      Height          =   435
      Left            =   2520
      TabIndex        =   1
      Top             =   1755
      Width           =   2580
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OptiType 3.0 PSC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   1380
      TabIndex        =   0
      Top             =   435
      Width           =   4650
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
form1.Show


End Sub

