VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OptiType 2.0 Credits"
   ClientHeight    =   2730
   ClientLeft      =   2715
   ClientTop       =   2610
   ClientWidth     =   3990
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0ECA
   ScaleHeight     =   2730
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmCredits.Hide

End Sub

Private Sub Form_Load()


lblCredits.Caption = "OptiType 2.0" & vbCrLf & "Door Benny Rossaer" & vbCrLf & vbCrLf & "Met dank aan: " & vbCrLf & "Brady Hegberg voor HTML conversie" & vbCrLf & "Jason Shimkoski voor Undo en Redo" & vbCrLf & vbCrLf & "Planet SourceCode" & vbCrLf & "(www.planet-source-code.com)"




End Sub

