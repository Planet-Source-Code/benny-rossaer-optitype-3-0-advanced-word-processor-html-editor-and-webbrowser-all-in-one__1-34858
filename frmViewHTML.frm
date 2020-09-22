VERSION 5.00
Begin VB.Form frmViewHTML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web page source (on-line files only)"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmViewHTML.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtViewHTML 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmViewHTML.frx":08CA
      Top             =   210
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frmViewHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmViewHTML

End Sub

Private Sub Form_Load()

txtViewHTML.Text = getsourcecode(frmBestand.WebBrowser1.LocationURL)



End Sub
Function getsourcecode(URL As String) As String

    getsourcecode = frmBestand.Inet1.OpenURL(URL)
End Function

