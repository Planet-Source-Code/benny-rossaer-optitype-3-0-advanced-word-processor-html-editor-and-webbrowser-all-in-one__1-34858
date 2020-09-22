VERSION 5.00
Begin VB.Form frmAfdrukken 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Afdrukken"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "Afdrukken.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Afdrukken.frx":08CA
   ScaleHeight     =   2655
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kleur"
      Height          =   255
      Left            =   3480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Wit/zwart"
      Height          =   255
      Left            =   3480
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Afdrukken"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Resolutie:"
      Height          =   1335
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hoog"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Gemiddeld"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Laag"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   5040
      TabIndex        =   16
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   3360
      TabIndex        =   15
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Paginastand:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aantal exemplaren:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Printing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Opti"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAfdrukken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo oeps

If Combo2.ListIndex = 0 Then Printer.Orientation = 1
If Combo2.ListIndex = 1 Then Printer.Orientation = 2

If Option1.Value = True Then Printer.ColorMode = 2
If Option2.Value = True Then Printer.ColorMode = 1

If Option3.Value = True Then Printer.PrintQuality = -2
If Option4.Value = True Then Printer.PrintQuality = -3
If Option5.Value = True Then Printer.PrintQuality = -4

For i = 1 To Combo1.ListIndex + 1
    Printer.Print ""
    form1.RichTextBox1.SelPrint Printer.hDC
    Printer.EndDoc
Next

frmAfdrukken.Hide
' err

Exit Sub

oeps:
If language = False Then MsgBox "Er is een fout opgetreden bij het printen." & vbCrLf & vbCrLf & Err & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Probeer het later opnieuw.", vbCritical, "Fout"
If language = True Then MsgBox "OptiType has encountered an error while trying to print." & vbCrLf & vbCrLf & Err & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Try again later.", vbCritical, "Error"


frmAfdrukken.Hide

End Sub

Private Sub Command2_Click()
frmAfdrukken.Hide
End Sub

Private Sub Form_Load()

For i = 1 To 50
    Combo1.AddItem i
Next i

Combo1.ListIndex = 0

Combo2.AddItem "Staand"
Combo2.AddItem "Liggend"
Combo2.ListIndex = 0

Option1.Value = True
Option4.Value = True



End Sub

