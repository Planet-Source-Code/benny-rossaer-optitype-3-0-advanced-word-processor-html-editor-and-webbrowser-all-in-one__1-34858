VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOpties 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OptiType 3.0 - Options and settings"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmOpties.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   4305
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "View"
      Height          =   3255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox Check4 
         Caption         =   "Show date"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Show status bar"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "These settings are saved and applied again the next time you start OptiType."
         Height          =   690
         Left            =   105
         TabIndex        =   15
         Top             =   990
         Width           =   4695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   4305
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   375
      Index           =   0
      Left            =   15
      TabIndex        =   10
      Top             =   3705
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Save"
      Height          =   3255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox Check3 
         Caption         =   "Ask before quick-saving"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Automatically quick-save the file when I've saved it manually"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   $"frmOpties.frx":08CA
         Height          =   1065
         Left            =   120
         TabIndex        =   16
         Top             =   2070
         Width           =   4830
      End
      Begin VB.Label Label2 
         Caption         =   "minutes"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Quick-save every"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   2835
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5055
      Begin VB.CheckBox Check6 
         Caption         =   "Show ""new document"" window when starting OptiType"
         Height          =   420
         Left            =   210
         TabIndex        =   17
         Top             =   1335
         Width           =   4410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Advanced settings"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show web page preview and webbrowser"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4125
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7276
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   "Algemeen"
            Object.Tag             =   "Algemeen"
            Object.ToolTipText     =   "Algemene opties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Saving"
            Key             =   "Opslaan"
            Object.Tag             =   "Opslaan"
            Object.ToolTipText     =   "Automatisch opslaan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View"
            Key             =   "Weergave"
            Object.Tag             =   "Weergave"
            Object.ToolTipText     =   "Weergave-opties"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOpties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mintCurFrame As Integer ' Current Frame visible



Private Sub Command1_Click()
ViewHTMLOpties.Show

End Sub

Private Sub Command2_Click()

optiestoepassen

Unload Me


End Sub



Private Sub Command4_Click()
setoptions

Unload Me

End Sub

Private Sub TabStrip1_Click()

If TabStrip1.SelectedItem.Index = mintCurFrame Then Exit Sub ' No need to change frame.
    ' Otherwise, hide old frame, show new.
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    Frame1(mintCurFrame).Visible = False
    ' Set mintCurFrame to new value.
    mintCurFrame = TabStrip1.SelectedItem.Index


End Sub


Private Sub optiestoepassen()

On Error Resume Next

If Check4.Value = 1 Then
    form1.lblDatum.Visible = True
Else
    form1.lblDatum.Visible = False
End If


If Check5.Value = 1 Then
    form1.StatusBar1.Visible = True
Else
    form1.StatusBar1.Visible = False
End If


If Check1.Value = 1 Then

    form1.RichTextBox1.Height = Fix(form1.RichTextBox1.Height / 2)
    form1.WebBrowser1.Visible = True
    form1.WebBrowser1.Height = form1.RichTextBox1.Height
    form1.htmlview.Checked = True
 
    frmOpties.Check1.Value = 1

If ViewHTMLOpties.Combo1.Text = "25%" Then
'Combo1.AddItem "25%"
'Combo1.AddItem "50%"
'Combo1.AddItem "75%"

form1.RichTextBox1.Height = ((form1.Height - 2640) * (3 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (1 / 4))


ElseIf ViewHTMLOpties.Combo1.Text = "50%" Then
    form1.RichTextBox1.Height = (form1.Height - 2640) / 2
    form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
    form1.WebBrowser1.Height = (form1.Height - 2640) / 2
ElseIf ViewHTMLOpties.Combo1.Text = "75%" Then
    form1.RichTextBox1.Height = ((form1.Height - 2640) * (1 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (3 / 4))

End If

   
    
Else   ' --------------------------------------------------------------
    form1.RichTextBox1.Height = form1.Height - 2640
    'form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False
    form1.htmlview.Checked = False

    
End If

Close
Open App.Path & "\OptiType.ini" For Output As #1
Print #1, Check4.Value
Print #1, Check5.Value
Print #1, Check2.Value
Print #1, Check3.Value
Print #1, Text1.Text
Print #1, Check6.Value
Close




End Sub
