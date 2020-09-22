VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNieuw 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OptiType 3.0 PSC - New file"
   ClientHeight    =   4725
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmNieuw.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNieuw.frx":08CA
   ScaleHeight     =   4725
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   315
      Top             =   3795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Always show this window when starting OptiType"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1755
      TabIndex        =   5
      Top             =   4410
      Value           =   1  'Checked
      Width           =   6660
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   5460
      TabIndex        =   4
      ToolTipText     =   "Dit venster sluiten."
      Top             =   3945
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open from disk"
      Height          =   375
      Left            =   5130
      TabIndex        =   3
      ToolTipText     =   "Een tekst- of RTF-document van schijf openen."
      Top             =   3435
      Width           =   1680
   End
   Begin VB.CommandButton cmdWebsite 
      Height          =   855
      Left            =   4185
      Picture         =   "frmNieuw.frx":65A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Een nieuwe website maken gebaseerd op een enkele webpagina of op meerdere pagina's met de Wizard Frames."
      Top             =   1785
      Width           =   1335
   End
   Begin VB.CommandButton cmdLeegDocument 
      Height          =   855
      Left            =   2160
      Picture         =   "frmNieuw.frx":6C6E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Een nieuw tekstdocument of een nieuwe brief met de Wizard Brief maken."
      Top             =   1785
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 3.0 PSC "
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5445
      TabIndex        =   6
      Top             =   90
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Klik op het item van uw keuze, of klik op 'Van schijf' om een eerder opgeslagen bestand te openen."
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "frmNieuw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrief_Click()


End Sub

Private Sub Check1_Click()

Close

Close
Open App.Path & "\OptiType.ini" For Input As #1
Input #1, datumweergeven
Input #1, statusbalkweergeven
Input #1, automatischopslaan
Input #1, automatischopslaanbevestiging
Input #1, autoopslaan
Input #1, shownew
Close


Open App.Path & "\optitype.ini" For Output As #1
Print #1, datumweergeven
Print #1, statusbalkweergeven
Print #1, automatischopslaan
Print #1, automatischopslaanbevestiging
Print #1, autoopslaan
Print #1, Check1.Value
Close


End Sub

Private Sub cmdLeegDocument_Click()


    form1.Enabled = False
              PopupMenu Form2.empty
              form1.Enabled = True
        
    On Error Resume Next: form1.RichTextBox1.SetFocus
        
Exit Sub


End Sub

Private Sub cmdWebsite_Click()


  
      form1.Enabled = False
              PopupMenu Form2.site
              
         form1.Enabled = True
        
On Error Resume Next: form1.RichTextBox1.SetFocus
Exit Sub
Exit Sub

End Sub

Private Sub Command4_Click()

   ' Set CancelError is True
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog2.Filter = "OptiType Rich Text (*.rtf)|*.rtf|Text files " & _
    "(*.txt)|*.txt|Webpage (*.htm, *.html)|*.ht*|Batch-file (*.bat)|*.bat|All files|*.*"
    ' Specify default filter
    CommonDialog2.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog2.ShowOpen
    ' Display name of selected file

    
   form1.Caption = "OptiType 3.0 - " & CommonDialog2.FileName
   
   
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "RTF" Then
    RichTextBox1.LoadFile CommonDialog2.FileName, 0
    RichTextBox1.SelBullet = False
    Else
    RichTextBox1.LoadFile CommonDialog2.FileName, 1
    RichTextBox1.SelBullet = False
    End If
    
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "BAT" Then
        batchpreview.Enabled = True
        form1.batchfile.Enabled = True
    Else
        batchpreview.Enabled = False
        form1.batchfile.Enabled = False
    End If
        
    If UCase$(Right$(CommonDialog2.FileName, 3)) = "HTM" Or UCase$(Right$(CommonDialog2.FileName, 3)) = "TML" Then
    nieuwhtmlmakenzonderwizard
    End If
    
    
        
    DocName.Text = CommonDialog2.FileName
    
    form1.AutoSave.Enabled = True
    Sluiten
            
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Sluiten
    Exit Sub
End Sub

Private Sub Command5_Click()
frmNieuw.Hide

End Sub

Private Sub Sluiten()
frmNieuw.Hide
End Sub


Private Sub Form_LostFocus()
On Error Resume Next: form1.RichTextBox1.SetFocus


    



End Sub

