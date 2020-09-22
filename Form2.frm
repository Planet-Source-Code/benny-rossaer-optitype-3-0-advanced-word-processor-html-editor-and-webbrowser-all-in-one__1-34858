VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu empty 
      Caption         =   "New text document"
      Begin VB.Menu doc 
         Caption         =   "&New document"
      End
      Begin VB.Menu letter 
         Caption         =   "&Letter wizard"
      End
      Begin VB.Menu batchfile 
         Caption         =   "&Batch file"
      End
   End
   Begin VB.Menu site 
      Caption         =   "New website"
      Begin VB.Menu newhtml 
         Caption         =   "&New HTML page"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub batchfile_Click()
hidewebbrowser2
form1.Caption = "OptiType 3.0 - Nieuw batchbestand"

form1.comboInvoegen.ListIndex = 1
dosinvoegen


On Error Resume Next
form1.batchfile.Enabled = True
form1.batchpreview.Enabled = True
form1.htmlview.Checked = False

form1.quicksave.Enabled = True
form1.mnuOpslaan.Enabled = True


    form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False

form1.RichTextBox1.Text = "@ECHO OFF" & vbCrLf


Unload frmNieuw
form1.Show
On Error Resume Next: form1.RichTextBox1.SetFocus

nieuwformaat


End Sub

Private Sub doc_Click()

leegdocumentmaken

End Sub

Private Sub frameswiz_Click()
hidewebbrowser2
form1.Caption = "OptiType 3.0 - New website with multiple frames"


form1.comboInvoegen.ListIndex = 0
htmltags





form1.htmlview.Checked = False
form1.batchfile.Enabled = False
form1.quicksave.Enabled = True
form1.mnuOpslaan.Enabled = True
  form1.RichTextBox1.Height = 5415

    form1.WebBrowser1.Visible = False
    form1.batchpreview.Enabled = False
    
form1.WindowState = vbMinimized

Dim opdracht As String
opdracht = App.Path & "\Wizard.exe"
Shell opdracht, 1


Unload frmNieuw
End Sub

Private Sub letter_Click()

hidewebbrowser2


form1.batchpreview.Enabled = False
form1.batchfile.Enabled = False


form1.quicksave.Enabled = True
form1.mnuOpslaan.Enabled = True


form1.htmlview.Checked = False

    
    form1.WebBrowser1.Visible = False

frmBrief1.Show vbModal

Unload frmNieuw
End Sub

Private Sub newhtml_Click()

nieuwhtmlmaken

End Sub
