VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Adresboek 
   BackColor       =   &H00800000&
   Caption         =   "Adress book"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3855
   Icon            =   "frmAdresBoek.frx":0000
   LinkTopic       =   "Adresboek"
   MaxButton       =   0   'False
   Picture         =   "frmAdresBoek.frx":0ECA
   ScaleHeight     =   4725
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   3840
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   3840
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1931
      _Version        =   393217
      TextRTF         =   $"frmAdresBoek.frx":33C2
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send e-mail to this person"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3525
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3045
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1125
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cell phone:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   225
      TabIndex        =   3
      Top             =   2085
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1605
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Adress:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "New contact..."
      End
      Begin VB.Menu delete 
         Caption         =   "Delete contact"
      End
      Begin VB.Menu de 
         Caption         =   "-"
      End
      Begin VB.Menu printit 
         Caption         =   "Print"
      End
      Begin VB.Menu cool 
         Caption         =   "-"
      End
      Begin VB.Menu endit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Adresboek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Combo1.ListIndex = Combo1.ListIndex
Combo2.ListIndex = Combo1.ListIndex
Combo3.ListIndex = Combo1.ListIndex
Combo4.ListIndex = Combo1.ListIndex
Combo5.ListIndex = Combo1.ListIndex
Combo6.ListIndex = Combo1.ListIndex
Combo7.ListIndex = Combo1.ListIndex
Combo8.ListIndex = Combo1.ListIndex


End Sub

Private Sub Combo2_Click()
Combo1.ListIndex = Combo2.ListIndex
Combo2.ListIndex = Combo2.ListIndex
Combo3.ListIndex = Combo2.ListIndex
Combo4.ListIndex = Combo2.ListIndex
Combo5.ListIndex = Combo2.ListIndex
Combo6.ListIndex = Combo2.ListIndex
Combo7.ListIndex = Combo2.ListIndex
Combo8.ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
Combo1.ListIndex = Combo3.ListIndex
Combo2.ListIndex = Combo3.ListIndex
Combo3.ListIndex = Combo3.ListIndex
Combo4.ListIndex = Combo3.ListIndex
Combo5.ListIndex = Combo3.ListIndex
Combo6.ListIndex = Combo3.ListIndex
Combo7.ListIndex = Combo3.ListIndex
Combo8.ListIndex = Combo3.ListIndex
End Sub

Private Sub Combo4_Click()
Combo1.ListIndex = Combo4.ListIndex
Combo2.ListIndex = Combo4.ListIndex
Combo3.ListIndex = Combo4.ListIndex
Combo4.ListIndex = Combo4.ListIndex
Combo5.ListIndex = Combo4.ListIndex
Combo6.ListIndex = Combo4.ListIndex
Combo7.ListIndex = Combo4.ListIndex
Combo8.ListIndex = Combo4.ListIndex
End Sub

Private Sub Combo5_Click()
Combo1.ListIndex = Combo5.ListIndex
Combo2.ListIndex = Combo5.ListIndex
Combo3.ListIndex = Combo5.ListIndex
Combo4.ListIndex = Combo5.ListIndex
Combo5.ListIndex = Combo5.ListIndex
Combo6.ListIndex = Combo5.ListIndex
Combo7.ListIndex = Combo5.ListIndex
Combo8.ListIndex = Combo5.ListIndex
End Sub

Private Sub Combo6_Click()
Combo1.ListIndex = Combo6.ListIndex
Combo2.ListIndex = Combo6.ListIndex
Combo3.ListIndex = Combo6.ListIndex
Combo4.ListIndex = Combo6.ListIndex
Combo5.ListIndex = Combo6.ListIndex
Combo6.ListIndex = Combo6.ListIndex
Combo7.ListIndex = Combo6.ListIndex
Combo8.ListIndex = Combo6.ListIndex
End Sub

Private Sub Combo7_Click()
Combo1.ListIndex = Combo7.ListIndex
Combo2.ListIndex = Combo7.ListIndex
Combo3.ListIndex = Combo7.ListIndex
Combo4.ListIndex = Combo7.ListIndex
Combo5.ListIndex = Combo7.ListIndex
Combo6.ListIndex = Combo7.ListIndex
Combo7.ListIndex = Combo7.ListIndex
Combo8.ListIndex = Combo7.ListIndex
End Sub

Private Sub Combo8_Click()
Combo1.ListIndex = Combo8.ListIndex
Combo2.ListIndex = Combo8.ListIndex
Combo3.ListIndex = Combo8.ListIndex
Combo4.ListIndex = Combo8.ListIndex
Combo5.ListIndex = Combo8.ListIndex
Combo6.ListIndex = Combo8.ListIndex
Combo7.ListIndex = Combo8.ListIndex
Combo8.ListIndex = Combo8.ListIndex
End Sub

Private Sub Command1_Click()
frmEmail.Text1.Text = Combo7.Text
Adresboek.Hide



End Sub

Private Sub Command2_Click()
Dim opdracht As String
opdracht = App.Path & "\OptiNet " & Combo8.Text
Shell opdracht, 1

End Sub

Private Sub Command3_Click()

frmAddContact.Show


End Sub

Private Sub Command4_Click()


a = MsgBox("Weet u zeker dat u " & Combo1.Text & " uit uw adresboek wilt verwijderen?", vbYesNo, "Opti2002")

If a = 6 Then

Close

Open App.Path & "\list.o2k" For Input As #1
Open App.Path & "\newlist.o2k" For Output As #2

Do Until EOF(1)
Line Input #1, regel

If regel <> Combo1.Text Then Print #2, regel
If regel = Combo1.Text Then Exit Do
Loop

If EOF(1) Then MsgBox "OptiType cannot remove this contact.  Your adress book could be damaged.  You should try uninstalling OptiType and then installing it again.  If the problem isn't solved, e-mail opti2002_support@pandora.be.", vbCritical, "Fout": Close: Exit Sub

Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy


Do Until EOF(1)

Line Input #1, regel
Print #2, regel
Loop

Close

Kill App.Path & "\list.o2k"
Name App.Path & "\newlist.o2k" As App.Path & "\list.o2k"



getbook

MsgBox "OptiType has deleted the contact.", vbInformation, "Opti2002"



End If

End Sub

Private Sub delete_Click()


a = MsgBox("Are you sure you want to remove " & Combo1.Text & " from your adress book?", vbYesNo, "OptiType")

If a = 6 Then

Close

Open App.Path & "\list.o2k" For Input As #1
Open App.Path & "\newlist.o2k" For Output As #2

Do Until EOF(1)
Line Input #1, regel

If regel <> Combo1.Text Then Print #2, regel
If regel = Combo1.Text Then Exit Do
Loop

If EOF(1) Then MsgBox "OptiType cannot remove this contact.  Your adress book could be damaged.  You should try uninstalling OptiType and then installing it again.  If the problem isn't solved, e-mail opti2002_support@pandora.be.", vbCritical, "Fout": Close: Exit Sub

Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy
Line Input #1, dummy


Do Until EOF(1)

Line Input #1, regel
Print #2, regel
Loop

Close

Kill App.Path & "\list.o2k"
Name App.Path & "\newlist.o2k" As App.Path & "\list.o2k"





MsgBox "OptiType has deleted the contact.", vbInformation, "Opti2002"

getbook

End If

End Sub

Private Sub endit_Click()
Adresboek.Hide
frmAddContact.Hide

End Sub

Private Sub Form_Load()
getbook



End Sub

Private Sub new_Click()
frmAddContact.Text1.Text = ""
frmAddContact.Text2.Text = ""
frmAddContact.Text3.Text = ""
frmAddContact.Text4.Text = ""
frmAddContact.Text5.Text = ""
frmAddContact.Text6.Text = ""
frmAddContact.Text7.Text = ""
frmAddContact.Text8.Text = ""


frmAddContact.Show
End Sub

Private Sub printit_Click()

Dim counter As Integer
counter = 0

RichTextBox1.Text = ""

Close

Open App.Path & "\list.o2k" For Input As #1

Do Until EOF(1)
Line Input #1, regel
counter = counter + 1
If counter = 4 Then RichTextBox1.Text = RichTextBox1.Text & "Tel: "
If counter = 5 Then RichTextBox1.Text = RichTextBox1.Text & "GSM: "
If counter = 6 Then RichTextBox1.Text = RichTextBox1.Text & "Fax: "
If counter = 7 Then RichTextBox1.Text = RichTextBox1.Text & "E-mail: "
If counter = 8 Then RichTextBox1.Text = RichTextBox1.Text & "Homepage: "
RichTextBox1.Text = RichTextBox1.Text & regel & vbCrLf
If counter = 8 Then
    RichTextBox1.Text = RichTextBox1.Text & vbCrLf
    counter = 0
End If
   
Loop

Close



PrintDialog.CancelError = True
On Error GoTo cantprintit

PrintDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If RichTextBox1.SelLength = 0 Then
        PrintDialog.Flags = PrintDialog.Flags + cdlPDAllPages
    Else
        PrintDialog.Flags = PrintDialog.Flags + cdlPDSelection
    End If
    PrintDialog.ShowPrinter
    'Printer.Print ""
    RichTextBox1.SelPrint PrintDialog.hdc

cantprintit:

End Sub
