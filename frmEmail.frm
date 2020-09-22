VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEmail 
   BackColor       =   &H00C0C0C0&
   Caption         =   "New e-mail message"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog printdialog 
      Left            =   3960
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog invoegen 
      Left            =   4800
      Top             =   4305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2475
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4905
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog bijlage 
      Left            =   5565
      Top             =   4275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Attach a file..."
      Height          =   510
      Left            =   1320
      TabIndex        =   4
      Top             =   4920
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3735
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmEmail.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Top             =   2040
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   1530
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1305
      TabIndex        =   2
      Top             =   600
      Width           =   4950
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   120
      Width           =   4950
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Benny Rossaer, 2000 - 2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4035
      TabIndex        =   10
      Top             =   5325
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   135
      TabIndex        =   8
      Top             =   1065
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail of receptient:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   1470
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Height          =   5685
      Left            =   -30
      TabIndex        =   11
      Top             =   -15
      Width           =   1830
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu dezdezdezdez 
         Caption         =   "-"
      End
      Begin VB.Menu printit 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu itsoutthere 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu fdez 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select all"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Insert"
      Begin VB.Menu txt 
         Caption         =   "&Text file"
      End
      Begin VB.Menu insertattach 
         Caption         =   "&Attachment..."
      End
      Begin VB.Menu tze 
         Caption         =   "-"
      End
      Begin VB.Menu invoegendatum 
         Caption         =   "&Date"
      End
      Begin VB.Menu invoegentijd 
         Caption         =   "&Time"
      End
   End
   Begin VB.Menu extr 
      Caption         =   "&Extra"
      Begin VB.Menu countwordsordie 
         Caption         =   "&Count words"
      End
   End
   Begin VB.Menu h 
      Caption         =   "&Help"
      Begin VB.Menu info 
         Caption         =   "&Info"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aantal As Integer
'aantal = 0



Private Sub Command1_Click()
On Error Resume Next: form1.MAPISession1.SignOn

If form1.MAPISession1.SessionID <> 0 Then

    With form1.MAPIMessages1
        .SessionID = form1.MAPISession1.SessionID
        .Compose
        
        For i = 0 To Combo1.ListCount - 1
            Combo1.ListIndex = i
            .AttachmentIndex = i
            .AttachmentPathName = Combo1.Text
            
          
    Next
    
            
        
        .RecipDisplayName = Text1.Text
        .RecipAddress = Text1.Text
        .MsgSubject = Text2.Text
        .MsgNoteText = RichTextBox1.Text
        .Send False
        
       ' MsgBox "Uw document is verstuurd naar " & Text1.Text & ".", vbOKOnly, "Document verzenden via e-mail"
        
        
        
    End With


    form1.MAPISession1.SignOff
End If

Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next: form1.MAPISession1.SignOff
Unload Me

End Sub

Private Sub Command3_Click()

aantal = aantal + 1

bijlage.DialogTitle = "Selecteer bijlage..."
bijlage.CancelError = False
bijlage.ShowOpen

Combo1.AddItem bijlage.FileName
Combo1.ListIndex = 0


End Sub

Private Sub Command4_Click()
Adresboek.Show
End Sub

Private Sub copy_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText

End Sub

Private Sub countwordsordie_Click()
MsgBox "Number of words in this e-mail: " & CountWords(1, RichTextBox1, False)
End Sub

Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End Sub

Private Sub delete_Click()
RichTextBox1.SelText = ""
End Sub

Private Sub Form_Load()

RichTextBox1.RightMargin = RichTextBox1.Width


Combo1.Clear
aantal = 0


End Sub

Private Sub info_Click()
MsgBox "OptiType v3.0 E-mail support" & vbCrLf & vbCrLf & "Benny Rossaer " & vbCrLf & vbCrLf, vbInformation, "Info"


End Sub

Private Sub insertattach_Click()
aantal = aantal + 1

bijlage.DialogTitle = "Select attachment..."
bijlage.CancelError = False
bijlage.ShowOpen

Combo1.AddItem bijlage.FileName
Combo1.ListIndex = 0
End Sub

Private Sub invoegendatum_Click()
RichTextBox1.SelText = RichTextBox1.SelText & Date$
End Sub

Private Sub invoegentijd_Click()
RichTextBox1.SelText = RichTextBox1.SelText & Time$
End Sub

Private Sub new_Click()
a = MsgBox("Bent u zeker dat u het huidige bericht wilt wissen en een nieuw bericht wilt starten?", vbYesNo, "Nieuw bericht")


If a = 6 Then
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""
End If
End Sub

Private Sub open_Click()
invoegen.CancelError = True
On Error GoTo nope
invoegen.DialogTitle = "Select file to open..."
invoegen.Filter = "Text file (*.txt)|*.txt|OptiType RTF|*.rtf|All files|*.*"
invoegen.FilterIndex = 1
invoegen.ShowOpen

RichTextBox1.FileName = invoegen.FileName


Exit Sub
nope:
Exit Sub

End Sub

Private Sub paste_Click()
If Clipboard.GetFormat(vbCFText) Then
selectie = Clipboard.GetText
RichTextBox1.SelText = RichTextBox1.SelText & selectie
End If


End Sub

Private Sub printit_Click()
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

Private Sub quit_Click()
On Error Resume Next: form1.MAPISession1.SignOff
frmEmail.Hide
End Sub

Private Sub save_Click()
invoegen.CancelError = True
On Error GoTo nope
invoegen.DialogTitle = "Select file to save to..."
invoegen.Filter = "Text file (*.txt)|*.txt|OptiType RTF|*.rtf"
invoegen.FilterIndex = 1
invoegen.ShowOpen

If invoegen.FilterIndex = 1 Then RichTextBox1.SaveFile invoegen.FileName, rtfText
If invoegen.FilterIndex = 2 Then RichTextBox1.SaveFile invoegen.FileName, rtfRTF



Exit Sub
nope:
Exit Sub
End Sub

Private Sub selectall_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub

Private Sub txt_Click()

invoegen.CancelError = True
On Error GoTo nope
invoegen.DialogTitle = "Select text file to insert333"
invoegen.Filter = "Tekstbestand (*.txt)|*.txt|All files|*.*"
invoegen.FilterIndex = 1
invoegen.ShowOpen

Dim regel As String

    Open invoegen.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, regel
    RichTextBox1.SelText = RichTextBox1.SelText & vbCrLf & regel
    Loop
    Close



Exit Sub
nope:
Exit Sub

End Sub


Public Function CountWords(countmodus As Integer, rtBox As RichTextBox, BuildList As Boolean, Optional TargetList, Optional SizeLimit = 0) As Long

    'Keith Gardner (kgard@mhonline.net) - November 1999
    'Required Globals(Just cut and past the following into General De
    '     clarations.)
    'Public WordList() As String
    'Public WordCount As Long
    'Public UnqWordCount As Long
    'Counters
    Dim X As Long, y As Long 'Loop counters
    'Flags
    Dim AddNow As Boolean 'Flag to indicate If it is time to add word
    Dim FoundIt As Boolean 'Flag to indicate If word was found in list
    'Temporary storage
    Dim ThisText As String 'Holds full text of the rich text control
    Dim ThisWord As String 'Holds current word as it is built
    Dim ThisChar As Integer 'Holds ASCII value of current character
    Dim PrevChar As Integer 'Holds ASCII value of previous character
    
    
    If countmodus = 1 Then ThisText = Trim(rtBox.Text)
    If countmodus = 2 Then ThisText = Trim(rtBox.SelText)
    


    If ThisText = "" Then
        WordCount = -1
    Else
        WordCount = 0
        UnqWordCount = 0


        If BuildList Then
            ReDim wordlist(2, 1)
        End If


        PrevChar = 0


        For X = 1 To Len(ThisText)
            ThisChar = Asc(Mid(ThisText, X, 1))


            Select Case ThisChar
                Case 13 'Line feed


                If PrevChar <> 10 Then
                    WordCount = WordCount + 1
                    AddNow = True
                End If


                Case 32 'Space
                WordCount = WordCount + 1
                AddNow = True
                Case 10, 33, 34, 39, 40, 41, 63 'Ignore LF, "!", """, "'", "(", ")", "?"
                Case 44, 46 'Ignore "," or "." unless it's in a number
                If PrevChar >= 48 And PrevChar <= 57 Then ThisWord = ThisWord & Chr(ThisChar)
                Case Else 'ThisChar not a delimiter
                ThisWord = ThisWord & Chr(ThisChar)


                If X = Len(ThisText) - 1 Then
                    AddNow = True 'Add last word in list
                End If


            End Select


        'Building WordList?


        If BuildList And AddNow Then
            'Look for the word in the list


            For y = 1 To UnqWordCount


                If ThisWord = wordlist(1, y) Then 'Found it!
                    FoundIt = True
                    wordlist(2, y) = wordlist(2, y) + 1
                End If


                If FoundIt Then Exit For
            Next




            If Not FoundIt Then
                UnqWordCount = UnqWordCount + 1
                ReDim Preserve wordlist(2, UnqWordCount)
                wordlist(1, y) = ThisWord
                wordlist(2, y) = 1
            End If


            FoundIt = False 'Reset flag
            ThisWord = ""
        End If


        AddNow = False
        PrevChar = ThisChar
    Next


    WordCount = WordCount + 1 'Add one to the word count


    If Not IsMissing(TargetList) Then


        With TargetList


            For X = 1 To UnqWordCount


                If Len(wordlist(1, X)) > SizeLimit Then
                    .AddItem wordlist(1, X) & " - (" & wordlist(2, X) & ")"
                End If


            Next


        End With


    End If


End If


CountWords = WordCount
End Function

