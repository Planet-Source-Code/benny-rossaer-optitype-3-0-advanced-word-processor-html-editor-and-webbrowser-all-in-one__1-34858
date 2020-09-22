Attribute VB_Name = "Module1"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public verversing As Integer
Public bron As Integer
Public ebron As Integer
Public shownew As Integer

Public lStart As Long

Public Const MAX_PATH = 260

Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long

Public Function Win32Keyword(ByVal URL As String) As Long
weburl = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function


Public Sub centerform()

    form1.Top = (Screen.Height - form1.Height) / 2
    form1.Left = (Screen.Width - form1.Width) / 2
End Sub


Public Sub nieuwformaat()


On Error Resume Next
centerform
On Error GoTo dontresize

' ----------------
If form1.WebBrowser1.Visible = False Then
weergaveinstellen

    form1.Label2.Height = form1.Height
   
    form1.RichTextBox1.Left = 2400
   form1.RichTextBox1.Width = form1.Width - form1.webbuttons.Width - 2600
   form1.WebBrowser2.Width = form1.Width - form1.webbuttons.Width - 2600
    'form1.Label2.Width = form1.Width - form1.RichTextBox1.Width - 420
    form1.Label2.Width = form1.Width - form1.RichTextBox1.Width - form1.webbuttons.Width - 220

    form1.comboInvoegen.Width = form1.Label2.Width - 100
    form1.List1.Width = form1.Label2.Width - 100


    form1.RichTextBox1.Top = form1.Toolbar2.Top + form1.Toolbar2.Height
    form1.RichTextBox1.Height = form1.Height - 2640
    form1.WebBrowser2.Height = form1.Height - 2640
    form1.Label1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
    form1.txtURL.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
    form1.Command2.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80


    form1.lblDatum.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
    form1.DocName.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80

    form1.WebBrowser1.Left = 2400
    form1.WebBrowser1.Width = form1.RichTextBox1.Width


    form1.WebBrowser1.Top = 850 + (form1.RichTextBox1.Height / 2)


    form1.WebBrowser1.Height = form1.RichTextBox1.Height / 2

Else

    If ViewHTMLOpties.Combo1.ListIndex = 1 Then  ' 50%; normale weergave

        weergaveinstellen

ElseIf ViewHTMLOpties.Combo1.ListIndex = 0 Then  ' 25%

weergaveinstellen
form1.RichTextBox1.Height = ((form1.Height - 2640) * (3 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (1 / 4))
    
ElseIf ViewHTMLOpties.Combo1.ListIndex = 2 Then  ' 75%

weergaveinstellen
form1.RichTextBox1.Height = ((form1.Height - 2640) * (1 / 4))
form1.WebBrowser1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height
form1.WebBrowser1.Height = ((form1.Height - 2640) * (3 / 4))




End If
End If
dontresize:


End Sub

Public Sub weergaveinstellen()

        form1.RichTextBox1.Left = 2400
        form1.RichTextBox1.Width = form1.Width - 2820
        form1.Label2.Width = form1.Width - form1.RichTextBox1.Width - 420
        form1.RichTextBox1.Top = form1.Toolbar2.Top + form1.Toolbar2.Height
        form1.RichTextBox1.Height = (form1.Height - 2640) / 2
        form1.WebBrowser1.Left = 2400
        form1.WebBrowser1.Width = form1.RichTextBox1.Width
        form1.WebBrowser1.Top = 850 + (form1.RichTextBox1.Height)
        form1.WebBrowser1.Height = form1.RichTextBox1.Height
        form1.Label1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + form1.WebBrowser1.Height + 80
        form1.txtURL.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + form1.WebBrowser1.Height + 80
        form1.Command2.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + form1.WebBrowser1.Height + 80
        form1.lblDatum.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + form1.WebBrowser1.Height + 80
        form1.DocName.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + form1.WebBrowser1.Height + 80
        form1.Label2.Height = form1.Height
        form1.comboInvoegen.Width = form1.Label2.Width - 100
        form1.List1.Width = form1.Label2.Width - 100


  '   form1.Label1.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
  '  form1.txtURL.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
  '  form1.Command2.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80


   ' form1.lblDatum.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
   ' form1.DocName.Top = form1.RichTextBox1.Top + form1.RichTextBox1.Height + 80
    
    'form1.WebBrowser1.Top = 850 + (form1.RichTextBox1.Height / 2)
    'form1.WebBrowser1.Height = form1.RichTextBox1.Height / 2

End Sub

Public Sub showHTMLPreview()


form1.RichTextBox1.SaveFile "c:\temp.htm", 1

Win32Keyword ("c:\temp.htm")


End Sub


Public Sub OpslaanAls()


    On Error GoTo ErrHandler
    
   ' Set CancelError is True
    form1.CommonDialog2.CancelError = True
    form1.CommonDialog2.FileName = ""
    Rem On Error GoTo ErrHandler
    ' Set flags
    form1.CommonDialog2.Flags = cdlOFNHideReadOnly
    ' Set filters
    form1.CommonDialog2.Filter = "OptiType Rich Text (*.rtf)|*.rtf|Text files" & _
    "(*.txt)|*.txt|Batch-file (*.bat)|*.bat|Webpage (*.htm)|*.htm|All files|*.*"
    ' Specify default filter
    If form1.batchpreview.Enabled = False Then form1.CommonDialog2.FilterIndex = 1
    If form1.batchpreview.Enabled = True Then form1.CommonDialog2.FilterIndex = 3
    ' Display the Open dialog box
    form1.CommonDialog2.ShowSave
    ' Display name of selected file

   form1.Caption = "OptiType 3.0 - " & form1.CommonDialog2.FileName
   
   
    If UCase$(Right$(form1.CommonDialog2.FileName, 3)) = "RTF" Then
    form1.RichTextBox1.SaveFile form1.CommonDialog2.FileName, 0
    Else
    form1.RichTextBox1.SaveFile form1.CommonDialog2.FileName, 1
    End If
    
    form1.DocName.Text = form1.CommonDialog2.FileName
    
    On Error Resume Next: form1.RichTextBox1.SetFocus
    
    form1.AutoSave.Enabled = True
        
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub



End Sub


Public Function GetWinPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetWinPath = ""
End If

End Function


Public Sub setoptions()
Close
Open App.Path & "\OptiType.ini" For Input As #1
Input #1, datumweergeven
Input #1, statusbalkweergeven
Input #1, automatischopslaan
Input #1, automatischopslaanbevestiging
Input #1, autoopslaan
Input #1, shownew
Close


If shownew = 1 Then frmOpties.Check6.Value = 1
If shownew = 0 Then frmOpties.Check6.Value = 0
If datumweergeven = 1 Then frmOpties.Check4.Value = 1  ' datum
If datumweergeven = 0 Then lblDatum.Visible = False
If statusbalkweergeven = 1 Then frmOpties.Check5.Value = 1  ' statusbalk
If statusbalkweergeven = 0 Then StatusBar1.Visible = False
If automatischopslaan = 1 Then frmOpties.Check2.Value = 1  ' automatisch opslaan
If automatischopslaanbevestiging = 1 Then frmOpties.Check3.Value = 1  ' bevestiging
frmOpties.Text1.Text = autoopslaan ' auto save om de zoveel minuten



End Sub


Sub htmltags()


form1.List1.Clear
form1.List1.AddItem "HTML-page header"
form1.List1.AddItem " "
form1.List1.AddItem "Line"
form1.List1.AddItem "Paragraph"
form1.List1.AddItem " "
form1.List1.AddItem "Image"
form1.List1.AddItem "Hyperlink"
form1.List1.AddItem "Horiz.rule"
form1.List1.AddItem "Applet"
form1.List1.AddItem " "
form1.List1.AddItem "Quote"
form1.List1.AddItem "Numbered list"
form1.List1.AddItem "List"
form1.List1.AddItem "Comment"
form1.List1.AddItem " "
form1.List1.AddItem "Bold"
form1.List1.AddItem "Italic"
form1.List1.AddItem "Underline"
form1.List1.AddItem "Subscript"
form1.List1.AddItem "Superscript"
form1.List1.AddItem "Center"
form1.List1.AddItem " "
form1.List1.AddItem "Font size 1"
form1.List1.AddItem "Font size 2"
form1.List1.AddItem "Font size 3"
form1.List1.AddItem "Font size 4"
form1.List1.AddItem "Font size 5"
form1.List1.AddItem "Font size 6"
form1.List1.AddItem "Font size 7"
form1.List1.AddItem " "
form1.List1.AddItem "Color: black"
form1.List1.AddItem "Color: white"
form1.List1.AddItem "Color: red"
form1.List1.AddItem "Color: yellow"
form1.List1.AddItem "Color: blue"
form1.List1.AddItem "Color: grey"
form1.List1.AddItem "Color: silver"
form1.List1.AddItem "Color: green"
form1.List1.AddItem " "
form1.List1.AddItem "Table"



End Sub

Sub ansi()
form1.List1.Clear

form1.List1.AddItem "33     !"
form1.List1.AddItem "34     """""
form1.List1.AddItem "35     #"
form1.List1.AddItem "36     $"
form1.List1.AddItem "37     %"
form1.List1.AddItem "38     &"
form1.List1.AddItem "39     '"
form1.List1.AddItem "40     ("
form1.List1.AddItem "41     )"
form1.List1.AddItem "42     *"
form1.List1.AddItem "43     +"
form1.List1.AddItem "44     ,"
form1.List1.AddItem "45     -"
form1.List1.AddItem "46     ."
form1.List1.AddItem "47     /"
form1.List1.AddItem " "
form1.List1.AddItem "58     :"
form1.List1.AddItem "59     ;"
form1.List1.AddItem "60     <"
form1.List1.AddItem "61     ="
form1.List1.AddItem "62     >"
form1.List1.AddItem "63     ?"
form1.List1.AddItem "64     @"
form1.List1.AddItem " "
form1.List1.AddItem "91     ["
form1.List1.AddItem "92     \"
form1.List1.AddItem "93     ]"
form1.List1.AddItem "94     ^"
form1.List1.AddItem "95     _"
form1.List1.AddItem "96     `"
form1.List1.AddItem " "
form1.List1.AddItem "123    {"
form1.List1.AddItem "124    |"
form1.List1.AddItem "125    }"
form1.List1.AddItem "126    ~"

form1.List1.AddItem "127    ‘"
form1.List1.AddItem "128    Ç"
form1.List1.AddItem "129    ü"
form1.List1.AddItem "130    é"
form1.List1.AddItem "131    â"
form1.List1.AddItem "132    ä"
form1.List1.AddItem "133    à"
form1.List1.AddItem "134    å"
form1.List1.AddItem "135    ç"
form1.List1.AddItem "136    ê"
form1.List1.AddItem "137    ë"
form1.List1.AddItem "138    è"
form1.List1.AddItem "139    ï"
form1.List1.AddItem "140    î"
form1.List1.AddItem "141    ì"
form1.List1.AddItem "142    Ä"
form1.List1.AddItem "143    Å"
form1.List1.AddItem "144    É"

form1.List1.AddItem "145    æ"
form1.List1.AddItem "146    Æ"
form1.List1.AddItem "147    ô"
form1.List1.AddItem "148    ö"
form1.List1.AddItem "149    ò"
form1.List1.AddItem "150    û"
form1.List1.AddItem "151    ù"
form1.List1.AddItem "152    ÿ"
form1.List1.AddItem "153    Ö"
form1.List1.AddItem "154    Ü"
form1.List1.AddItem "155    ø"
form1.List1.AddItem "156    £"
form1.List1.AddItem "157    Ø"
form1.List1.AddItem "158    ×"
form1.List1.AddItem "159    ƒ"
form1.List1.AddItem "160    á"
form1.List1.AddItem "161    í"
form1.List1.AddItem "162    ó"
form1.List1.AddItem "163    ú"
form1.List1.AddItem "164    ñ"
form1.List1.AddItem "165    Ñ"
form1.List1.AddItem "166    ª"
form1.List1.AddItem "167    º"
form1.List1.AddItem "168    ¿"
form1.List1.AddItem "169    ©"
form1.List1.AddItem "170    ¬"
form1.List1.AddItem "171    ½"
form1.List1.AddItem "172    ¼"
form1.List1.AddItem "173    ¡"
form1.List1.AddItem "174    «"
form1.List1.AddItem "175    »"
form1.List1.AddItem "176    °"
form1.List1.AddItem "177    ±"
form1.List1.AddItem "178    ²"
form1.List1.AddItem "179    ³"
form1.List1.AddItem "180    ´"
form1.List1.AddItem "181    µ"
form1.List1.AddItem "182    ¶"
form1.List1.AddItem "183    ·"
form1.List1.AddItem "184    ¸"
form1.List1.AddItem "185    ¹"
form1.List1.AddItem "186    º"
form1.List1.AddItem "187    »"
form1.List1.AddItem "188    ¼"
form1.List1.AddItem "189    ½"
form1.List1.AddItem "190    ¾"
form1.List1.AddItem "191    ¿"
form1.List1.AddItem "192    À"
form1.List1.AddItem "193    Á"
form1.List1.AddItem "194    Â"
form1.List1.AddItem "195    Ã"
form1.List1.AddItem "196    Ä"
form1.List1.AddItem "197    Å"
form1.List1.AddItem "198    Æ"
form1.List1.AddItem "199    Ç"
form1.List1.AddItem "200    È"
form1.List1.AddItem "201    É"
form1.List1.AddItem "202    Ê"
form1.List1.AddItem "203    Ë"
form1.List1.AddItem "204    Ì"
form1.List1.AddItem "205    Í"
form1.List1.AddItem "206    Î"
form1.List1.AddItem "207    Ï"
form1.List1.AddItem "208    Ð"
form1.List1.AddItem "209    Ñ"
form1.List1.AddItem "210    Ò"
form1.List1.AddItem "211    Ó"
form1.List1.AddItem "212    Õ"
form1.List1.AddItem "213    i"
form1.List1.AddItem "214    Õ"
form1.List1.AddItem "215    ×"
form1.List1.AddItem "216    Ø"
form1.List1.AddItem "217    Ù"
form1.List1.AddItem "218    +"
form1.List1.AddItem "219    Û"
form1.List1.AddItem "220    Ü"
form1.List1.AddItem "221    ¦"
form1.List1.AddItem "222    Ì"
form1.List1.AddItem "223    ß"
form1.List1.AddItem "224    à"
form1.List1.AddItem "225    á"
form1.List1.AddItem "226    â"
form1.List1.AddItem "227    ã"
form1.List1.AddItem "228    ä"
form1.List1.AddItem "229    å"
form1.List1.AddItem "230    µ"
form1.List1.AddItem "231    þ"
form1.List1.AddItem "232    Þ"
form1.List1.AddItem "233    Ú"
form1.List1.AddItem "234    Û"
form1.List1.AddItem "235    Ù"
form1.List1.AddItem "236    ý"
form1.List1.AddItem "237    Ý"
form1.List1.AddItem "238    î"
form1.List1.AddItem "239    ï"
form1.List1.AddItem "240    ð"
form1.List1.AddItem "241    ñ"
form1.List1.AddItem "242    ò"
form1.List1.AddItem "243    ó"
form1.List1.AddItem "244    ô"
form1.List1.AddItem "245    õ"
form1.List1.AddItem "246    ö"
form1.List1.AddItem "247    ÷"
form1.List1.AddItem "248    ø"
form1.List1.AddItem "249    ù"
form1.List1.AddItem "250    ú"
form1.List1.AddItem "251    û"
form1.List1.AddItem "252    ü"
form1.List1.AddItem "253    ý"
form1.List1.AddItem "254    þ"
form1.List1.AddItem "255    ÿ"



End Sub

Sub dosinvoegen()
form1.List1.Clear

form1.List1.AddItem "ATTRIB"
form1.List1.AddItem "BREAK"
form1.List1.AddItem "CALL"
form1.List1.AddItem "CD"
form1.List1.AddItem "CHDIR"
form1.List1.AddItem "CHKDSK"
form1.List1.AddItem "CLS"
form1.List1.AddItem "COLOR"
form1.List1.AddItem "COPY"
form1.List1.AddItem "DATE"
form1.List1.AddItem "DEL"
form1.List1.AddItem "DIR"
form1.List1.AddItem "DISKCOMP"
form1.List1.AddItem "DISKCOPY"
form1.List1.AddItem "DOSKEY"
form1.List1.AddItem "ECHO"
form1.List1.AddItem "ERASE"
form1.List1.AddItem "EXIT"
form1.List1.AddItem "FIND"
form1.List1.AddItem "FOR"
form1.List1.AddItem "FORMAT"
form1.List1.AddItem "GOTO"
form1.List1.AddItem "IF"
form1.List1.AddItem "LABEL"
form1.List1.AddItem "MD"
form1.List1.AddItem "MKDIR"
form1.List1.AddItem "MODE"
form1.List1.AddItem "MOVE"
form1.List1.AddItem "PATH"
form1.List1.AddItem "PAUSE"
form1.List1.AddItem "PROMPT"
form1.List1.AddItem "RD"
form1.List1.AddItem "Rem"
form1.List1.AddItem "REN"
form1.List1.AddItem "RENAME"
form1.List1.AddItem "REPLACE"
form1.List1.AddItem "RMDIR"
form1.List1.AddItem "SET"
form1.List1.AddItem "START"
form1.List1.AddItem "SUBST"
form1.List1.AddItem "TIME"
form1.List1.AddItem "TREE"
form1.List1.AddItem "TYPE"
form1.List1.AddItem "VER"
form1.List1.AddItem "VOL"
form1.List1.AddItem "XCOPY"

End Sub

Sub autotekstinvoegen()


form1.comboInvoegen.ListIndex = 3

form1.List1.Clear


form1.List1.AddItem "Dear sir,"
form1.List1.AddItem "Dear parents,"
form1.List1.AddItem "Ladies and gentlemen ,"
form1.List1.AddItem " "
form1.List1.AddItem "With kind regards,"
form1.List1.AddItem "Thank you in advance,"
form1.List1.AddItem "Thank you very much,"
form1.List1.AddItem "See you soon,"
form1.List1.AddItem "Love,"
form1.List1.AddItem "Regards,"



End Sub

Public Sub leegdocumentmaken()

hidewebbrowser2
form1.comboInvoegen.ListIndex = 2
ansi



form1.Caption = "OptiType 3.0 - New text file"

form1.RichTextBox1.Left = 2400
'form1.RichTextBox1.Width = form1.Width - 2820 - form1.webbuttons.Width
form1.RichTextBox1.Width = form1.Width - form1.webbuttons.Width - 2600


form1.RichTextBox1.Top = form1.Toolbar2.Top + form1.Toolbar2.Height
form1.RichTextBox1.Height = form1.Height - 2640

    
    'form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False
    form1.htmlview.Checked = False
  
    


On Error Resume Next
form1.batchpreview.Enabled = False
form1.batchfile.Enabled = False
form1.htmlview.Checked = False

form1.quicksave.Enabled = True
form1.mnuOpslaan.Enabled = True



  '  form1.RichTextBox1.Height = 5415
    form1.WebBrowser1.Visible = False
    frmOpties.Check1.Value = 0
    
On Error Resume Next
Unload frmNieuw
form1.Show
form1.RichTextBox1.SetFocus
End Sub


Public Sub nieuwhtmlmaken()
hidewebbrowser2
form1.comboInvoegen.ListIndex = 0
htmltags

form1.Caption = "OptiType 3.0 - New website"

On Error Resume Next



Close
form1.batchfile.Enabled = False
form1.batchpreview.Enabled = False
form1.quicksave.Enabled = False
form1.mnuOpslaan.Enabled = False


form1.htmlview.Checked = True
form1.RichTextBox1.Height = (form1.Height - 2640) / 2   ' NEW CODE! BUG FIXED!!!

   form1.WebBrowser1.Visible = True
    
    frmOpties.Check1.Value = 1
form1.Caption = "OptiType - HTML Modus"


Open "c:\temp.htm" For Output As #1
Print #1, ""
Close



frmWebsite1.Show vbModal




On Error Resume Next
Unload frmNieuw
form1.Show
On Error Resume Next: form1.RichTextBox1.SetFocus

End Sub


Public Sub hidewebbrowser2()
If form1.WebBrowser2.Visible = True Then

    form1.pagina.Visible = False
    form1.mnuBestand.Visible = True
    form1.mnuBewerken.Visible = True
    form1.beeld.Visible = True
    form1.mnuOpmaak.Visible = True
    form1.Invoegen.Visible = True
    form1.letsdothesel.Visible = True
    form1.Extra.Visible = True
    
    
    form1.WebBrowser2.Visible = False
    form1.webbuttons.Buttons.Item(9).Image = 6
    
       
    form1.RichTextBox1.Visible = True
    form1.WebBrowser2.Visible = False
    form1.webbuttons.Buttons.Item(9).Image = 6
    leegdocumentmaken
   
    
End If

End Sub


Public Sub showwebbrowser2()
On Error Resume Next

    form1.pagina.Visible = True
    form1.mnuBestand.Visible = False
    form1.mnuBewerken.Visible = False
    form1.beeld.Visible = False
    form1.mnuOpmaak.Visible = False
    form1.Invoegen.Visible = False
    form1.letsdothesel.Visible = False
    form1.Extra.Visible = False


            ' naar openstaande locatie gaan
            If form1.WebBrowser1.Visible = True Then              'And form1.txtURL.Text <> "Typ hier een internetadres in..." And form1.txtURL.Text <> "file:///C:/temp.htm" Then
                form1.WebBrowser2.Navigate form1.txtURL.Text
            End If
    
            leegdocumentmaken
            form1.WebBrowser1.Visible = False
            form1.RichTextBox1.Visible = False
            ' --------
            form1.WebBrowser2.Left = 2400
            form1.WebBrowser2.Width = form1.Width - form1.webbuttons.Width - 2600
            form1.WebBrowser2.Top = form1.Toolbar2.Top + form1.Toolbar2.Height
            form1.WebBrowser2.Height = form1.Height - 2640
            
            form1.WebBrowser2.Visible = True
            
            
                form1.webbuttons.Buttons.Item(9).Image = 9
              '  webbuttons.Refresh
                
End Sub


Public Function getsourcecode(URL As String) As String


    getsourcecode = form1.Inet1.OpenURL(URL)
End Function


Public Sub focusherstellen()


form1.WebBrowser1.Navigate "c:\temp.htm"

form1.showpopups.Enabled = False
form1.webbuttons.Buttons.Item(11).Value = tbrUnpressed
form1.browseoffline.Enabled = False
form1.back.Enabled = False
form1.gotonext.Enabled = False
form1.refreshbrowser.Enabled = False
form1.webbuttons.Buttons.Item(11).Value = tbrUnpressed
form1.searchforapage.Enabled = False



End Sub

Public Sub nieuwhtmlmakenzonderwizard()
hidewebbrowser2
form1.comboInvoegen.ListIndex = 0
htmltags

form1.Caption = "OptiType 3.0 - New website"


On Error Resume Next


Close
form1.batchfile.Enabled = False
form1.batchpreview.Enabled = False
form1.quicksave.Enabled = False
form1.mnuOpslaan.Enabled = False


form1.htmlview.Checked = True
form1.RichTextBox1.Height = (form1.Height - 2640) / 2   ' NEW CODE! BUG FIXED!!!

    form1.WebBrowser1.Visible = True
    
    frmOpties.Check1.Value = 1
form1.Caption = "OptiType - HTML Mode"


Open "c:\temp.htm" For Output As #1
Print #1, ""
Close


On Error Resume Next
Unload frmNieuw
form1.Show
On Error Resume Next: form1.RichTextBox1.SetFocus

End Sub
