VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quick Crack"
   ClientHeight    =   4425
   ClientLeft      =   1095
   ClientTop       =   1755
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   5490
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   4230
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Max             =   255
   End
   Begin VB.TextBox txtNote 
      Height          =   3675
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   540
      Width           =   6675
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1,17491e-38
   End
   Begin VB.Label Label1 
      Caption         =   "Did NOT Find the code? Click Here!"
      Height          =   375
      Left            =   4140
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblFile 
      Caption         =   "Type your note and then save it to disk."
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3825
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open Encrypted File..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Crack"
         Begin VB.Menu mnuCrackItem 
            Caption         =   "&Crack Encrypted File..."
         End
         Begin VB.Menu mnuResumeCrackItem 
            Caption         =   "&Resume cracking"
         End
      End
      Begin VB.Menu mnuItemSave 
         Caption         =   "&Save Encrypted File..."
      End
      Begin VB.Menu mnuItemDate 
         Caption         =   "Insert &Date"
      End
      Begin VB.Menu mnuItemExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Crck As Boolean
Dim LVC, TTmp As Long
Dim Suspect
Dim MultilineTest As Boolean

Private Sub Command1_Click()
Decode LVC, CommonDialog1.FileName
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtNote.Height = Me.Height - lblFile.Height - 900
txtNote.Width = Me.Width - 140
ProgressBar1.Top = Me.Height - ProgressBar1.Height - lblFile.Height - 160
ProgressBar1.Width = Me.Width - 140
End Sub

Private Sub mnuCrackItem_Click()
MsgBox "This Feature gives you less than 40 possible codes, (if you know only 1 letter) and displays the codes to you. You get the one you want. If you don't know one of the letters contain this will also find it but the number of preveiw messages will be 256"
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT|Inifiles|*.ini"
    CommonDialog1.ShowOpen      'display Open dialog box
Suspect = InputBox("Do you suspect any letter? 1 caracter (if it is a text document note numbers concider the ""."" ")
Tmp = MsgBox("Does this file contain more than 1 line?" & Wrap$ & "If you don't know click no", vbYesNo)
If Tmp = 6 Then MultilineTest = True
    If CommonDialog1.FileName <> "" Then
    Crck = True
    LVC = 0
    Command1.Visible = False
    Label1.Visible = False
      lblFile.Caption = CommonDialog1.FileName 'set caption
Decode 1, CommonDialog1.FileName
End If
End Sub
Sub Decode(ByVal Bp As Integer, Flenme As String)
 Dim Ct, Ml As Boolean
 Dim ApproxSize As Integer
 Dim Possiblities, Comment As String
 Dim BoxesNotShown, Perc, BoxesShown, Tmp, EnterCound As Integer
 Dim G%, K%, code, decrypt$, e$
  Wrap$ = Chr$(13) + Chr$(10) 'create wrap character
      Form1.MousePointer = 11 'display hourglass
      For K% = Bp To Bp + 255
        LVC = K%
        'get encryption code to convert coded file to text
        'code = InputBox("Enter encryption code", , 1)
        code = K%
        Open Flenme For Input As #1 'open file
        On Error Resume Next
        decrypt$ = ""   'initialize string for decryption
        Do Until EOF(1)         'until end of file reached
            Input #1, Number&   'read encrypted numbers
            ApproxSize = ApproxSize + 1
            e$ = Chr$(Number& Xor code) 'convert with Xor
            decrypt$ = decrypt$ & e$    'and build string
        If (UCase(e$) = UCase(Suspect)) Then Ct = True
        Loop
        txtNote.Text = decrypt$ 'display converted string
        'txtNote.Enabled = True  'and enable scroll bars
CleanUp:                        'when finished...
       Close #1                'close file
ProgressBar1.Value = K% - Bp
If MultilineTest Then
 For G% = 1 To Len(txtNote.Text) - 1
  If Mid(txtNote.Text, G%, 2) = Wrap$ Then
   EnterCound = EnterCound + 1
  End If
 Next G%
If EnterCound > 1 Then Ml = True
End If
EnterCound = 0
  If Ct Or (Suspect = "") Then
   If ApproxSize = Len(txtNote.Text) Then
    If (MultilineTest) Then
     If Ml Then
      MsgBox K%
      BoxesShown = BoxesShown + 1
      Possiblities = Possiblities & K% & ","
     End If
    Else ' Do same but not check ml
      MsgBox K%
      BoxesShown = BoxesShown + 1
      Possiblities = Possiblities & K% & ","
    End If
    Ct = False
   End If
  End If
 ApproxSize = 0
 Ml = False
 Next K%
Form1.MousePointer = 0  'reset mouse
BoxesNotShown = 255 - BoxesShown
Perc = Int((BoxesNotShown * 100) / 255) + 1 '+ 1 the 1 suarly shown(correct answer)
Command1.Visible = True
Label1.Visible = True
If Suspect <> "" Then
 If Perc > 94 Then
  Comment = " Perfect!"
 ElseIf Perc > 79 Then
  Comment = " Great Work!"
 ElseIf Perc > 69 Then
  Comment = " Ok!"
 ElseIf Perc > 50 Then
  Comment = " Sorry i was n't any help!"
 Else: Comment = " You could be mor specific!"
 End If
Else
If MultilineTest = False Then Comment = " Try guessig a number or letter that is maybe in the text for faster and more accurate search also notice that it could be mre than one line, if you give this information the search will be 50% faster!"
If MultilineTest Then Comment = " Try guessig a number or letter that is maybe in the text for faster and more accurate search"
End If
If Perc <> 101 Then MsgBox "Matching codes: " & BoxesShown & " Unmatching codes: " & BoxesNotShown & " Saved trouble: " & Perc & "%" & Comment & Wrap$ & Possiblities
TTmp = LVC + 255
Command1.Caption = LVC & "-" & TTmp
    Exit Sub
    Resume CleanUp:   'finally, finish with CleanUp routine
Exit Sub
End Sub

Private Sub mnuItemDate_Click()
    Wrap$ = Chr$(13) & Chr$(10) 'add date to string
    txtNote.Text = Date$ & Wrap$ & txtNote.Text
End Sub

Private Sub mnuItemExit_Click()
    End                         'quit program
End Sub

Private Sub mnuItemSave_Click()
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT"
    CommonDialog1.ShowSave      'display Save dialog
    If CommonDialog1.FileName <> "" Then
        'get encryption code and use it to encrypt file
        code = InputBox("Enter Encryption Code", , 1)
        If code = "" Then Exit Sub  'if Cancel chosen, exit sub
        Form1.MousePointer = 11     'display hourglass
        charsInFile% = Len(txtNote.Text) 'find string length
        Open CommonDialog1.FileName For Output As #1 'open file
        For i% = 1 To charsInFile%  'for each character in file
            letter$ = Mid(txtNote.Text, i%, 1) 'read next char
            'convert to number w/ Asc, then use Xor to encrypt
            Print #1, Asc(letter$) Xor code; 'and save in file
        Next i%
        Close #1                'close file when finished
        CommonDialog1.FileName = ""  'clear filename
        Form1.MousePointer = 0  'reset mouse
    End If
End Sub

Private Sub mnuOpenItem_Click()
    Wrap$ = Chr$(13) + Chr$(10) 'create wrap character
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT|Inifiles|*.ini"
    CommonDialog1.ShowOpen      'display Open dialog box
    If CommonDialog1.FileName <> "" Then
        'get encryption code to convert coded file to text
        code = InputBox("Enter encryption code", , 1)
        If code = "" Then Exit Sub 'if Cancel chosen, exit sub
        Form1.MousePointer = 11 'display hourglass
        Open CommonDialog1.FileName For Input As #1 'open file
        On Error GoTo problem:  'set error handler
        decrypt$ = ""   'initialize string for decryption
        Do Until EOF(1)         'until end of file reached
            Input #1, Number&   'read encrypted numbers
            e$ = Chr$(Number& Xor code) 'convert with Xor
            decrypt$ = decrypt$ & e$    'and build string
        Loop
        lblFile.Caption = CommonDialog1.FileName 'set caption
        txtNote.Text = decrypt$ 'display converted string
        txtNote.Enabled = True  'and enable scroll bars
CleanUp:                        'when finished...
        Form1.MousePointer = 0  'reset mouse
        Close #1                'close file
        'CommonDialog1.FileName = ""  'clear filename
    End If
  Exit Sub
problem:  'if there is a problem, display appropriate message
    If Err.Number = 5 Then  'Chr$ problem means bad key
        Resume Next
        'MsgBox ("Incorrect Encryption Key")
    Else  'for other problems (like file too big) show error
        MsgBox "Error Opening File", , Err.Description
    End If
    Resume CleanUp:   'finally, finish with CleanUp routine
End Sub

Private Sub mnuResumeCrackItem_Click()
Dim BegnPoint As Integer
MsgBox "This will resume any crack serch that you left in the middle, you must rember the point you left it thow eg. 256"
    CommonDialog1.Filter = "Text files (*.TXT)|*.TXT|Inifiles|*.ini"
    CommonDialog1.ShowOpen      'display Open dialog box
Suspect = InputBox("Do you suspect any letter? 1 caracter (if it is a text document note numbers concider the ""."" ")
Tmp = MsgBox("Does this file contain more than 1 line?" & Wrap$ & "If you don't know click no", vbYesNo)
If Tmp = 6 Then MultilineTest = True
    If CommonDialog1.FileName <> "" Then
    Crck = True
    LVC = 0
    Command1.Visible = False
    Label1.Visible = False
      lblFile.Caption = CommonDialog1.FileName 'set caption
BegnPoint = InputBox("Reume point? 256 - 100000", "Resume search")
Decode BegnPoint, CommonDialog1.FileName
End If
End Sub
