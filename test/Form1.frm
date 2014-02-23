VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#75.0#0"; "scivb_lite.ocx"
Begin VB.Form d 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   1710
      TabIndex        =   2
      Top             =   135
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1230
   End
   Begin SCIVB_LITE.SciSimple SciSimple1 
      Height          =   3840
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6773
   End
End
Attribute VB_Name = "d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Command1_Click()
     Dim x
     Dim results() As Long
     Dim ret As String
     Dim y
     With SciSimple1
        'x = .FirstVisibleLine
        '.ReplaceAll "already", ""
        '.Text = Replace(.Text, "already", "")
        '.FirstVisibleLine = x
     
      'tested does not reset current line..
      MsgBox .ReplaceAll("test", "blah!")
      
'      If .FindAll("test", results) = -1 Then
'            MsgBox "not found!"
'      Else
'            For Each x In results
'                ret = ret & x & ", "
'            Next
'            MsgBox UBound(results) & " hits start offsets:" & ret
'      End If
       
                 
     End With
     
     'MsgBox x
End Sub

Private Sub Command2_Click()
    SciSimple1.hilightClear
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Sub Form_Load()

'    Dim pth As String
'    pth = "C:\Documents and Settings\david\Desktop\scivb\highlighters\java.bin"
'    LoadHighlighter pth
'
'    For i = 0 To 127
'        Highlighter.StyleFont(i) = "Courier New"
'        Highlighter.StyleSize(i) = 12
'    Next
'
'    SaveHighlighter pth

    With SciSimple1
        .codePage = SC_CP_UTF8
        .WordWrap = True
        .ShowIndentationGuide = True
        .Folding = True
        .Text = Replace("this func_01 v1 is a simple testaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\nIf it were for real it would not be a test\nBut you already knew that i know\nif(a){\n  alert(2)\n}\n\nadd a dot after this: fsonow type controlh after fso.app\n\n\n\n\ng\ng\ng\ng\njkfdljsfkl\nkdjskldjsl\ndjklfjkds\nhfjdfhjd\nfjkdljfk\njfkdlsjfkl\n", "\n", vbCrLf)
        .Text = .Text & .Text & .Text & .Text & .Text & .Text
        .LoadCallTips App.Path & "\..\js_api.txt"
        .AddCallTip "appendfile(blah,blah)"
        .AddCallTip "func_01(test)"
        .ShowIndentationGuide = True
        '.AutoCompleteOnCTRLSpace = False
        '.EnableArrowKeys
        '.SetFocusSci
         
    End With
    
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    SciSimple1.Width = Me.Width - SciSimple1.Left - 240
    SciSimple1.Height = Me.Height - SciSimple1.Top - 600
End Sub

Private Sub SciSimple1_AutoCompleteEvent(className As String)
    
    Dim prevWord As String
    prevWord = SciSimple1.PreviousWord()
    
    Debug.Print "AutoCompleteEvent: ClassName: " & className & " Prevword: " & prevWord
        
    'scintinella is smart enough to autoscroll the autocomplete list to the partial match of the curWord :)
    'so fso.app CTRL+H will send us curword=app prevword=fso and sci will scroll to appendfile at top of list.
    
    If className = "tb" Or prevWord = "tb" Then
        SciSimple1.ShowAutoComplete "Save2Clipboard GetClipboard t eval unescape alert Hexdump WriteFile ReadFile HexString2Bytes Disasm pad EscapeHexString GetStream CRC getPageNumWords GetPageNthWord"
    End If
    
    If className = "fso" Or prevWord = "fso" Then
        SciSimple1.ShowAutoComplete "readfile writefile appendfile fileexists deletefile"
    ElseIf className = "ida" Or prevWord = "ida" Then
        'do i want to break these up into smaller chunks for intellisense?
        SciSimple1.ShowAutoComplete "imagebase() loadedfile() jump patchbyte originalbyte readbyte inttohex refresh() " & _
                               "numfuncs() functionstart functionend functionname getasm instsize xrefsto " & _
                               "xrefsfrom undefine getname jumprva screenea() funccount() find " & _
                               "hideea showea hideblock showblock removename setname makecode " & _
                               "getcomment addcomment addcodexref adddataxref delcodexref deldataxref " & _
                               "funcindexfromva funcvabyname nextea prevea patchstring makestr makeunk " & _
                               "renamefunc decompile"
                               
    ElseIf className = "list" Or prevWord = "list" Then
        SciSimple1.ShowAutoComplete "additem clear"
    ElseIf className = "app" Or prevWord = "app" Then
        SciSimple1.ShowAutoComplete "getclipboard setclipboard askvalue openfiledialog savefiledialog exec list benchmark enableIDADebugMessages"
    End If
        
    
End Sub

Private Sub SciSimple1_OnError(Number As String, Description As String)
    MsgBox "SciSimple Error: " & Description
End Sub

Private Sub SciSimple1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    SciSimple1.hilightClear
    Me.Caption = SciSimple1.hilightWord(SciSimple1.CurrentWord) & " word matches"
End Sub
