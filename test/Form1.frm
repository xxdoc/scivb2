VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#71.0#0"; "scivb_lite.ocx"
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
 'WH_KEYBOARD_LL
  
Private Sub Command1_Click()
     Dim x As Long
     With SciSimple1
        'x = .FirstVisibleLine
        '.ReplaceAll "already", ""
        '.Text = Replace(.Text, "already", "")
        '.FirstVisibleLine = x
        
        .ShowFindReplace
     End With
     
     'MsgBox x
End Sub

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
        .WordWrap = noWrap
        .ShowIndentationGuide = True
        .Folding = True
        .Text = Replace("this is a simple test\nIf it were for real it would not be a test\nBut you already knew that i know\nif(a){\n  alert(2)\n}\n\nadd a dot after this: fsonow type controlh after fso.app\n\n\n\n\ng\ng\ng\ng\njkfdljsfkl\nkdjskldjsl\ndjklfjkds\nhfjdfhjd\nfjkdljfk\njfkdlsjfkl\n", "\n", vbCrLf)
        .Text = .Text & .Text & .Text & .Text & .Text & .Text
        .LoadCallTips App.Path & "\..\js_api.txt"
        .AddCallTip "appendfile(blah,blah)"
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
