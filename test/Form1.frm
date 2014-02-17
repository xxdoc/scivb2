VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#42.0#0"; "dSCIVB.ocx"
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
 
 
Private Type Highlighter
  StyleBold(127) As Long
  StyleItalic(127) As Long
  StyleUnderline(127) As Long
  StyleVisible(127) As Long
  StyleEOLFilled(127) As Long
  StyleFore(127) As Long
  StyleBack(127) As Long
  StyleSize(127) As Long
  StyleFont(127) As String
  StyleName(127) As String
  Keywords(7) As String
  strFilter As String
  strComment As String
  strName As String
  iLang As Long
  strFile As String
End Type

Private Highlighter As Highlighter
                                     
Public Function LoadHighlighter(strFile As String)
  Dim fFile As Integer
  fFile = FreeFile
  Open strFile For Binary Access Read As #fFile
  Get #fFile, , Highlighter
  Close #fFile
  FreeFile fFile
End Function
 
Public Function SaveHighlighter(strFile As String)
  On Error Resume Next
  Dim fFile As Integer
  fFile = FreeFile
  Kill strFile
  Open strFile For Binary Access Write As #fFile
  Put #fFile, , Highlighter
  Close #fFile
  FreeFile fFile
End Function
 
Private Sub Command1_Click()
    SciSimple1.MarkAll "test"
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
        .WordWrap = noWrap
        .IndentationGuide = True
        .Folding = True
        .Text = Replace("this is a simple test\nIf it were for real it would not be a test\nBut you already knew that i know\nif(a){\n  alert(2)\n}\n\nadd a dot after this: fso", "\n", vbCrLf)
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
    Debug.Print "AutoCompleteEvent: " & className
    
    
    If className = "fso" Then
        SciSimple1.ShowAutoComplete "readfile writefile appendfile fileexists deletefile"
    ElseIf className = "ida" Then
        'do i want to break these up into smaller chunks for intellisense?
        SciSimple1.ShowAutoComplete "imagebase() loadedfile() jump patchbyte originalbyte readbyte inttohex refresh() " & _
                               "numfuncs() functionstart functionend functionname getasm instsize xrefsto " & _
                               "xrefsfrom undefine getname jumprva screenea() funccount() find " & _
                               "hideea showea hideblock showblock removename setname makecode " & _
                               "getcomment addcomment addcodexref adddataxref delcodexref deldataxref " & _
                               "funcindexfromva funcvabyname nextea prevea patchstring makestr makeunk " & _
                               "renamefunc decompile"
    ElseIf className = "list" Then
        SciSimple1.ShowAutoComplete "additem clear"
    ElseIf className = "app" Then
        SciSimple1.ShowAutoComplete "getclipboard setclipboard askvalue openfiledialog savefiledialog exec list benchmark enableIDADebugMessages"
    End If
    
    
End Sub

Private Sub SciSimple1_OnError(Number As String, Description As String)
    MsgBox "SciSimple Error: " & Description
End Sub
