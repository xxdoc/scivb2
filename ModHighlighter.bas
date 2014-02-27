Attribute VB_Name = "ModHighlighter"
Option Explicit

'+---------------------------------------------------------------------------+
'| modHighlighter.bas                                                        |
'+---------------------------------------------------------------------------+
'| This is a basic module to provide very basic highlighter loading support. |
'| In reality I wouldn't really recomend using this as a basis for your      |
'| editor but it should give you some idea's.  The biggest reason I did not  |
'| want to bundle the code to read highlighter files into the class itself   |
'| is for performance reasons.  With this setup you can load the files one   |
'| time, and then just set each editor.  For the demo application this is a  |
'| fairly useless feature but if your dealing with a MDI application it's    |
'| going to make a world of difference.  If it was bundled directly into the |
'| class quite litterly every document you create would load every single    |
'| file.  That would be very poor use of system resources :)                 |
'+---------------------------------------------------------------------------+
Public Declare Function GetTickCount Lib "kernel32" () As Long

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


Private Highlighters() As Highlighter 'can not be public

'these are for the export to html function..they all need a redo..
Private sBuffer As String
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Property Get HighLightersCount() As Long
    If hAryIsEmpty(Highlighters) Then
        HighLightersCount = -1
    Else
        HighLightersCount = UBound(Highlighters)
    End If
End Property

Public Function CompileVersionInfo(owner As scisimple) As String
    On Error Resume Next
    Dim dllVer As String
    Dim dllPath As String
    Dim ret() As String
    Dim hIndex As Long
    Dim hlNames As String
    Dim i As Long
    
    push ret, "scivb_lite: " & App.Major & "." & App.Minor & "." & App.Revision & "  (" & FileSize(App.path & "\scivb_lite.ocx") & ")"
    
    dllPath = GetLoadedSciLexerPath()
    If FileExists(dllPath) Then
        dllVer = GetFileVersion(dllPath)
        If Len(dllVer) > 0 Then push ret, "SciLexer:   " & dllVer & "    (" & FileSize(dllPath) & ")"
        push ret, "SciVB Path: " & App.path
        push ret, "Lexer Path: " & dllPath
        
    Else
        push ret, "SciVB Path: " & App.path
        push ret, "SciLexer:   NOT FOUND!"
    End If
    
    push ret(), ""
    
    If hAryIsEmpty(Highlighters) Then
        push ret(), "Highlighters loaded: None"
    Else
        For i = 0 To UBound(Highlighters)
            hlNames = hlNames & Highlighters(i).strName & ", "
        Next
        hlNames = Trim(hlNames)
        If Len(hlNames) > 1 Then hlNames = VBA.Left(hlNames, Len(hlNames) - 1)
        push ret(), UBound(Highlighters) + 1 & " highlighter(s) loaded: " & hlNames
        hIndex = owner.currentHighlighter
        push ret(), "Active Highlighter: " & Highlighters(hIndex).strFile
    End If
    
    CompileVersionInfo = Join(ret, vbCrLf)
    
End Function

Private Sub ReInit()
    sBuffer = ""
    lOffset = 0
End Sub

Private Function GetString() As String
    GetString = Left$(sBuffer, lOffset)
    sBuffer = ""  'reset
    lOffset = 0
End Function

Private Sub SConcat(ByRef Source As String)
    Dim lBufferLen As Long
    lBufferLen = Len(Source)
    'Allocate more space in buffer if needed
    If (lOffset + lBufferLen) >= Len(sBuffer) Then
       If lBufferLen > lOffset Then
          sBuffer = sBuffer & String$(lBufferLen, 0)
       Else
          sBuffer = sBuffer & String$(ciIncriment, 0)
       End If
    End If
    Mid$(sBuffer, lOffset + 1, lBufferLen) = Source
    lOffset = lOffset + lBufferLen
End Sub
'------------------------------------------------------------------

Private Function FindHighlighter(strLangName As String) As Integer
  Dim i As Integer

  FindHighlighter = -1
  If hAryIsEmpty(Highlighters) Then Exit Function
  
  For i = 0 To UBound(Highlighters)
      If UCase(Highlighters(i).strName) = UCase(strLangName) Then
            FindHighlighter = i
            Exit Function
      End If
  Next i
    
End Function

Public Function SetHighlighter(owner As scisimple, strHighlighter As String) As Boolean
  Dim i As Long, X As Long
  
  On Error GoTo hell
  
  X = FindHighlighter(strHighlighter)
  If X = -1 Then Exit Function
  
  With owner
     .DirectSCI.ClearDocumentStyle
     
     If LCase(strHighlighter) = "html" Then
           .DirectSCI.StyleSetBits 7
     Else
           .DirectSCI.StyleSetBits 5
     End If
     
     .DirectSCI.SetLexer Highlighters(X).iLang
     For i = 0 To 7
           If Highlighters(X).Keywords(i) <> "" Then .DirectSCI.SetKeyWords i, Highlighters(X).Keywords(i)
     Next i
    
     .DirectSCI.StyleSetBack 32, Highlighters(X).StyleBack(32)
     .DirectSCI.StyleSetFore 32, Highlighters(X).StyleFore(32)
     .DirectSCI.StyleSetVisible 32, CLng(Highlighters(X).StyleVisible(32))
     .DirectSCI.StyleSetEOLFilled 32, CLng(Highlighters(X).StyleEOLFilled(32))
     .DirectSCI.StyleSetBold 32, CLng(Highlighters(X).StyleBold(32))
     .DirectSCI.StyleSetItalic 32, CLng(Highlighters(X).StyleItalic(32))
     .DirectSCI.StyleSetUnderline 32, CLng(Highlighters(X).StyleUnderline(32))
     .DirectSCI.StyleSetFont 32, Highlighters(X).StyleFont(32)
     .DirectSCI.StyleSetSize 32, Highlighters(X).StyleSize(32)
     .DirectSCI.StyleClearAll
     
     For i = 0 To 127
           .DirectSCI.StyleSetBold i, CLng(Highlighters(X).StyleBold(i))
           .DirectSCI.StyleSetItalic i, CLng(Highlighters(X).StyleItalic(i))
           .DirectSCI.StyleSetUnderline i, CLng(Highlighters(X).StyleUnderline(i))
           .DirectSCI.StyleSetVisible i, CLng(Highlighters(X).StyleVisible(i))
           If Highlighters(X).StyleFont(i) <> "" Then .DirectSCI.StyleSetFont i, Highlighters(X).StyleFont(i)
           .DirectSCI.StyleSetFore i, CLng(Highlighters(X).StyleFore(i))
           .DirectSCI.StyleSetBack i, CLng(Highlighters(X).StyleBack(i))
           .DirectSCI.StyleSetSize i, CLng(Highlighters(X).StyleSize(i))
           .DirectSCI.StyleSetEOLFilled i, CLng(Highlighters(X).StyleEOLFilled(i))
     Next i
     
     .DirectSCI.StyleSetFore 35, .misc.BraceBadFore
     .DirectSCI.StyleSetFore 34, .misc.BraceMatchFore
     .DirectSCI.StyleSetBack 35, .misc.BraceBadBack
     .DirectSCI.StyleSetBack 34, .misc.BraceMatchBack
     .DirectSCI.StyleSetBold 35, .misc.BraceMatchBold
     .DirectSCI.StyleSetBold 34, .misc.BraceMatchBold
     .DirectSCI.StyleSetItalic 35, .misc.BraceMatchItalic
     .DirectSCI.StyleSetItalic 34, .misc.BraceMatchItalic
     .DirectSCI.StyleSetUnderline 35, .misc.BraceMatchUnderline
     .DirectSCI.StyleSetUnderline 34, .misc.BraceMatchUnderline
     
     .DirectSCI.Colourise 0, -1
     .currentHighlighter = strHighlighter
  End With
  
  SetHighlighter = True
  Exit Function
hell:
End Function

Public Function LoadHighlighter(strFile As String) As Boolean
  Dim fFile As Integer
  Dim h As Highlighter
  
  If Not FileExists(strFile) Then Exit Function
  
  fFile = FreeFile
  
  Open strFile For Binary Access Read As #fFile
  Get #fFile, , h
  h.strFile = strFile
  Close #fFile
  
  If FindHighlighter(h.strName) <> -1 Then Exit Function 'dont add again if its already loaded..
  
  hpush Highlighters(), h
  LoadHighlighter = True
  
End Function

Public Sub LoadDirectory(strDir As String)
  
  Dim str As String, i As Long
  
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  str = Dir(strDir & "\*bin")
  
  Erase Highlighters
  
  Do Until str = ""
    LoadHighlighter strDir & "\" & str
    str = Dir
  Loop
  
End Sub

Public Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = " "
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function

Public Function HighlighterForExtension(file As String) As String
    
    Dim ext  As String, X As Long
    
    On Error GoTo hell
    If hAryIsEmpty(Highlighters) Then Exit Function
    
    ext = LCase$(Mid$(file, InStrRev(file, ".") + 1, Len(file) - InStrRev(file, ".")))
    ext = "." & ext
    
    For X = 0 To UBound(Highlighters)
        If InStr(1, Highlighters(X).strFilter, ext) Then
            HighlighterForExtension = Highlighters(X).strName
            Exit For
        End If
    Next X
    
hell:
    
End Function

Private Function hAryIsEmpty(ary() As Highlighter) As Boolean
  On Error GoTo oops
  Dim X As Long
    X = UBound(ary)
    hAryIsEmpty = False
  Exit Function
oops: hAryIsEmpty = True
End Function

Private Sub hpush(ary() As Highlighter, Value As Highlighter) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub



Public Function ExportToHTML2(strFile As String, scisimple As scisimple) As Boolean
  
  On Error Resume Next
  
  ' This function will output the source to HTML with the styling
  ' It is far from perfect and frankly it's slower than hell if you ask me
  ' It takes it a solid 7-8 seconds to output this file (modHighlighter.bas)
  ' So if anyone can think of ways to improve it's speed.  At least its
  ' better than what it initially was (about 19 seconds for this file)
  ' thanks to a simple concatation function and comparing the long value's
  ' of the characters in question instead of a string to string comparison.
  ' but otherwise still slow :)
  
  Dim iLen As Long
  Dim strOutput As String
  Dim strCSS As String
  Dim lPrevStyle As Long
  Dim lStyle As Long
  Dim Style(127) As Boolean
  Dim prevStyle As Long
  Dim curStyle As Long
  Dim nextStyle As Long
  Dim i As Long
  Dim strTotal As String
  Dim strStyle As String
  Dim currentHighlighter As Integer
  
  currentHighlighter = FindHighlighter(scisimple.currentHighlighter)
  If currentHighlighter = -1 Then Exit Function
  
  scisimple.DirectSCI.Colourise 0, -1
  
  For i = 0 To 127
        Style(i) = False
  Next i
  
  For i = 0 To Len(scisimple.Text)
        lStyle = scisimple.DirectSCI.GetStyleAt(i)
        Style(lStyle) = True
  Next
  
  strCSS = ""
  strTotal = "<HTML>" & vbCrLf & "  <HEAD>" & vbCrLf & "    <Meta Generator=" & """" & "scisimple Class (http://www.ceditmx.com)" & """" & ">" & vbCrLf
  strCSS = "<style type=" & """" & "text/css" & """" & ">" & vbCrLf
  
  For i = 0 To 127
    If Style(i) = True Then
        With Highlighters(currentHighlighter)
            strCSS = strCSS & ".c" & i & " {" & vbCrLf
            
            If .StyleFont(i) <> "" Then
                strCSS = strCSS & "font-family: " & "'" & .StyleFont(i) & "'" & ";" & vbCrLf
            End If
            
            If .StyleFore(i) <> 0 Then
                strCSS = strCSS & "color: " & DectoHex(.StyleFore(i)) & ";" & vbCrLf
            End If
            
            If .StyleBack(i) <> 0 Then
                strCSS = strCSS & "background: " & DectoHex(.StyleBack(i)) & ";" & vbCrLf
            End If
            
            If .StyleSize(i) <> 0 Then
                strCSS = strCSS & "font-size: " & .StyleSize(i) & "pt" & ";" & vbCrLf
            End If
            
            If .StyleBold(i) = 0 Then
                 strCSS = strCSS & "font-weight: 400;" & vbCrLf
            Else
                strCSS = strCSS & "font-weight: 700;" & vbCrLf
            End If
            
            strStyle = ""
            If .StyleItalic(i) <> 0 Then
                strStyle = "text-decoration: italic;"
            End If
            
            If .StyleUnderline(i) <> 0 Then
              If strStyle = "" Then
                    strStyle = "text-decoration: underline;"
              Else
                    strStyle = strStyle & ", underline;"
              End If
            End If
            
            If strStyle <> "" Then
                strCSS = strCSS & strStyle & vbCrLf
            End If
            
            strCSS = strCSS & "}" & vbCrLf
        End With
    End If
  Next i
  
  strCSS = strCSS & "</style>" & vbCrLf
  strTotal = strTotal & strCSS
  strTotal = strTotal & "  </HEAD>" & vbCrLf & "  <BODY BGCOLOR=#FFFFFF TEXT=#000000>"
  strOutput = ""
  sBuffer = ""
  iLen = scisimple.DirectSCI.GetLength
  
  For i = 0 To iLen
        curStyle = scisimple.DirectSCI.GetStyleAt(i)
        If (i + 1) < iLen Then
            nextStyle = scisimple.DirectSCI.GetStyleAt(i + 1)
        End If
        If curStyle <> prevStyle Then
        
                SConcat "<span class=c" & curStyle & ">"
                'strOutput = strOutput & "<span class=c" & curStyle & ">"
                If scisimple.DirectSCI.GetCharAt(i) <> 13 And scisimple.DirectSCI.GetCharAt(i) <> 10 And scisimple.DirectSCI.GetCharAt(i) <> 60 And scisimple.DirectSCI.GetCharAt(i) <> 62 Then
                    If scisimple.DirectSCI.GetCharAt(i) = 32 Then
                        'strOutput = strOutput & "&nbsp;"
                        SConcat "&nbsp;"
                    Else
                        SConcat Chr(scisimple.DirectSCI.GetCharAt(i))
                        'strOutput = strOutput & scisimple.GetCharAt(i)
                    End If
                Else
                    If scisimple.DirectSCI.GetCharAt(i) = 13 Then
                        If scisimple.DirectSCI.GetCharAt(i + 1) <> 10 Then
                          SConcat "<BR>"
                          SConcat vbCrLf
                        End If
                    ElseIf scisimple.DirectSCI.GetCharAt(i) = 10 Then
                        SConcat "<BR>"
                        SConcat vbCrLf
                    ElseIf scisimple.DirectSCI.GetCharAt(i) = 60 Then
                        SConcat "&LT;"
                    ElseIf scisimple.DirectSCI.GetCharAt(i) = 62 Then
                        SConcat "&GT;"
                    End If
                    'strOutput = strOutput & "<BR>"
                End If
                If i = iLen Or nextStyle <> curStyle Then
                    SConcat "</span>"
                    'strOutput = strOutput & "</span>"
                End If
              Else
              
                If scisimple.DirectSCI.GetCharAt(i) <> 13 And scisimple.DirectSCI.GetCharAt(i) <> 10 And scisimple.DirectSCI.GetCharAt(i) <> 60 And scisimple.DirectSCI.GetCharAt(i) <> 62 Then
                  If scisimple.DirectSCI.GetCharAt(i) = 32 Then
                        SConcat "&nbsp;"
                        'strOutput = strOutput & "&nbsp;"
                  Else
                        SConcat Chr(scisimple.DirectSCI.GetCharAt(i))
                        'strOutput = strOutput & scisimple.GetCharAt(i)
                  End If
                Else
                  If scisimple.DirectSCI.GetCharAt(i) = 13 Then
                        If scisimple.DirectSCI.GetCharAt(i + 1) <> 10 Then
                            SConcat "<BR>"
                            SConcat vbCrLf
                        End If
                  ElseIf scisimple.DirectSCI.GetCharAt(i) = 10 Then
                        SConcat "<BR>"
                        SConcat vbCrLf
                  ElseIf scisimple.DirectSCI.GetCharAt(i) = 60 Then
                        SConcat "&LT;"
                  ElseIf scisimple.DirectSCI.GetCharAt(i) = 62 Then
                        SConcat "&GT;"
                  End If
                  'strOutput = strOutput & "<BR>"
            End If
            
            If i = iLen Or nextStyle <> curStyle Then
                  SConcat "</span>"
                  'strOutput = strOutput & "</span>"
            End If
        End If
        prevStyle = curStyle
  Next i
  
  strOutput = GetString
  strTotal = strTotal & strOutput
  strTotal = strTotal & vbCrLf & "  </BODY>" & vbCrLf & "</HTML>"
  
  i = FreeFile
  Open strFile For Output As #i
  Print #i, strTotal
  Close #i
  strOutput = ""
  
  ExportToHTML2 = True
  
End Function

'Convert decimal colour to hex
Public Function DectoHex(lngColour As Long) As String
    Dim strColour As String
    
    strColour = Hex(lngColour)
    
    'Add leading zero's
    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop

    'Reverse the bgr string pairs to rgb
    DectoHex = "#" & Right$(strColour, 2) & _
                     Mid$(strColour, 3, 2) & _
                     Left$(strColour, 2)
                        
End Function

Public Sub CommentBlock2(SCI As scisimple)
  On Error GoTo errHandler
  Dim i As Long
  Dim lStart As Long, lEnd As Long
  Dim lLineStart As Long, lLineEnd As Long
  Dim strCmp As String, strTmp As String, strHold As String
  Dim lTmp As Long 'If the line sel is reversed
  Dim ua() As String
  Dim strSplit As String
  Dim X As Long
  Dim lAdd As Long
  Dim currentHighlighter As Integer
  
  lStart = SCI.SelStart
  lEnd = SCI.SelEnd
  lLineStart = SCI.DirectSCI.LineFromPosition(lStart)
  lLineEnd = SCI.DirectSCI.LineFromPosition(lEnd)
  strCmp = ""
  currentHighlighter = FindHighlighter(SCI.currentHighlighter)
  
  strCmp = SCI.SelText
  If InStr(1, strCmp, Chr(13)) > 1 Then
    If InStr(1, strCmp, Chr(10)) > 1 Then
      strCmp = Replace(strCmp, Chr(13), "")
      ua() = Split(strCmp, Chr(10))
      strSplit = vbCrLf
    Else
      ua() = Split(strCmp, Chr(13))
      strSplit = vbCr
    End If
  ElseIf InStr(1, strCmp, Chr(13)) = 0 Then
    If InStr(1, strCmp, Chr(10)) > 1 Then
      ua() = Split(strCmp, Chr(10))
      strSplit = vbLf
    Else
      ReDim ua(0)
      ua(0) = strCmp
    End If

  End If
  strCmp = ""
  For i = 0 To UBound(ua)
    strCmp = strCmp & Highlighters(currentHighlighter).strComment & ua(i)
    If i < UBound(ua) Then strCmp = strCmp & strSplit
  Next i
  If UBound(ua) > 0 Then
    lAdd = ((UBound(ua) + 1) * Len(Highlighters(currentHighlighter).strComment)) ' + (Len(strSplit) * (UBound(ua) - 1))
  Else
    lAdd = Len(Highlighters(currentHighlighter).strComment)
  End If
  SCI.DirectSCI.SetSelText strCmp
  SCI.SelStart = lStart
  SCI.SelEnd = lEnd + lAdd
  Erase ua()
  Exit Sub
errHandler:
  Erase ua()    ' Just in case it breaks off somewhere erase the
                ' array so it's not taking up unneccisary memory.
End Sub

Public Sub UncommentBlock2(SCI As scisimple)
  On Error Resume Next
  Dim str As String
  Dim lStart As Long, lEnd As Long
  Dim ua() As String
  Dim currentHighlighter As Integer
  
  str = SCI.SelText
  currentHighlighter = FindHighlighter(SCI.currentHighlighter)
  lStart = SCI.SelStart
  lEnd = SCI.SelEnd
  ua() = Split(str, Highlighters(currentHighlighter).strComment)
  str = Replace(str, Highlighters(currentHighlighter).strComment, "")
  SCI.SelText = str
  SCI.DirectSCI.SetSel lStart, lEnd - (UBound(ua) * Len(Highlighters(currentHighlighter).strComment))
  Erase ua()

End Sub


