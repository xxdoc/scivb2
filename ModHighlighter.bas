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

Public hlCount As Long

Public Type Highlighter
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

Public HCount As Integer

Public Highlighters() As Highlighter ' Make it publicly exposed so the app can
                                     ' read off name's for menu's or such
Private CurrentHighlighter As Integer

Private sBuffer As String
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Public Sub ReInit()
sBuffer = ""
lOffset = 0
End Sub

Public Function GetString() As String
GetString = Left$(sBuffer, lOffset)
sBuffer = ""  'reset
lOffset = 0

End Function

'This function lets you assign a string to the concating buffer.
Public Sub SetString(ByRef Source As String)
sBuffer = Source & String$(ciIncriment, 0)
End Sub

Public Sub SConcat(ByRef Source As String)
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

Private Function FindHighlighter(strLangName As String) As Integer
  Dim i As Integer
  Dim L As Long
  L = GetTickCount
   For i = 0 To hlCount - 1 ' UBound(Highlighters) - 1
    
    'If Len(strLangName) = Len(Highlighters(i)) Then
      If UCase(Highlighters(i).strName) = UCase(strLangName) Then
        FindHighlighter = i
        Exit Function
      End If
'    End If
  Next i
    
End Function


Public Function IsSet(X As Long, i As Long) As Boolean
  'If Highlighters(x).Keywords(i) <> "" Then
    If Highlighters(X).StyleBack(i) <> Highlighters(X).StyleBack(32) _
      Or Highlighters(X).StyleBold(i) <> Highlighters(X).StyleBold(32) _
        Or Highlighters(X).StyleEOLFilled(i) <> Highlighters(X).StyleEOLFilled(32) _
          Or Highlighters(X).StyleFont(i) <> Highlighters(X).StyleFont(32) _
            Or Highlighters(X).StyleFore(i) <> Highlighters(X).StyleFore(32) _
              Or Highlighters(X).StyleItalic(i) <> Highlighters(X).StyleItalic(32) _
                Or Highlighters(X).StyleSize(i) <> Highlighters(X).StyleSize(32) _
                  Or Highlighters(X).StyleUnderline(i) <> Highlighters(X).StyleUnderline(32) _
                    Or Highlighters(X).StyleVisible(i) <> Highlighters(X).StyleVisible(32) Then

'                      Debug.Print "Changing Style On Highlighter " & Highlighters(x).strName & " | Style Number: " & i
                      IsSet = True
    End If
End Function

Public Function SetHighlighters(scisimple As scisimple, strHighlighter As String, Optional lMarginBack As Long, Optional lMarginFore As Long)
  Dim i As Long, X As Long
  
  scisimple.DirectSCI.ClearDocumentStyle
  
  X = FindHighlighter(strHighlighter)
  
  If LCase(strHighlighter) = "html" Then
        scisimple.DirectSCI.StyleSetBits 7
  Else
        scisimple.DirectSCI.StyleSetBits 5
  End If
  
  scisimple.DirectSCI.SetLexer Highlighters(X).iLang
  For i = 0 To 7
        If Highlighters(X).Keywords(i) <> "" Then scisimple.DirectSCI.SetKeyWords i, Highlighters(X).Keywords(i)
  Next i
 
  scisimple.DirectSCI.StyleSetBack 32, Highlighters(X).StyleBack(32)
  scisimple.DirectSCI.StyleSetFore 32, Highlighters(X).StyleFore(32)
  scisimple.DirectSCI.StyleSetVisible 32, CLng(Highlighters(X).StyleVisible(32))
  scisimple.DirectSCI.StyleSetEOLFilled 32, CLng(Highlighters(X).StyleEOLFilled(32))
  scisimple.DirectSCI.StyleSetBold 32, CLng(Highlighters(X).StyleBold(32))
  scisimple.DirectSCI.StyleSetItalic 32, CLng(Highlighters(X).StyleItalic(32))
  scisimple.DirectSCI.StyleSetUnderline 32, CLng(Highlighters(X).StyleUnderline(32))
  scisimple.DirectSCI.StyleSetFont 32, Highlighters(X).StyleFont(32)
  scisimple.DirectSCI.StyleSetSize 32, Highlighters(X).StyleSize(32)
  scisimple.DirectSCI.StyleClearAll
  
  For i = 0 To 127
        scisimple.DirectSCI.StyleSetBold i, CLng(Highlighters(X).StyleBold(i))
        scisimple.DirectSCI.StyleSetItalic i, CLng(Highlighters(X).StyleItalic(i))
        scisimple.DirectSCI.StyleSetUnderline i, CLng(Highlighters(X).StyleUnderline(i))
        scisimple.DirectSCI.StyleSetVisible i, CLng(Highlighters(X).StyleVisible(i))
        If Highlighters(X).StyleFont(i) <> "" Then scisimple.DirectSCI.StyleSetFont i, Highlighters(X).StyleFont(i)
        scisimple.DirectSCI.StyleSetFore i, CLng(Highlighters(X).StyleFore(i))
        scisimple.DirectSCI.StyleSetBack i, CLng(Highlighters(X).StyleBack(i))
        scisimple.DirectSCI.StyleSetSize i, CLng(Highlighters(X).StyleSize(i))
        scisimple.DirectSCI.StyleSetEOLFilled i, CLng(Highlighters(X).StyleEOLFilled(i))
  Next i
  
  scisimple.DirectSCI.StyleSetFore 35, scisimple.misc.BraceBadFore
  scisimple.DirectSCI.StyleSetFore 34, scisimple.misc.BraceMatchFore
  scisimple.DirectSCI.StyleSetBack 35, scisimple.misc.BraceBadBack
  scisimple.DirectSCI.StyleSetBack 34, scisimple.misc.BraceMatchBack
  scisimple.DirectSCI.StyleSetBold 35, scisimple.misc.BraceMatchBold
  scisimple.DirectSCI.StyleSetBold 34, scisimple.misc.BraceMatchBold
  scisimple.DirectSCI.StyleSetItalic 35, scisimple.misc.BraceMatchItalic
  scisimple.DirectSCI.StyleSetItalic 34, scisimple.misc.BraceMatchItalic
  scisimple.DirectSCI.StyleSetUnderline 35, scisimple.misc.BraceMatchUnderline
  scisimple.DirectSCI.StyleSetUnderline 34, scisimple.misc.BraceMatchUnderline
  
  CurrentHighlighter = X
  scisimple.DirectSCI.Colourise 0, -1
  scisimple.CurrentHighlighter = strHighlighter
  
End Function

Public Function LoadHighlighter(strFile As String)
  Dim fFile As Integer
  
  If Not FileExists(strFile) Then Exit Function
  
  fFile = FreeFile
  ReDim Preserve Highlighters(0 To HCount + 1)
  
  Open strFile For Binary Access Read As #fFile
  Get #fFile, , Highlighters(HCount)
  Highlighters(HCount).strFile = strFile
  Close #fFile
  
  FreeFile fFile
  HCount = HCount + 1
  
End Function

Public Sub LoadDirectory(strDir As String)
  Dim str As String, i As Long
  hlCount = 0
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  str = Dir(strDir & "\*bin")
  Erase Highlighters
  HCount = 0
  Do Until str = ""
    hlCount = hlCount + 1
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

Public Function SetHighlighterBasedOnExtension(file As String, Optional lMarginBack As Long, Optional lMarginFore As Long) As String
  Dim Extension As String, ua() As String, ClrExt As String, X As Long
  Extension = LCase$(Mid$(file, InStrRev(file, ".") + 1, Len(file) - InStrRev(file, ".")))
  Extension = "." & Extension
  For X = 0 To hlCount - 1 'UBound(Highlighters)
    If InStr(1, Highlighters(X).strFilter, Extension) Then
      SetHighlighterBasedOnExtension = Highlighters(X).strName
      Exit For
    End If
  Next X
End Function

Public Function ExportToHTML2(strFile As String, scisimple As scisimple)
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
  CurrentHighlighter = FindHighlighter(scisimple.CurrentHighlighter)
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
      With Highlighters(CurrentHighlighter)
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
End Function


Public Function DectoHex(lngColour As Long) As String

    '     *********
    Dim strColour As String
    'Convert decimal colour to hex
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

Function AddToString(St As String, ToAdd As String, Optional NumTimes As Long = 1) As String

    Dim LC As Long, StrLoc As Long
    AddToString = String$((Len(ToAdd) * NumTimes) + Len(St), 0) 'For CopyMemory() to work, the string must be padded With nulls to the desired size
    CopyMemory ByVal StrPtr(AddToString), ByVal StrPtr(St), LenB(St) 'Copy the original string to the return code
    StrLoc = StrPtr(AddToString) + LenB(St) 'Memory Location = Location of return code + size of original string
    'We use LenB() because strings are actua
    '     lly twice as long as Len() says when sto
    '     red in memory

    For LC = 1 To NumTimes
        CopyMemory ByVal StrLoc, ByVal StrPtr(ToAdd), LenB(ToAdd) 'Copy the source String to the return code
        StrLoc = StrLoc + LenB(ToAdd) 'Add the size of the String to the pointer


        DoEvents 'Comment this out If you don't plan To use huge repeat values, you'll Get a nice speed boost
        Next LC

 End Function

'Public Function DoSyntaxOptions(strDir As String, hl As SCIHighlighter) As Boolean
'  'Load frmOptions
'  With frmOptions
'    Set .hlMain = hl
'    .hlPath = strDir
'    '.ListLangs strDir
'    .show vbModal
'    If .WhatToDo = 1 Then
'      DoSyntaxOptions = True
'    Else
'      DoSyntaxOptions = False
'    End If
'  End With
'End Function

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
  lStart = SCI.SelStart
  lEnd = SCI.SelEnd
  lLineStart = SCI.DirectSCI.LineFromPosition(lStart)
  lLineEnd = SCI.DirectSCI.LineFromPosition(lEnd)
  strCmp = ""
  CurrentHighlighter = FindHighlighter(SCI.CurrentHighlighter)
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
    strCmp = strCmp & Highlighters(CurrentHighlighter).strComment & ua(i)
    If i < UBound(ua) Then strCmp = strCmp & strSplit
  Next i
  If UBound(ua) > 0 Then
    lAdd = ((UBound(ua) + 1) * Len(Highlighters(CurrentHighlighter).strComment)) ' + (Len(strSplit) * (UBound(ua) - 1))
  Else
    lAdd = Len(Highlighters(CurrentHighlighter).strComment)
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
  str = SCI.SelText
  CurrentHighlighter = FindHighlighter(SCI.CurrentHighlighter)
  lStart = SCI.SelStart
  lEnd = SCI.SelEnd
  ua() = Split(str, Highlighters(CurrentHighlighter).strComment)
  str = Replace(str, Highlighters(CurrentHighlighter).strComment, "")
  SCI.SelText = str
  SCI.DirectSCI.SetSel lStart, lEnd - (UBound(ua) * Len(Highlighters(CurrentHighlighter).strComment))
  Erase ua()

End Sub


