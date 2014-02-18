VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMiscSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'these have been seperated out to lighten main user control from rarely used stuff
'they have not been left to DirectSCI access because they require intimiate knowledge of sci

Private owner As scisimple

Dim m_BraceMatchBold As Boolean
Dim m_BraceMatchItalic As Boolean
Dim m_BraceMatchUnderline As Boolean
Dim m_BraceMatchBack As OLE_COLOR
Dim m_BraceBadBack As OLE_COLOR
Dim m_BraceMatch As OLE_COLOR
Dim m_BraceBad As OLE_COLOR

Dim m_Gutter0Type As GutterTypes
Dim m_Gutter0Width As Long
Dim m_Gutter1Type As GutterTypes
Dim m_Gutter1Width As Long
Dim m_Gutter2Type As GutterTypes
Dim m_Gutter2Width As Long

Dim m_MarkerBack As OLE_COLOR
Dim m_MarkerFore As OLE_COLOR
'Dim m_CaretForeColor As OLE_COLOR

' Gutter Type Enum (Using a symbol or linenumber gutter style.)
Public Enum GutterTypes
  GutSymbol = 0
  GutLineNumber = 1
End Enum

Public Enum Gutters
    gut0 = 0
    gut1 = 1
    gut2 = 2
End Enum

Sub Initilize(o As scisimple)
    Set owner = o
End Sub

'Public Property Let CaretForeColor(ByVal New_CaretForeColor As OLE_COLOR)
'    m_CaretForeColor = New_CaretForeColor
'    owner.DirectSCI.SetCaretFore New_CaretForeColor
'End Property

'============================ [ braces ] =============================================
Public Property Get BraceMatchFore() As OLE_COLOR
    BraceMatchFore = m_BraceMatch
End Property

Public Property Let BraceMatchFore(ByVal New_BraceMatch As OLE_COLOR)
    m_BraceMatch = New_BraceMatch
    'PropertyChanged "BraceMatch"
    owner.DirectSCI.StyleSetFore 34, New_BraceMatch
End Property

Public Property Get BraceBadFore() As OLE_COLOR
    BraceBadFore = m_BraceBad
End Property

Public Property Let BraceBadFore(ByVal New_BraceBad As OLE_COLOR)
    m_BraceBad = New_BraceBad
    'PropertyChanged "BraceBad"
    owner.DirectSCI.StyleSetFore 35, New_BraceBad
End Property

Public Property Get BraceMatchBold() As Boolean
    BraceMatchBold = m_BraceMatchBold
End Property

Public Property Let BraceMatchBold(ByVal New_BraceMatchBold As Boolean)
    m_BraceMatchBold = New_BraceMatchBold
    'PropertyChanged "BraceMatchBold"
    owner.DirectSCI.StyleSetBold 35, New_BraceMatchBold
    owner.DirectSCI.StyleSetBold 34, New_BraceMatchBold
End Property

Public Property Get BraceMatchItalic() As Boolean
    BraceMatchItalic = m_BraceMatchItalic
End Property

Public Property Let BraceMatchItalic(ByVal New_BraceMatchItalic As Boolean)
    m_BraceMatchItalic = New_BraceMatchItalic
    'PropertyChanged "BraceMatchItalic"
    owner.DirectSCI.StyleSetItalic 35, New_BraceMatchItalic
    owner.DirectSCI.StyleSetItalic 34, New_BraceMatchItalic
End Property

Public Property Get BraceMatchUnderline() As Boolean
    BraceMatchUnderline = m_BraceMatchUnderline
End Property

Public Property Let BraceMatchUnderline(ByVal New_BraceMatchUnderline As Boolean)
    m_BraceMatchUnderline = New_BraceMatchUnderline
    'PropertyChanged "BraceMatchUnderline"
    owner.DirectSCI.StyleSetUnderline 35, New_BraceMatchUnderline
    owner.DirectSCI.StyleSetUnderline 34, New_BraceMatchUnderline
End Property

Public Property Get BraceMatchBack() As OLE_COLOR
    BraceMatchBack = m_BraceMatchBack
End Property

Public Property Let BraceMatchBack(ByVal New_BraceMatchBack As OLE_COLOR)
    m_BraceMatchBack = New_BraceMatchBack
    'PropertyChanged "BraceMatchBack"
    owner.DirectSCI.StyleSetBack 34, New_BraceMatchBack
End Property

Public Property Get BraceBadBack() As OLE_COLOR
    BraceBadBack = m_BraceBadBack
End Property

Public Property Let BraceBadBack(ByVal New_BraceBadBack As OLE_COLOR)
    m_BraceBadBack = New_BraceBadBack
    'PropertyChanged "BraceBadBack"
    owner.DirectSCI.StyleSetBack 35, New_BraceBadBack
End Property


'============================= [ gutters ] =============================================
Property Get GutterType(g As Gutters) As GutterTypes
    If g = gut0 Then GutterType = m_Gutter0Type
    If g = gut1 Then GutterType = m_Gutter1Type
    If g = gut2 Then GutterType = m_Gutter2Type
End Property

Property Get GutterWidth(g As Gutters) As Long
    If g = gut0 Then GutterWidth = m_Gutter0Width
    If g = gut1 Then GutterWidth = m_Gutter1Width
    If g = gut2 Then GutterWidth = m_Gutter2Width
End Property

Property Let GutterType(g As Gutters, v As GutterTypes)
    
    If g = gut0 Then
        m_Gutter0Type = v
        owner.DirectSCI.SetMarginTypeN 0, m_Gutter0Type
    End If
    
    If g = gut1 Then
        m_Gutter1Type = v
        owner.DirectSCI.SetMarginTypeN 1, m_Gutter1Type
        If owner.ShowFlags = True Then
            owner.DirectSCI.SetMarginWidthN 1, m_Gutter1Type '<-- this looks like a bug? -dzzie
        Else
            owner.DirectSCI.SetMarginWidthN 1, 0
        End If
    End If
    
    If g = gut2 Then
        m_Gutter2Type = v
        owner.DirectSCI.SetMarginTypeN 2, m_Gutter2Type
    End If
    
End Property

Property Let GutterWidth(g As Gutters, v As Long)
    
    If g = gut0 Then
        m_Gutter0Width = v
        owner.DirectSCI.SetMarginWidthN 0, v
        If owner.LineNumbers = True Then
          owner.DirectSCI.SetMarginWidthN 0, v
        Else
          owner.DirectSCI.SetMarginWidthN 0, 0
        End If
    End If
    
    If g = gut1 Then
        m_Gutter1Width = v
        If owner.Folding = True Then
            owner.DirectSCI.SetMarginWidthN 2, v
        Else
            owner.DirectSCI.SetMarginWidthN 2, 0
        End If
    End If
    
    If g = gut2 Then
        m_Gutter2Width = v
        owner.DirectSCI.SetMarginWidthN 2, v
    End If
    
End Property
 
'==================================== [ markers ] ===============================
Public Property Get MarkerBack() As OLE_COLOR   'Allows you to configure the backcolor of the folding markers.
    MarkerBack = m_MarkerBack
End Property

Public Property Let MarkerBack(ByVal New_MarkerBack As OLE_COLOR)
    m_MarkerBack = New_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPEN, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDER, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERMIDTAIL, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERSUB, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDERTAIL, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPEN, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEROPENMID, m_MarkerBack
    owner.DirectSCI.MarkerSetBack SC_MARKNUM_FOLDEREND, m_MarkerBack
End Property

Public Property Get MarkerFore() As OLE_COLOR   'Allows you to configure the forecolor of the folding marker.
    MarkerFore = m_MarkerFore
End Property

Public Property Let MarkerFore(ByVal New_MarkerFore As OLE_COLOR)
    m_MarkerFore = New_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPEN, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDER, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERMIDTAIL, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERSUB, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDERTAIL, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPEN, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEROPENMID, m_MarkerFore
    owner.DirectSCI.MarkerSetFore SC_MARKNUM_FOLDEREND, m_MarkerFore
End Property



 Public Function CurrentFunction()

    Dim line As String
    Dim i As Integer, i2 As Integer, X As Integer
    line = owner.GetLineText(owner.CurrentLine())

  CurrentFunction = ""
  X = owner.GetCaretInLine

  For i = X To 1 Step -1
        If Mid(line, i, 1) = "(" Then
            For i2 = i - 1 To 1 Step -1
                If Mid(line, i2, 1) < 33 And CurrentFunction <> "" Then    ' ignore whitespace before (
                    Exit For
                Else
                    If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                        CurrentFunction = Mid(line, i2, 1) & CurrentFunction
                    Else
                        If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                            Exit For
                        End If
                    End If
                End If
            Next i2
        End If
    
        If CurrentFunction <> "" Then
            Exit For
        End If
  Next i

  ' Cant find a function going backwards - check forwards instead ?
  If CurrentFunction = "" Then
        For i = X To Len(line)
            If Mid(line, i, 1) = "(" Then
                For i2 = i - 1 To 1 Step -1
                    If Mid(line, i2, 1) < 33 And CurrentFunction <> "" Then    ' ignore whitespace before (
                        Exit For
                    Else
                        If InStr(1, CallTipWordCharacters, Mid(line, i2, 1)) > 0 Then
                            CurrentFunction = Mid(line, i2, 1) & CurrentFunction
                        Else
                            If Asc(Mid(line, i2, 1)) > 33 Then   ' not valid character (and not whitespace)
                                Exit For
                            End If
                        End If
                    End If
                Next i2
            End If
            If CurrentFunction <> "" Then
                Exit For
            End If
        Next i
  End If

End Function
 


