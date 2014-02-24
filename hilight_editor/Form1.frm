VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "SCIVB Highlighter Editor"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLang 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6165
      TabIndex        =   35
      Top             =   1305
      Width           =   600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set all to current "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2970
      TabIndex        =   31
      Top             =   8730
      Width           =   3480
      Begin VB.CommandButton cmdResetSizes 
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   33
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdResetFonts 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.TextBox txtPropName 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3825
      TabIndex        =   30
      Top             =   5715
      Width           =   4380
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   8280
      TabIndex        =   28
      Top             =   9180
      Width           =   1590
   End
   Begin VB.CommandButton cmdUpdateStyle 
      Caption         =   "Update"
      Height          =   420
      Left            =   8280
      TabIndex        =   27
      Top             =   5670
      Width           =   1500
   End
   Begin VB.CommandButton cmdUpdateKeyWord 
      Caption         =   "Update"
      Height          =   375
      Left            =   8235
      TabIndex        =   26
      Top             =   2250
      Width           =   1365
   End
   Begin VB.TextBox txtKeyWords 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   2835
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   2250
      Width           =   5325
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1530
      TabIndex        =   22
      Top             =   1305
      Width           =   2625
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1530
      TabIndex        =   20
      Top             =   900
      Width           =   6585
   End
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1530
      TabIndex        =   17
      Top             =   495
      Width           =   6585
   End
   Begin VB.CheckBox chkEOLFilled 
      Caption         =   "EOLFilled"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6165
      TabIndex        =   16
      Top             =   8145
      Width           =   2130
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "Visible"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6165
      TabIndex        =   15
      Top             =   7695
      Width           =   1680
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6165
      TabIndex        =   14
      Top             =   7110
      Width           =   1635
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6165
      TabIndex        =   13
      Top             =   6615
      Width           =   1410
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6165
      TabIndex        =   12
      Top             =   6120
      Width           =   1365
   End
   Begin Project1.ArielColorBox clrFore 
      Height          =   315
      Left            =   3825
      TabIndex        =   11
      Top             =   6930
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ArielColorBox clrBack 
      Height          =   315
      Left            =   3825
      TabIndex        =   10
      Top             =   6525
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSize 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Text            =   "10"
      Top             =   7380
      Width           =   2220
   End
   Begin VB.ComboBox cmbFont 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3825
      TabIndex        =   4
      Top             =   6120
      Width           =   2175
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4110
      Left            =   90
      TabIndex        =   3
      Top             =   5535
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   7250
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   8145
      TabIndex        =   2
      Top             =   45
      Width           =   1320
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1530
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   45
      Width           =   6540
   End
   Begin MSComctlLib.ListView lvKeyWords 
      Height          =   3210
      Left            =   45
      TabIndex        =   23
      Top             =   2250
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Language"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   5040
      TabIndex        =   34
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Index           =   6
      Left            =   3330
      TabIndex        =   29
      Top             =   5760
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "KeyWords: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   45
      TabIndex        =   24
      Top             =   1890
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Name: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   675
      TabIndex        =   21
      Top             =   1395
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Comment: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   495
      TabIndex        =   19
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Filter: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   495
      TabIndex        =   18
      Top             =   540
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Size:"
      Height          =   255
      Left            =   3285
      TabIndex        =   9
      Top             =   7425
      Width           =   450
   End
   Begin VB.Label Label3 
      Caption         =   "Forecolor:"
      Height          =   375
      Left            =   2925
      TabIndex        =   8
      Top             =   6975
      Width           =   870
   End
   Begin VB.Label Label4 
      Caption         =   "Backcolor:"
      Height          =   240
      Left            =   2925
      TabIndex        =   7
      Top             =   6615
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Index           =   1
      Left            =   3285
      TabIndex        =   5
      Top             =   6165
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Hilighter file: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'basic editor for scivb hilighter files - dzzie@yahoo.com
'Ariel colorbox user control by T De Lange, 2000

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

Private H As Highlighter
Private selKeyWord As ListItem
Private selItem As ListItem

Private Sub cmdResetFonts_Click()
    For i = 0 To 127
        H.StyleFont(i) = cmbFont.Text
    Next
End Sub

Private Sub cmdResetSizes_Click()
    For i = 0 To 127
        H.StyleSize(i) = txtSize.Text
    Next
End Sub

Private Sub cmdSave_Click()
    SaveHighlighter txtFile
End Sub

Private Sub cmdUpdateKeyWord_Click()
    
    If selKeyWord Is Nothing Then Exit Sub
    
    Dim i As Long
    i = selKeyWord.Tag
    H.Keywords(i) = txtKeyWords
    selKeyWord.Text = i & ") Length: " & Len(H.Keywords(i))
    
End Sub

Private Sub cmdUpdateStyle_Click()
    
    If selItem Is Nothing Then Exit Sub
    
    Dim i As Long
    i = selItem.Tag
    selItem.Text = txtPropName
    
    With H
         .StyleName(i) = txtPropName
         .StyleFont(i) = cmbFont.Text
         .StyleSize(i) = txtSize
         .StyleBack(i) = clrBack.SelectedColor
         .StyleFore(i) = clrFore.SelectedColor
         .StyleBold(i) = chkBold.Value
         .StyleItalic(i) = chkItalic.Value
         .StyleUnderline(i) = chkUnderline.Value
         .StyleVisible(i) = chkVisible.Value
         .StyleEOLFilled(i) = chkEOLFilled.Value
    End With
    
    
End Sub

Private Sub lvKeyWords_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Long
    i = Item.Tag
    Set selKeyWord = Item
    txtKeyWords = H.Keywords(i)
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim i As Long
    i = Item.Tag
    Set selItem = Item
    
    With H
        txtPropName = .StyleName(i)
        cmbFont.Text = .StyleFont(i)
        txtSize = .StyleSize(i)
        clrBack.SelectedColor = .StyleBack(i)
        clrFore.SelectedColor = .StyleFore(i)
        chkBold.Value = .StyleBold(i)
        chkItalic.Value = .StyleItalic(i)
        chkUnderline.Value = .StyleUnderline(i)
        chkVisible.Value = .StyleVisible(i)
        chkEOLFilled.Value = .StyleEOLFilled(i)
    End With
    
End Sub


Private Sub cmdLoad_Click()
    
    Dim li As ListItem
    
    If Not FileExists(txtFile) Then
        MsgBox "File not found", vbInformation
        Exit Sub
    End If
    
    LoadHighlighter txtFile
    
    txtFilter = H.strFilter
    txtComment = H.strComment
    txtName = H.strName
    txtLang = H.iLang
    
    lv.ListItems.Clear
    lvKeyWords.ListItems.Clear
    For i = 0 To 7
        Set li = lvKeyWords.ListItems.Add(, , i & ") Length: " & Len(H.Keywords(i)))
        li.Tag = i
    Next
    
    For i = 0 To 127
        Set li = lv.ListItems.Add(, , H.StyleName(i))
        li.Tag = i
    Next
    
    lv_ItemClick lv.ListItems(1)
    lvKeyWords_ItemClick lvKeyWords.ListItems(1)
    
End Sub

                                     
Public Function LoadHighlighter(strFile As String)
  Dim fFile As Integer
  fFile = FreeFile
  Open strFile For Binary Access Read As #fFile
  Get #fFile, , H
  Close #fFile
  FreeFile fFile
End Function
 
Public Function SaveHighlighter(strFile As String)
  On Error Resume Next
  Dim fFile As Integer
  fFile = FreeFile
  Kill strFile
  Open strFile For Binary Access Write As #fFile
  Put #fFile, , H
  Close #fFile
  FreeFile fFile
End Function

Private Sub Form_Load()
    lv.ColumnHeaders(1).Width = lv.Width - 150
    lvKeyWords.ColumnHeaders(1).Width = lvKeyWords.Width - 150
    txtFile = App.Path & "\..\java.hilighter"
    Dim i As Long
    For i = 0 To Screen.FontCount - 1
      cmbFont.AddItem Screen.Fonts(i)
    Next i
End Sub

 Public Function FileExists(strFile As String) As Boolean
 If Len(strFile) = 0 Then Exit Function
  If Dir(strFile) = "" Then
    FileExists = False
  Else
    FileExists = True
  End If
End Function


Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtFile = Data.Files(1)
    cmdLoad_Click
End Sub
