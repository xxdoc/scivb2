VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SCIVB"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7650
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":0000
      Top             =   720
      Width           =   8565
   End
   Begin VB.Label lblURL 
      Caption         =   "http://www.scintilla.org"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label lblURL 
      Caption         =   "https://github.com/dzzie/scivb2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label lblTop 
      Caption         =   "SCIVB_LITE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3705
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0006
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
      Unload Me
End Sub

Public Function LaunchForm(owner As scisimple)

     txtDesc = CompileVersionInfo(owner) & vbCrLf & vbCrLf & _
            "SCIVB is an easy to use ActiveX control that wraps Scintilla." & vbCrLf & _
            vbCrLf & _
            "Scintilla is an excellent opensource component which " & _
            "supports syntax highlighting, folding, code tips, and much more." & vbCrLf & _
            vbCrLf & _
            "SCIVB Created by Stu Collier and Stewart, mods by dzzie"
            
End Function

 
Private Sub lblURL_Click(index As Integer)
        ShellDocument lblURL(index).Caption  '"http://www.ceditmx.com" 'original authors website is no longer up
End Sub
