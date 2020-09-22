VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Example"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DON'T STOP THE APP. BY CLICKING END BUTTON IN VISUAL BASIC - THE FONT WILL REMAIN IN MEMORY!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BEFORE USING MODULE 'modFont' IN YOUR APP. PLEASE READ
'>>> readme.txt

'DON'T STOP THE APP. BY CLICKING END BUTTON IN VISUAL BASIC
'- THE FONT WILL REMAIN IN MEMORY!

Dim fntFileName01 As String, fntFileName02 As String
Dim fntName01 As String, fntName02 As String

Private Sub Form_Load()

fntFileName01 = "Ffxrndlt.ttf" 'full path to the font file
fntFileName02 = "Microsbe.ttf"

fntName01 = UseFont(fntFileName01)
fntName02 = UseFont(fntFileName02)

Label1 = "This text use a font that is not installed on this system" & vbNewLine & "FontName: " & GetFontName(fntFileName01)
Label1.FontName = fntName01
Label1.FontSize = 16

Label2 = "Button: " & GetFontName(fntFileName02)
Label2.FontName = fntName02
Label2.FontSize = 9

Command1.FontName = fntName02
Command1.FontSize = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
'REMEMBER to remove the font(s) that you used otherwise the font(s)
'will temporary remain in your system and you will not be able to
'move or delete this file(s) until you restart the computer.
'
'Still if you don't remove the font, programs such as Word will
'recognize it in the font list. After restart the font will
'dissapear from the list.
'
'Remove only the font(s) that you added !
RemoveFont (fntFileName01)
RemoveFont (fntFileName02)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


