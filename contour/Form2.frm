VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmrichtext 
   Caption         =   "richtext"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"Form2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmrichtext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim ch As String
Dim cur As Integer
Dim j As Integer

Private Sub Command1_Click()
        With text1
           ' .SelStart = 2
            .SelLength = 1
            .SelColor = vbRed
        End With
End Sub

Private Sub Form_Load()
j = 1
With text1
    .SelStart = 0
    .SelLength = 0
    .SelColor = vbBlue
End With

End Sub

Private Sub text1_Change()
cur = Len(text1.Text)
For i = cur To 100
    ch = Mid(text1.Text, i, j)
        
    Select Case ch
     Case "*"
            With text1
           .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbRed
        End With
            Call currpos
        Case "+"
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbRed
        End With
        Call currpos
        Case "-"
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbRed
        End With
        Call currpos
        Case "/"
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbRed
        End With
        Call currpos
        Case "^"
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbRed
        End With
        Call currpos
        Case "("
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbYellow
        End With
        Call currpos
        Case ")"
         With text1
            .SelStart = i - 1
            .SelLength = 1
            .SelColor = vbYellow
        End With
        Call currpos
        
   End Select
Next i
End Sub


Public Sub currpos()
With text1
    .SelStart = cur
    .SelLength = 0
    .SelColor = vbBlue
End With
End Sub

