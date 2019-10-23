VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   720
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "OUTPUT"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2085
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "UPPER VALUE"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   1440
         Width           =   500
      End
      Begin VB.TextBox Text2 
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Text            =   "297.532"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOWER VALUE"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   120
         TabIndex        =   1
         Text            =   "294.286"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DISTANCE"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
      Begin VB.TextBox Text3 
         Height          =   525
         Left            =   120
         TabIndex        =   8
         Text            =   "5"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&file"
      Begin VB.Menu mnuci 
         Caption         =   "contour interval"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusolve 
         Caption         =   "&Calculate"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuinput 
         Caption         =   "&inputdata"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu 
         Caption         =   "e&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuin 
      Caption         =   "in"
   End
   Begin VB.Menu aw 
      Caption         =   "a"
      Begin VB.Menu aa 
         Caption         =   "a"
      End
      Begin VB.Menu asd 
         Caption         =   "asd"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, a, b, d, k, p, q, j, t1, t2, t3, t4, c As Integer
Dim X(1000), Y(1000), z(1000) As Single
Dim CI As Variant
Dim abc As Variant

Private Sub aa_Click()
frmAbout.Show
End Sub

Private Sub asd_Click()
frmrichtext.Show
End Sub

Private Sub Form_Resize()
Text5.Width = Form1.Width - 450 - 240
Frame4.Height = Form1.Height - (7980 - 2295)
Text5.Height = Form1.Height - (7980 - 2295) - (2295 - 1725)
text1.Width = Form1.Width - 450 - 240
Text2.Width = Form1.Width - 450 - 240
Text3.Width = Form1.Width - 450 - 240
Frame1.Width = Form1.Width - 450
Frame2.Width = Form1.Width - 450
Frame3.Width = Form1.Width - 450
Frame4.Width = Form1.Width - 450
CI = 1
End Sub

Private Sub Label3_Click()

End Sub

Private Sub mnu_Click()
End
End Sub

Private Sub mnuci_Click()
'CI = InputBox("Enter Contour Inteval ", "CI")
CI = 0
Do While Not IsNumeric(CI) Or CI <= 0
'If CI <= 0 Or Not IsNumeric(CI) Then
CI = InputBox("Enter Contour Inteval ", "CI")
'End If
Loop
End Sub

Private Sub mnuin_Click()
Randomize
abc = InputBox("save loc : ", "Save", DateValue(Now) & "-" & TimeValue(Now) & Int(10 * Rnd))
End Sub

Private Sub mnuinput_Click()
text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
text1.SetFocus
End Sub

Private Sub mnusolve_Click()
Text5.Text = ""
a = Val(text1.Text)
 b = Val(Text2.Text)
 d = Val(Text3.Text)
 p = Int(a)
 q = Int(b)
 'i = q - p
'For k = 0 To i - 1 Step 1
  '  Y(k) = p + 1 + k
'Next k
Y(0) = p
c = 0
k = 0
Do While Y(k) <= (q + 1)
k = k + 1
Y(k) = Y(k - 1) + CI
If Y(k) >= a And Y(k) <= b Then z(c) = Y(k): c = c + 1
Loop
'For j = 0 To i - 1
 '   X(j) = (d / (b - a)) * (Y(j) - a)
  '  Text5.Text = Text5.Text & vbCrLf & j + 1 & "     y(" & (j + 1) & ") = " & Y(j) & "     x(" & (j + 1) & ") = " & X(j)
'Next j
'Text6.Text = c & " ci  " & CI
For j = 0 To c
If z(j) >= a And z(j) <= b Then
'Text5.Text = Text5.Text & vbCrLf & " y(" & (j + 1) & ") = " & z(j)
 X(j) = (d / (b - a)) * (z(j) - a)
 Text5.Text = Text5.Text & vbCrLf & j + 1 & "     y(" & (j + 1) & ") = " & z(j) & "     x(" & (j + 1) & ") = " & X(j)
End If
Next j
End Sub

Private Sub SonicClick1_Click()
End
End Sub
