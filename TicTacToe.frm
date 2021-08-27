VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TicTacToe"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "TicTacToe"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   5655
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "ridhoyanualif copyright;2021"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hope you enjoyed the game :)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   2250
      End
      Begin VB.Label Label3 
         Caption         =   "*And then Player O will turn into Player X"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "*If someone was won. then Player X will turn into Player O"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "The Winner Is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Private Sub Command1_Click()
If flag = False Then
Command1.Caption = "X"
flag = True
Else
Command1.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command10_Click()
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
Text1.Text = ""
End Sub

Private Sub Command2_Click()
If flag = False Then
Command2.Caption = "X"
flag = True
Else
Command2.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command3_Click()
If flag = False Then
Command3.Caption = "X"
flag = True
Else
Command3.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command4_Click()
If flag = False Then
Command4.Caption = "X"
flag = True
Else
Command4.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command5_Click()
If flag = False Then
Command5.Caption = "X"
flag = True
Else
Command5.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command6_Click()
If flag = False Then
Command6.Caption = "X"
flag = True
Else
Command6.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command7_Click()
If flag = False Then
Command7.Caption = "X"
flag = True
Else
Command7.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command8_Click()
If flag = False Then
Command8.Caption = "X"
flag = True
Else
Command8.Caption = "O"
flag = False
End If
win
End Sub

Private Sub Command9_Click()
If flag = False Then
Command9.Caption = "X"
flag = True
Else
Command9.Caption = "O"
flag = False
End If
win
End Sub

Private Sub win()
If Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command2.Caption = "X" And Command5.Caption = "X" And Command8.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command3.Caption = "X" And Command6.Caption = "X" And Command9.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command3.Caption = "X" And Command5.Caption = "X" And Command7.Caption = "X" Then Text1.Text = ("Player X Win!")
If Command1.Caption = "O" And Command2.Caption = "O" And Command3.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command4.Caption = "O" And Command5.Caption = "O" And Command6.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command7.Caption = "O" And Command8.Caption = "O" And Command9.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command1.Caption = "O" And Command4.Caption = "O" And Command7.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command2.Caption = "O" And Command5.Caption = "O" And Command8.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command3.Caption = "O" And Command6.Caption = "O" And Command9.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command1.Caption = "O" And Command5.Caption = "O" And Command9.Caption = "O" Then Text1.Text = ("Player O Win!")
If Command3.Caption = "O" And Command5.Caption = "O" And Command7.Caption = "O" Then Text1.Text = ("Player O Win!")
End Sub

