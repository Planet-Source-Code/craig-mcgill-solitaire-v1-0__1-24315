VERSION 5.00
Begin VB.Form frmColour 
   BackColor       =   &H00FF8080&
   Caption         =   "Choose your marble !"
   ClientHeight    =   3990
   ClientLeft      =   6435
   ClientTop       =   2850
   ClientWidth     =   4995
   Icon            =   "frmColour.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4995
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   840
      Picture         =   "frmColour.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   1680
      Picture         =   "frmColour.frx":1444
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   3480
      Picture         =   "frmColour.frx":2446
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   2640
      Picture         =   "frmColour.frx":3448
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3960
      Picture         =   "frmColour.frx":444A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   3720
      Picture         =   "frmColour.frx":544C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2760
      Picture         =   "frmColour.frx":644E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1800
      Picture         =   "frmColour.frx":7450
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   840
      Picture         =   "frmColour.frx":8452
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox piccolour 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "frmColour.frx":9454
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "frmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim letter As String

Private Sub Command1_Click()
    frmSolitaire.choice = letter
    Unload Me
End Sub

Private Sub Form_Load()
    frmColour.BackColor = RGB(135, 135, 255)
    Select Case Mid(Right(frmSolitaire.currentMarble, 5), 1, 1)
        Case "a": piccolour(2).BorderStyle = 1
        Case "b": piccolour(3).BorderStyle = 1
        Case "c": piccolour(4).BorderStyle = 1
        Case "d": piccolour(7).BorderStyle = 1
        Case "e": piccolour(8).BorderStyle = 1
        Case "f": piccolour(9).BorderStyle = 1
        Case "g": piccolour(1).BorderStyle = 1
        Case "h": piccolour(5).BorderStyle = 1
        Case "i": piccolour(6).BorderStyle = 1
        Case "j": piccolour(0).BorderStyle = 1
    End Select
End Sub

Private Sub picColour_Click(Index As Integer)
Dim iter As Integer

    Select Case Index
        Case 0: letter = "j"
        Case 1: letter = "g"
        Case 2: letter = "a"
        Case 3: letter = "b"
        Case 4: letter = "c"
        Case 5: letter = "h"
        Case 6: letter = "i"
        Case 7: letter = "d"
        Case 8: letter = "e"
        Case 9: letter = "f"
    End Select
    piccolour(Index).BorderStyle = 1
    For iter = 0 To 9
        If Not (iter = Index) Then
            piccolour(iter).BorderStyle = 0
        End If
    Next iter
End Sub
