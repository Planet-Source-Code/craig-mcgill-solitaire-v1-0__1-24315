VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSolitaire 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solitaire"
   ClientHeight    =   8085
   ClientLeft      =   2490
   ClientTop       =   2715
   ClientWidth     =   6810
   FillColor       =   &H00FF8080&
   FillStyle       =   0  'Solid
   Icon            =   "frmSolitaire.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   6810
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   7800
      TabIndex        =   83
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   40
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7200
      Top             =   2040
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1535
      ButtonWidth     =   1588
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Game"
            Object.ToolTipText     =   "Start a New Game"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Undo"
            Object.ToolTipText     =   "Undo the last move"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "marbles"
            Object.ToolTipText     =   "Change the style of marble"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Easy/Hard"
            Object.ToolTipText     =   "Choose game type"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About this Program"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit the program"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   81
      Top             =   7680
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6059
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   884
            MinWidth        =   884
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2295
            MinWidth        =   2295
            TextSave        =   "1:37 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   2206
            MinWidth        =   2206
            TextSave        =   "6/22/01"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolitaire.frx":19DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   63
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   62
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   61
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   60
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   59
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   58
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   57
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   5160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   56
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   55
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   54
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   53
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   52
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   51
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   50
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   49
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   48
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   47
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   46
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   45
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   44
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   43
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   42
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   41
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   39
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   24
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   38
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   37
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   36
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   34
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   33
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   32
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   31
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   30
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   29
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   28
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   27
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   26
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   25
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   24
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   23
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   22
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   21
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   20
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   42
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   19
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   18
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   17
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   16
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   15
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   14
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   48
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   13
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   49
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   12
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   50
      Top             =   2160
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   11
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   10
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   9
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   8
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   54
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   7
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   6
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   56
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   5
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   4
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   58
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   3
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   59
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   2
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   60
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   1
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   61
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   0
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   62
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   35
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   63
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   64
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   64
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   65
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   65
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   66
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   66
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   67
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   67
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   68
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   68
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   69
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   69
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   70
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   70
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   71
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   71
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   72
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   72
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   73
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   73
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   74
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   74
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   75
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   75
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   76
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   76
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   77
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   77
      Top             =   6360
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   78
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   78
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   79
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox space 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   495
      Index           =   80
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   80
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H0080C0FF&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Shape shpEdge 
      BorderWidth     =   2
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   6495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "frmSolitaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Type cell
    filled As Boolean
    spaceNo As Integer
End Type

Private Type coOrd
    x As Integer
    y As Integer
End Type

Private Type turn
    oldX As Integer
    oldY As Integer
    newX As Integer
    newY As Integer
    victimX As Integer
    victimY As Integer
End Type

Private Type positions
    x As Integer
    y As Integer
    spaceNo As Integer
    speedLeft As Integer
    speedTop As Integer
    directionLeft As Integer
    directionRight As Integer
End Type

Private Const animateSpeed = 1000000
Public currentMarble As String
Public choice As String
Public picPath As String
Dim gameType As String
Dim history() As turn
Dim oldPos() As positions
Dim oldCount As Integer
Dim spaceXY() As coOrd
Dim spaceCount As Integer
Dim board() As cell
Dim Gridlength As Integer

Dim filter As Integer
Dim validMove As Boolean
Dim over As Boolean
Dim marblesLeft As Integer
Dim turnsTaken As Integer
Dim blankPath, filledPath, anipath As String, clearPath As String
Dim countPieces, oldCountPieces, lastClick, OldLastClick As Integer
Dim someoneWon As Boolean
Dim onTarget As Boolean
Dim bTop As Integer
Dim bMiddle1 As Integer
Dim bMiddle2 As Integer
Dim bBottom1  As Integer
Dim bBottom2  As Integer

Private Sub Form_Load()
    Dim iter, cellNo, x, y As Integer
    
    ReDim spaceXY(80)
    ReDim board(1 To 9, 1 To 9)
    
    If gameType = "" Then
        gameType = "hard"
    End If
    
    spaceCount = 48
    Gridlength = 9
    
    If oldCount > 1 Then
        For iter = 1 To UBound(oldPos)
            space(oldPos(iter).spaceNo).Left = oldPos(iter).x
            space(oldPos(iter).spaceNo).Top = oldPos(iter).y
        Next iter
    End If
    filter = 1
    oldCount = 0
    space(40).Top = 3960
    space(40).Height = 495
    space(40).Width = 495
    space(40).Left = 3120
    space(40).Refresh
    
    frmSolitaire.Refresh
    
    picPath = App.Path & "\graphics\solitaireFilled1"
    If currentMarble = "" Then
        currentMarble = picPath & "d.bmp"
    End If
    For iter = 0 To 80
        space(iter).DragIcon = LoadPicture(App.Path & "\graphics\misc15.ico")
    Next iter
    turnsTaken = 0
    StatusBar1.Font.Bold = True
    StatusBar1.Font.Size = 8
    StatusBar1.Panels(1).Text = "You'll never do it !!"
    If gameType = "easy" Then
        marblesLeft = 32
    ElseIf gameType = "hard" Then
        marblesLeft = 44
    End If
    StatusBar1.Panels(2).Text = CStr(marblesLeft)
    shpBorder.FillColor = RGB(128, 128, 255)
    someoneWon = False
                            
    blankPath = App.Path & "\graphics\SolitaireBlank.bmp"
    filledPath = App.Path & "\graphics\solitaireFilled1.bmp"
    anipath = App.Path & "\graphics\solitaireFilled"
    clearPath = App.Path & "\graphics\SolitaireClear.bmp"
                            
    cellNo = 0
    For y = 1 To Gridlength
        For x = 1 To Gridlength
            board(x, y).filled = False
            board(x, y).spaceNo = cellNo
            spaceXY(cellNo).x = x
            spaceXY(cellNo).y = y
            cellNo = cellNo + 1
        Next x
    Next y
    
    If gameType = "easy" Then
    
        For x = 1 To 9
            board(x, 1).filled = True
            board(x, 9).filled = True
            space(board(x, 1).spaceNo).Picture = LoadPicture(clearPath)
            
            space(board(x, 9).spaceNo).Picture = LoadPicture(clearPath)
            space(board(x, 1).spaceNo).Visible = False
            space(board(x, 9).spaceNo).Visible = False
        Next x
        
        For y = 2 To 8
            board(1, y).filled = True
            board(9, y).filled = True
            space(board(1, y).spaceNo).Picture = LoadPicture(clearPath)

            space(board(9, y).spaceNo).Picture = LoadPicture(clearPath)
            space(board(1, y).spaceNo).Visible = False
            space(board(9, y).spaceNo).Visible = False
        Next y

    End If
    
    Select Case gameType
        Case "easy"
            bTop = 2
            bMiddle1 = 2
            bMiddle2 = 8
            bBottom1 = 6
            bBottom2 = 8
        Case "hard"
            bTop = 1
            bMiddle1 = 1
            bMiddle2 = 9
            bBottom1 = 6
            bBottom2 = 9
    End Select
    
    space(board(5, 5).spaceNo).Picture = LoadPicture(blankPath)
    space(board(5, 5).spaceNo).Visible = True
    For x = 4 To 6
        For y = bTop To 3
            board(x, y).filled = True
            space(board(x, y).spaceNo).Picture = LoadPicture(currentMarble)
            space(board(x, y).spaceNo).Visible = True
        Next y
    Next x
    For x = bMiddle1 To bMiddle2
        For y = 4 To 6
            If Not (x = 5 And y = 5) Then
                board(x, y).filled = True
                space(board(x, y).spaceNo).Picture = LoadPicture(currentMarble)
                space(board(x, y).spaceNo).Visible = True
            End If
        Next y
    Next x
    For x = 4 To 6
        For y = bBottom1 To bBottom2
            board(x, y).filled = True
            space(board(x, y).spaceNo).Picture = LoadPicture(currentMarble)
            space(board(x, y).spaceNo).Visible = True
        Next y
    Next x
    space(board(5, 5).spaceNo).Picture = LoadPicture(blankPath)
    space(board(5, 5).spaceNo).Visible = True
     
End Sub

Private Sub space_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
Dim validMove As Boolean
Dim newX As Integer
Dim newY As Integer
Dim oldX As Integer
Dim oldY As Integer
Dim goUp, goDown, goLeft, goRight As Boolean

    newX = spaceXY(Index).x
    newY = spaceXY(Index).y
    oldX = spaceXY(Source.Index).x
    oldY = spaceXY(Source.Index).y
    
    If (newX = oldX And newY = oldY - 2 And newY < 9) Then
        If board(newX, newY + 1).filled = True Then
            goUp = True
        End If
    ElseIf (newX = oldX And newY = oldY + 2 And newY > 1) Then
        If board(newX, newY - 1).filled = True Then
            goDown = True
        End If
    ElseIf (newY = oldY And newX = oldX - 2 And newX < 9) Then
        If board(newX + 1, newY).filled = True Then
            goLeft = True
        End If
    ElseIf (newY = oldY And newX = oldX + 2 And newX > 1) Then
        If board(newX - 1, newY).filled = True Then
            goRight = True
        End If
    End If
    
    validMove = (board(newX, newY).filled = False) And _
                    (goUp Or goDown Or goLeft Or goRight)
    If validMove Then
        turnsTaken = turnsTaken + 1
        ReDim Preserve history(turnsTaken)
        history(turnsTaken).oldX = oldX
        history(turnsTaken).oldY = oldY
        history(turnsTaken).newX = newX
        history(turnsTaken).newY = newY
        marblesLeft = marblesLeft - 1
        StatusBar1.Panels(2).Text = CStr(marblesLeft)
        board(newX, newY).filled = True
        board(oldX, oldY).filled = False
        
        space(board(newX, newY).spaceNo).Picture = LoadPicture(currentMarble)
        space(board(oldX, oldY).spaceNo).Picture = LoadPicture(blankPath)
        If goUp Then
            board(newX, newY + 1).filled = False
            animate newX, newY + 1, "forward"
            history(turnsTaken).victimX = newX
            history(turnsTaken).victimY = newY + 1
        ElseIf goDown Then
            board(newX, newY - 1).filled = False
            animate newX, newY - 1, "forward"
            history(turnsTaken).victimX = newX
            history(turnsTaken).victimY = newY - 1
        ElseIf goLeft Then
            board(newX + 1, newY).filled = False
            animate newX + 1, newY, "forward"
            history(turnsTaken).victimX = newX + 1
            history(turnsTaken).victimY = newY
        ElseIf goRight Then
            board(newX - 1, newY).filled = False
            animate newX - 1, newY, "forward"
            history(turnsTaken).victimX = newX - 1
            history(turnsTaken).victimY = newY
        End If
        over = checkForGameOver
        displayMessage
    End If
End Sub

Private Function checkForGameOver() As Boolean
Dim x, y As Integer
Dim iter As Integer

    For iter = 0 To 80
        x = spaceXY(iter).x
        y = spaceXY(iter).y
        If space(iter).Visible = True And board(x, y).filled = True Then
            If y - 2 > 0 Then
                If board(x, y - 1).filled = True And _
                    (board(x, y - 2).filled = False) And _
                        space(board(x, y - 2).spaceNo).Visible = True Then
                    Exit Function
                End If
            End If
            If x - 2 > 0 Then
                If board(x - 1, y).filled = True And _
                    (board(x - 2, y).filled = False) And _
                        space(board(x - 2, y).spaceNo).Visible = True Then
                    Exit Function
                End If
            End If
            If x + 2 < 10 Then
                If board(x + 1, y).filled = True And _
                    (board(x + 2, y).filled = False) And _
                        space(board(x + 2, y).spaceNo).Visible = True Then
                    Exit Function
                End If
            End If
            If y + 2 < 10 Then
                If board(x, y + 1).filled = True And _
                    (board(x, y + 2).filled = False) And _
                        space(board(x, y + 2).spaceNo).Visible = True Then
                    Exit Function
                End If
            End If
        End If
    Next iter
    checkForGameOver = True
End Function

Private Sub animate(xCo As Integer, yCo As Integer, direct As String)
Dim iter As Integer
Dim wait As Long

    If direct = "forward" Then
        For iter = 1 To 7
            space(board(xCo, yCo).spaceNo).Picture = _
                        LoadPicture(anipath & CStr(iter) & ".bmp")
            For wait = 1 To animateSpeed
            Next
        Next iter
        space(board(xCo, yCo).spaceNo).Picture = LoadPicture(blankPath)
    Else
        For iter = 7 To 1 Step -1
            space(board(xCo, yCo).spaceNo).Picture = _
                        LoadPicture(anipath & CStr(iter) & ".bmp")
            For wait = 1 To animateSpeed
            Next
        Next iter
        space(board(xCo, yCo).spaceNo).Picture = LoadPicture(currentMarble)
     End If
End Sub

Private Sub undoMove()
Dim iter As Integer
Dim x, y As Integer

    If over Then
        For x = 4 To 6
            For y = bTop To 3
                space(board(x, y).spaceNo).Visible = True
                board(x, y).filled = False
            Next y
        Next x
        For x = bMiddle1 To bMiddle2
            For y = 4 To 6
                    space(board(x, y).spaceNo).Visible = True
                    board(x, y).filled = False
            Next y
        Next x
        For x = 4 To 6
            For y = bBottom1 To bBottom2
                space(board(x, y).spaceNo).Visible = True
                board(x, y).filled = False
            Next y
        Next x
        If oldCount > 0 Then
            For iter = 1 To UBound(oldPos)
                space(oldPos(iter).spaceNo).Left = oldPos(iter).x
                space(oldPos(iter).spaceNo).Top = oldPos(iter).y
                'board(oldPos(iter).x, oldPos(iter).y).filled = True
                board(spaceXY(oldPos(iter).spaceNo).x, spaceXY(oldPos(iter).spaceNo).y).filled = True
            Next iter
        End If
    End If

    frmSolitaire.BackColor = &HFFC0C0
    shpEdge.FillColor = &HC0C000
    space(40).Top = 3960
    space(40).Height = 495
    space(40).Width = 495
    space(40).Left = 3120
    space(40).Refresh
    frmSolitaire.Refresh

    over = False
    Timer1.Enabled = False
    If turnsTaken > 0 Then
        board(history(turnsTaken).newX, history(turnsTaken).newY).filled = False
        board(history(turnsTaken).oldX, history(turnsTaken).oldY).filled = True
        board(history(turnsTaken).victimX, history(turnsTaken).victimY).filled = True
        space(board(history(turnsTaken).newX, history(turnsTaken).newY).spaceNo).Picture = LoadPicture(blankPath)
        space(board(history(turnsTaken).oldX, history(turnsTaken).oldY).spaceNo).Picture = LoadPicture(currentMarble)
        animate history(turnsTaken).victimX, history(turnsTaken).victimY, "back"
        space(board(history(turnsTaken).victimX, history(turnsTaken).victimY).spaceNo).Picture = LoadPicture(currentMarble)
        turnsTaken = turnsTaken - 1
        ReDim Preserve history(turnsTaken)
        marblesLeft = marblesLeft + 1
        StatusBar1.Panels(2).Text = CStr(marblesLeft)
        displayMessage
    End If
End Sub

Private Sub displayMessage()
    If over And marblesLeft > 1 Then
        StatusBar1.Panels(1).Text = " Better luck next time !"
        PrepareTimer
        Timer1.Enabled = True
    ElseIf marblesLeft = 1 Then
        over = True
        StatusBar1.Panels(1).Text = " You have proved me wrong, wise one."
        winnerPic
    ElseIf marblesLeft <= 5 Then
        StatusBar1.Panels(1).Text = " Surely not..."
    ElseIf marblesLeft <= 10 Then
        StatusBar1.Panels(1).Text = " The tension is mounting !"
    ElseIf marblesLeft > 10 Then
        StatusBar1.Panels(1).Text = " You'll never do it !!"
    End If
End Sub

Private Sub winnerPic()
Dim crazyLoop As Integer

    space(40).Picture = LoadPicture(App.Path & "\graphics\winnerpic.bmp")
    For crazyLoop = 1 To 770
        space(40).Width = space(40).Width + 9
        space(40).Left = space(40).Left - 4
        space(40).Height = space(40).Height + 8
        space(40).Top = space(40).Top - 4
        space(40).Refresh
    Next crazyLoop
End Sub

Private Sub Timer1_Timer()
Dim iter As Integer
Dim lft As Double
Dim tp As Double

    For iter = 1 To UBound(oldPos)
        lft = Abs(3200 - space(oldPos(iter).spaceNo).Left) + 1
        tp = Abs(3400 - space(oldPos(iter).spaceNo).Top) + 1
        If Sqr((lft ^ 2) + (tp ^ 2)) >= 2400 Then
            oldPos(iter).directionLeft = oldPos(iter).directionLeft * -1
            oldPos(iter).directionRight = oldPos(iter).directionRight * -1
        End If
        space(oldPos(iter).spaceNo).Left = _
                        space(oldPos(iter).spaceNo).Left + _
                            oldPos(iter).speedLeft * oldPos(iter).directionLeft
        space(oldPos(iter).spaceNo).Top = _
                        space(oldPos(iter).spaceNo).Top + _
                            oldPos(iter).speedTop * oldPos(iter).directionRight
    Next iter
End Sub

Private Sub PrepareTimer()
Dim iter As Integer
Dim temp As Integer

    Randomize
    'oldCount = 0
    For iter = 0 To 80
        temp = (Rnd * 4) + 1
        If board(spaceXY(iter).x, spaceXY(iter).y).filled = False Then
            space(iter).Visible = False
        Else
            oldCount = oldCount + 1
            ReDim Preserve oldPos(oldCount)
            oldPos(oldCount).x = space(iter).Left
            oldPos(oldCount).y = space(iter).Top
            oldPos(oldCount).spaceNo = iter
            oldPos(oldCount).speedLeft = (Rnd * 40) + 1
            oldPos(oldCount).speedTop = (Rnd * 40) + 1
            If temp = 1 Then
                oldPos(oldCount).directionLeft = -1
                oldPos(oldCount).directionRight = -1
            ElseIf temp = 2 Then
                oldPos(oldCount).directionLeft = 1
                oldPos(oldCount).directionRight = 1
            ElseIf temp = 3 Then
                oldPos(oldCount).directionLeft = -1
                oldPos(oldCount).directionRight = 1
            ElseIf temp = 4 Then
                oldPos(oldCount).directionLeft = 1
                oldPos(oldCount).directionRight = -1
            Else
                oldPos(oldCount).directionLeft = 1
                oldPos(oldCount).directionRight = 1
            End If
        End If
    Next iter
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: Timer1.Enabled = False
                Form_Load
        Case 2: undoMove
        Case 3: frmColour.Show vbModal
                changeColour
        Case 4: changeGame
        Case 5: frmAbout.Show vbModal
        Case 6: Unload Me
    End Select
End Sub

Private Sub changeGame()
    Select Case gameType
        Case "easy"
            If marblesLeft = 32 Then
                changeLevel ("toHard")
            Else
                MsgBox "Not once you've started !", vbExclamation
            End If
        Case "hard"
            If marblesLeft = 44 Then
                Call changeLevel("toEasy")
            Else
                MsgBox "Not once you've started !", vbExclamation
            End If
    End Select
End Sub

Private Sub changeLevel(ByVal whichWay As String)
    Dim x As Integer, y As Integer
    
    Select Case whichWay
        Case "toHard"
                bTop = 1
                bMiddle1 = 1
                bMiddle2 = 9
                bBottom1 = 6
                bBottom2 = 9
            For x = 4 To 6
                board(x, 1).filled = True
                board(x, 9).filled = True
                space(board(x, 1).spaceNo).Picture = LoadPicture(currentMarble)
                space(board(x, 9).spaceNo).Picture = LoadPicture(currentMarble)
                space(board(x, 1).spaceNo).Visible = True
                space(board(x, 9).spaceNo).Visible = True

            Next x

            For y = 4 To 6
                board(1, y).filled = True
                board(9, y).filled = True
                space(board(1, y).spaceNo).Picture = LoadPicture(currentMarble)
                space(board(9, y).spaceNo).Picture = LoadPicture(currentMarble)
                space(board(1, y).spaceNo).Visible = True
                space(board(9, y).spaceNo).Visible = True
            Next y
            marblesLeft = 44
            gameType = "hard"

        Case "toEasy"
            bTop = 2
            bMiddle1 = 2
            bMiddle2 = 8
            bBottom1 = 6
            bBottom2 = 8
            For x = 1 To 9
                board(x, 1).filled = True
                board(x, 9).filled = True
                space(board(x, 1).spaceNo).Picture = LoadPicture(clearPath)
                space(board(x, 9).spaceNo).Picture = LoadPicture(clearPath)
                space(board(x, 1).spaceNo).Visible = False
                space(board(x, 9).spaceNo).Visible = False
            Next x

            For y = 2 To 8
                board(1, y).filled = True
                board(9, y).filled = True
                space(board(1, y).spaceNo).Picture = LoadPicture(clearPath)
                space(board(9, y).spaceNo).Picture = LoadPicture(clearPath)
                space(board(1, y).spaceNo).Visible = False
                space(board(9, y).spaceNo).Visible = False
            Next y
            marblesLeft = 32
            gameType = "easy"
    End Select
End Sub

Private Sub changeColour()
Dim xPos, yPos As Integer
Dim iter As Integer
    
    If choice <> "" Then
        currentMarble = picPath & _
                                choice & ".bmp"
        For xPos = 1 To 9
            For yPos = 1 To 9
                If board(xPos, yPos).filled = True And _
                        space(board(xPos, yPos).spaceNo).Visible = True Then
                    space(board(xPos, yPos).spaceNo).Picture = LoadPicture _
                                                    (picPath & choice & ".bmp")
                End If
            Next yPos
        Next xPos
    End If
End Sub

Private Sub Form_Resize()
    If frmSolitaire.WindowState = 2 Or frmSolitaire.WindowState = 0 Then
        frmSolitaire.Icon = LoadPicture(App.Path & "\graphics\misc05.ico")
        frmSolitaire.Caption = "Solitaire"
    Else
        frmSolitaire.Icon = LoadPicture(App.Path & "\graphics\graph07.ico")
        frmSolitaire.Caption = "Spread 1"
    End If
End Sub

Private Sub Command1_Click()
    frmSolitaire.WindowState = 1
    frmSolitaire.Icon = LoadPicture(App.Path & "\graphics\graph07.ico")
    frmSolitaire.Caption = "Spread 1"
End Sub
