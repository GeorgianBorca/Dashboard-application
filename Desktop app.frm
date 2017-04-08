VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   15360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   19466.14
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Reboot"
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
      Left            =   120
      TabIndex        =   12
      Top             =   13680
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Power OFF"
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
      Left            =   120
      TabIndex        =   11
      Top             =   12840
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   120
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   7200
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Term"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MPALB IPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pic Programming"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Serial Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   18840
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Serial tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "Putty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Programs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
         Begin VB.CommandButton Command7 
            Caption         =   "Hyper Terminal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   19.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   3375
         End
      End
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   14520
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command4_Click() ' Open MPLAB IPE jar file
Shell "C:\Program Files\java\jre1.7.0_79\bin\javaw.exe -jar ""C:\Program Files\Microchip\MPLABX\v3.55/mplab_ipe/ipe.jar"" "
End Sub

Private Sub Command7_Click() ' Open hyper terminal (windows app)
Shell "C:\Program Files\Windows NT\hypertrm.exe"
End Sub

Private Sub Command8_Click() 'power OFF
Call Shell("shutdown /s")
End Sub

Private Sub Command9_Click() 'reboot
Call Shell("shutdown /r")
End Sub

Private Sub Form_Load()
'Form1.Width = 1280 * Screen.TwipsPerPixelX
'Form1.Height = 1024 * Screen.TwipsPerPixelY
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub


