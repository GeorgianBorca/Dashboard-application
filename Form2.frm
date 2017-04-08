VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   15360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   LinkTopic       =   "Form2"
   ScaleHeight     =   93584.95
   ScaleMode       =   0  'User
   ScaleWidth      =   5251.282
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Text            =   "Form2.frx":0000
      Top             =   3360
      Width           =   11055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Send"
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
      Left            =   5400
      TabIndex        =   16
      ToolTipText     =   "send"
      Top             =   2280
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command8 
      Caption         =   "\r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "\n\r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "MACROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   15600
      TabIndex        =   11
      Top             =   1800
      Width           =   2775
      Begin VB.CommandButton Command6 
         Caption         =   "\n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "CTRL+Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
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
      Left            =   13560
      TabIndex        =   10
      ToolTipText     =   "Clear text area"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4920
      TabIndex        =   4
      Text            =   "9600"
      ToolTipText     =   "Baud rate"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
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
      Left            =   9360
      TabIndex        =   3
      ToolTipText     =   "CONNECT"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scan"
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
      Left            =   7560
      TabIndex        =   2
      ToolTipText     =   "Re-scan for ports"
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "chose one port"
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11400
      TabIndex        =   0
      ToolTipText     =   "Back"
      Top             =   480
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Availabale Ports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "BAUD RATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7440
      TabIndex        =   7
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton Command13 
         Caption         =   "BTN13"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "BTN12"
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
         TabIndex        =   21
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   20
         ToolTipText     =   "DISCONNECT"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "DTR"
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
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Text control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Enable double click to select all text"
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   7095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "Form2.frx":0013
         ToolTipText     =   "text to send"
         Top             =   480
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const MAX_NUM = 25  'Maximum port number

Dim selectAll As Boolean

Public Function COMAvailable(port As Integer) As Boolean

If MSComm1.CommPort = port Then
    COMAvailable = True
    Exit Function
Else
    COMAvailable = False
    Exit Function
End If

If MSComm1.PortOpen Then
    MSComm1.PortOpen = False
End If
     
End Function

Private Sub Check1_Click()

If Check1.Value = 1 Then
    selectAll = True
Else
    selectAll = False
End If

End Sub

Private Sub Command1_Click()

If MSComm1.PortOpen Then
    MSComm1.PortOpen = False
End If

Unload Me

End Sub

Private Sub Command10_Click()

If MSComm1.DTREnable Then
    MSComm1.DTREnable = False
Else
    MSComm1.DTREnable = True
End If

End Sub

Private Sub Command11_Click()
If MSComm1.PortOpen Then
    MSComm1.PortOpen = False
End If
Command11.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command2_Click() 'SCAN
Combo1.Clear
Dim i As Integer

For i = 1 To MAX_NUM
    On Error Resume Next
    MSComm1.CommPort = i
    On Error Resume Next
    MSComm1.PortOpen = True
    On Error Resume Next
    MSComm1.PortOpen = False
    If Err.Number = 0 And i <> 3 Then
        Combo1.AddItem "COM" & i
    End If
Next i



End Sub

Private Sub Command3_Click() 'Connect

Dim comPort As Integer

If Not Combo1.Text = "" Then
    comPort = Ret(Combo1.Text)
    
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If

    MSComm1.CommPort = comPort
    MSComm1.CommPort = comPort
    MSComm1.Settings = Combo2.Text & ",N,8,1"
    MSComm1.PortOpen = True

    If MSComm1.PortOpen Then
        Command3.Enabled = False
        Command11.Enabled = True
    End If

Else
    MsgBox "Please seect a COM port first!" & vbCrLf & "COMx", vbInformation, "Error"
End If

Text2.Text = ""

End Sub

Public Function Ret(ByVal sString As String) As String
Dim i As Integer
For i = 1 To Len(sString)
    If Mid(sString, i, 1) Like "[0-9]" Then
        Ret = Ret + Mid(sString, i, 1)
    End If
Next i

End Function

Private Sub Command4_Click()
Text2.Text = ""
End Sub

Private Sub Command9_Click() 'Send
Text2.Text = Text2.Text & Text1.Text & vbCrLf
Text2.SelStart = Len(Text2.Text)
Text1.Text = ""
End Sub

Private Sub Form_Load()

Text2.SelStart = Len(Text2.Text)

MSComm1.RTSEnable = False
MSComm1.DTREnable = True
MSComm1.RThreshold = 1

Command3.Enabled = True
Command11.Enabled = False

'Add baud rates to combo2
Combo2.AddItem "250000"
Combo2.AddItem "230400"
Combo2.AddItem "115200"
Combo2.AddItem "74800"
Combo2.AddItem "57600"
Combo2.AddItem "38400"
Combo2.AddItem "19200"
Combo2.AddItem "9600"
Combo2.AddItem "4800"
Combo2.AddItem "2400"
Combo2.AddItem "1200"
Combo2.AddItem "300"

'Scan the available serial ports
Combo1.Clear
Dim i As Integer

For i = 1 To MAX_NUM
    On Error Resume Next
    MSComm1.CommPort = i
    On Error Resume Next
    MSComm1.PortOpen = True
    On Error Resume Next
    MSComm1.PortOpen = False
    If Err.Number = 0 And i <> 3 Then
        Combo1.AddItem "COM" & i
    End If
Next i

Combo1.ListIndex = Combo1.ListCount - 1



End Sub

Private Sub MSComm1_OnComm()
Dim buffer As String
buffer = MSComm1.Input
If MSComm1.CommEvent = comEvReceive Then
    Text2.Text = Text2.Text & buffer
    Text2.SelStart = Len(Text2.Text)
End If

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shifts As Integer)
If KeyCode = 13 Then
    'MSComm1.Output = Text1.Text & vbCrLf
    Text2.Text = Text2.Text & Text1.Text & vbCrLf
    Text2.SelStart = Len(Text2.Text)
    Text1.Text = ""
End If

End Sub

Private Sub Text2_DblClick()
If selectAll Then
    Dim txt As TextBox
    If KeyAscii = ASC_CTRL_A Then
        If TypeOf ActiveControl Is TextBox Then
            Set txt = ActiveControl
            txt.SelStart = 0
            txt.SelLength = Len(txt.Text)
        End If
    End If
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Const ASC_CTRL_A As Integer = 1
Dim txt As TextBox
If KeyAscii = ASC_CTRL_A Then
    If TypeOf ActiveControl Is TextBox Then
        Set txt = ActiveControl
        txt.SelStart = 0
        txt.SelLength = Len(txt.Text)
    End If
End If
End Sub
