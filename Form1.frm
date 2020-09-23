VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Port Scanner"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtLimit 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "32000"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtTimeout 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   2280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox chkBeep 
      Caption         =   " Beep For Port"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Host To Scan:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Stop Scan On"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Timeout"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Open Ports"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Scanning"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblOpenPorts 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblScannedPorts 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port As Integer

'Use FREELY!


Private Sub cmdScan_Click()
If cmdScan.Caption = "Stop" Then
Timer1.Enabled = False
cmdScan.Caption = "Scan"
Exit Sub
ElseIf cmdScan.Caption = "Scan" Then
Port = 0
lblOpenPorts = "0"
lblScannedPorts = "0"
List1.Clear
Timer1.Interval = txtTimeout
cmdScan.Caption = "Stop"
Timer1.Enabled = True
End If
End Sub

Private Sub Form_Load()
Port = 0
txtIP = sck.LocalIP
End Sub

Private Sub Timer1_Timer()
lblScannedPorts = Port
If PortNumber = Not txtLimit Then
Timer1.Enabled = False
cmdScan.Caption = "Scan"
Exit Sub
ElseIf sck.State = 7 Then
List1.AddItem sck.RemotePort
lblOpenPorts = lblOpenPorts + 1
If chkBeep = 1 Then Beep
sck.Close: Port = Port + 1
sck.Connect txtIP.Text, Port
Exit Sub
ElseIf Not sck.State Then sck.Close
sck.Connect txtIP.Text, Port
Port = Port + 1
End If
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdScan_Click
End Sub
