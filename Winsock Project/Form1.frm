VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Winsock Example"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Text            =   "None"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "None"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh IP"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close Connection"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Host"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Port In Use:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "IP You Are Connected To:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Your IP:"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Hosting As Boolean
Dim CloseConnection As Boolean
Dim ConnectionCheck As String

Private Sub Command1_Click()

Host = InputBox("Enter the host's computer name or ip address:")
Port = InputBox("Enter the host's port to connect to:")

If Host = "" Or Port = "" Then
    
    MsgBox ("Error Establishing Connection; Port or Host IP Must Be Entered Correctly")
    GoTo SubEnd
    
End If

On Error GoTo Error
Winsock1.Connect Host, Port
Text3.Text = Port
Text2.Text = Host
GoTo SubEnd

Error:
    MsgBox ("Uknown Error; Could Not Establish The Connection")

SubEnd:
End Sub

Private Sub Command2_Click()

Port = InputBox("What port do you want to host on?")

If Port = "" Then
    MsgBox ("Please Enter A Port That Can Be Connected To")
    GoTo SubEnd
End If

On Error GoTo Error
Winsock1.LocalPort = Port
Winsock1.Listen
Text3.Text = Port
Hosting = True
GoTo SubEnd

Error:
    MsgBox ("Error Listening To Port: " & Port & "; Make Sure You Are Connected To The Internet")

SubEnd:
End Sub

Private Sub Command3_Click()

Text = InputBox("Enter Text To Send")
On Error GoTo Error
Winsock1.SendData Text

Error:
    MsgBox ("Error Sending Data; Make Sure You Are Connected")

End Sub

Private Sub Command4_Click()

CloseConnection = True
Text = "Close Connection"
On Error GoTo Error
Winsock1.SendData Text
Winsock1.Close

Error:
    MsgBox ("Error Closing Connection; Make Sure Connection Isn't Already Closed")
    
End Sub

Private Sub Command5_Click()

Text1.Text = Winsock1.LocalIP
MsgBox ("IP Refreshed")

End Sub

Private Sub Form_Load()

ConnectionCheck = "Check?"
CloseConnection = False
Text1.Text = Winsock1.LocalIP

End Sub

Private Sub Winsock1_Close()
   
MsgBox ("Connection Has Been Closed")

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

If Winsock1.State <> sckClosed Then
    MsgBox ("Error With Connection; Winsock Control Already In Use")
    Winsock1.Close
End If

Winsock1.Accept requestID
Winsock1.SendData ConnectionCheck
Text2.Text = Winsock1.RemoteHostIP

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Winsock1.GetData Text, vbString, 10000

If Text = "Check?" Then
    Text = "Check!"
    Winsock1.SendData Text

ElseIf Text = "Check!" Then
    MsgBox ("You Are Connected!")

ElseIf Text = "Close Connection" Then
    CloseConnection = True
    If Hosting = True Then
        MsgBox ("User Has Closed The Connection")
        Winsock1.Close
    Else
        MsgBox ("Host Has Closed The Connection")
        Winsock1.Close
    End If
    
Else
    MsgBox (Text)

End If

End Sub

