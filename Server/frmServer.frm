VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Application   (File Transfer using MS Winsock 6.0) "
   ClientHeight    =   4500
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   252
      Left            =   4320
      TabIndex        =   8
      Top             =   3840
      Width           =   972
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   252
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   252
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   5172
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   4800
      Top             =   600
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   4800
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   252
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   972
   End
   Begin VB.TextBox txtView 
      Height          =   1332
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   5292
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Browse for a file to transfer to the Client. Make sure you are connected , then send."
      ForeColor       =   &H00000000&
      Height          =   684
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   3480
   End
   Begin VB.Shape Shape1 
      Height          =   972
      Left            =   120
      Top             =   840
      Width           =   3972
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to send:"
      Height          =   204
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status : Listening......."
      Height          =   252
      Left            =   48
      TabIndex        =   3
      Top             =   4200
      Width           =   5292
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Data View Port:"
      Height          =   204
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1860
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtView = ""
    txtFileName = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim FName_Only As String
    
    If txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If frmServer.tcpServer.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(txtFileName)
          SendFile FName_Only$
       End If
    End If
End Sub

Private Sub Form_Load()
    ' connect to the port
    tcpServer.LocalPort = Port
    ' Listen for incoming data
    tcpServer.Listen
    
    bInconnection = False
    
    Status "Listening.... (Not Connected)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' alert the client the server has been disconnected
    SendData "ServerClosed,"
    Pause 500
    tcpServer.Close
    End
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   G E N E R A L   W I N S O C K   P R O C W D U R E S
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub tcpServer_Close()
    '
    'Socket got a close call so close it if it's not already closed
    If tcpServer.State <> sckClosed Then tcpServer.Close
    Form_Load      ' resume listening
    
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
    '
     On Error GoTo IDERROR
     If tcpServer.State <> sckClosed Then tcpServer.Close ' close Connection
     tcpServer.Accept requestID    'Make the connection
     
     bInconnection = True
     Status "Listening... Connected."
     SendData "Accepted,"
     Exit Sub
     
IDERROR:
     MsgBox Err.Description, vbCritical
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
    '
    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String
    Static DataCnt   As Long
    
    tcpServer.GetData NewArrival$, vbString
    
    ' Extract the command from the Left
    ' of the comma (default divider)
    Command = EvalData(NewArrival$, 1)
    ' extract the data being sent from the
    ' right of the comma (default divider)
    Data$ = EvalData(NewArrival$, 2)
    
    ' execute according to command sent
    Select Case Command$
                  
        Case "OpenFile"  ' open the file
           Dim Fname As String
           
           ' the file name only should've been sent
           Fname$ = App.Path & "\" & Data$
           Open Fname$ For Binary As #1
           ' file now opened to recieve input
           Status "File opened.... " & Data$
               
        Case "CloseFile" ' close the file
           ' all data has been sent, close the file
           Close #1
           Status "File Transfer complete..."
           Pause 3000
           Status "Listening... (Connected)"
           
       ' when sending a file.... it is best not to Name
       ' the Case instead use ELse for file transfer
        
        Case Else ' a 4169 byte string of incoming data
           ' write the incoming chunk of data to the
           ' opened file
           Put #1, , NewArrival$
           ' update the view port with the new addition
' ** // ** '
' IMPORTANT: comment out the code below when sending files
' larger than 500Kb. It makes the function CRAWL otherwise
              
           txtView = txtView & NewArrival$
' comment the above line to increase speed

           ' count and report the incoming chunks
           DataCnt& = DataCnt& + 1
           Status "Recieving Data... " & (MAX_CHUNK * DataCnt&) & " bytes"
              
    End Select
    
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'  end  G E N E R A L   W I N S O C K   P R O C W D U R E S  end
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\







Private Sub cmdBrowse_Click()
    ' show the Open Dialog for the user to select a file.
    cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtFileName = cdOpen.filename
    End If
    
End Sub
