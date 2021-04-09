VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Client Application   (File Transfer using MS Winsock 6.0) "
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   252
      Left            =   4320
      TabIndex        =   10
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox txtView 
      Height          =   1332
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2400
      Width           =   5292
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   252
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   252
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   252
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   252
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5172
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   4440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   4800
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   252
      Left            =   4200
      TabIndex        =   0
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Browse for a file to transfer to the server. Connect, then send."
      ForeColor       =   &H00FFFFFF&
      Height          =   684
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Data View Port:"
      ForeColor       =   &H00FFFFFF&
      Height          =   204
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   972
      Left            =   120
      Top             =   840
      Width           =   3972
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to send:"
      ForeColor       =   &H00FFFFFF&
      Height          =   204
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status : Disconnected"
      Height          =   252
      Left            =   48
      TabIndex        =   1
      Top             =   4200
      Width           =   5292
   End
End
Attribute VB_Name = "frmClient"
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
    End
End Sub

Private Sub cmdConnect_Click()
        
    'try to make a connection to the Server.
    bReplied = False
    tcpClient.Connect "127.0.0.1", 1256
    
    lTIme = 0
    
    While (Not bReplied) And (lTIme < 100000)
        DoEvents
        lTIme = lTIme + 1
    Wend
    
    
    If lTIme >= 100000 Then
        'Didn't reply or timed out. close the connection
        MsgBox "Unable to connect to remote server", vbCritical, "Connection Error"
        
        tcpClient.Close
        Exit Sub
    End If
    
End Sub



Private Sub cmdDisconnect_Click()
    tcpClient.Close
    Form_Load
End Sub

Private Sub cmdSend_Click()
    Dim FName_Only As String
    
    If txtFileName = "" Then
       MsgBox "No file selected to send...", vbCritical
    Else ' send the file, if connected
       If tcpClient.State <> sckClosed Then
          ' send only the file name because it will
          ' be stored in another area than the source
          FName_Only$ = GetFileName(txtFileName)
          SendFile FName_Only$
       End If
    End If
End Sub

Private Sub Form_Load()
    Status "Disconnected."
    bReplied = False
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   G E N E R A L   W I N S O C K   P R O C W D U R E S
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub tcpClient_Close()
    '
    'Socket got a close call so close it if it's not already closed
    If tcpClient.State <> sckClosed Then tcpClient.Close
    
End Sub



Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    '
    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String
    Static DataCnt   As Long
    
    tcpClient.GetData NewArrival$, vbString
    
    ' Extract the command from the Left
    ' of the comma (default divider)
    Command$ = EvalData(NewArrival$, 1)
    ' extract the data being sent from the
    ' right of the comma (default divider)
    Data$ = EvalData(NewArrival$, 2)
    
    ' execute according to command sent
    Select Case Command
        Case "Accepted"          ' server accepted connection
             bReplied = True
             Status "Connected."
             
             ' this is a good practice.
             ' when the server has been closed
             ' theclient is notified here.
             ' and immediatley disconnected.
        Case "ServerClosed"
             Form_Load
             tcpClient.Close
             
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
           Status "Connected."
                
       ' when sending a file.... it is best not to Name
       ' the Case instead use ELse for file transfer
            
        Case Else
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
       txtFileName = cdOpen.FileName
    End If
    
End Sub
