VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckFtp 
      Left            =   2640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "request"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "txtfile"
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "disconnect"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "connect"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tFtpHeader
    headerlength As Integer
    type As Integer
    filesize As Long
    adbannerid As Long
    adfileextension As Long
    filesft As FILETIME
    filename As String
    '
    '
    headpos As Integer
    headerfull As Boolean
    filebuffer As String
End Type
Private CurrFile As tFtpHeader

Private ftpheadcount As Long
Private ftpgothead As Boolean
Private ftpfilebuffer As String

Private bsecondtime As Boolean

Private Sub AddText(ByVal strMsg As String)
    Text1.Text = Text1.Text & strMsg & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Command1_Click()
    If sckFtp.State <> sckConnected Then
        Call AddText("Attempting to connect.")
        sckFtp.Connect
    Else
        Call AddText("some fucked up error try the disconnect buton.")
    End If
End Sub

Private Sub Command2_Click()
    If sckFtp.State <> sckClosed Then
        Call sckFtp.Close
        Call AddText("You have closed the ftp socket.")
        bsecondtime = False
    End If
End Sub

Private Sub Command3_Click()
    If sckFtp.State = sckClosed Then
        Call AddText("Connect the socket before requesting a file.")
        Exit Sub
    End If
    
    'are we currently downloading
    If sckFtp.State = sckConnected Then
        'are we finnished with the file?
        If CurrFile.filesize = Len(CurrFile.filebuffer) Then
            'file is complete
            'if this isnt our first send dont send the x2
            If bsecondtime Then
            Else
                bsecondtime = True
                sckFtp.SendData Chr(2)
            End If
        Else
            'were mid transfer
            Exit Sub
        End If
    End If
    
    'reset the current file data
    CurrFile.headerlength = 0
    CurrFile.type = 0
    CurrFile.filesize = 0
    CurrFile.adbannerid = 0
    CurrFile.adfileextension = 0
    CurrFile.filesft.dwHighDateTime = 0
    CurrFile.filesft.dwLowDateTime = 0
    CurrFile.filename = ""
    '
    '
    CurrFile.headpos = 0
    CurrFile.headerfull = False
    CurrFile.filebuffer = ""
    
    'build our packet for sending the request
    Dim packetdata As String
    Dim packetbody As String
    
    packetbody = MakeWORD(&H100)
    packetbody = packetbody & "68XI"
    packetbody = packetbody & "LTRD"
    packetbody = packetbody & MakeDWORD(0)
    packetbody = packetbody & MakeDWORD(0)
    packetbody = packetbody & MakeDWORD(0)
    'ft
    packetbody = packetbody & MakeDWORD(0)
    packetbody = packetbody & MakeDWORD(0)
    'fname
    packetbody = packetbody & txtfile.Text & Chr(0)
    
    packetdata = packetdata & MakeWORD(Len(packetbody) + 2) & packetbody
    
    sckFtp.SendData packetdata
    
End Sub

Private Sub Form_Load()
    bsecondtime = False

    sckFtp.RemoteHost = "vultr-chi.bnetdocs.org"
    sckFtp.RemotePort = 6112
    
    ftpheadcount = 0
    ftpgothead = False
    
    Text1.Text = ""
    txtfile.Text = ""
End Sub

Private Sub sckFtp_Close()
    Call AddText("Disconnected.")
    bsecondtime = False
End Sub

Private Sub sckFtp_Connect()
    Call AddText("Connected, Ready for file requests.")
End Sub

Private Sub sckFtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call AddText("Something fucked up.")
End Sub

Private Sub sckFtp_DataArrival(ByVal bytesTotal As Long)
    sckFtp.GetData inBuf, vbString
    CurrFile.filebuffer = CurrFile.filebuffer & inBuf
    
    Do While CurrFile.filebuffer <> ""
        Select Case CurrFile.headpos
            Case 0
                If Len(CurrFile.filebuffer) < 2 Then
                    Exit Sub
                End If
                CurrFile.headerlength = MakeShort(Mid(CurrFile.filebuffer, 1, 2))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 3)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader length: " & CurrFile.headerlength)

            Case 1
                If Len(CurrFile.filebuffer) < 2 Then
                    Exit Sub
                End If
                CurrFile.type = MakeShort(Mid(CurrFile.filebuffer, 1, 2))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 3)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader type: " & CurrFile.type)

            Case 2
                If Len(CurrFile.filebuffer) < 4 Then
                    Exit Sub
                End If
                CurrFile.filesize = MakeLong(Mid(CurrFile.filebuffer, 1, 4))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 5)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader filesize: " & CurrFile.filesize)
                
            Case 3
                If Len(CurrFile.filebuffer) < 4 Then
                    Exit Sub
                End If
                CurrFile.adbannerid = MakeLong(Mid(CurrFile.filebuffer, 1, 4))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 5)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader id: " & CurrFile.adbannerid)
                
            Case 4
                If Len(CurrFile.filebuffer) < 4 Then
                    Exit Sub
                End If
                CurrFile.adfileextension = MakeLong(Mid(CurrFile.filebuffer, 1, 4))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 5)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader ext: " & CurrFile.adfileextension)
                
            Case 5
                If Len(CurrFile.filebuffer) < 8 Then
                    Exit Sub
                End If
                CurrFile.filesft = MakeFileTime(Mid(CurrFile.filebuffer, 1, 8))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, 9)
                CurrFile.headpos = CurrFile.headpos + 1
                
                Call AddText("fileheader ft h: " & CurrFile.filesft.dwHighDateTime)
                Call AddText("fileheader ft l: " & CurrFile.filesft.dwLowDateTime)
                
            Case 6
                If Len(CurrFile.filebuffer) < (CurrFile.headerlength - 24) Then
                    Exit Sub
                End If
                CurrFile.filename = Mid(CurrFile.filebuffer, 1, (CurrFile.headerlength - 24))
                CurrFile.filebuffer = Mid(CurrFile.filebuffer, (CurrFile.headerlength - 24) + 1)
                CurrFile.headpos = CurrFile.headpos + 1
                
                'remove the null
                CurrFile.filename = Mid(CurrFile.filename, 1, Len(CurrFile.filename) - 1)
                
                Call AddText("fileheader file: " & CurrFile.filename)
                
            Case 7
                'this is all file data from here on
                If Len(CurrFile.filebuffer) < CurrFile.filesize Then
                    '%% (v2 / v1) * 100.0 '[\ 1]=vb magic
                    Call AddText("Downloading: " & (((Len(CurrFile.filebuffer) / CurrFile.filesize) * 100) \ 1) & "%")
                    Exit Sub
                End If
                If Len(CurrFile.filebuffer) = CurrFile.filesize Then
                    Call AddText("Download: Complete.")
                    Exit Sub
                Else
                    'something fucked up here
                    Exit Sub
                End If
        End Select
    Loop
End Sub

