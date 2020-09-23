VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Telynet"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12300
   Begin VB.Timer tmrKeepAlive 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Log As"
      Filter          =   "Text Files (*.txt)|*.txt"
      Flags           =   2054
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuOptBlank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuLocalEcho 
         Caption         =   "&Local Echo"
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "&Font Size"
      End
      Begin VB.Menu mnuKeepAlive 
         Caption         =   "&Keep Alive"
      End
      Begin VB.Menu mnuOptBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetWinFont 
         Caption         =   "&Reset Window && Font"
      End
      Begin VB.Menu mnuNewWindow 
         Caption         =   "&New Window"
      End
   End
   Begin VB.Menu mnuMacros 
      Caption         =   "&Macros"
      Begin VB.Menu mnuEditMacros 
         Caption         =   "&Edit Macros"
      End
      Begin VB.Menu mnuMacBlank 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMacro 
         Caption         =   "X"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Telnet main commands  -  list copied from somewhere on the net - thank you to them
Const TN_IAC = 255   ' Interpret as command escape sequence, Prefix to all telnet commands. 1, 2 or sometimes more commands normally follow this character
Const TN_DONT = 254  ' You are not to use this option
Const TN_DO = 253    ' Please, you use this option
Const TN_WONT = 252  ' I won't use option
Const TN_WILL = 251  ' I will use option
Const TN_SB = 250    ' Subnegotiate, X number of commands follow
Const TN_GA = 249    ' Go ahead
Const TN_EL = 248    ' Erase line
Const TN_EC = 247    ' Erase character
Const TN_AYT = 246   ' Are you there
Const TN_AO = 245    ' Abort output
Const TN_IP = 244    ' Interrupt process
Const TN_BRK = 243   ' Break
Const TN_DM = 242    ' Data mark
Const TN_NOP = 241   ' No operation.
Const TN_SE = 240    ' End of subnegotiation, from above
Const TN_EOR = 239   ' End of record
Const TN_ABORT = 238 ' About process
Const TN_SUSP = 237  ' Suspend process
Const TO_EOF = 236  ' End of file

' Telnet (option) mainly return commands from above
Const TN_BIN = 0     ' Binary transmission
Const TN_ECHO = 1    ' Echo
Const TN_RECN = 2    ' Reconnection
Const TN_SUPP = 3    ' Suppress go ahead
Const TN_APRX = 4    ' Approx message size negotiation
Const TN_STAT = 5    ' Status
Const TN_TIM = 6     ' Timing mark
Const TN_REM = 7     ' Remote controlled trans/echo
Const TN_OLW = 8     ' Output line width
Const TN_OPS = 9     ' Output page size
Const TN_OCRD = 10   ' Out carriage-return disposition
Const TN_OHT = 11    ' Output horizontal tabstops
Const TN_OHTD = 12   ' Out horizontal tab disposition
Const TN_OFD = 13    ' Output formfeed disposition
Const TN_OVT = 14    ' Output vertical tabstops
Const TN_OVTD = 15   ' Output vertical tab disposition
Const TN_OLD = 16    ' Output linefeed disposition
Const TN_EXT = 17    ' Extended ascii character set
Const TN_LOGO = 18   ' Logout
Const TN_BYTE = 19   ' Byte macro
Const TN_DATA = 20   ' Data entry terminal
Const TN_SUP = 21    ' supdup protocol
Const TN_SUPO = 22   ' supdup output
Const TN_SNDL = 23   ' Send location
Const TN_TERM = 24   ' Terminal type
Const TO_EOR = 25    ' End of record
Const TN_TACACS = 26 ' Tacacs user identification
Const TN_OM = 27     ' Output marking
Const TN_TLN = 28    ' Terminal location number
Const TN_3270 = 29   ' Telnet 3270 regime
Const TN_X3 = 30     ' X.3 PAD
Const TN_NAWS = 31   ' Negotiate about window size
Const TN_TS = 32     ' Terminal speed
Const TN_RFC = 33    ' Remote flow control
Const TN_LINE = 34   ' Linemode
Const TN_XDL = 35    ' X display location
Const TN_ENVIR = 36  ' Telnet environment option
Const TN_AUTH = 37   ' Telnet authentication option
Const TN_NENVIR = 39 ' Telnet environment option
Const TN_EXTOP = 25  ' Extended-options-list

Const Dash = " - "
Dim LastHost As String ' Current and Last Connected Host

Function MainConnect(Host As String, Port As Long, Prot As Integer) ' Str or IP , 64K limit , 0 or 1
    On Error Resume Next
    Me.Caption = App.Title & Dash & Host & Dash & "Connecting..."
    
    
    ' Wait till current connection is disconnected
    If Winsock.State <> 0 Then
        Winsock.Close
        Do
            DoEvents
        Loop Until Winsock.State = 0
    End If
    
    
    Winsock.Protocol = Prot ' Sets TCP/UDP
    Winsock.RemoteHost = Host
    Winsock.RemotePort = Port
    Select Case Prot
    Case 0: Winsock.Connect
    Case 1: Winsock.Bind ' Incompleate
    End Select
    
    
    ' Reset Keep Alive
    If tmrKeepAlive.Enabled = True Then
        tmrKeepAlive.Enabled = False
        tmrKeepAlive.Enabled = True
        IntCount = 0
    End If
End Function

Function ListMacros()
    Dim I As Long
    Dim tmpName As String
    
    
    ' Clear up old macro menu list
    mnuMacro(0).Visible = False: mnuMacBlank.Visible = False
    For I = 1 To mnuMacro().UBound
        Unload mnuMacro(I)
    Next I
    
    
    I = 0
    ' Get Macro list from the reg.
    Do
        tmpName = GetSetting(App.Title, "Macros", (I) & " Name", "")
        If tmpName = "" Then Exit Do
        
        If I > 0 Then Load mnuMacro(I) Else mnuMacBlank.Visible = True
        mnuMacro(I).Visible = True
        mnuMacro(I).Caption = (I + 1) & ".   " & tmpName
        I = I + 1
    Loop Until I = 65535
End Function

Private Sub Form_Load()
    On Error Resume Next
    
    
    ' Set Window position
    Me.WindowState = GetSetting(App.Title, "Position", "Window State", Me.WindowState)
    Me.Left = GetSetting(App.Title, "Position", "Left", Me.Left)
    Me.Top = GetSetting(App.Title, "Position", "Top", Me.Top)
    Me.Width = GetSetting(App.Title, "Position", "Width", Me.Width)
    Me.Height = GetSetting(App.Title, "Position", "Height", Me.Height)
    
    
    ' Set other options
    strKeepAlive = GetSetting(App.Title, "Settings", "Keep Alive String", Chr$(13))
    KeepAliveInt = GetSetting(App.Title, "Settings", "Keep Alive Interval", 3)
    tmrKeepAlive.Enabled = -CInt(GetSetting(App.Title, "Settings", "Keep Alive On", 0))
    If tmrKeepAlive.Enabled = True Then mnuKeepAlive.Checked = True
    mnuLocalEcho.Checked = -CInt(GetSetting(App.Title, "Settings", "Local Echo", 0))
    txtMain.FontSize = GetSetting(App.Title, "Settings", "Font Size", 12)
    
    
    ListMacros
    Me.Show
    frmConnect.Show 1, Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    Static DoNothing As Boolean
    
'    If DoNothing = True Then DoNothing = False: Exit Sub
'    Select Case Me.WindowState
'    Case 0
        'If Me.Width > 12420 Then Me.Width = 12420
        txtMain.Width = Me.Width - 120
        txtMain.Height = Me.Height - 885 + 120 + 60 + 15
'        DoNothing = True
'    Case 2
'        Me.WindowState = 0
'        txtMain.Top = 0
'        txtMain.Height = Screen.Height
'        Me.Width = 10020
'        DoNothing = True
'    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save Window position
    SaveSetting App.Title, "Position", "Window State", Me.WindowState
    If Me.WindowState = 0 Then
        SaveSetting App.Title, "Position", "Left", Me.Left
        SaveSetting App.Title, "Position", "Top", Me.Top
        SaveSetting App.Title, "Position", "Width", Me.Width
        SaveSetting App.Title, "Position", "Height", Me.Height
    End If
    
    
    SaveSetting App.Title, "Settings", "Font Size", txtMain.FontSize
    SaveSetting App.Title, "Settings", "Local Echo", -(mnuLocalEcho.Checked)
    Winsock.Close
    DoEvents
    End
End Sub

Private Sub mnuConnect_Click()
    frmConnect.Show 1, Me
End Sub

Private Sub mnuDisconnect_Click()
    Winsock.Close
    Me.Caption = App.Title
End Sub

Private Sub mnuEditMacros_Click()
    frmMacro.Show 1, Me
End Sub

Private Sub mnuFontSize_Click()
    Dim tmpVal As String
    
    tmpVal = InputBox("Please enter a new Font size", "Font Size", txtMain.FontSize)
    If tmpVal = "" Then Exit Sub
    
    If Not IsNumeric(tmpVal) Then Beep: Exit Sub
    txtMain.FontSize = Val(tmpVal)
End Sub

Private Sub mnuKeepAlive_Click()
    frmKeepAlive.Show 1, Me
End Sub

Private Sub mnuLocalEcho_Click()
    mnuLocalEcho.Checked = Not mnuLocalEcho.Checked
End Sub

Private Sub mnuMacro_Click(Index As Integer)
    On Error Resume Next ' Disconnected
    Dim strData As String, tmpStr As String
    Dim I As Long, Pos As Long
    Dim Tmr As Single
    
    
    ' Process the Macro from the Menu
    strData = GetSetting(App.Title, "Macros", Index & " Data")
    For I = 1 To Len(strData)
        
        ' Look for the start of a unprintable character in paratheses
        Pos = InStr(I, strData, "{")
        If Pos = 0 Then Winsock.SendData Mid$(strData, I, Len(strData) - Pos): Exit For ' Send last string and exit
        
        If (Pos - 1) > I Then Winsock.SendData Mid$(strData, I, Pos - I) ' Send intermidiate string
        DoEvents
        
        ' Wait/Send the weird character
        I = Pos + 1
        Pos = InStr(I, strData, "}")
        If Mid$(strData, I, 4) = "WAIT" Then
            Tmr = Timer
            Do
                DoEvents
            Loop Until Timer > (Tmr + Mid$(strData, I + 4, Pos - I - 4))
        Else
            Select Case Mid$(strData, I, Pos - I)
            Case "CR+LF": tmpStr = vbCrLf
            Case "TAB": tmpStr = vbTab
            Case "ESC": tmpStr = Chr$(27)
            Case "CR": tmpStr = vbCr
            Case "LF": tmpStr = vbLf
            Case Else: tmpStr = "": MsgBox "Command Error", vbExclamation, "Macro"
            End Select
        End If
        
        Winsock.SendData tmpStr ' Send unprintable character
        DoEvents
        
        I = Pos
    Next I
End Sub

Private Sub mnuNewWindow_Click()
    On Error Resume Next
    Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
End Sub

Private Sub mnuResetWinFont_Click()
    txtMain.FontSize = 12
    Me.Height = 7545
    Me.Width = 12420
    
    If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - 7545 - 480
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - 12420
End Sub

Private Sub mnuSave_Click()
    On Error Resume Next
    Dim FileName As String
    
    CommonDialog.FileName = LastHost & Dash & Format(Date, "DD-MM-YY") & ".txt"
    CommonDialog.ShowSave
    If Err Then Exit Sub
    
    
    Open CommonDialog.FileName For Output As #1
    Print #1, txtMain.Text
    Close #1
End Sub

Private Sub tmrKeepAlive_Timer()
    On Error Resume Next
    Dim I As Long
    
    If Winsock.State = 0 Then Exit Sub
    
    
    ' Make sure more than just 1 min has passed
    IntCount = IntCount + 1
    If IntCount <> KeepAliveInt Then Exit Sub
    
    
    ' Convert 3 digit number(s) to a character and send
    For I = 1 To Len(strKeepAlive) Step 3
        Winsock.SendData Chr$(Mid$(strKeepAlive, I, 3))
    Next I
    IntCount = 0
End Sub

Private Sub txtMain_Change()
    Dim I As Long, LineCount As Integer, Pos As Long
    
    
    ' Check 819 line limit (64K / 80) ' Doh add CRLF
    For I = Len(txtMain.Text) To 1 Step -1
        Pos = InStrRev(txtMain.Text, Chr$(13), I)
        If Pos = 0 Then Exit For
        LineCount = LineCount + 1
        If LineCount > 794 Then txtMain.Text = Right$(txtMain.Text, Len(txtMain.Text) - I - 2): Exit For ' 795
        I = Pos
    Next I
    
    
    txtMain.SelStart = Len(txtMain.Text) ' Position cursor at end of the text
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next ' Disconnected
    Dim tmpStr As String
    Dim I As Long
    
    
    ' Cut out some key commands - problematic for running programs in a telnet session
    Select Case KeyCode
    Case 8, 46: KeyCode = 0 ' Backspace, Delete ' 9=Tab
    Case 86 ' Paste
        If Shift = 2 Then
            tmpStr = Clipboard.GetText(1)
            If tmpStr <> "" Then
                For I = 1 To Len(tmpStr)
                    Winsock.SendData Mid$(tmpStr, I, 1)
                    If Asc(Mid$(tmpStr, I, 1)) = 8 Then txtMain_KeyPress 8 ' Do a backspace - but might be too eairly
                Next I
            End If
        End If
        KeyCode = 0
        Shift = 0
    End Select
    
    
    ' Reset Keep Alive
    If tmrKeepAlive.Enabled = True Then
        tmrKeepAlive.Enabled = False
        tmrKeepAlive.Enabled = True
        IntCount = 0
    End If
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer) ' 80x24
    On Error Resume Next ' Disconnected
    Dim KeyPress As Integer
    
    
    ' Ignore some key strokes & print others
    Select Case KeyAscii
    Case 24, 3, 22: KeyAscii = 0 ' Cut, Copy, Paste
    Case 8 ' Backspace
        txtMain.Text = Left$(txtMain.Text, Len(txtMain.Text) - 1)
    Case Else
        KeyPress = KeyAscii
    End Select
    
    
    If Winsock.State <> 0 Then Winsock.SendData Chr$(KeyAscii)
    If mnuLocalEcho.Checked = False Then KeyAscii = 0
End Sub

Private Sub Winsock_Close()
    Me.Caption = App.Title & Dash & "Disconnected"
    txtMain.SelStart = Len(txtMain.Text)
    
    Winsock.Close
End Sub

Private Sub Winsock_Connect()
    Me.Caption = App.Title & Dash & Winsock.RemoteHost & Dash & "Connected"
    If Len(txtMain.Text) > 2 Then txtMain.Text = txtMain.Text & vbCrLf & vbCrLf
    LastHost = Winsock.RemoteHost
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim D() As Byte
    ReDim D(bytesTotal - 1) As Byte
    Dim I As Long
    Dim strToDisp As String
    
    Winsock.GetData D(), vbByte
    
    For I = 0 To bytesTotal - 1
        
        ' The following is very incompleate, it just rejects any kinda options the server throws us
        If D(I) = TN_IAC Then ' Start of command
            I = I + 1
            
            Select Case D(I)
            Case TN_SB ' Subnegotiate
                MsgBox "Damm, check ""Winsock_DataArrival"" - ""TN_SB""" & vbCr & "Subnegotiate Incompleate...", vbExclamation, "Error"
                Do ' Dumb waster
                    I = I + 1
                Loop Until D(I) = TN_SE Or I = bytesTotal - 1
                
                
            Case TN_DO ' Server asking can you use
                I = I + 1
                'TN_AUTH - Win XP - Think for using windows passwords
                'TN_NENVIR - Win XP
                'TN_NAWS - Win XP
                'TN_BIN - Win XP
                'After Logon to Win XP Telnet Service - TN_TERM - Could indicate for a clear screen
                Select Case D(I)
                'Case TN_BIN, TN_ECHO: Winsock.SendData Chr$(TN_IAC) & Chr$(TN_WILL) + Chr$(D(I))
                'Case TN_TERM: Winsock.SendData Chr$(TN_IAC) & Chr$(TN_WILL) + Chr$(D(I)) ' Subnegotiate
                Case Else: Winsock.SendData Chr$(TN_IAC) & Chr$(TN_WONT) + Chr$(D(I))
                End Select
                
                
            Case TN_DONT ' Confirming not to use - bad phrasing
                I = I + 1
                'TN_TERM - Cisco Router - Just feed back?
                'TN_NAWS - Cisco Router - Just feed back?
                Winsock.SendData Chr$(TN_IAC) & Chr$(TN_WONT) + Chr$(D(I))
                
                
            Case TN_WILL ' Server can use - F off
                I = I + 1
                'TN_ECHO - Win XP
                'TN_SUPP - Win XP
                'TN_BIN - Win XP
                Winsock.SendData Chr$(TN_IAC) & Chr$(TN_DONT) + Chr$(D(I))
                
                
            Case TN_WONT ' Server wont use - Cool
                I = I + 1
                'TN_ECHO - Cisco Router - Just feed back?
                'TN_SUPP - Cisco Router - Just feed back?
                Winsock.SendData Chr$(TN_IAC) & Chr$(TN_DONT) + Chr$(D(I))
                
                
            'Case Else
            '    Stop
            '    strToDisp = strToDisp & Chr$(D(I))
                
            End Select
            
            
        'ElseIf D(I) = 12 Then ' Clear screen ?
        '    txtMain.Text = ""
            
        Else
            
            ' Not a command - the real ASCII text (luckerly bit 7 should be 0)
            If D(I) = 10 Or D(I) = 8 Or D(I) = 9 Then ' LF, Backspace, Tab
                strToDisp = strToDisp
            ElseIf D(I) = 13 Then
                strToDisp = strToDisp & vbCrLf
            Else
                strToDisp = strToDisp & Chr$(D(I))
            End If
        End If
    Next I
    
    txtMain.Text = txtMain.Text & strToDisp ' Dump text
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.Caption = App.Title & Dash & "Error, Disconnected"
    MsgBox Description, vbExclamation, "Error - " & Number
    
    Winsock.Close
End Sub
