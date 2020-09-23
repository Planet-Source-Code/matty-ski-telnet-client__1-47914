VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optPort 
      Caption         =   "Custom"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "POP3 (110)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   8
      Tag             =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "TFTP (69)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "TFTP Can't handle file transfer"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "SMTP (25)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Tag             =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "FTP (21)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Tag             =   "0"
      ToolTipText     =   "FTP Can't handle file transfer"
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Telnet (23)"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Tag             =   "0"
      Top             =   360
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox lstPorts 
      Height          =   1035
      Left            =   3720
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   13
      Text            =   "23"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame frmProtocol 
      Caption         =   "P&rotocol"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   975
      Begin VB.OptionButton optProt 
         Caption         =   "UDP"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optProt 
         Caption         =   "TCP"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.ListBox lstHosts 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "&Port Number:"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Host:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)
    Dim I As Long
    Dim Match As Boolean
    
    
    If Val(txtPort) > 65535 Then MsgBox "Invalid Port Number", vbExclamation, "Error": Exit Sub
    
    
    If Index = 0 Then ' OK button
        If Trim$(txtHost) = "" Or Trim$(txtPort) = "" Then Beep: Exit Sub
        
        
        ' Save Host list
        For I = 0 To lstHosts.ListCount - 1
            SaveSetting App.Title, "Recent Hosts", I, lstHosts.List(I)
            If lstHosts.List(I) = txtHost Then Match = True
        Next I
        If Match = 0 Then SaveSetting App.Title, "Recent Hosts", I, txtHost
        
        
        ' Save Port number list
        If optPort(5).Value = True Then
            Match = 0
            For I = 0 To lstPorts.ListCount - 1
                SaveSetting App.Title, "Recent Ports", I, lstPorts.List(I)
                If lstPorts.List(I) = txtPort Then Match = True
            Next I
            If Match = 0 Then SaveSetting App.Title, "Recent Ports", I, txtPort
        End If
        
        
        SaveSetting App.Title, "Recent Hosts", "Last Host", txtHost.Text
        frmMain.MainConnect txtHost, txtPort, -(optProt(1).Value) ' Pass info to the main form
    End If
    
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim tmpHost As String, tmpPort As String
    
    
    ' Retreve recents Hosts and Port numbers
    Do
        tmpHost = GetSetting(App.Title, "Recent Hosts", I, "")
        tmpPort = GetSetting(App.Title, "Recent Ports", I, "")
        If tmpHost <> "" Then lstHosts.AddItem tmpHost
        If tmpPort <> "" Then lstPorts.AddItem tmpPort
        I = I + 1
    Loop Until tmpHost = "" And tmpPort = ""
    
    
    txtHost.Text = GetSetting(App.Title, "Recent Hosts", "Last Host", "")
    txtHost.SelLength = Len(txtHost.Text)
End Sub

Private Sub lstHosts_DblClick()
    txtHost.Text = lstHosts.List(lstHosts.ListIndex)
End Sub

Private Sub lstPorts_DblClick()
    txtPort.Text = lstPorts.List(lstPorts.ListIndex)
End Sub

Private Sub optPort_Click(Index As Integer)
    Dim tmpStr As String
    
    ' Set 'txtPort' to the port number
    If Index < 5 Then
        ' Get port number from the label
        tmpStr = Right$(optPort(Index).Caption, 4)
        If Left$(tmpStr, 1) = "(" Then txtPort = Mid$(tmpStr, 2, 2) Else txtPort = Mid$(tmpStr, 1, 3)
        optProt(CInt(optPort(Index).Tag)).Value = True
    Else
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort.Text)
        txtPort.SetFocus
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    optPort(5).Value = True
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 10
End Sub
