VERSION 5.00
Begin VB.Form frmKeepAlive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Keep Alive Settings"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "3"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtChar 
      Height          =   285
      Left            =   1200
      MaxLength       =   21845
      TabIndex        =   8
      ToolTipText     =   "Enter the characters as you see them"
      Top             =   2040
      Width           =   855
   End
   Begin VB.OptionButton optChar 
      Caption         =   "Custom"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton optChar 
      Caption         =   "LF + CR (10 + 13)"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Tag             =   "010013"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton optChar 
      Caption         =   "CR + LF (13 + 10)"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Tag             =   "013010"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton optChar 
      Caption         =   "Carrage Return (13)"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Tag             =   "013"
      Top             =   960
      Width           =   1935
   End
   Begin VB.OptionButton optChar 
      Caption         =   "Line Feed (10)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Tag             =   "010"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton optChar 
      Caption         =   "Space + Backspace"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "032008"
      Top             =   720
      Width           =   1935
   End
   Begin VB.OptionButton optChar 
      Caption         =   "Space (32)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "032"
      Top             =   480
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "O&N"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "O&FF"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "&Send Interval in mins:"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "&Character(s) to send:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmKeepAlive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)
    Dim I As Long
    
    Select Case Index
    Case 0 ' Reset Keep Alive
        frmMain.tmrKeepAlive.Enabled = False
        frmMain.tmrKeepAlive.Enabled = True
        IntCount = 0
    Case 1: frmMain.tmrKeepAlive.Enabled = False ' Off
    Case 2: GoTo EndOfSub
    End Select
    
    
    ' Get fixed characters from label tags
    For I = 0 To 5
        If optChar(I).Value = True Then strKeepAlive = optChar(I).Tag: Exit For
    Next I
    ' Get custom characters from 'txtChar'
    If I = 6 Then
        strKeepAlive = ""
        For I = 1 To Len(txtChar.Text)
            strKeepAlive = strKeepAlive & Format$(Asc(Mid$(txtChar.Text, I, 1)), "000")
        Next I
    End If
    
    
    ' Save settings
    If frmMain.tmrKeepAlive.Enabled = True Then frmMain.mnuKeepAlive.Checked = True Else frmMain.mnuKeepAlive.Checked = False
    KeepAliveInt = CInt(txtInterval)
    SaveSetting App.Title, "Settings", "Keep Alive String", strKeepAlive
    SaveSetting App.Title, "Settings", "Keep Alive Interval", KeepAliveInt
    SaveSetting App.Title, "Settings", "Keep Alive On", -frmMain.tmrKeepAlive.Enabled
    
EndOfSub:
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Long
    
    
    ' Get select correct option from 'strKeepAlive'
    For I = 0 To 5
        If optChar(I).Tag = strKeepAlive Then optChar(I).Value = True: Exit For
    Next I
    If I = 6 Then
        For I = 1 To Len(strKeepAlive) Step 3
             txtChar.Text = txtChar.Text & Chr$(CLng(Mid$(strKeepAlive, I, 3)))
        Next I
    End If
    
    txtInterval.Text = KeepAliveInt
End Sub

Private Sub txtChar_KeyPress(KeyAscii As Integer)
    optChar(6).Value = True
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 10
End Sub
