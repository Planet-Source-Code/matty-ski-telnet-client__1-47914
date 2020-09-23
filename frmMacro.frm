VERSION 5.00
Begin VB.Form frmMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Macros"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "< Del."
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "< View"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "< Add"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "3"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "<"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "\/"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtCommands 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Position Cursor to insert Words / Commands"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.ListBox lstMacros 
      Height          =   2400
      ItemData        =   "frmMacro.frx":0000
      Left            =   120
      List            =   "frmMacro.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox lstCmds 
      Height          =   2400
      ItemData        =   "frmMacro.frx":0004
      Left            =   2880
      List            =   "frmMacro.frx":0029
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "&Command List:"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Add &Wait (Secs):"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "&Macro List:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MacroInfo() As String ' 1,X (X starts from 1)
                          ' 0=Name
                          ' 1=Data

Private Sub cmdAct_Click(Index As Integer)
    Dim tmpStr As String
    Dim tmpVal As Long
    
    Select Case Index
    Case 0 ' Add New Macro to list and store data
        If Trim$(txtCommands) = "" Then Beep: Exit Sub
        tmpStr = InputBox("Enter a name / phrase for this new Macro", "New Macro", "")
        If tmpStr = "" Then Exit Sub
        
        lstMacros.AddItem tmpStr
        tmpVal = UBound(MacroInfo(), 2) + 1
        ReDim Preserve MacroInfo(1, tmpVal) As String
        MacroInfo(0, tmpVal) = tmpStr
        MacroInfo(1, tmpVal) = txtCommands.Text
        
        ' Cosmetics
        lstMacros.ListIndex = lstMacros.ListCount - 1
        txtCommands.SelStart = 0
        txtCommands.SelLength = Len(txtCommands.Text)
        
    Case 1 ' View current selected Macro
        If lstMacros.ListIndex = -1 Then Beep: Exit Sub
        txtCommands.Text = MacroInfo(1, lstMacros.ListIndex + 1)
        txtCommands.SelStart = Len(txtCommands.Text)
        
    Case 2 ' Mark Macro for Deletion
        If lstMacros.ListIndex = -1 Or lstMacros.List(lstMacros.ListIndex) = Chr$(1) Then Beep: Exit Sub
        If MsgBox("Are you sure you want to delete:" & vbCr & lstMacros.List(lstMacros.ListIndex), vbInformation + vbYesNo + vbDefaultButton2, "Delete") = vbNo Then Exit Sub
        MacroInfo(0, lstMacros.ListIndex + 1) = Chr$(1) ' Delete Mark Character
        lstMacros.List(lstMacros.ListIndex) = "* Deleted *"
    End Select
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Dim I As Long, II As Long
    
    If Index = 0 Then
        Do ' Save Macro list to reg.
            I = I + 1
            If MacroInfo(0, I) = "" Then Exit Do ' None left
            If Left$(MacroInfo(0, I), 1) <> Chr$(1) Then ' Not Marked for Deleting
                SaveSetting App.Title, "Macros", (II) & " Name", MacroInfo(0, I)
                SaveSetting App.Title, "Macros", (II) & " Data", MacroInfo(1, I)
                II = II + 1
            End If
        Loop Until I = 65535 Or (I = UBound(MacroInfo(), 2))
        SaveSetting App.Title, "Macros", (II) & " Name", "" ' For ignoring old data
        frmMain.ListMacros
    End If
    
    Unload Me
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim tmpStr As String
    
    ' The '\/' or '<' buttons
    Select Case Index
    Case 0: tmpStr = "{WAIT" & Trim$(txtWait.Text) & "}"
    Case 1: tmpStr = lstCmds.List(lstCmds.ListIndex)
    End Select
    txtCommands.SelText = tmpStr
End Sub

Private Sub Form_Load()
    ReDim MacroInfo(1, 0) As String
    Dim I As Long
    Dim tmpName As String
    
    Do ' Get Macro list from the reg.
        tmpName = GetSetting(App.Title, "Macros", (I) & " Name", "")
        If tmpName = "" Then Exit Do
        
        I = I + 1
        ReDim Preserve MacroInfo(1, I) As String
        lstMacros.AddItem tmpName
        MacroInfo(0, I) = tmpName
        MacroInfo(1, I) = GetSetting(App.Title, "Macros", (I - 1) & " Data", "")
    Loop Until I = 65535
End Sub

Private Sub lstCmds_DblClick()
    txtCommands.SelText = lstCmds.List(lstCmds.ListIndex)
End Sub

Private Sub lstMacros_DblClick()
    cmdAct_Click 1
End Sub

Private Sub txtWait_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 10
End Sub
