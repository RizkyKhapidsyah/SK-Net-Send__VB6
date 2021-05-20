VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNetSend 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Message"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   15
      TabIndex        =   6
      Top             =   -75
      Width           =   6765
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1320
         Left            =   1545
         MultiLine       =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Enter your message"
         Top             =   1065
         Width           =   5115
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1545
         TabIndex        =   0
         ToolTipText     =   "Enter computer name and Press Enter Key"
         Top             =   240
         Width           =   2505
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   4335
         TabIndex        =   7
         ToolTipText     =   "List of computer names"
         Top             =   180
         Width           =   2340
      End
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5205
         TabIndex        =   4
         Top             =   2520
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3870
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Send Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   285
         Width           =   1380
      End
   End
   Begin MSComctlLib.StatusBar ssb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3060
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12348
            MinWidth        =   12348
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   450
      Top             =   2610
   End
   Begin VB.Timer Timer2 
      Left            =   30
      Top             =   2610
   End
End
Attribute VB_Name = "frmNetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExecuteVal As Variant
Dim Dest As String
Dim Strval As String
Dim bYes As Boolean
Private Sub Command1_Click()
Dim i As Integer
On Error GoTo ram
    Dim t, m
    t = Trim(Text2.Text)
    m = Trim(Text1.Text)
    If List1.ListCount <= 0 Then
    'If t = "" Then '" & Chr(13) & "
        MsgBox "Enter Remote Computer Name(s) For this message", vbInformation, "Message"
        Text2.SetFocus
        Exit Sub
    ElseIf m = "" Then
        MsgBox "Enter Some Message", vbInformation, "Message"
        Text1.SetFocus
        Exit Sub
    End If
    'Strval = "net send " & t & " " & m
    For i = 0 To List1.ListCount - 1
        Strval = "net send " & List1.List(i) & " " & m
        ExecuteVal = Shell(Strval)
    Next i
ram:
If Err.Number <> 0 Then MsgBox Err.Number
End Sub
Private Sub Command1_GotFocus()
    'ssb.Panels(1).Text = "Click me to send your message"
End Sub
Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    List1.Clear
    Text2.SetFocus
End Sub
Private Sub Command2_GotFocus()
    'ssb.Panels(1).Text = "Click me to clear data"
End Sub
Private Sub Command3_Click()
    'Me.bandh
    Unload Me
End Sub
Private Sub Command3_GotFocus()
    'ssb.Panels(1).Text = "Ckick me to Exit or Press Escape"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me 'Me.bandh
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Me.bandh
End Sub
Private Sub List1_Click()
   ' Text2.Text = List1.List(List1.ListIndex)
End Sub
Private Sub List1_DblClick()
On Error Resume Next
    Text2.Text = List1.Text
    List1.RemoveItem (List1.ListIndex)
End Sub
Private Sub List1_GotFocus()
    ssb.Panels(1).Text = "Dispalys list of Computers to send your message"
End Sub

Private Sub Text1_GotFocus()
    ssb.Panels(1).Text = "Enter Your Message"
End Sub

Private Sub Text2_GotFocus()
    ssb.Panels(1).Text = "Enter Computer Name and Press Enter Key"
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    bYes = False
    If KeyAscii = 13 Then
        If Trim(Text2.Text) <> "" Then
            For i = 0 To List1.ListCount
                If Trim(UCase(Text2.Text)) = UCase(List1.List(i)) Then bYes = True: Exit For
            Next i
            If bYes = False Then List1.AddItem Trim(Text2.Text): Text2.Text = ""
        End If
    End If
End Sub
Private Sub Text2_LostFocus()
Dim i As Integer
    bYes = False
    If Trim(Text2.Text) <> "" Then
        For i = 0 To List1.ListCount
            If Trim(UCase(Text2.Text)) = UCase(List1.List(i)) Then bYes = True: Exit For
        Next i
        If bYes = False Then List1.AddItem Trim(Text2.Text): Text2.Text = ""
    End If
End Sub

Private Sub Timer1_Timer()
    'Label4.Caption = "Bye Rams !!"
End Sub
Private Sub Timer2_Timer()
    End
End Sub
Public Sub bandh()
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer1.Interval = 200
    Timer2.Interval = 1400
End Sub
