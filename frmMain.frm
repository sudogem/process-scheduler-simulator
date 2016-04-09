VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Proc"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSimulate 
      Caption         =   "&Simulate"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cboAlgorithm 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0013
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Process"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lbl 
      Caption         =   "Choose Algorithm:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  frmAddProcess.Show vbModal, Me
End Sub

Private Sub cmdClear_Click()
  ClearProcess
  ClearProcessesInList List1
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdSimulate_Click()
  Select Case cboAlgorithm.Text
    Case "FCFS"
      Scheduler = FCFS
    Case "SJF"
      Scheduler = SJF
    Case "SRTF"
      Scheduler = SRTF
    Case "Priority"
      Scheduler = Priority
    Case "RoundRobin"
      Scheduler = RoundRobin
    Case Else
      Scheduler = NONE
  End Select
  
  If Scheduler = NONE Or List1.ListCount < 1 Then
    Dim str As String
    str = ""
    If Scheduler = NONE Then
      str = "Please Choose an Algorithm to Use!"
    End If
    
    If List1.ListCount < 1 Then
      str = str & IIf(str = "", "", vbCrLf & "And Also ") & "Please Add Process/es to Schedule!"
    End If
    MsgBox str, vbQuestion + vbOKOnly, "No Algorithm Chosen!"
    Exit Sub
  Else
    frmGantt.Command1.Caption = "Start"
    frmGantt.Show vbModal, Me
  End If
End Sub


