VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGantt 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4380
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   4740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox CPUbox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmGantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllInQueue As Boolean
Dim AllInQueueIsDone As Boolean

Private Sub Command1_Click()
  If Command1.Caption = "Start" Then
    Command1.Enabled = False
    AllInQueue = False
    AllInQueueIsDone = False
    MyCPU.ProcessName = -1
    MyCPU.CpuTimer = 0
    ClearQueue
    ResetProcess
    Timer1.Enabled = True
  ElseIf Command1.Caption = "Finish" Then
    Unload Me
  End If
  'Randomize
  'DrawBox CPUbox, RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256)), vbWhite
End Sub

Private Sub Form_Load()
  SetCoords
End Sub

Private Sub Timer1_Timer()
  If AllInQueue = False Then
    'check for newly arrived processes!
    If PutInQueue(MyCPU.CpuTimer) = False Then
      Timer1.Enabled = False
      Exit Sub
    End If
    AllInQueue = AllProcInQueue
  Else
    AllInQueueIsDone = CheckIfAllProcInQueueIsDone
  End If
  
  If AllInQueueIsDone = True Then
    Timer1.Enabled = False
    Command1.Caption = "Finish"
    Command1.Enabled = True
    Exit Sub
  End If
  
  Select Case Scheduler
    Case FCFS
      SortQueueFCFS
    Case SJF
      SortQueueSJF
    Case SRTF
      SortQueueSRTF
    Case Priority
      SortQueuePriority
    Case RoundRobin
      SortQueueFCFS
  End Select
  
  'if queue is empty
  If CheckMyQueue = False Then
    MyCPU.CpuTimer = MyCPU.CpuTimer + 1
    'draw box to represent proc in cpu!
    DrawBox CPUbox, vbBlack, vbRed
  'if queue is not empty
  Else
    Dim i As Long
    'if notpreemptive
    If Scheduler = FCFS Or Scheduler = SJF Then
      Dim QueueIndex As Long
      QueueIndex = GetQueueIndexGivenName(MyCPU.ProcessName)
      
      'if cpu is empty or process in cpu is finish
      If QueueIndex = -1 Then
        MyCPU.ProcessName = -1
        For i = LBound(myQueue) To UBound(myQueue)
          'put process in cpu for processing
          If myQueue(i).Done = False Then
            MyCPU.ProcessName = myQueue(i).Name
            Exit For
          End If
        Next i
      ElseIf myQueue(GetQueueIndexGivenName(MyCPU.ProcessName)).Done = True Then
        MyCPU.ProcessName = -1
        For i = LBound(myQueue) To UBound(myQueue)
          'put process in cpu for processing
          If myQueue(i).Done = False Then
            MyCPU.ProcessName = myQueue(i).Name
            Exit For
          End If
        Next i
        
      'if there is an unfinished process in cpu do nothing
      End If
            
    'if preemptive
    ElseIf Scheduler = SRTF Or Scheduler = RoundRobin Then
      
    End If
    
    
    
    'increment cpu timer
    MyCPU.CpuTimer = MyCPU.CpuTimer + 1
    
    'increment waiting time for those process that are not in cpu
    'decrement remainingtime of process in cpu
    For i = LBound(myQueue) To UBound(myQueue)
      If myQueue(i).Done = False Then
        If (myQueue(i).Name <> MyCPU.ProcessName) Then
          myQueue(i).WaitingTime = myQueue(i).WaitingTime + 1
        Else
          myQueue(i).RemainingTime = myQueue(i).RemainingTime - 1
          If myQueue(i).RemainingTime = 0 Then
            myQueue(i).Done = True
          End If
        End If
      End If
    Next i
    'draw box to represent proc in cpu!
    If MyCPU.ProcessName = -1 Then
      DrawBox CPUbox, vbBlack, vbRed
    Else
      DrawBox CPUbox, myQueue(GetQueueIndexGivenName(MyCPU.ProcessName)).Color, vbWhite
    End If
  End If
End Sub

