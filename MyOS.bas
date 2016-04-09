Attribute VB_Name = "MyOS"
Option Explicit

'Constants, types and other Global Variables

Public Enum ProcSchedules
  NONE = 0
  FCFS = 1
  SJF = 2
  SRTF = 3
  Priority = 4
  RoundRobin = 5
End Enum

Public Scheduler As ProcSchedules

Public Type Process
  CPUTime As Long
  ArrivalTime As Long
  Priority As Long
  Name As Long
  InQueue As Boolean
  Color As Long
End Type

Public Type QueueProcess
  Name As String
  CPUTime As Long
  ArrivalTime As Long
  Priority As Long
  RemainingTime As Long
  WaitingTime As Long
  Done As Boolean
  Color As Long
End Type

Public Type CPU
  ProcessName As Long
  CpuTimer As Long
End Type

Public myProc() As Process
Public myQueue() As QueueProcess
Public MyCPU As CPU

Public TimeSlice As Long

Public Const aswIsVacant As Long = -1
Public Const aswFinish As Long = -2
Public Const aswNoneAvailable As Long = -3

Public BoxTop As Long
Public BoxLeft As Long
Public Const BoxStartTop As Long = 100
Public Const BoxStartLeft As Long = 100
Public Const HIncrement As Long = 400
Public Const WIncrement As Long = 200

'Graphics of My Program!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Public Sub DrawBox(box As PictureBox, FillColor As Long, Border As Long)
Dim i As Long
Static RefreshProc As Boolean
  'draw rectangle
  If RefreshProc = True Then
    BoxTop = BoxStartTop
    BoxLeft = BoxStartLeft
    box.Cls
    RefreshProc = False
  End If
  
  box.ForeColor = FillColor
  For i = BoxLeft To BoxLeft + WIncrement
    box.Line (i, BoxTop)-(i, BoxTop + HIncrement)
  Next i
  
  'draw border
  box.ForeColor = Border
  box.Line (BoxLeft, BoxTop)-(BoxLeft + WIncrement, BoxTop)
  box.Line (BoxLeft, BoxTop + HIncrement)-(BoxLeft + WIncrement, BoxTop + HIncrement)
  box.Line (BoxLeft, BoxTop)-(BoxLeft, BoxTop + HIncrement)
  box.Line (BoxLeft + WIncrement, BoxTop)-(BoxLeft + WIncrement, BoxTop + HIncrement)
  
  'increment coords
  If (BoxLeft + WIncrement) >= box.ScaleWidth - WIncrement Then
    BoxLeft = BoxStartLeft
    BoxTop = BoxTop + HIncrement + 100
    If BoxTop >= box.ScaleHeight - HIncrement Then
      RefreshProc = True
    Else
      RefreshProc = False
    End If
  Else
    BoxLeft = BoxLeft + WIncrement
  End If
End Sub

Public Sub SetCoords()
  BoxLeft = BoxStartLeft
  BoxTop = BoxStartTop
End Sub

'Process Entry Codes
Public Function CheckMyProc() As Boolean
  On Error GoTo NotInitialized
  Dim i As Long
  i = UBound(myProc)
  CheckMyProc = True
  Exit Function
NotInitialized:
  CheckMyProc = False
End Function

Public Sub AddProcess(CPUTime As Long, ArrivalTime As Long, Priority As Long)
  If CheckMyProc = True Then
    ReDim Preserve myProc(UBound(myProc) + 1)
  Else
    ReDim myProc(0)
  End If
  Randomize
  With myProc(UBound(myProc))
    .CPUTime = CPUTime
    .ArrivalTime = ArrivalTime
    .Priority = Priority
    .Name = UBound(myProc) + 1
    .InQueue = False
    .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
  End With
End Sub

Public Sub ShowProcessesInList(list As ListBox)
  Dim i As Long
  list.AddItem ("NAME" & vbTab & vbTab & "CPUTime" & vbTab & vbTab & "ArrivalTime" & vbTab & "Priority")
  For i = LBound(myProc) To UBound(myProc)
    With myProc(i)
      list.AddItem ("Process " & .Name & vbTab & .CPUTime & vbTab & vbTab & .ArrivalTime & vbTab & vbTab & .Priority)
    End With
  Next i
End Sub

Public Sub ClearProcessesInList(list As ListBox)
  While list.ListCount > 0
    list.RemoveItem 0
  Wend
End Sub

Public Sub ResetQueue()
  Dim i As Long
  For i = LBound(myQueue) To UBound(myQueue)
    With myQueue(i)
      .RemainingTime = .CPUTime
      .WaitingTime = 0
      .Done = False
    End With
  Next i
End Sub

Public Sub ResetProcess()
  Dim i As Long
  For i = LBound(myProc) To UBound(myProc)
    With myProc(i)
      .InQueue = False
    End With
  Next i
End Sub

Public Sub ClearProcess()
  Erase myProc
  Erase myQueue
End Sub

Public Sub ClearQueue()
  Erase myQueue
End Sub

Public Function PutInQueue(CpuTimer As Long) As Boolean
  Dim i As Long
  If CheckMyProc = False Then
    PutInQueue = False
    Exit Function
  End If
  For i = LBound(myProc) To UBound(myProc)
    If myProc(i).ArrivalTime <= CpuTimer And myProc(i).InQueue = False Then
      AddQueueProc i
      myProc(i).InQueue = True
    End If
  Next i
  PutInQueue = True
End Function

Public Sub AddQueueProc(Proc As Long)
  If CheckMyQueue = False Then
    ReDim myQueue(0)
  Else
    ReDim Preserve myQueue(UBound(myQueue) + 1)
  End If
  With myQueue(UBound(myQueue))
    .CPUTime = myProc(Proc).CPUTime
    .ArrivalTime = myProc(Proc).ArrivalTime
    .Priority = myProc(Proc).Priority
    .Color = myProc(Proc).Color
    .Name = myProc(Proc).Name
    .RemainingTime = .CPUTime
    .Done = False
    .WaitingTime = 0
  End With
End Sub

Public Function CheckMyQueue() As Boolean
  On Error GoTo NotInitialized
  Dim i As Long
  i = UBound(myQueue)
  CheckMyQueue = True
  Exit Function
NotInitialized:
  CheckMyQueue = False
End Function

Public Function CheckIfAllProcInQueueIsDone() As Boolean
  If CheckMyQueue = False Then
    CheckIfAllProcInQueueIsDone = True
    Exit Function
  Else
    Dim i As Long
    For i = LBound(myQueue) To UBound(myQueue)
      If myQueue(i).Done = False Then
        CheckIfAllProcInQueueIsDone = False
        Exit Function
      End If
    Next i
    CheckIfAllProcInQueueIsDone = True
    Exit Function
  End If
End Function

'Process Scheduling Algorithms!
Public Function SortQueueFCFS() As Boolean
  On Error GoTo ErrorCode
  Dim temp As QueueProcess
  
  Dim i As Long, j As Long
  For i = LBound(myQueue) To (UBound(myQueue) - 1)
    For j = i + 1 To UBound(myQueue)
      If myQueue(i).ArrivalTime > myQueue(j).ArrivalTime Then
        temp = myQueue(i)
        myQueue(i) = myQueue(j)
        myQueue(j) = temp
      ElseIf myQueue(i).ArrivalTime = myQueue(j).ArrivalTime Then
        If myQueue(i).Name > myQueue(j).Name Then
          temp = myQueue(i)
          myQueue(i) = myQueue(j)
          myQueue(j) = temp
        End If
      End If
    Next j
  Next i
  SortQueueFCFS = True
  Exit Function
ErrorCode:
  SortQueueFCFS = False
End Function

Public Function SortQueueSJF() As Boolean
  On Error GoTo ErrorCode
  Dim temp As QueueProcess
  
  Dim i As Long, j As Long
  For i = LBound(myQueue) To (UBound(myQueue) - 1)
    For j = i + 1 To UBound(myQueue)
      If myQueue(i).CPUTime > myQueue(j).CPUTime Then
        temp = myQueue(i)
        myQueue(i) = myQueue(j)
        myQueue(j) = temp
      ElseIf myQueue(i).CPUTime = myQueue(j).CPUTime Then
        If myQueue(i).ArrivalTime > myQueue(j).ArrivalTime Then
          temp = myQueue(i)
          myQueue(i) = myQueue(j)
          myQueue(j) = temp
        ElseIf myQueue(i).ArrivalTime = myQueue(j).ArrivalTime Then
          If myQueue(i).Name > myQueue(j).Name Then
            temp = myQueue(i)
            myQueue(i) = myQueue(j)
            myQueue(j) = temp
          End If
        End If
      End If
    Next j
  Next i
  SortQueueSJF = True
  Exit Function
ErrorCode:
  SortQueueSJF = False
End Function
  
Public Function SortQueueSRTF() As Boolean
  On Error GoTo ErrorCode
  Dim temp As QueueProcess
  Dim i As Long, j As Long
  For i = LBound(myQueue) To (UBound(myQueue) - 1)
    For j = i + 1 To UBound(myQueue)
      If myQueue(i).RemainingTime > myQueue(j).RemainingTime Then
        temp = myQueue(i)
        myQueue(i) = myQueue(j)
        myQueue(j) = temp
      End If
    Next j
  Next i
  SortQueueSRTF = True
  Exit Function
ErrorCode:
  SortQueueSRTF = False
End Function

Public Function SortQueuePriority() As Boolean
  On Error GoTo ErrorCode
  Dim temp As QueueProcess
  Dim i As Long, j As Long
  For i = LBound(myQueue) To (UBound(myQueue) - 1)
    For j = i + 1 To UBound(myQueue)
      If myQueue(i).Priority > myQueue(j).Priority Then
        temp = myQueue(i)
        myQueue(i) = myQueue(j)
        myQueue(j) = temp
      End If
    Next j
  Next i
  SortQueuePriority = True
  Exit Function
ErrorCode:
  SortQueuePriority = False
End Function

Public Function AllProcInQueue() As Boolean
  On Error GoTo ErrorCode
  Dim i As Long
  Dim Result As Boolean
  
  Result = True
  For i = LBound(myProc) To UBound(myProc)
    If myProc(i).InQueue = False Then
      Result = False
    End If
  Next i
  AllProcInQueue = Result
  Exit Function
ErrorCode:
  AllProcInQueue = True
End Function


Public Sub main()
  Scheduler = NONE
  frmMain.Show
End Sub

Public Function GetQueueIndexGivenName(Name As Long) As Long
  Dim i As Long
  If (Name = -1) Or (CheckMyQueue = False) Then
    GetQueueIndexGivenName = -1
    Exit Function
  End If
  For i = LBound(myQueue) To UBound(myQueue)
    If myQueue(i).Name = Name Then
      GetQueueIndexGivenName = i
      Exit Function
    End If
  Next i
  GetQueueIndexGivenName = -1
End Function
