VERSION 5.00
Begin VB.Form frmAddProcess 
   Caption         =   "Add Process Form"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   3915
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPriority 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Enter a value Between 0-10"
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtArrival 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Enter a value Between 0-100"
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCPU 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Enter a value Between 1 - 20"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "&Priority"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Arrival&Time"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "C&PUTime"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim str As String
  str = CheckProcessEntries()
  If Trim(str) = "" Then
    AddProcess Val(txtCPU.Text), Val(txtArrival.Text), Val(txtPriority.Text)
    ClearProcessesInList frmMain.List1
    ShowProcessesInList frmMain.List1
    Unload Me
  Else
    MsgBox str
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Function CheckProcessEntries() As String
    Dim str As String
    str = ""
    If Val(txtCPU.Text) < 1 Or Val(txtCPU.Text) > 20 Then
       str = "CPU time need to be Between 1 and 20" & vbCrLf
    End If
    If (Val(txtArrival.Text) < 0 Or Val(txtArrival.Text) > 100) And IsNumeric(txtArrival.Text) Then
       str = str & "Arrival time need to be Between 0 and 100"
    End If
    If (Val(txtPriority.Text) < 0 Or Val(txtPriority.Text) > 10) And IsNumeric(txtPriority.Text) Then
       str = str & "Priority need to be Between 0 and 10"
    End If
    CheckProcessEntries = str
End Function

