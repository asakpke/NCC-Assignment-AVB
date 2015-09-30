VERSION 5.00
Begin VB.Form frmData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4245
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelFind 
      Caption         =   "Show All Records"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Age"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Location"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Death Date"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Birth Date"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As ADODB.Connection
Public Rs As ADODB.Recordset
Dim AP As Long

Public Sub DbOpen(ByVal strPath As String)
On Error GoTo Err_DbOpen
    Set Con = New ADODB.Connection
    Con.ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & strPath & "\Data.mdb"
    Con.Open
    Set Rs = New ADODB.Recordset
    's
    Dim strSQL As String
    strSQL = "SELECT Data.*, (DateDiff(""yyyy"",[BirthDate],[DeathDate])) AS Age " _
           & "FROM Data;"
    'MsgBox strSQL
    'e
    Rs.Open strSQL, Con, adOpenKeyset, adLockOptimistic
    'Rs.Open "SELECT * FROM Data", Con, adOpenKeyset, adLockOptimistic
    'Rs.AbsolutePosition = AP
    Exit Sub
    
Err_DbOpen:
    MsgBox Err.Description

End Sub
Public Sub mFirst()
    Rs.MoveFirst
    Call ShowData
End Sub
Public Sub mNext()
    Rs.MoveNext
    If Rs.EOF Then
        Rs.MoveLast
    End If
    Call ShowData
End Sub
Public Sub mPrev()
    Rs.MovePrevious
    If Rs.BOF Then
        Rs.MoveFirst
    End If
    Call ShowData
End Sub
Public Sub mLast()
    Rs.MoveLast
    Call ShowData
End Sub

Public Sub ShowData()
  Dim L As Integer
  For L = 0 To 5
    txtData(L) = Rs(L)
  Next
  'txtData(4) = DateDiff("yyyy", txtData(2), txtData(3))
End Sub

Public Sub cmdCancelFind_Click()
  Rs.Close
  Dim strSQL As String
  strSQL = "SELECT Data.*, (DateDiff(""yyyy"",[BirthDate],[DeathDate])) AS Age " _
           & "FROM Data;"
  Rs.Open strSQL, Con, adOpenKeyset, adLockOptimistic
  Call ShowData
  Call SwitchFind
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdFind_Click()
  frmFind.Show
End Sub
Public Sub SwitchFind()
  cmdFind.Visible = Not cmdFind.Visible
  cmdCancelFind.Visible = Not cmdCancelFind.Visible
End Sub
Private Sub cmdFirst_Click()
  Call mFirst
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
  frmHelp.WebBrowser1.Navigate App.Path & "\help.htm#frmData"
End Sub

Private Sub cmdLast_Click()
  Call mLast
End Sub

Private Sub cmdNext_Click()
  Call mNext
End Sub

Private Sub cmdPrev_Click()
  Call mPrev
End Sub

Private Sub Form_Load()
  Call DbOpen(App.Path)
  Call ShowData
End Sub

