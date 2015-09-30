VERSION 5.00
Begin VB.Form frmWordDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8340
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Caption         =   "Address"
      Height          =   495
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Caption         =   "Class"
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Caption         =   "Roll No."
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Caption         =   "Name"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmWordDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wdA As Word.Application
Dim wdD As Word.Document
Dim wdTbl As Word.Table
Dim nRow As Integer
Dim bIsNewRec As Boolean

Private Sub cmdCancel_Click()
On Error GoTo Err_cmdCancel_Click
  Call ShowData
  Call LockCmd
  Call LockTxt
  
  Exit Sub
Err_cmdCancel_Click:
  MsgBox Err.Description
End Sub

Private Sub cmdDel_Click()
On Error GoTo Err_cmdDel_Click
  If IsDataExist = True Then
    Dim nResp As Integer
    nResp = MsgBox("Want to Delete current record?", vbYesNo, "Delete?")
    If nResp = vbYes Then
      wdTbl.Rows(nRow).Delete
      If nRow > wdTbl.Rows.Count Then nRow = wdTbl.Rows.Count
      Call ShowData
    End If
  End If
  
  Exit Sub
Err_cmdDel_Click:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
If IsDataExist Then
  Call LockTxt
  Call LockCmd
End If
End Sub

Private Sub cmdExit_Click()
On Error GoTo Err_cmdExit_Click
'  wdA.Visible = False
'  wdD.Close
'  wdA.Quit
'  Set wdA = Nothing
  'End
  Unload Me

Err_cmdExit_Click:
  If Err.Number = 0 Then
    Exit Sub
  End If
  MsgBox Err.Description
End Sub

Private Sub cmdFind_Click()
On Error GoTo Err_cmdFind_Click
  Dim nR As Integer
  Dim nC As Integer
  Dim nLn As Integer
  Dim nFindVal As String
  nFindVal = InputBox("Enter Roll No.", "Roll No?", 1)
  If nFindVal = "" Then
    Exit Sub
  End If
  If IsNumeric(nFindVal) Then
    For nR = 2 To wdTbl.Rows.Count
      For nC = 1 To wdTbl.Columns.Count
        nLn = Len(wdTbl.Rows(nR).Cells(nC).Range.Text)
        If CInt(nFindVal) = Mid(wdTbl.Rows(nR).Cells(nC).Range.Text, 1, nLn - 2) Then
          nRow = nR
          Call ShowData
          Exit Sub
        End If
      Next nC
    Next nR
    MsgBox "Value not found"
  Else
    MsgBox "No a valid Number"
  End If
  
  Exit Sub
Err_cmdFind_Click:
  MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
 ' Selection.HomeKey wdStory
 nRow = 2
 Call ShowData
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
  frmHelp.WebBrowser1.Navigate App.Path & "\help.htm#frmWordDB"
End Sub

Private Sub cmdLast_Click()
On Error GoTo Err_cmdLast_Click
  nRow = wdTbl.Rows.Count
  Call ShowData
  
  Exit Sub
Err_cmdLast_Click:
  MsgBox Err.Description
End Sub

Private Sub cmdNew_Click()
  Call EmptyTxt
  Call LockCmd
  Call LockTxt
  'nRow = wdTbl.Rows.Count
  bIsNewRec = True
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_cmdNext_Click
  If nRow >= wdTbl.Rows.Count Then
    MsgBox "You are at the end!", vbCritical
    Exit Sub
  End If
  
  nRow = nRow + 1
  Call ShowData
  
  Exit Sub
Err_cmdNext_Click:
  MsgBox Err.Description
End Sub

Private Sub cmdPrev_Click()
  If nRow <= 2 Then
    MsgBox "You are at the start", vbCritical
    Exit Sub
  End If
  
  nRow = nRow - 1
  Call ShowData

End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_cmdSave_Click
  Call LockTxt
  If bIsNewRec = True Then
    Selection.EndKey wdStory
    Selection.MoveUp wdLine
    Selection.InsertRowsBelow 1
    bIsNewRec = False
    nRow = wdTbl.Rows.Count
  End If
  
  
  If IsDataExist Then
    Dim nCol As Integer
    For nCol = 1 To wdTbl.Columns.Count
      wdTbl.Rows(nRow).Cells(nCol).Range.Text = txtData(nCol - 1)
    Next nCol
    wdD.Save
    'MsgBox "Rec is saved!"
  End If
  Call ShowData
  Call LockCmd
  
  Exit Sub
Err_cmdSave_Click:
  MsgBox Err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
'MsgBox Left("ABC", 2)
bIsNewRec = False

Set wdA = New Word.Application
'Set wdD = New Word.Document

Set wdD = wdA.Application.Documents.Open(App.Path & "\data.doc")

Set wdTbl = wdD.Tables(1)
nRow = 1
Call ShowHeader
nRow = 2
Call ShowData
wdA.Visible = True
Me.Show
Exit Sub
'MsgBox wdTbl.Rows.Count
'End
'Dim Str As String
'Dim Ln As Integer
''While (1)
'  Dim N As Integer, C As Integer
'  For N = 1 To wdTbl.Rows.Count
'    For C = 1 To wdTbl.Columns.Count
'      Ln = Len(wdTbl.Rows(N).Cells(C).Range.Text)
'      Str = Str & " " & Mid(wdTbl.Rows(N).Cells(C).Range.Text, 1, Ln - 2)
'    Next C
'    Str = Str & vbCrLf
'  Next N
'  MsgBox Str
'Wend
'wdD.Close
'wdA.Quit

'Set wdD = Nothing
'Set wdA = Nothing
Err_Form_Load:
  MsgBox Err.Number & ". " & Err.Description
End Sub

Private Sub ShowData()
On Error GoTo Err_ShowData
  If IsDataExist = True Then
    Dim nC As Integer
    Dim nLn As Integer
    Dim sStr As String
    For nC = 1 To wdTbl.Columns.Count
        sStr = wdTbl.Rows(nRow).Cells(nC).Range.Text
        nLn = Len(sStr)
        txtData(nC - 1) = Mid(sStr, 1, nLn - 2)
    Next nC
  End If
  
  Exit Sub
Err_ShowData:
  MsgBox Err.Description
End Sub

Private Sub ShowHeader()
On Error GoTo Err_ShowHeader
  Dim nC As Integer
  Dim nLn As Integer
  For nC = 1 To wdTbl.Columns.Count
      nLn = Len(wdTbl.Rows(nRow).Cells(nC).Range.Text)
      lblHeader(nC - 1) = Mid(wdTbl.Rows(nRow).Cells(nC).Range.Text, 1, nLn - 2)
  Next nC
  
  Exit Sub
Err_ShowHeader:
  MsgBox Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload
  wdA.Visible = False
  wdD.Close
  wdA.Quit
  Set wdA = Nothing
  Exit Sub
  
Err_Form_Unload:
  MsgBox Err.Description
End Sub

Private Sub EmptyTxt()
  Dim L As Byte
  For L = 0 To 3
    txtData(L) = ""
  Next L
End Sub

Private Sub LockCmd()
  cmdPrev.Enabled = Not cmdPrev.Enabled
  cmdNext.Enabled = Not cmdNext.Enabled
  cmdExit.Enabled = Not cmdExit.Enabled
  cmdNew.Enabled = Not cmdNew.Enabled
  cmdSave.Enabled = Not cmdSave.Enabled
  cmdCancel.Enabled = Not cmdCancel.Enabled
  cmdFirst.Enabled = Not cmdFirst.Enabled
  cmdLast.Enabled = Not cmdLast.Enabled
  cmdDel.Enabled = Not cmdDel.Enabled
  cmdFind.Enabled = Not cmdFind.Enabled
  cmdEdit.Enabled = Not cmdEdit.Enabled
End Sub

Private Sub LockTxt()
Dim L As Byte
For L = 0 To 3
  txtData(L).Enabled = Not txtData(L).Enabled
Next L
End Sub


Private Function IsDataExist() As Boolean
On Error GoTo Err_IsDataExist
  If wdTbl.Rows.Count <= 1 Then
    MsgBox "No record found"
    IsDataExist = False
  Else
    IsDataExist = True
  End If
  
  Exit Function
Err_IsDataExist:
  MsgBox Err.Description
End Function

