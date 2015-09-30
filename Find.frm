VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtAge 
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Data Form,"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtLoc 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtDeathDateStart 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDeathDateEnd 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtBirthDateEnd 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtBirthDateStart 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Age"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "To"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Location "
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Death Date Range"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Birth Date Range"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Call frmData.ShowData
    'MsgBox frmData.Visible
    Call frmData.cmdCancelFind_Click
    Call frmData.SwitchFind
    Unload Me
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
  frmHelp.WebBrowser1.Navigate App.Path & "\help.htm#frmSearch"
End Sub

Private Sub cmdSearch_Click()
  frmData.Show
    
  Dim strSQL As String
  strSQL = ""
  
  If txtName <> "" Then
    strSQL = "Name = '" & txtName & "'"
  End If
  
  If IsDate(txtBirthDateStart) Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "BirthDate >= #" & txtBirthDateStart & "#"
  End If
  
  If IsDate(txtBirthDateEnd) Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "BirthDate <= #" & txtBirthDateEnd & "#"
  End If
  
  's
  If IsDate(txtDeathDateStart) Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "DeathDate >= #" & txtDeathDateStart & "#"
  End If
  
  If IsDate(txtDeathDateEnd) Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "DeathDate <= #" & txtDeathDateEnd & "#"
  End If
  'e
  
  If txtLoc <> "" Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "Location = '" & txtLoc & "'"
  End If
  
  If txtAge <> "" Then
    If strSQL <> "" Then
      strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & "DateDiff(""yyyy"",[BirthDate],[DeathDate])= " & txtAge
  End If
  
  If strSQL <> "" Then
    strSQL = " WHERE " & strSQL
  End If
  
  frmData.Rs.Close
  
  'MsgBox strSQL
  
  frmData.Rs.Open "SELECT Data.*, (DateDiff(""yyyy"",[BirthDate],[DeathDate])) AS Age FROM Data" & strSQL, frmData.Con, adOpenKeyset, adLockOptimistic
  'SELECT Data.*, (DateDiff(""yyyy"",[BirthDate],[DeathDate])) AS Age
  If frmData.Rs.EOF Then
    MsgBox "No Record Found"
  Else
    'MsgBox "find"
    Call frmData.ShowData
    Call frmData.SwitchFind
    Unload Me
  End If
End Sub

Private Sub Text2_Change()

End Sub

