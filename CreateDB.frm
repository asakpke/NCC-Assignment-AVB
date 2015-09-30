VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateDb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   5040
      Width           =   735
   End
   Begin VB.ListBox lstInd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   7680
      Style           =   1  'Checkbox
      TabIndex        =   25
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstType 
      Height          =   3900
      Left            =   6480
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstFld 
      Height          =   3900
      ItemData        =   "CreateDB.frx":0000
      Left            =   5040
      List            =   "CreateDB.frx":0002
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame fraFld 
      Caption         =   "Field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkIndex 
         Caption         =   "Indexed Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkPk 
         Caption         =   "Primary Key"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CreateDB.frx":0004
         Left            =   3000
         List            =   "CreateDB.frx":0006
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddFld 
         Caption         =   "Add Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtFld 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdNewTbl 
         Caption         =   "New Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Fld "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cdlSave 
      Left            =   2280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNewDb 
      Caption         =   "New Database"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame fraTbl 
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtTbl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCreateTbl 
         Caption         =   "Create Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tbl Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame FraDb 
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCreateDb 
         Caption         =   "Create Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      Caption         =   "If you want to save table && exit  You must click finish to finalize the Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Indexed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Select and right click to delete a field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   24
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Added Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblTbl 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDb 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4455
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Del"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuDelFld 
         Caption         =   "Delete Field"
      End
   End
End
Attribute VB_Name = "frmCreateDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ws As DAO.Workspace
Dim Db As DAO.Database
Dim Tbl As DAO.TableDef
Dim Fld As DAO.Field
Dim Ind As DAO.Index

Private Sub chkIndex_Click()
  'chkBlank.Enabled = True
  chkPk.Enabled = Not chkPk.Enabled
  'chkReq.Enabled = True
End Sub

Private Sub cmdAddFld_Click()
  If txtFld.Text = "" Then
    MsgBox "Enter field name"
    Exit Sub
  End If
  Dim Lv As Integer
'  MsgBox lstFld.ListCount
  If lstFld.ListCount > 0 Then
    For Lv = 0 To lstFld.ListCount - 1
      If txtFld.Text = lstFld.List(Lv) Then
        MsgBox "You already have a field name " & txtFld
        Exit Sub
      End If
      If lstInd.List(Lv) = "PrimaryKey" And chkPk.Value = vbChecked Then
        MsgBox "Primary Key already defined"
        Exit Sub
      End If
    Next Lv
  End If
  Set Fld = Tbl.CreateField(txtFld, cboType.ItemData(cboType.ListIndex))
  
  txtFld.SetFocus
  Tbl.Fields.Append Fld
  
  lstFld.AddItem txtFld
  lstType.AddItem cboType.Text
  lstInd.AddItem " "
  
  If chkIndex.Value = vbChecked Then
    Set Ind = Tbl.CreateIndex(txtFld)
    Set Fld = Ind.CreateField(txtFld, cboType.ItemData(cboType.ListIndex))
    Ind.Fields.Append Fld
    lstInd.Selected(lstInd.NewIndex) = True
    Ind.Primary = IIf(chkPk.Value, True, False)
    If Ind.Primary = True Then
      lstInd.List(lstInd.NewIndex) = "PrimaryKey"
      chkPk.Value = vbUnchecked
      'chkPk.Enabled = False
    End If
    'Ind.IgnoreNulls = IIf(chkBlank.Value, True, False)
    'Ind.Required = IIf(chkReq.Value, True, False)
    
    Tbl.Indexes.Append Ind
  End If
  txtFld = ""
  'chkPk.Value = vbUnchecked
  'chkBlank.Value = vbUnchecked
  'chkReq.Value = vbUnchecked
  chkIndex.Value = vbUnchecked
  'chkPk.Enabled = False
  'chkBlank.Enabled = False
  'chkReq.Enabled = False
End Sub

Private Sub cmdCancel_Click()
  cmdFinish.Enabled = False
  fraFld.Visible = False
  fraTbl.Visible = True
  txtTbl = ""
  txtTbl.SetFocus
  fraTbl.Caption = "Table"
  lstFld.Clear
  lstType.Clear
  lstInd.Clear
End Sub

Private Sub cmdCreateDb_Click()
On Error GoTo Err_Db
  cdlSave.Filter = "MS-Access Database Files|*.mdb"
  cdlSave.ShowSave
  'MsgBox cdlSave.FileName
  If cdlSave.FileName <> "" And Right(cdlSave.FileName, 4) <> ".mdb" Then
    MsgBox "There should be no file extention OR only 'mdb' file extention is allowed"
    Exit Sub
  End If
  'End
  If cdlSave.FileName <> "" Then
    Set Ws = DBEngine.Workspaces(0)
    Set Db = Ws.CreateDatabase(cdlSave.FileName, dbLangGeneral)
    FraDb.Visible = False
    fraTbl.Visible = True
    lblDb.Caption = "Database Name: " & cdlSave.FileName
    txtTbl.SetFocus
  End If
  Exit Sub
  
Err_Db:
  Select Case Err.Number
    Case 3204
      Dim nResp As Long
      nResp = MsgBox("Database already exists. Delete it or no", vbYesNo)
      If nResp = vbYes Then
        Kill (cdlSave.FileName) '& ".mdb")
        Resume
      End If
  End Select
End Sub

Private Sub cmdCreateTbl_Click()
  If txtTbl.Text = "" Then
    MsgBox "Enter Table name"
    Exit Sub
  End If
  Dim Lv As Integer
  'MsgBox Db.TableDefs.Count
  For Lv = 0 To Db.TableDefs.Count - 1
    'MsgBox Db.TableDefs(Lv).Name
    If txtTbl.Text = Db.TableDefs(Lv).Name Then
      MsgBox "Table already exits!", vbCritical
      Exit Sub
    End If
  Next Lv
  Set Tbl = Db.CreateTableDef(txtTbl)
  fraTbl.Visible = False
  fraFld.Visible = True
  cmdFinish.Enabled = True
  cmdNewDb.Enabled = True
  txtFld.SetFocus
  lblTbl.Caption = "Table Name: " & txtTbl
  lblTbl.Visible = True
End Sub

Private Sub cmdExit_Click()
  'End
  Unload Me
End Sub

Private Sub cmdFinish_Click()
On Error GoTo Err_NoFld
  Db.TableDefs.Append Tbl
  'End
  Unload Me
  
Err_NoFld:
  Select Case Err.Number
    Case 3264
      MsgBox "Add at least one field"
    Case Else
      MsgBox "An error occures i.e " & Err.Description
  End Select
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
  frmHelp.WebBrowser1.Navigate App.Path & "\help.htm#frmCreateDB"
End Sub

Private Sub cmdNewDb_Click()
On Error GoTo Err_NewDb
  lblTbl.Visible = False
  Db.TableDefs.Append Tbl
  Ws.Close
  FraDb.Visible = True
  fraTbl.Visible = False
  fraFld.Visible = False
  cmdFinish.Enabled = False
  cmdNewDb.Enabled = False
  'txtDb = ""
  txtTbl = ""
  FraDb.Caption = "Database"
  Exit Sub
  
Err_NewDb:
  Select Case Err.Number
    Case 3264
      MsgBox "Add at least one field"
    Case Else
      MsgBox "An error occures i.e " & Err.Description
  End Select
End Sub

Private Sub cmdNewTbl_Click()
On Error GoTo Err_NewTbl
  Db.TableDefs.Append Tbl
  cmdFinish.Enabled = False
  fraFld.Visible = False
  fraTbl.Visible = True
  txtTbl = ""
  txtTbl.SetFocus
  fraTbl.Caption = "Table"
  lstFld.Clear
  lstType.Clear
  lstInd.Clear
  Exit Sub
  
Err_NewTbl:
 Select Case Err.Number
    Case 3264
      MsgBox "Add at least one field"
    Case Else
      MsgBox "An error occures i.e " & Err.Description
  End Select
End Sub

Private Sub Form_Load()

  cboType.AddItem "Text"
  cboType.ItemData(0) = 10
  
  cboType.AddItem "Integer"
  cboType.ItemData(1) = 3
  
  cboType.ListIndex = 0
End Sub

Private Sub lstFld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuDel
  End If
End Sub

Private Sub mnuDelFld_Click()
  If lstFld.ListIndex > -1 Then
    'MsgBox "Field name: " & lstFld.List(lstFld.ListIndex) & " Will be delete
    Tbl.Fields.Delete lstFld.List(lstFld.ListIndex)
    
    MsgBox lstFld.List(lstFld.ListIndex) & ", Field is deleted"
    If lstInd.Selected(lstFld.ListIndex) = True Then
      Tbl.Indexes.Delete lstFld.List(lstFld.ListIndex)
    End If
    lstInd.RemoveItem lstFld.ListIndex
    lstType.RemoveItem lstFld.ListIndex
    lstFld.RemoveItem lstFld.ListIndex
  Else
    MsgBox "First select a field"
  End If
End Sub
