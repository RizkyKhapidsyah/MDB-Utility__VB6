VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form DbDefs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDB Viewer"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DbDefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame6 
      Caption         =   "Print Table Definition.."
      Height          =   2415
      Left            =   2790
      TabIndex        =   15
      Top             =   780
      Width           =   2475
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   270
         TabIndex        =   19
         Top             =   1770
         Width           =   1905
      End
      Begin VB.CheckBox chkFldSize 
         Caption         =   "Field Size"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1140
         Width           =   2205
      End
      Begin VB.CheckBox chkFldType 
         Caption         =   "Field Type"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2205
      End
      Begin VB.CheckBox chkFldName 
         Caption         =   "Field Name"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   2205
      End
      Begin VB.Shape Shape3 
         Height          =   645
         Left            =   120
         Top             =   1620
         Width           =   2265
      End
   End
   Begin TabDlg.SSTab SS 
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   3210
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Table View"
      TabPicture(0)   =   "DbDefs.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Generate SQL"
      TabPicture(1)   =   "DbDefs.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(1)=   "txtSQL"
      Tab(1).Control(2)=   "cmdSelect"
      Tab(1).Control(3)=   "cmdUpdate"
      Tab(1).Control(4)=   "cmdInsert"
      Tab(1).Control(5)=   "cmdWhere"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdWhere 
         Caption         =   "Where Clause"
         Height          =   405
         Left            =   -67710
         TabIndex        =   25
         Top             =   2340
         Width           =   1515
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert"
         Height          =   405
         Left            =   -70050
         TabIndex        =   23
         Top             =   2340
         Width           =   1515
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   405
         Left            =   -72480
         TabIndex        =   22
         Top             =   2340
         Width           =   1515
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   405
         Left            =   -74700
         TabIndex        =   21
         Top             =   2340
         Width           =   1515
      End
      Begin VB.TextBox txtSQL 
         Height          =   1725
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   390
         Width           =   8925
      End
      Begin VB.Frame Frame5 
         Caption         =   "Table.."
         Height          =   2505
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   8955
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   765
            Left            =   2220
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   750
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   345
            Left            =   3780
            TabIndex        =   12
            Top             =   1980
            Width           =   1365
         End
         Begin VB.PictureBox DB1 
            Height          =   1605
            Left            =   90
            ScaleHeight     =   1545
            ScaleWidth      =   8685
            TabIndex        =   13
            Top             =   180
            Width           =   8745
         End
         Begin VB.Shape Shape1 
            Height          =   555
            Left            =   90
            Top             =   1860
            Width           =   8745
         End
      End
      Begin VB.Shape Shape2 
         Height          =   765
         Left            =   -74910
         Top             =   2160
         Width           =   8955
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controls.."
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   9195
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   270
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Table Definition.."
      Height          =   2415
      Left            =   5310
      TabIndex        =   6
      Top             =   780
      Width           =   3885
      Begin VB.CheckBox chkSelAll 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1950
         Width           =   3615
      End
      Begin MSComctlLib.ListView lstFlds 
         Height          =   1695
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2990
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fields"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2025
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   3572
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Table.."
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   2745
      Begin VB.ListBox lstTable 
         Height          =   2010
         ItemData        =   "DbDefs.frx":047A
         Left            =   120
         List            =   "DbDefs.frx":047C
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Database.."
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.CommandButton cmdSelDB 
         Caption         =   "..."
         Height          =   315
         Left            =   8490
         TabIndex        =   3
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox txtDB 
         Height          =   345
         Left            =   1500
         TabIndex        =   1
         Top             =   300
         Width           =   6975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Database"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1500
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "DbDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'References Taken:
'
'Visual Basic For Applications
'Visual Basic runtime objects and procedures
'Visual Basic objects and procedures
'OLE Automation
'Microsoft DAO 2.5/3.51 Compatibility Library
'Microsoft Scripting Runtime
'
'Components Used:
'Microsoft Common Dialog Control
'Microsoft Data Bound Grid Control
'Microsoft Datagrid Control 6.0 (OLEDB)
'Microsoft Flexgrid Control 6.0
'Microsoft Tabbed Dialog Control 6.0
'Microsoft Windows Common Controls 6.0

Option Explicit
Dim DB As Database
Dim TD As TableDef
Dim rs As Recordset
Dim lstItem As ListItem


Private Sub cmdEdit_Click()
On Error Resume Next
Dim ddd As Integer
'If DB1.AllowUpdate = False Then
 '   ddd = MsgBox("Any changes you make here will be directly reflected in the database. Are you sure you want to continue", vbYesNo)
  '  If ddd = vbYes Then
   '     DB1.AllowUpdate = True
   ' Else
   '     Exit Sub
   ' End If
'cmdEdit.Enabled = False
'End If


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
If chkList() = False Then Exit Sub
GenerateSQL cmdInsert.Caption
cmdWhere.Enabled = False
End Sub

Private Sub cmdPrint_Click()
If chkFldName.Value = 0 And chkFldSize.Value = 0 And chkFldType.Value = 0 Then
    MsgBox "Please select atleast one field for printing."
    chkFldName.SetFocus
    Exit Sub
End If
Call PrintTableDef
End Sub

Private Sub cmdSelDB_Click()
CD1.InitDir = App.Path
CD1.Filter = "Microsoft Access Database(*.mdb)|*.mdb"
CD1.ShowOpen
txtDB.Text = CD1.FileName
If txtDB.Text <> "" Then
    Call PopList
    lstTable.Enabled = True
    cmdPrint.Enabled = True
    chkFldName.Enabled = True
    chkFldType.Enabled = True
    chkFldSize.Enabled = True

Else
    cmdSelDB.SetFocus
    Exit Sub
End If
End Sub

Public Sub PopList()
Dim Ctr As Integer

Set DB = OpenDatabase(txtDB.Text)
lstTable.Clear
For Ctr = 0 To DB.TableDefs.Count - 1
    Set TD = DB.TableDefs(Ctr)
    lstTable.AddItem TD.Name
Next Ctr

Set TD = Nothing

        
End Sub

Public Sub PopFlexGrid()
Dim CtrFld As Integer
Dim lstItem As ListItem
Dim I As Integer
lstFlds.ListItems.Clear
lstFlds.View = lvwReport

Set TD = DB.TableDefs(lstTable.ListIndex)

MSF1.Rows = TD.Fields.Count + 1

For CtrFld = 0 To TD.Fields.Count - 1
    MSF1.TextMatrix((CtrFld + 1), 0) = TD.Fields(CtrFld).Name
    MSF1.TextMatrix((CtrFld + 1), 1) = GetActualVal((TD.Fields(CtrFld).Type))
    MSF1.TextMatrix((CtrFld + 1), 2) = TD.Fields(CtrFld).Size
Next CtrFld
Call PopListView
End Sub

Private Sub cmdSelect_Click()
If chkList() = False Then Exit Sub
GenerateSQL cmdSelect.Caption
cmdWhere.Enabled = True
Call ClearList
End Sub

Private Sub cmdUpdate_Click()
If chkList() = False Then Exit Sub
GenerateSQL cmdUpdate.Caption
cmdWhere.Enabled = True
Call ClearList
End Sub

Private Sub cmdWhere_Click()
Dim SelCnt As Integer
Dim I As Integer
SelCnt = 0

For I = 1 To lstFlds.ListItems.Count
    If lstFlds.ListItems(I).Checked = True Then
        SelCnt = SelCnt + 1
    End If
Next I

If SelCnt = 0 Then
    MsgBox "Please select fields for building where clause.", vbInformation
    lstFlds.SetFocus
    'chkList = False
    Exit Sub
End If

Call Update_Where
cmdWhere.Enabled = False

End Sub

Private Sub Form_Load()
cmdEdit.Enabled = False
DB1.AllowUpdate = False
lstTable.Enabled = False
chkFldName.Enabled = False
chkFldType.Enabled = False
chkFldSize.Enabled = False
cmdPrint.Enabled = False
txtDB.Locked = True
chkSelAll.Caption = "Select All"
MSF1.FormatString = "^ Field Name |^   Dataype   |^Size"
cmdSelect.Enabled = False
cmdInsert.Enabled = False
cmdUpdate.Enabled = False
cmdWhere.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not IsObject(DB) Then
DB.Close
Set DB = Nothing
End If
End Sub

Private Sub lstTable_Click()
Call PopFlexGrid
Call PopDBGrid
Frame3.Caption = "Table Definition for : " & lstTable.Text
Frame5.Caption = "Table : " & lstTable.Text
DB1.AllowUpdate = False
chkSelAll.Value = 0
chkFldName.Value = 1
chkFldSize.Value = 1
chkFldType.Value = 1
cmdEdit.Enabled = True
txtSQL.Text = ""
cmdSelect.Enabled = True
cmdInsert.Enabled = True
cmdUpdate.Enabled = True
cmdWhere.Enabled = False
End Sub

Public Function GetActualVal(TypeNo As Integer) As String

Select Case TypeNo

Case 1
    GetActualVal = "Yes/No"
Case 2
    GetActualVal = "Byte"
Case 3
    GetActualVal = "Integer"
Case 4
    GetActualVal = "Long"
Case 5
    GetActualVal = "Currency"
Case 6
    GetActualVal = "Single"

Case 7
    GetActualVal = "Double"
Case 8
    GetActualVal = "Date"
Case 9
Case 10
    GetActualVal = "Text"
Case 11
    GetActualVal = "OLE Object"
Case 12
    GetActualVal = "Memo"
Case 13
Case 14
Case 15
    GetActualVal = "REPLICATIONID"
Case 16
Case 17

End Select

End Function

Public Sub PopDBGrid()

'Set rs = DB.OpenRecordset("select * from " & lstTable.Text)
Data1.DatabaseName = txtDB.Text
Data1.RecordSource = "select * from " & lstTable.Text
Data1.Refresh

End Sub

Public Sub PrintTableDef()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim I As Long
Dim str As String

Dim fld1 As String * 25
Dim fld2 As String * 25
Dim fld3 As String * 25

fld1 = Space$(25)
fld2 = Space$(25)
fld3 = Space$(25)

str = ""
Dim file As TextStream

Set file = fso.CreateTextFile(App.Path & "\" & lstTable.Text & ".rtf", 2)

file.WriteLine "Table Name : " & lstTable.Text
file.WriteLine "**************************************************************"
file.WriteBlankLines (2)



For I = 0 To lstFlds.ListItems.Count - 1
    If chkFldName.Value = 1 Then
        fld1 = StrConv(MSF1.TextMatrix(I, 0), vbUpperCase)
    Else
        fld1 = ""
    End If
    If chkFldType.Value = 1 Then
        fld2 = StrConv(MSF1.TextMatrix(I, 1), vbUpperCase)
    Else
        fld2 = ""
    End If
    If chkFldSize.Value = 1 Then
        fld3 = StrConv(MSF1.TextMatrix(I, 2), vbUpperCase)
    Else
        fld3 = ""
    End If
    
    If chkFldName.Value = 1 Then
        str = str & fld1
    End If
    If chkFldType.Value = 1 Then
        str = str & fld2
    End If
    If chkFldSize.Value = 1 Then
        str = str & fld3
    End If

    file.WriteLine str
    str = ""
    
    If I = 0 Then
        file.WriteLine "-------------------------------------------------------------"
    End If
Next I

file.WriteLine "-------------------------------------------------------------"

file.Close

Set file = Nothing
Set fso = Nothing

MsgBox "File written to : " & App.Path & "\" & lstTable.Text & ".rtf"

End Sub


Public Sub PopListView()
Dim CtrFld As Integer
lstFlds.ListItems.Clear
lstFlds.View = lvwReport

For CtrFld = 0 To TD.Fields.Count - 1
    Set lstItem = lstFlds.ListItems.Add(CtrFld + 1, , TD.Fields(CtrFld).Name)
    lstItem.SubItems(1) = GetActualVal((TD.Fields(CtrFld).Type))
    lstItem.SubItems(2) = TD.Fields(CtrFld).Size
Next CtrFld

End Sub

Private Sub SS_Click(PreviousTab As Integer)
Select Case PreviousTab
    Case 0
    Case 1
End Select
End Sub

Private Sub chkSelAll_Click()
Dim I As Integer
If chkSelAll.Value = 1 Then
    For I = 1 To lstFlds.ListItems.Count
        lstFlds.ListItems.Item(I).Checked = True
    Next I
    chkSelAll.Caption = "Deselect All"
Else
    For I = 1 To lstFlds.ListItems.Count
        lstFlds.ListItems.Item(I).Checked = False
    Next I
    chkSelAll.Caption = "Select All"
End If
End Sub

Public Sub GenerateSQL(SQL As String)

Dim strSQL As String
Dim tempSQL As String
Dim appendStr As String
Dim ValSQL As String
Dim liCtr As Integer

txtSQL.Text = ""

For liCtr = 1 To lstFlds.ListItems.Count
    If lstFlds.ListItems(liCtr).Checked = True Then
        appendStr = appendStr & " " & lstFlds.ListItems(liCtr).Text & ","
    End If
Next

tempSQL = Mid(appendStr, 1, Len(appendStr) - 1)




Select Case SQL

Case "Select"

    
    strSQL = "Select " & tempSQL & " from " & lstTable.Text
    txtSQL.Text = Chr(34) & strSQL & Chr(34)

Case "Update"
    For liCtr = 1 To lstFlds.ListItems.Count
        If lstFlds.ListItems(liCtr).Checked = True Then
            If lstFlds.ListItems(liCtr).SubItems(1) = "Text" Or lstFlds.ListItems(liCtr).SubItems(1) = "Memo" Or lstFlds.ListItems(liCtr).SubItems(1) = "Date" Then
                ValSQL = ValSQL & Space(1) & lstFlds.ListItems(liCtr).Text & " = " & "'" & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & "',"
            ElseIf lstFlds.ListItems(liCtr).SubItems(1) = "Integer" Or lstFlds.ListItems(liCtr).SubItems(1) = "Long" Or lstFlds.ListItems(liCtr).SubItems(1) = "Double" Or lstFlds.ListItems(liCtr).SubItems(1) = "Currency" Then
                ValSQL = ValSQL & Space(1) & lstFlds.ListItems(liCtr).Text & " = " & "" & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & ","
'            ElseIf lstFlds.ListItems(liCtr).SubItems(1) Then
'                ValSQL = ValSQL & "'" & lstFlds.ListItems(liCtr).Text & " = " & "'" & lstFlds.ListItems(liCtr).Text & "'" & vbCrLf & ","
            End If
        End If
    Next
    
    ValSQL = Mid(ValSQL, 1, Len(ValSQL) - 1)
    strSQL = "Update " & lstTable.Text & " set " & ValSQL
    txtSQL.Text = Chr(34) & strSQL & Chr(34)
    
Case "Insert"
    For liCtr = 1 To lstFlds.ListItems.Count
        If lstFlds.ListItems(liCtr).Checked = True Then
            If lstFlds.ListItems(liCtr).SubItems(1) = "Text" Or lstFlds.ListItems(liCtr).SubItems(1) = "Memo" Or lstFlds.ListItems(liCtr).SubItems(1) = "Date" Then
                ValSQL = ValSQL & "'" & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & "',"
            ElseIf lstFlds.ListItems(liCtr).SubItems(1) = "Integer" Or lstFlds.ListItems(liCtr).SubItems(1) = "Long" Or lstFlds.ListItems(liCtr).SubItems(1) = "Double" Or lstFlds.ListItems(liCtr).SubItems(1) = "Currency" Then
                ValSQL = ValSQL & " " & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & ","
Rem            ElseIf lstFlds.ListItems(liCtr).SubItems(1) Then
Rem                ValSQL = ValSQL & "'" & lstFlds.ListItems(liCtr).Text & "',"
            End If
        End If
    Next
    
    ValSQL = Mid(ValSQL, 1, Len(ValSQL) - 1)

    strSQL = "Insert into " & lstTable.Text & " (" & tempSQL & ") " & "values " & "(" & ValSQL & ")"
    txtSQL.Text = Chr(34) & strSQL & Chr(34)
End Select

End Sub

Public Function chkList() As Boolean
        Dim SelCnt As Integer
        Dim I As Integer
        SelCnt = 0
        For I = 1 To lstFlds.ListItems.Count
            If lstFlds.ListItems(I).Checked = True Then
                SelCnt = SelCnt + 1
            End If
        Next I
        
        If SelCnt = 0 Then
            MsgBox "Please select fields for building query", vbInformation
            lstFlds.SetFocus
            chkList = False
            Exit Function
        End If
        
chkList = True

End Function

Public Sub ClearList()
Dim I As Integer
For I = 1 To lstFlds.ListItems.Count
    lstFlds.ListItems.Item(I).Checked = False
Next I
chkSelAll.Value = 0
chkSelAll.Caption = "Select All"
lstFlds.SetFocus
End Sub

Public Sub Update_Where()
Dim tempSQL As String
Dim appendStr As String
Dim ValSQL As String
Dim liCtr As Integer
Dim strSQL As String

For liCtr = 1 To lstFlds.ListItems.Count
    If lstFlds.ListItems(liCtr).Checked = True Then
        If lstFlds.ListItems(liCtr).SubItems(1) = "Text" Or lstFlds.ListItems(liCtr).SubItems(1) = "Memo" Or lstFlds.ListItems(liCtr).SubItems(1) = "Date" Then
            ValSQL = ValSQL & Space(1) & lstFlds.ListItems(liCtr).Text & " = " & "'" & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & "',"
        ElseIf lstFlds.ListItems(liCtr).SubItems(1) = "Integer" Or lstFlds.ListItems(liCtr).SubItems(1) = "Long" Or lstFlds.ListItems(liCtr).SubItems(1) = "Double" Or lstFlds.ListItems(liCtr).SubItems(1) = "Currency" Then
            ValSQL = ValSQL & Space(1) & lstFlds.ListItems(liCtr).Text & " = " & "" & Chr(34) & Chr(38) & Space(1) & lstFlds.ListItems(liCtr).Text & Space(1) & Chr(38) & Chr(34) & ","
        End If
    End If
Next
    
ValSQL = Mid(ValSQL, 1, Len(ValSQL) - 1)
strSQL = " where " & ValSQL
txtSQL.Text = txtSQL.Text & Space(1) & Chr(38) & Space(1) & Chr(34) & strSQL & Chr(34)

End Sub
