VERSION 5.00
Begin VB.Form Data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Information"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   2955
   LinkTopic       =   "Data"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   8
      Left            =   75
      TabIndex        =   9
      Top             =   3465
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   7
      Left            =   75
      TabIndex        =   8
      Top             =   3075
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   6
      Left            =   75
      TabIndex        =   7
      Top             =   2670
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   5
      Left            =   75
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   4
      Left            =   75
      TabIndex        =   5
      Top             =   1890
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   3
      Left            =   75
      TabIndex        =   4
      Top             =   1500
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox Categories 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox DataName 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Menu menuData 
      Caption         =   "DataBase"
      Begin VB.Menu menuCategory 
         Caption         =   "New Category"
      End
      Begin VB.Menu menuDelCategory 
         Caption         =   "Delete Category"
      End
      Begin VB.Menu menuRecord 
         Caption         =   "New Record"
      End
      Begin VB.Menu menuDelRecord 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu menuField 
         Caption         =   "New Field"
      End
   End
End
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Counter As Integer
Dim i As Integer
Dim DB As Database
Dim Tbl As TableDef
Dim Fld As Field
Dim RS As Recordset
Private Sub AddField()
If DataName.ListCount = 0 Then Exit Sub
For i = 1 To 8
   If Not Value(i).Visible Then
   Value(i).Visible = True
   Value(i).SetFocus
   Exit Sub
   End If
Next
End Sub

Private Sub AddRecord()
If Categories = "" Then Exit Sub
Dim Temp As String
Temp = InputBox("Enter Name", "ADD NEW RECORD")
Temp = Trim(Temp)
If Temp = "" Then Exit Sub
Set RS = DB.OpenRecordset(Categories)
With RS
If .RecordCount <> 0 Then
For Counter = 1 To .RecordCount
    If .Fields("Name") = Temp Then
    MsgBox "An entry by this name already exists."
        For i = 0 To DataName.ListCount - 1
        DataName.ListIndex = i
        If DataName = Temp Then Exit For
        Next
    DataName_Click
    Exit Sub
    End If
.MoveNext
If .EOF Then Exit For
Next
End If

DataName.AddItem Temp
DataName.ListIndex = DataName.ListCount - 1
DataName.Visible = True
ClearBoxes
.AddNew
.Fields("Name") = DataName
For i = 1 To 8
.Fields("Value" & i) = " "
Next
.Update
.Close
End With
DataName_Click
End Sub

Private Sub AddTable()
Dim Find As TableDef
Dim Temp As String
Temp = InputBox("Enter Name", "ADD NEW CATEGORY")
Temp = Trim(Temp)
If Temp = "" Then Exit Sub
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = Temp Then
      MsgBox "A category by that name already exists."
         For i = 0 To Categories.ListCount - 1
         Categories.ListIndex = i
            If Temp = i Then
            Set Find = Nothing
            Exit For
            End If
         Next
      Exit Sub
      End If
   Next
Set Tbl = DB.CreateTableDef(Temp)
Set Fld = Tbl.CreateField("Name", dbText, 100)
Tbl.Fields.Append Fld
For i = 1 To 8
Set Fld = Tbl.CreateField("Value" & i, dbText, 100)
Tbl.Fields.Append Fld
Next
DB.TableDefs.Append Tbl
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub

Private Sub DeleteRecord()
If Categories.ListCount = 0 Then Exit Sub
If DataName.ListCount = 0 Then Exit Sub
Dim Action As VbMsgBoxResult
Action = MsgBox("Do you really want to delete the records " & """" & DataName & """" & " and all the data it contains.", vbYesNo, "DELETE RECORDSET")
If Action = vbYes Then
Set RS = DB.OpenRecordset(Categories)
With RS
   If .RecordCount <> 0 Then
      For Counter = 1 To .RecordCount
         If .Fields("Name") = DataName Then
         .Delete
         Exit For
         End If
      .MoveNext
      If .EOF Then Exit For
      Next
   End If
.Close
End With
LoadRecords
End If
End Sub

Private Sub DeleteTable()
Dim Action As VbMsgBoxResult
If Categories.ListCount = 0 Then Exit Sub
Action = MsgBox("Do you really want to delete the category " & """" & Categories & """" & " and all the data it contains.", vbYesNo, "DELETE TABLE")
If Action = vbYes Then
DB.TableDefs.Delete Categories
LoadTables
End If
End Sub

Private Sub LoadTables()
Dim Temp As String
Categories.Clear
DataName.Clear
DataName.Visible = False
ClearBoxes
Set Tbl = New TableDef
For i = 0 To DB.TableDefs.Count - 1
Temp = DB.TableDefs(i).Name
   If Left(Temp, 4) <> "MSys" Then
   Categories.AddItem Temp
   End If
Next
   If Categories.ListCount > 0 Then
   Categories.ListIndex = 0
   Categories.Visible = True
   Else
   Categories.Visible = False
   End If
End Sub
Private Sub LoadRecords()
DataName.Clear
ClearBoxes
If Categories.ListCount > 0 Then
Set RS = DB.OpenRecordset(Categories)
With RS
   If .RecordCount <> 0 Then
      For Counter = 1 To .RecordCount
      DataName.AddItem .Fields("Name")
      .MoveNext
      If .EOF Then Exit For
      Next
   End If
.Close
End With
End If
If DataName.ListCount > 0 Then
DataName.ListIndex = 0
DataName.Visible = True
Else
DataName.Visible = False
End If
End Sub


Private Sub SaveData(Index As Integer)
If DataName.ListCount = 0 Then Exit Sub
Set RS = DB.OpenRecordset(Categories)
With RS
If .RecordCount <> 0 Then
For Counter = 1 To .RecordCount
   If .Fields("Name") = DataName Then
   .Edit
      If Trim(Value(Index)) = "" Then
      .Fields("Value" & Index) = " "
      Else
      .Fields("Value" & Index) = Trim(Value(Index))
      End If
   .Update
   Exit For
   End If
  .MoveNext
   If .EOF Then Exit For
Next
End If
.Close
End With
End Sub

Private Sub ClearBoxes()
For i = 1 To 8
Value(i) = ""
Value(i).Visible = False
Next
End Sub

Private Sub Categories_Click()
LoadRecords
End Sub



Private Sub menuCategory_Click()
AddTable
End Sub

Private Sub menuDelCategory_Click()
DeleteTable
End Sub

Private Sub menuDelRecord_Click()
DeleteRecord
End Sub


Private Sub menuField_Click()
AddField
End Sub

Private Sub menuRecord_Click()
AddRecord
End Sub




Private Sub DataName_Click()
Dim n As Integer
ClearBoxes
If Categories.ListCount = 0 Then Exit Sub
Set RS = DB.OpenRecordset(Categories)
With RS
If .RecordCount <> 0 Then
For Counter = 1 To .RecordCount
    If .Fields("Name") = DataName Then
       For n = 1 To .Fields.Count - 1
       Value(n) = Trim(.Fields("Value" & n))
          If Trim(Value(n)) <> "" Then
          Value(n).Visible = True
          End If
       Next
    Exit For
    End If
.MoveNext
If .EOF Then Exit For
Next
End If
.Close
End With
End Sub


Public Sub OpenDB()
If IsFile(App.Path + "\DataFile.MDB") Then GoTo SkipCreation
Set DB = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\DataFile.MDB", dbLangGeneral)
SkipCreation:
Set DB = OpenDatabase(App.Path + "\DataFile.MDB")
End Sub

Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer
On Error Resume Next
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
If Err Then
IsFile = False
Exit Function
End If
IsFile = True
Close #FileNumber
End Function
Private Sub Form_Load()
OpenDB
LoadTables
End Sub











Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DB.Close
Set Fld = Nothing
Set RS = Nothing
Set Tbl = Nothing
Set DB = Nothing
End Sub

Private Sub Value_LostFocus(Index As Integer)
SaveData Index
End Sub


