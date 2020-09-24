VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MySQL Dummy - My Learning Visual Basic with MySQL"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   6960
      TabIndex        =   14
      Top             =   120
      Width           =   1335
      Begin VB.TextBox Text4 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewTable 
         Caption         =   "Create Table and New data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSqlAdd 
         Caption         =   "Add New"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSqlDel 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSqlExecuteUpd 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Enter"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4000
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Search"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   375
         Left            =   120
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   375
         Left            =   120
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Execute Example"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001A63EC&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdSqlUpdate 
         Caption         =   "SQL Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   140
         Picture         =   "Main.frx":0000
         Top             =   3630
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   375
         Left            =   120
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "RecordSet Example"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001A63EC&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   3840
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         RowDividerStyle =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1054
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1054
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EEE4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EEE4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "Main.frx":056A
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EEE4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   550
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Product Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB4D68&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB4D68&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Product ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB4D68&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========Last Update Search Function 8 May 2004 ===========================
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=53471&lngWId=1
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=53570&lngWId=1
'MySQL Server configuration
'username : root
'password:
'database: test
'port: 3306
'======================================
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim Rx As Long
Dim AddNewStatus As Boolean
Dim xCount As Integer
Dim db_name, db_server, db_port, db_user, db_pass, constr As String

Private Sub Form_Load()
    On Error GoTo DBerror
    db_name = "test"
    db_server = "localhost"
    db_port = ""    'default port is 3306
    db_user = "root"
    db_pass = ""
    xExt = ".jpg"
    'ConnServer ' Open with ODBC in Control Panel
    OpenServer ' Open without ODBC in Control Panel
    Rx = 0
    ShowData
    ShowGrid
    frmPic.Show
    DisableX frmPic
    ShowImage Text1
    Exit Sub
DBerror:
    CreateMyTable
    ShowData
    ShowGrid
End Sub

Private Sub ConnServer()
  'connect to MySQL server using MySQL ODBC 3.51 Driver
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=localhost;" _
                        & " DATABASE=TEST;" _
                        & "UID=root;PWD=; OPTION=3"
conn.Open
End Sub

Private Sub OpenServer() 'Connect MySQL Server Without ODBC setup
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
End Sub

Private Sub cmdAdd_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    AddNewStatus = True
    ChangeMode True, False, &H80000005
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    CtrlButton False
End Sub

Private Sub cmdCancel_Click()
    ChangeMode False, True, &HF1EEE4
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    CtrlButton True
    Label6.Caption = ""
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    AddNewStatus = False
    ShowData
    Set DataGrid1.DataSource = rs1
End Sub

Private Sub cmdDelete_Click()
    DelData
    Rx = Rx - 1
    ShowData
    Set DataGrid1.DataSource = rs1
    ShowGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdEdit_Click()
    Label6.Caption = "You can choose left Update or Right Update "
    ChangeMode False, False, &H80000005
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    CtrlButton False
    Shape1.Visible = True
    Shape3.Visible = True
    AddNewStatus = False
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    Rx = 0
    ShowData
    ShowGrid
    DataGrid1.Row = Rx
    ShowImage Text1
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowData
    ShowGrid
    DataGrid1.Row = Rx
    ShowImage Text1
End Sub

Private Sub cmdSave_Click()
    If DupCheck(Text1.Text) = True And AddNewStatus = True Then
        MsgBox "Duplicate Record ", , "Warning"
        ShowData
        Set DataGrid1.DataSource = rs1
        DataGrid1.Row = Rx
    Else
        SaveData
    End If
        SetButton
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowData
    ShowGrid
    DataGrid1.Row = Rx
    ShowImage Text1
End Sub

Private Sub cmdPrev_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowData
    ShowGrid
    DataGrid1.Row = Rx
    ShowImage Text1
End Sub

Private Sub cmdSqlAdd_Click()
    If DupCheck(Text1.Text) = True Then
        MsgBox "Duplicate Record ", , "Warning"
    Else
        conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('" & Text1 & "','" & Text2 & "','" & Text3 & "')"
    End If
    SetButton
    cmdCancel_Click
End Sub

Private Sub cmdNewTable_Click()
    CreateMyTable
    ShowGrid
End Sub

Private Sub cmdSqlDel_Click()
    conn.Execute "DELETE FROM MyTable Where ProductID = '" & Text1 & "'"
    Rx = Rx - 1
    ShowData
    Set DataGrid1.DataSource = rs1
    ShowGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdSqlExecuteUpd_Click()
    conn.Execute "UPDATE MyTable SET ProductName = '" & Text2.Text & "' ,ProductPrice = '" & Text3.Text & "' Where ProductID='" & Text1.Text & "'"
    ChangeMode False, True, &HF1EEE4
    ShowData
    Set DataGrid1.DataSource = rs1
    ShowGrid
    DataGrid1.Row = Rx
    CtrlButton True
    Shape1.Visible = False
    Shape3.Visible = False
    cmdCancel_Click
End Sub

Private Sub SaveData()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
            If AddNewStatus = True Then
                rs.Open "select * from MyTable", conn, adOpenStatic, adLockOptimistic, adCmdText
                AddNewStatus = False
                rs.AddNew
            Else
                sql = "select * from MyTable where ProductID = '" & Text1.Text & "'"
                rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
            End If
    rs!ProductID = Text1.Text
    rs!ProductName = Text2.Text
    rs!ProductPrice = Text3.Text
    rs.Update
    rs.Close
    Set rs = Nothing
    cmdSave.Enabled = False
End Sub

Private Sub cmdSqlUpdate_Click()
    Dim rsTemp As ADODB.Recordset
    Dim sql As String
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    sql = "UPDATE MyTable SET ProductName = '" & Text2.Text & "' , ProductPrice = '" & Text3.Text & "' Where ProductID='" & Text1.Text & "'"
    rsTemp.Open sql, conn, adOpenKeyset, adLockOptimistic
    ChangeMode False, True, &HF1EEE4
    ShowData
    Set DataGrid1.DataSource = rs1
    ShowGrid
    DataGrid1.Row = Rx
    CtrlButton True
End Sub

Private Sub DelData()
    Dim rsTemp As ADODB.Recordset
    Dim sql As String
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    sql = "DELETE FROM MyTable Where ProductID='" & Text1.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    rsTemp.Open sql, conn, adOpenKeyset, adLockOptimistic
    Set rsTemp = Nothing
End Sub

Private Sub ShowData()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT * FROM MyTable", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid1.DataSource = rs
    Text1.Text = rs!ProductID
    Text2.Text = rs!ProductName
    Text3.Text = rs!ProductPrice
    'Text3.Text = Format(rs!ProductPrice, "###,###,###,##0.00")
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ShowGrid()
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.LockType = adLockReadOnly
    rs1.Open "SELECT * FROM MyTable", conn
    Set DataGrid1.DataSource = rs1
End Sub

Private Sub CreateMyTable()
On Error GoTo ServerErr
    conn.Execute "DROP TABLE IF EXISTS MyTable"

    conn.Execute "CREATE TABLE MyTable " _
    & "(ProductID varchar(10) NOT NULL PRIMARY KEY," _
    & "ProductName varchar(255),ProductPrice INT," _
    & "file_name VARCHAR(64) NOT NULL," _
    & "file MEDIUMBLOB NOT NULL)", , adExecuteNoRecords

    conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('A01','MB1ASUS-1',1000)", , adExecuteNoRecords
    conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('B02','MB2GIGABYTE-2',2000)", , adExecuteNoRecords
    conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('C03','MB3SIS-3',3000)", , adExecuteNoRecords
    conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('D04','MB4AOPEN-4',4000)", , adExecuteNoRecords
    conn.Execute "INSERT INTO MyTable(ProductID,ProductName,ProductPrice) values('E05','Mainboard-5',5000)", , adExecuteNoRecords
    frmPic.Show
    Exit Sub
ServerErr:
    MsgBox "Can't connect to MySQL server or create new database please call program again !"
    
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    conn.Execute "DROP DATABASE IF EXISTS test"
    conn.Execute "Create Database " & Trim$("test"), , adExecuteNoRecords
    End
End Sub

Private Sub ChangeMode(AddMode As Boolean, LockMode As Boolean, BackColor As Variant)
    If AddMode = True Then
        Text1.Locked = False
        Text1.BackColor = &H80000005
    Else
        Text1.Locked = True
        Text1.BackColor = &HF1EEE4
    End If

    Text2.Locked = LockMode
    Text3.Locked = LockMode
    Text2.BackColor = BackColor
    Text3.BackColor = BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs1.Close
    Set rs1 = Nothing
    conn.Close
    Set conn = Nothing
    End
End Sub

Private Sub Text1_Change()
    CheckEmpty
End Sub

Private Sub Text2_Change()
    CheckEmpty
End Sub

Private Sub Text3_Change()
    CheckEmpty
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc(vbCr)
            KeyAscii = 0
        Case 8, 46
        Case 47 To 58
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CheckEmpty()
    If AddNewStatus = True Then
        If Len(Text1.Text) = 0 Or Len(Text2.Text) = 0 Or Len(Text3.Text) = 0 Then
            Label6.Caption = ""
            Shape1.Visible = False
            Shape2.Visible = False
            cmdSave.Enabled = False
            cmdSqlAdd.Enabled = False
        Else
            Label6.Caption = "You can choose Save or Add "
            Shape1.Visible = True
            Shape2.Visible = True
            cmdSave.Enabled = True
            cmdSqlAdd.Enabled = True
        End If
    End If
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    sql = "select * from MyTable where ProductID = '" & chkID & "'"
    rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If chkID = rs!ProductID Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function CodeSearch(xSearch As String) As Boolean
On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    sql = "select * from Stock where ProductID LIKE '" & xSearch & "%'"
    rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    'rs.Find "ProductID LIKE '" & xSearch & "*'"
    
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If IsNull(rs!ProductName) Then
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        CodeSearch = False
    Else
        Text1.Text = rs!ProductID
        Text2.Text = rs!ProductName
        Text3.Text = rs!ProductPrice
         'Text3.Text = Format(rs!ProductPrice, "###,###,###,##0.00")
        CodeSearch = True
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function NameSearch(xSearch As String) As Boolean
On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    'sql = "select * from Stock"
    sql = "select * from Stock where ProductName LIKE '" & xSearch & "%'"
    rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
    'rs.Find "ProductName LIKE '" & xSearch & "*'"
    
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If IsNull(rs!ProductName) Then
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        NameSearch = False
    Else
        Text1.Text = rs!ProductID
        Text2.Text = rs!ProductName
        Text3.Text = rs!ProductPrice
         'Text3.Text = Format(rs!ProductPrice, "###,###,###,##0.00")
        NameSearch = True
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub CtrlButton(bMode As Boolean)
    cmdNext.Enabled = bMode
    cmdPrev.Enabled = bMode
    cmdFirst.Enabled = bMode
    cmdLast.Enabled = bMode
End Sub

Private Sub SetButton()
    ChangeMode False, True, &HF1EEE4
    cmdSave.Enabled = False
    cmdSqlAdd.Enabled = False
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    CtrlButton True
    Label6.Caption = ""
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
End Sub

Private Sub Text4_Change()
    'DataSearch Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CodeSearch(Text4) = False Then
            NameSearch Text4
        End If
        Text4 = ""
    End If
End Sub

Private Sub Label8_Click()
    CodeSearch Text4
    Text4 = ""
End Sub

Private Sub ShowImage(pID As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim mystream As ADODB.Stream
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
conn.CursorLocation = adUseClient
Dim sql As String
sql = "Select * from MyTable WHERE MyTable.ProductID = '" & pID & "'"
rs.Open sql, conn
xExt = Right(rs!file_name, 4)
mystream.Open
mystream.Write rs!file
mystream.SaveToFile App.Path & "\Temp" & xExt, adSaveCreateOverWrite

If mystream.Size > 0 Then
frmPic.Image1.Visible = True
frmPic.Image1.Picture = LoadPicture(App.Path & "\Temp" & xExt)
Kill App.Path & "\Temp" & xExt
Else
frmPic.Image1.Visible = False
End If
mystream.Close
rs.Close
End Sub

Public Sub ImageUpdate(pID As String, PicPath As String, PicName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim mystream As ADODB.Stream
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
Dim sql As String

sql = "Select * from MyTable WHERE MyTable.ProductID = '" & pID & "'"
rs.Open sql, conn, adOpenStatic, adLockOptimistic

mystream.Open
mystream.LoadFromFile PicPath
rs!file_name = PicName
rs!file = mystream.Read
rs.Update

mystream.Close
rs.Close
End Sub

