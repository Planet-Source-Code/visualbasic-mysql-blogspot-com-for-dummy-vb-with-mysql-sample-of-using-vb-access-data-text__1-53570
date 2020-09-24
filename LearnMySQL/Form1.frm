VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton myodbc_ado_Click 
      Caption         =   "myodbc_ado_Click"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub myodbc_ado_Click_Click()
Dim conn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim sql As String

  'connect to MySQL server using MySQL ODBC 3.51 Driver
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=localhost;" _
                        & " DATABASE=DB4SQL;" _
                        & "UID=root;PWD=; OPTION=3"

  conn.Open

  'create table
  'conn.Execute "DROP TABLE IF EXISTS my_ado"
  'conn.Execute "CREATE TABLE my_ado(id int not null primary key, name varchar(20)," _
                                 & "txt text, dt date, tm time, ts timestamp)"

  'direct insert
  'conn.Execute "INSERT INTO my_ado(id,name,txt) values(1,100,'venu')"
  'conn.Execute "INSERT INTO my_ado(id,name,txt) values(2,200,'MySQL')"
  'conn.Execute "INSERT INTO my_ado(id,name,txt) values(3,300,'Delete')"

  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseServer

  'fetch the initial table ..
  rs.Open "SELECT * FROM Stock", conn
    MsgBox rs!ProductID
    Debug.Print rs.RecordCount
    rs.MoveFirst
    Debug.Print String(50, "-") & "Initial my_ado Result Set " & String(50, "-")
    For Each fld In rs.Fields
      Debug.Print fld.Name,
      Next
      Debug.Print

      Do Until rs.EOF
      For Each fld In rs.Fields
      Debug.Print fld.Value,
      Next
      rs.MoveNext
      Debug.Print
    Loop
  rs.Close

  'rs insert
  rs.Open "select * from Stock", conn, adOpenDynamic, adLockOptimistic
  rs.AddNew
  rs!ProductID = "Monty"
  rs!ProductName = "Insert row"
  rs.Update
  rs.Close

  'rs update
  rs.Open "SELECT * FROM Stock"
  rs!ProductID = "update"
  rs!ProductName = "updated-row"
  rs.Update
  rs.Close

  'rs update second time..
  rs.Open "SELECT * FROM Stock"
  rs!ProductID = "update"
  rs!ProductName = "updated-second-time"
  rs.Update
  rs.Close

  'rs delete
  rs.Open "SELECT * FROM Stock"
  rs.MoveNext
  rs.MoveNext
  rs.Delete
  rs.Close

  'fetch the updated table ..
  rs.Open "SELECT * FROM Stock", conn
    Debug.Print rs.RecordCount
    rs.MoveFirst
    Debug.Print String(50, "-") & "Updated my_ado Result Set " & String(50, "-")
    For Each fld In rs.Fields
      Debug.Print fld.Name,
      Next
      Debug.Print

      Do Until rs.EOF
      For Each fld In rs.Fields
      Debug.Print fld.Value,
      Next
      rs.MoveNext
      Debug.Print
    Loop
  rs.Close
  conn.Close
End Sub
