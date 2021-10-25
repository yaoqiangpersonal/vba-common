VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "D:\vba\purchase_details.VBAProject"
Option Explicit
Private dataBase As String
Private username As String
Private password As String
Private conn As Connection
Private sql As String
Private serverip As String
Private rs As Recordset
Public Property Let letServerip(ip As String)
    serverip = ip
End Property

Public Property Let letDataBase(strDataBase As String)
    dataBase = strDataBase
End Property

Public Property Let letSql(s As String)
    sql = s
    Debug.Print sql
End Property

Public Property Let letUsername(u As String)
    username = u
End Property

Public Property Let letPassword(p As String)
    password = p
End Property

Public Property Get getConnection()
    Set getConnection = conn
End Property

'获取Excel数据库
Public Sub getExcelConnection()
    Set conn = CreateObject("adodb.connection")
    With conn
        .Provider = "Microsoft.ace.OLEDB.12.0"
        .ConnectionString = "Data Source=" & dataBase & ";Extended Properties=Excel 12.0"
        .Open
    End With
End Sub
'获取Mysql数据库
Public Sub getMysqlConnection()
    Set conn = CreateObject("adodb.connection")
    With conn
        .ConnectionString = "DRIVER=MySQL ODBC 5.3 Unicode Driver;SERVER=" & serverip & ";Database=" & dataBase & ";Uid=" & username & ";Pwd=" & password
        .Open
    End With
End Sub
'获取SqlServver数据库
Public Sub getSqlServerConnection()
    Set conn = CreateObject("adodb.connection")
    With conn
        .ConnectionString = "Provider=SQLOLEDB;Initial Catalog=" & dataBase & ";User ID=" & username & ";Password=" & password & ";Data Source=" & serverip
        'Debug.Print "Provider=SQLOLEDB;Initial Catalog=" & dataBase & ";User ID=" & user & ";Password=" & password & ";Data Source=" & serverip
        .Open
    End With
End Sub
'获取Access数据库
Public Sub getAccessConnection()
    Set conn = CreateObject("adodb.connection")
    With conn
        .Provider = "Microsoft.ace.OLEDB.12.0"
        .ConnectionString = "Data Source=" & dataBase
        .Open
    End With
End Sub
