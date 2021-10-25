VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBaseUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "D:\vba\vba-common"
Option Explicit
'操作数据库
Private record As Recordset
Private dataBase As New dataBase
Private sql As String
Private conn As Connection
Private Const sqlServerUsername = "sa"
Private Const mysqlUsername = "root"
Private Const mysqlPassword = "root123456"
Private Const accessPath = "D:\cgtools.mdb"
Public Property Let letDataBase(db As dataBase)
    Set dataBase = db
End Property

Public Property Let letSql(s)
    sql = s
End Property

Public Property Get getRecordSet()
     If checkRecordNotIsColse() Then
        Set getRecordSet = record
     End If
End Property

'初始化方法
Private Sub init()
    loadConnection
    loadRecordSet
    destroyDataBase
End Sub

'加载记录集
Private Sub loadRecordSet()
    Set record = conn.Execute(sql)
    If checkRecordNotIsColse() Then
        With record
            .Close
            .LockType = adLockBatchOptimistic
            .CursorLocation = adUseClientBatch
            .CursorType = adOpenStatic
            .Open
        End With
    End If
End Sub

'检验recordSet是否不为空
Function checkRecordNotIsColse()
    checkRecordNotIsColse = record.State <> 0
End Function

'加载链接
Private Sub loadConnection()
    Set conn = dataBase.getConnection
End Sub

'销毁dataBase
Private Sub destroyDataBase()
    Set dataBase = Nothing
End Sub

'根据IP后3位和密码获取Connection
Private Sub getSqlServerDataBase(ipLast, password)
    With dataBase
        .letDataBase = "mpdisk"
        .letServerip = "192.168.1." & ipLast
        .letUsername = sqlServerUsername
        .letPassword = password
        .getSqlServerConnection
    End With
End Sub

'获取Excel为数据库
Private Sub getExcelDataBase(path)
    With dataBase
        .letDataBase = path
        .getExcelConnection
    End With
End Sub

'获取mdb文件
Private Sub getAccessDataBase()
    With dataBase
        .letDataBase = accessPath
        .getAccessConnection
    End With

End Sub

'获取mysql的dataBase
Private Sub getMysqlDataBase(ipLast, dabaBaseName As String)
    With dataBase
        .letDataBase = dabaBaseName
        .letServerip = "192.168.1." & ipLast
        .letUsername = mysqlUsername
        .letPassword = mysqlPassword
        .getMysqlConnection
    End With
End Sub

'执行Sql获取其中数据，并默认放置在[a1]
Sub extractDataInR1C1()
    getFieldName [a1]
    executeDQLSql [a2]
End Sub

'链接sqlServer数据库，需要传入sql语句
Sub getDataBaseUtilForSqlServer(strSql, ipLast, password)
    getSqlServerDataBase ipLast, password
    sql = strSql
    init
End Sub


'链接Excel
Sub getDataBaseUtilForExcel(strSql, path)
    getExcelDataBase (path)
    sql = strSql
    init
End Sub

'链接Access
Sub getDataBaseUtilForAccess(strSql)
    getAccessDataBase
    sql = strSql
    init
End Sub

'链接mysql数据库，需要传入sql语句 和ip后3位
Sub getDataBaseUtilForMysql(strSql, ipLast, dabaBaseName As String)
    getMysqlDataBase ipLast, dabaBaseName
    sql = strSql
    init
End Sub

'添加新的淘宝asin
Sub saveOrUpdateAsin(sku, asin, site)
    If checkAsinIsExistInRecordset(asin) Then
        updateAsin sku, asin, site
    Else
        addNewAsin
        updateAsin sku, asin, site
    End If
End Sub

'向记录集中添加sku和asin
Sub addNewAsin()
    record.AddNew
End Sub

Sub updateAsin(sku, asin, site)
     With record
        .Fields("sku") = sku
        .Fields("asin") = asin
        .Fields("site") = site
    End With
    updateRecordSet
End Sub

'检验recordSet中是否存在asin
Function checkAsinIsExistInRecordset(asinId)
    record.MoveFirst
    record.Find ("asin=" & asinId)
    If record.EOF Then
        checkAsinIsExistInRecordset = True
    Else
        checkAsinIsExistInRecordset = False
    End If
End Function

'单条更新
Sub updateRecordSet()
    record.Update
End Sub

'批量更新到数据库
Sub updateRecordSetBatch()
    record.UpdateBatch
End Sub

'关闭链接
Public Sub closeConnection()
    If checkRecordNotIsColse() Then
        record.Close
        Set record = Nothing
    End If
    conn.Close
    Set conn = Nothing
End Sub

'抽取查询数据到指定单元格
Public Sub executeDQLSql(range As range)
    range.CopyFromRecordset record

End Sub

'执行DML语句
Public Sub executeDMLSql()
    conn.Execute (sql)

End Sub

'获取字段名，并放置到指定单元格
Public Sub getFieldName(rng As range)
    Dim fieldNum%
    Dim fieldsName$
    For fieldNum = 0 To record.Fields.Count - 1
        fieldsName = fieldsName & record.Fields.Item(fieldNum).name & ","
    Next
    'Debug.Print fieldsName
    rng.Resize(, UBound(Split(fieldsName, ","))) = Application.Transpose(Application.Transpose(Split(fieldsName, ",")))
End Sub

'以sku为主键,更新结果
Public Sub updataRecord(sku As Variant, name As String, param As Variant)
    With record
        .Find "ssku=" & sku
        .Fields(name) = .Fields(name).value & "," & param
        .MoveFirst
    End With
End Sub

