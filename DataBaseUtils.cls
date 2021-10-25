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
'�������ݿ�
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

'��ʼ������
Private Sub init()
    loadConnection
    loadRecordSet
    destroyDataBase
End Sub

'���ؼ�¼��
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

'����recordSet�Ƿ�Ϊ��
Function checkRecordNotIsColse()
    checkRecordNotIsColse = record.State <> 0
End Function

'��������
Private Sub loadConnection()
    Set conn = dataBase.getConnection
End Sub

'����dataBase
Private Sub destroyDataBase()
    Set dataBase = Nothing
End Sub

'����IP��3λ�������ȡConnection
Private Sub getSqlServerDataBase(ipLast, password)
    With dataBase
        .letDataBase = "mpdisk"
        .letServerip = "192.168.1." & ipLast
        .letUsername = sqlServerUsername
        .letPassword = password
        .getSqlServerConnection
    End With
End Sub

'��ȡExcelΪ���ݿ�
Private Sub getExcelDataBase(path)
    With dataBase
        .letDataBase = path
        .getExcelConnection
    End With
End Sub

'��ȡmdb�ļ�
Private Sub getAccessDataBase()
    With dataBase
        .letDataBase = accessPath
        .getAccessConnection
    End With

End Sub

'��ȡmysql��dataBase
Private Sub getMysqlDataBase(ipLast, dabaBaseName As String)
    With dataBase
        .letDataBase = dabaBaseName
        .letServerip = "192.168.1." & ipLast
        .letUsername = mysqlUsername
        .letPassword = mysqlPassword
        .getMysqlConnection
    End With
End Sub

'ִ��Sql��ȡ�������ݣ���Ĭ�Ϸ�����[a1]
Sub extractDataInR1C1()
    getFieldName [a1]
    executeDQLSql [a2]
End Sub

'����sqlServer���ݿ⣬��Ҫ����sql���
Sub getDataBaseUtilForSqlServer(strSql, ipLast, password)
    getSqlServerDataBase ipLast, password
    sql = strSql
    init
End Sub


'����Excel
Sub getDataBaseUtilForExcel(strSql, path)
    getExcelDataBase (path)
    sql = strSql
    init
End Sub

'����Access
Sub getDataBaseUtilForAccess(strSql)
    getAccessDataBase
    sql = strSql
    init
End Sub

'����mysql���ݿ⣬��Ҫ����sql��� ��ip��3λ
Sub getDataBaseUtilForMysql(strSql, ipLast, dabaBaseName As String)
    getMysqlDataBase ipLast, dabaBaseName
    sql = strSql
    init
End Sub

'�����µ��Ա�asin
Sub saveOrUpdateAsin(sku, asin, site)
    If checkAsinIsExistInRecordset(asin) Then
        updateAsin sku, asin, site
    Else
        addNewAsin
        updateAsin sku, asin, site
    End If
End Sub

'���¼��������sku��asin
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

'����recordSet���Ƿ����asin
Function checkAsinIsExistInRecordset(asinId)
    record.MoveFirst
    record.Find ("asin=" & asinId)
    If record.EOF Then
        checkAsinIsExistInRecordset = True
    Else
        checkAsinIsExistInRecordset = False
    End If
End Function

'��������
Sub updateRecordSet()
    record.Update
End Sub

'�������µ����ݿ�
Sub updateRecordSetBatch()
    record.UpdateBatch
End Sub

'�ر�����
Public Sub closeConnection()
    If checkRecordNotIsColse() Then
        record.Close
        Set record = Nothing
    End If
    conn.Close
    Set conn = Nothing
End Sub

'��ȡ��ѯ���ݵ�ָ����Ԫ��
Public Sub executeDQLSql(range As range)
    range.CopyFromRecordset record

End Sub

'ִ��DML���
Public Sub executeDMLSql()
    conn.Execute (sql)

End Sub

'��ȡ�ֶ����������õ�ָ����Ԫ��
Public Sub getFieldName(rng As range)
    Dim fieldNum%
    Dim fieldsName$
    For fieldNum = 0 To record.Fields.Count - 1
        fieldsName = fieldsName & record.Fields.Item(fieldNum).name & ","
    Next
    'Debug.Print fieldsName
    rng.Resize(, UBound(Split(fieldsName, ","))) = Application.Transpose(Application.Transpose(Split(fieldsName, ",")))
End Sub

'��skuΪ����,���½��
Public Sub updataRecord(sku As Variant, name As String, param As Variant)
    With record
        .Find "ssku=" & sku
        .Fields(name) = .Fields(name).value & "," & param
        .MoveFirst
    End With
End Sub
