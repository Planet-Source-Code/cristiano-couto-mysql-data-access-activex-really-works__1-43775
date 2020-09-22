VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyConnection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MyDatabases"
Attribute VB_Ext_KEY = "Member1" ,"MyProcessLists"
Attribute VB_Ext_KEY = "Member2" ,"MyServerStatus"
Attribute VB_Ext_KEY = "Member3" ,"MyServerVariables"
Attribute VB_Ext_KEY = "Member4" ,"MyServerVariables"
'local variable(s) to hold property value(s)
Private mvarMyUsername As String 'local copy
Private mvarMyPassword As String 'local copy
Private mvarMyServerAddress As String 'local copy
Public Enum MyEnum_Status
    MyStatusConnected = 1
    MyStatusDisconnected = 2
End Enum

Public Enum MyEnum_Operation
    MyOperationCreateNew = 1
    MyOperationCreateAlter = 8
End Enum

Public Enum MyEnum_TableType
    ISAM = 1
    MYISAM = 2
    HEAP = 3
    MERGE = 4
    INNODB = 5
    BDB = 6
    AUTO = 7
End Enum


Private mvarMyStatus As MyEnum_Status 'local copy

Dim MyConHandle As Long
Dim MyRec As API_MYSQL         'valid record to hold ourselves
Dim myRec_res As API_MYSQL_RES 'this is we we will store results for ourselves
Dim myRec_field As API_MYSQL_FIELD
Dim myRec_rows As API_MYSQL_ROWS
'local variable(s) to hold property value(s)
Private mvarMyServerVersion As String 'local copy
Private mvarMyServerStatus As MyRecordset
Private mvarMyProcessLists As MyRecordset
Private mvarMyServerVariables As MyRecordset
'local variable(s) to hold property value(s)
Private mvarMyDatabaseName As String 'local copy
Public Function MyDumpStructure(Optional pMyTableName As String) As String
Dim TmpRs1 As MyRecordset
Dim TmpRs2 As MyRecordset
Dim RetDump As String

RetDump = ""

If pMyTableName <> "" Then

    Set TmpRs1 = MyExecute("SHOW CREATE TABLE `" & pMyTableName & "`")
    Do While Not TmpRs1.MyEOF
        RetDump = RetDump & TmpRs1("Create table")
        TmpRs1.MyMoveNext
    Loop
    
Else
    Set TmpRs2 = MyExecute("SHOW TABLES FROM `" & MyDatabaseName & "`")
    Do While Not TmpRs2.MyEOF
        Set TmpRs1 = MyExecute("SHOW CREATE TABLE `" & TmpRs2("Tables_in_" & MyDatabaseName) & "`")
        Do While Not TmpRs1.MyEOF
            RetDump = RetDump & TmpRs1("Create table") & ";" & vbCrLf & vbCrLf
            TmpRs1.MyMoveNext
        Loop
        TmpRs2.MyMoveNext
    Loop
End If

MyDumpStructure = RetDump

End Function

Public Function MyRenameField(MyTableName As String, MyOldName As String, MyNewName As String) As Boolean
Dim TmpRs As MyRecordset

Set TmpRs = MyExecute("show fields from `" & MyTableName & "`")
Do While Not TmpRs.MyEOF
    If MyOldName = TmpRs!field Then
        Err.Clear
        MyExecute "ALTER TABLE `" & MyTableName & "` CHANGE `" & MyOldName & "` `" & MyNewName & "` " & TmpRs!type & " " & IIf(TmpRs!Null <> "YES", "NOT NULL", "") & " " & TmpRs!Extra
        If Err Then MyRenameField = False Else MyRenameField = True
        Exit Do
    End If
    TmpRs.MyMoveNext
Loop
End Function

Public Function MyRenameTable(MyOldName As String, MyNewName As String) As Boolean
Err.Clear
MyExecute "ALTER TABLE `" & MyOldName & "` RENAME `" & MyNewName & "`"
If Err Then MyRenameTable = False Else MyRenameTable = True
End Function

Public Property Let MyDatabaseName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyDatabaseName = 5
    mvarMyDatabaseName = vData
    MySelectDatabase mvarMyDatabaseName
End Property


Public Property Get MyDatabaseName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyDatabaseName
    MyDatabaseName = mvarMyDatabaseName
End Property



Public Function MyCloseConnection() As Boolean
    If MyConHandle <> 0 Then Call API_mysql_close(MyRec)
    MyCloseConnection = True
End Function

Public Function MyDropIndex(pMyTableName As String, pMyIndexName As String) As Boolean
If pMyTableName = "" Then
    def_err vbObjectError & Len("MyDropIndex"), "MySQLDTA:MyDropIndex", "Empty table name"
    MyDropIndex = False
Else
    If pMyIndexName = "" Then
        def_err vbObjectError & Len("MyDropIndex"), "MySQLDTA:MyDropField", "Empty field name"
        MyDropIndex = False
    Else
        Err.Clear
        MyExecute "ALTER TABLE `" & pMyTableName & "` DROP INDEX `" & pMyIndexName & "`"
        If Err Then MyDropIndex = False Else MyDropIndex = True
    End If
End If
End Function

Public Function MyDropField(pMyTableName As String, pMyFieldName As String) As Boolean
If pMyTableName = "" Then
    def_err vbObjectError & Len("MyDropField"), "MySQLDTA:MyDropField", "Empty table name"
    MyDropField = False
Else
    If pMyFieldName = "" Then
        def_err vbObjectError & Len("MyDropField"), "MySQLDTA:MyDropField", "Empty field name"
        MyDropField = False
    Else
        Err.Clear
        MyExecute "ALTER TABLE `" & pMyTableName & "` DROP `" & pMyFieldName & "`"
        If Err Then MyDropField = False Else MyDropField = True
    End If
End If
End Function

Public Function MyDropTable(pMyTableName As String) As Boolean
If pMyTableName = "" Then
    def_err vbObjectError & Len("MyDropTable"), "MySQLDTA:MyDropTable", "Empty table name"
    MyDropTable = False
Else
    Err.Clear
    MyExecute "DROP TABLE IF EXISTS `" & pMyTableName & "`"
    If Err Then MyDropTable = False Else MyDropTable = True
End If
End Function

Public Function MyDropDatabase(pMyDatabaseName As String) As Boolean
If pMyDatabaseName = "" Then
    def_err vbObjectError & Len("MyDropDatabase"), "MySQLDTA:MyDropDatabase", "Empty Database name"
    MyDropDatabase = False
Else
    Err.Clear
    MyExecute "DROP DATABASE IF EXISTS `" & pMyDatabaseName & "`"
    If Err Then MyDropDatabase = False Else MyDropDatabase = True
End If
End Function

Public Function MyCreateIndex(pMyTableName As String, pMyIndexes As MyIndexes) As Boolean
Dim lIdx As MyIndex
Dim lFldIdx As MyField
Dim TmpRs1 As MyRecordset
Dim SqlBuildTable As String

'Alter table
SqlBuildTable = "ALTER TABLE `" & pMyTableName & "`, "

For Each lIdx In pMyIndexes
    If lIdx.MyIndexPrimary Then
        SqlBuildTable = SqlBuildTable & " ADD PRIMARY KEY `" & lIdx.MyIndexName & "` ("
    ElseIf lIdx.MyIndexUnique Then
        SqlBuildTable = SqlBuildTable & " ADD UNIQUE `" & lIdx.MyIndexName & "` ("
    ElseIf lIdx.MyIndexFullText Then
        SqlBuildTable = SqlBuildTable & " ADD FULLTEXT `" & lIdx.MyIndexName & "` ("
    Else
        SqlBuildTable = SqlBuildTable & " ADD INDEX `" & lIdx.MyIndexName & "` ("
    End If
    
    For Each lFldIdx In lIdx.MyIndexFields
        SqlBuildTable = SqlBuildTable & "`" & lFldIdx.MyFieldName & "`, "
    Next
    
    SqlBuildTable = fn_remove_last_char(SqlBuildTable, ",")
    SqlBuildTable = SqlBuildTable & "), "
    
Next

SqlBuildTable = fn_remove_last_char(SqlBuildTable, ",")

MsgBox SqlBuildTable

Err.Clear
MyExecute SqlBuildTable
If Err Then MyCreateIndex = False Else MyCreateIndex = True

End Function

Public Function MyCreateField(pMyTableName As String, pMyFields As MyFields) As Boolean
Dim lFld As MyField
Dim FieldExists As Boolean
Dim TmpRs1 As MyRecordset
Dim SqlBuildTable As String

SqlBuildTable = "ALTER TABLE `" & pMyTableName & "` "

Set TmpRs1 = MyExecute("show fields from `" & pMyTableName & "`")

For Each lFld In pMyFields
    FieldExists = False
    TmpRs1.MyMoveFirst
    
    Do While Not TmpRs1.MyEOF
        If lFld.MyFieldName = TmpRs1!field Then
            FieldExists = True
            Exit Do
        End If
        TmpRs1.MyMoveNext
    Loop
    
    If FieldExists Then
        SqlBuildTable = SqlBuildTable & " CHANGE `" & lFld.MyFieldName & "` `" & lFld.MyFieldName & "` " & fn_stru_field_to_query(lFld) & ", "
    Else
        SqlBuildTable = SqlBuildTable & " ADD `" & lFld.MyFieldName & "` " & fn_stru_field_to_query(lFld) & ", "
    End If
Next
SqlBuildTable = fn_remove_last_char(SqlBuildTable, ",")

Err.Clear
MyExecute SqlBuildTable
If Err Then MyCreateField = False Else MyCreateField = True

End Function

Public Function MyCreateTable(pMyTableName As String, pMyFields As MyFields, pMyIndexes As MyIndexes, pMyTemporary As Boolean, pMyTableType As MyEnum_TableType, pMyOperation As MyEnum_Operation) As Boolean
Dim lFld As MyField
Dim lIdx As MyIndex
Dim lFldIdx As MyField
Dim TableExists As Boolean
Dim FieldExists As Boolean
Dim TmpRs1 As MyRecordset
Dim SqlBuildTable As String
Dim SqlBuildIndex As String

Set TmpRs1 = MyExecute("show tables from `" & MyDatabaseName & "`")
Do While Not TmpRs1.MyEOF
    If TmpRs1!tables_in_teste = pMyTableName Then
        TableExists = True
        Exit Do
    End If
    TmpRs1.MyMoveNext
Loop

    
'Create new table
SqlBuildTable = "CREATE " & IIf(pMyTemporary = True, "TEMPORARY", "") & " TABLE `" & pMyTableName & "` ("

For Each lFld In pMyFields
    SqlBuildTable = SqlBuildTable & "`" & lFld.MyFieldName & "` " & fn_stru_field_to_query(lFld) & ", "
Next

SqlBuildTable = fn_remove_last_char(SqlBuildTable, ",")

SqlBuildIndex = ""
For Each lIdx In pMyIndexes
    If lIdx.MyIndexPrimary Then
        SqlBuildIndex = SqlBuildIndex & " PRIMARY KEY `" & lIdx.MyIndexName & "`("
    ElseIf lIdx.MyIndexUnique Then
        SqlBuildIndex = SqlBuildIndex & " UNIQUE `" & lIdx.MyIndexName & "`("
    ElseIf lIdx.MyIndexFullText Then
        SqlBuildIndex = SqlBuildIndex & " FULLTEXT `" & lIdx.MyIndexName & "`("
    Else
        SqlBuildIndex = SqlBuildIndex & " INDEX `" & lIdx.MyIndexName & "`("
    End If
    
    For Each lFldIdx In lIdx.MyIndexFields
        SqlBuildIndex = SqlBuildIndex & "`" & lFldIdx.MyFieldName & "`, "
    Next
    
    SqlBuildIndex = fn_remove_last_char(SqlBuildIndex, ",")
    SqlBuildIndex = SqlBuildIndex & "), "
    
Next

SqlBuildIndex = fn_remove_last_char(SqlBuildIndex, ",")

If SqlBuildIndex <> "" Then
    SqlBuildTable = SqlBuildTable & ", " & SqlBuildIndex & ")"
Else
    SqlBuildTable = SqlBuildTable & ")"
End If
SqlBuildTable = SqlBuildTable & " Type = " & fn_stru_table(pMyTableType)
    
If TableExists Then
    If pMyOperation = MyOperationCreateNew Then
        SqlBuildTable = "DROP TABLE IF EXISTS `" & pMyTableName & "`;" & vbCrLf & SqlBuildTable
    Else
        
        SqlBuildTable = "ALTER TABLE `" & pMyTableName & "` "
        
        
        Set TmpRs1 = MyExecute("show fields from `" & pMyTableName & "`")
        
        For Each lFld In pMyFields
            FieldExists = False
            TmpRs1.MyMoveFirst
            
            Do While Not TmpRs1.MyEOF
                If lFld.MyFieldName = TmpRs1!field Then
                    FieldExists = True
                    Exit Do
                End If
                TmpRs1.MyMoveNext
            Loop
            
            If FieldExists Then
                SqlBuildTable = SqlBuildTable & " CHANGE `" & lFld.MyFieldName & "` `" & lFld.MyFieldName & "` " & fn_stru_field_to_query(lFld) & ", "
            Else
                SqlBuildTable = SqlBuildTable & " ADD `" & lFld.MyFieldName & "` " & fn_stru_field_to_query(lFld) & ", "
            End If
        Next
        SqlBuildTable = fn_remove_last_char(SqlBuildTable, ",")
        SqlBuildTable = SqlBuildTable & ", Type = " & fn_stru_table(pMyTableType)
    
    End If
End If

Err.Clear
MyExecute SqlBuildTable
If Err Then MyCreateTable = False Else MyCreateTable = True

End Function

Public Function MyCreateDatabase(pMyDatabaseName As String, Optional pMyOverwrite As Boolean = False) As Boolean
If pMyDatabaseName = "" Then
    def_err vbObjectError & Len("MyCreateDatabase"), "MySQLDTA:MyCreateDatabase", "Empty Database name"
    MyCreateDatabase = False
Else
    If pMyOverwrite Then
        MyExecute "DROP DATABASE IF EXISTS `" & pMyDatabaseName & "`"
    End If
    
    Err.Clear
    MyExecute "CREATE DATABASE `" & pMyDatabaseName & "`"
    If Err Then MyCreateDatabase = False Else MyCreateDatabase = True
End If
End Function

Public Function MyExecute(pSQLScript As String, Optional pMyRetArray) As MyRecordset
Dim lcRecordset As New MyRecordset
Dim pResultArray()

Dim lScripts() As String

lScripts = Split(pSQLScript, ";")

For x = LBound(lScripts) To UBound(lScripts)

    fn_execute lScripts(x), MyRec, lcRecordset, pResultArray
    
    Set lcRecordset.MyActiveConnection = Me
    
    lcRecordset.MySQLScript = lScripts(x)
    lcRecordset.MyStoreResult pResultArray
    
    Set MyExecute = lcRecordset

    If Not IsMissing(pMyRetArray) Then pMyRetArray = pResultArray

Next

End Function

Public Function MySelectDatabase(pMyDatabaseName As String) As Boolean
Dim Ret As Long
If MyConHandle <> 0 Then
    Ret = API_mysql_select_db(MyRec, StrPtr(StrConv(pMyDatabaseName, vbFromUnicode)))
    If Ret <> 0 Then
        Err.Raise API_mysql_errno(MyRec), "MySQLDTA:MySelectDatabase", ptr2str(API_mysql_error(MyRec))
        MySelectDatabase = False
    Else
        mvarMyDatabaseName = pMyDatabaseName
        MySelectDatabase = True
    End If
End If
End Function



Public Property Get MyServerVariables() As MyRecordset
    
    Set mvarMyServerVariables = MyExecute("show variables")
    
    Set MyServerVariables = mvarMyServerVariables
End Property


Public Property Set MyServerVariables(vData As MyRecordset)
    Set mvarMyServerVariables = vData
End Property



Public Property Get MyProcessLists() As MyRecordset
        
    Set mvarMyProcessLists = MyExecute("show full processlist")
    
    Set MyProcessLists = mvarMyProcessLists

End Property


Public Property Set MyProcessLists(vData As MyRecordset)
    Set mvarMyProcessLists = vData
End Property



Public Property Get MyServerStatus() As MyRecordset
    
    Set mvarMyServerStatus = MyExecute("show status")
    
    Set MyServerStatus = mvarMyServerStatus
End Property


Public Property Set MyServerStatus(vData As MyRecordset)
    Set mvarMyServerStatus = vData
End Property



Private Sub Class_Initialize()
    mvarMyStatus = MyStatusDisconnected
End Sub

Private Sub Class_Terminate()
  MyCloseConnection
  Set mvarMyServerVariables = Nothing
  Set mvarMyProcessLists = Nothing
  Set mvarMyServerStatus = Nothing
End Sub



Public Property Let MyServerVersion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyServerVersion = 5
    mvarMyServerVersion = vData
End Property


Public Property Get MyServerVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyServerVersion
    MyServerVersion = mvarMyServerVersion
End Property




Public Function MyConnect(Optional pMyUsername As String, Optional pMyPassword As String, Optional pMyServerAddress As String) As Boolean
Dim lMyUsername As String
Dim lMyPassword As String
Dim lMyServerAddress As String

If pMyUsername = "" Then lMyUsername = MyUsername Else lMyUsername = pMyUsername
If pMyPassword = "" Then lMyPassword = MyPassword Else lMyPassword = pMyPassword
If pMyServerAddress = "" Then lMyServerAddress = MyServerAddress Else lMyServerAddress = pMyServerAddress

MyConHandle = API_mysql_connect(MyRec, StrPtr(StrConv(lMyServerAddress, vbFromUnicode)), StrPtr(StrConv(lMyUsername, vbFromUnicode)), StrPtr(StrConv(lMyPassword, vbFromUnicode)))

MyServerVersion = ptr2str(MyRec.server_version)

If MyConHandle Then
    CopyMemory MyRec, ByVal MyConHandle, LenB(MyRec) 'again we have to mooo-ve the memory
    MyConnect = True
    MyUsername = lMyUsername
    MyPassword = lMyUsername
    MyServerAddress = lMyServerAddress
    MyStatus = MyStatusConnected
Else
    MyStatus = MyStatusDisconnected
    def_err API_mysql_errno(MyRec), "MySQLDTA:MyConnect", ptr2str(API_mysql_error(MyRec))
    MyConnect = False
End If

End Function

Public Property Let MyStatus(ByVal vData As MyEnum_Status)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyStatus = 5
    mvarMyStatus = vData
End Property


Public Property Get MyStatus() As MyEnum_Status
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyStatus
    MyStatus = mvarMyStatus
End Property




Public Property Let MyServerAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyServerAddress = 5
    mvarMyServerAddress = vData
End Property


Public Property Get MyServerAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyServerAddress
    MyServerAddress = mvarMyServerAddress
End Property



Public Property Let MyPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyPassword = 5
    mvarMyPassword = vData
End Property


Public Property Get MyPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyPassword
    MyPassword = mvarMyPassword
End Property



Public Property Let MyUsername(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyUsername = 5
    mvarMyUsername = vData
End Property


Public Property Get MyUsername() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyUsername
    MyUsername = mvarMyUsername
End Property



