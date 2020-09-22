Attribute VB_Name = "MYSQLDTA_MODULE"
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Sub def_err(num_err As Long, src_err, desc_err)
    Err.Raise num_err, src_err, desc_err
End Sub

Sub fn_conv_valor(pFld As MyField, pValor)
Select Case pFld.MyFieldType
    Case FIELD_TYPE_DATE, FIELD_TYPE_TIME, FIELD_TYPE_DATETIME
        If IsDate(pValor) Then
            pFld.MyFieldValue = CDate(pValor)
        Else
            pFld.MyFieldValue = CDate("01/01/01")
        End If
    Case FIELD_TYPE_DECIMAL, FIELD_TYPE_DOUBLE, FIELD_TYPE_FLOAT, FIELD_TYPE_INT24, FIELD_TYPE_LONG, FIELD_TYPE_LONGLONG
        pFld.MyFieldValue = Val(pValor)
    Case Else
        pFld.MyFieldValue = pValor
End Select
End Sub

Function fn_execute(pSQL As String, pApiRec As API_MYSQL, pRecordset As MyRecordset, pRetArray()) As Double

Dim myRec_res As API_MYSQL_RES
Dim myRec_field As API_MYSQL_FIELD
Dim myRec_rows As API_MYSQL_ROWS

Dim Ret As Long
Dim m_fieldcount As Long
Dim m_rowcount As Long
Dim i As Long
Dim j As Long
Dim s As String
Dim PickUp() As Long

Dim lFields As MyFields

Set pRecordset.MyRsFields = New MyFields
Set lFields = pRecordset.MyRsFields

Ret = API_mysql_query(pApiRec, StrPtr(StrConv(pSQL, vbFromUnicode)))
If Ret = 0 Then 'query was good
    Ret = API_mysql_store_result(pApiRec)
    If Ret Then
        CopyMemory myRec_res, ByVal Ret, LenB(myRec_res)
    
        m_fieldcount = myRec_res.field_count
        m_rowcount = convert642l(myRec_res.row_count)
        
        pRecordset.MyNumRows = m_rowcount
        
        If m_rowcount > 0 Then
        
            ReDim PickUp(1 To m_fieldcount)
            ReDim pRetArray(1 To m_rowcount, 1 To m_fieldcount)
            
            For i = 1 To m_fieldcount
                Ret = API_mysql_fetch_field(myRec_res)
                If Ret Then
                    CopyMemory myRec_field, ByVal Ret, LenB(myRec_field)
                
                    With lFields.Add(ptr2str(myRec_field.name))
                        .MyFieldName = ptr2str(myRec_field.name)
                        .MyFieldDefault = ptr2str(myRec_field.def)
                        '.MyFieldAutoIncrement = ptr2str(myRec_field.Flags)
                        .MyFieldSize = myRec_field.length
                        .MyFieldType = myRec_field.type
                    End With
                    
                End If
            Next
                      
            For j = 1 To m_rowcount  'append rows to the recordset
                Ret = API_mysql_fetch_row(myRec_res) 'fetch a row
                If Ret Then
                    CopyMemory PickUp(1), ByVal Ret, SIZE_OF_CHAR * m_fieldcount 'copy it into array so we can pick it up
                    For i = 1 To m_fieldcount
                         s = ptr2str(PickUp(i))
                         pRetArray(j, i) = s
                    Next i
                End If
            Next j
        
        End If
    
        Ret = API_mysql_free_result(myRec_res)
    
        fn_execute = m_rowcount
    Else
        fn_execute = convert642l(pApiRec.affected_rows)
    End If

Else
    def_err API_mysql_errno(pApiRec), "MySQLDTA:MyExecute", ptr2str(API_mysql_error(pApiRec))
End If

End Function


Function fn_remove_last_char(pStr As String, pChar As String) As String
Dim TmpStr As String
Dim NumCicle As Long

NumCicle = 0

TmpStr = pStr

For x = Len(TmpStr) To 1 Step -1
    NumCicle = NumCicle + 1
    If NumCicle > 4 Then
        fn_remove_last_char = TmpStr
        Exit For
    End If
    
    If Mid(TmpStr, x, 1) = pChar Then
        If x = 1 Then
            fn_remove_last_char = ""
        Else
            fn_remove_last_char = Left(TmpStr, x - 1)
        End If
        Exit For
    End If
Next

End Function

Function fn_stru_field_to_query(pFld As MyField) As String
Dim TmpSql As String

If pFld.MyFieldType = FIELD_TYPE_BLOB Then
    If pFld.MyFieldNull = False Then
        TmpSql = "BLOB NOT NULL"
    Else
        TmpSql = "BLOB"
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_DATE Or pFld.MyFieldType = FIELD_TYPE_NEWDATE Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DATE NOT NULL DEFAULT '0000-00-00'"
        Else
            TmpSql = "DATE NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DATE"
        Else
            TmpSql = "DATE DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_DATETIME Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00'"
        Else
            TmpSql = "DATETIME NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DATETIME"
        Else
            TmpSql = "DATETIME DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_DOUBLE Or pFld.MyFieldType = FIELD_TYPE_DECIMAL Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DECIMAL(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "DECIMAL(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "DECIMAL(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "DECIMAL(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_FLOAT Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "FLOAT(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "FLOAT(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "FLOAT(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "FLOAT(" & pFld.MyFieldSize & "," & pFld.MyFieldDecimals & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_ENUM Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "ENUM(" & pFld.MyFieldEnumDef & ")" & " NOT NULL"
        Else
            TmpSql = "ENUM(" & pFld.MyFieldEnumDef & ")" & " NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "ENUM(" & pFld.MyFieldEnumDef & ")"
        Else
            TmpSql = "ENUM(" & pFld.MyFieldEnumDef & ")" & " DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_LONG Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "BIGINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "BIGINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "BIGINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "BIGINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_INT24 Or pFld.MyFieldType = FIELD_TYPE_SHORT Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "INT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "INT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "INT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "INT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_LONG_BLOB Then
    If pFld.MyFieldNull = False Then
        TmpSql = "BLOB NOT NULL"
    Else
        TmpSql = "BLOB"
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_LONGLONG Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "MEDIUMINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "MEDIUMINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "MEDIUMINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "MEDIUMINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_MEDIUM_BLOB Then
    If pFld.MyFieldNull = False Then
        TmpSql = "MEDIUMBLOB NOT NULL"
    Else
        TmpSql = "MEDIUMBLOB"
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_SET Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "SET(" & pFld.MyFieldEnumDef & ")" & " NOT NULL"
        Else
            TmpSql = "SET(" & pFld.MyFieldEnumDef & ")" & " NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "SET(" & pFld.MyFieldEnumDef & ")"
        Else
            TmpSql = "SET(" & pFld.MyFieldEnumDef & ")" & " DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_TIME Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "TIME NOT NULL DEFAULT '00:00:00'"
        Else
            TmpSql = "TIME NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "TIME"
        Else
            TmpSql = "TIME DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If


If pFld.MyFieldType = FIELD_TYPE_TIMESTAMP Then
    TmpSql = "TIMESTAMP(" & pFld.MyFieldSize & ")"
End If

If pFld.MyFieldType = FIELD_TYPE_TINY Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "TINYINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '0'"
        Else
            TmpSql = "TINYINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " NOT NULL " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "TINYINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "")
        Else
            TmpSql = "TINYINT(" & pFld.MyFieldSize & ")" & IIf(pFld.MyUnsigned = True, " UNSIGNED ", "") & IIf(pFld.MyZeroFill = True, " ZEROFILL ", "") & " " & IIf(pFld.MyFieldAutoIncrement, " AUTO_INCREMENT ", "") & "DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_TINY_BLOB Then
    If pFld.MyFieldNull = False Then
        TmpSql = "TINYBLOB NOT NULL"
    Else
        TmpSql = "TINYBLOB"
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_VAR_STRING Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "VARCHAR(" & pFld.MyFieldSize & ")" & " NOT NULL"
        Else
            TmpSql = "VARCHAR(" & pFld.MyFieldSize & ")" & " NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "VARCHAR(" & pFld.MyFieldSize & ")"
        Else
            TmpSql = "VARCHAR(" & pFld.MyFieldSize & ")" & " DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_STRING Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "CHAR(" & pFld.MyFieldSize & ")" & " NOT NULL"
        Else
            TmpSql = "CHAR(" & pFld.MyFieldSize & ")" & " NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "CHAR(" & pFld.MyFieldSize & ")"
        Else
            TmpSql = "CHAR(" & pFld.MyFieldSize & ")" & " DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

If pFld.MyFieldType = FIELD_TYPE_YEAR Then
    If pFld.MyFieldNull = False Then
        If pFld.MyFieldDefault = "" Then
            TmpSql = "YEAR(" & pFld.MyFieldSize & ")" & " NOT NULL"
        Else
            TmpSql = "YEAR(" & pFld.MyFieldSize & ")" & " NOT NULL DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    Else
        If pFld.MyFieldDefault = "" Then
            TmpSql = "YEAR(" & pFld.MyFieldSize & ")"
        Else
            TmpSql = "YEAR(" & pFld.MyFieldSize & ")" & " DEFAULT '" & pFld.MyFieldDefault & "'"
        End If
    End If
End If

fn_stru_field_to_query = TmpSql

End Function


Function fn_stru_index_to_query(pFld As MyField) As String

End Function

Function fn_stru_table(pTblType As MyEnum_TableType) As String
If pTblType = AUTO Then
    fn_stru_table = "MyISAM"
ElseIf pTblType = BDB Then
    fn_stru_table = "BDB"
ElseIf pTblType = HEAP Then
    fn_stru_table = "HEAP"
ElseIf pTblType = INNODB Then
    fn_stru_table = "InnoDB"
ElseIf pTblType = ISAM Then
    fn_stru_table = "ISAM"
ElseIf pTblType = MERGE Then
    fn_stru_table = "MERGE"
ElseIf pTblType = MYISAM Then
    fn_stru_table = "MyISAM"
End If
End Function

Sub Main()
RestoreMySQLApi
End Sub

Sub RestoreMySQLApi()
Dim lBytes() As Byte
Dim fFile As Long

lBytes = LoadResData("LIBMYSQL", "DLL")

fFile = FreeFile

If Dir(SysDir & "libmysql.dll") = "" Then
    Open SysDir & "libmysql.dll" For Binary Shared As fFile
    Put fFile, , lBytes
    Close fFile
End If
End Sub

Function WinDir() As String
Dim TmpStr As String
Dim Ret As Long

TmpStr = Space(255)
Ret = GetWindowsDirectory(TmpStr, Len(TmpStr))
WinDir = IIf(Right(Left(TmpStr, Ret), 1) = "\", Left(TmpStr, Ret), Left(TmpStr, Ret) & "\")

End Function


Function SysDir() As String
Dim TmpStr As String
Dim Ret As Long

TmpStr = Space(255)
Ret = GetSystemDirectory(TmpStr, Len(TmpStr))
SysDir = IIf(Right(Left(TmpStr, Ret), 1) = "\", Left(TmpStr, Ret), Left(TmpStr, Ret) & "\")

End Function

