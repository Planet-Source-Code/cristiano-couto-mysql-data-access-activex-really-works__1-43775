VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyField"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMyFieldName As String 'local copy
Private mvarMyFieldType As API_enum_field_types 'local copy
Private mvarMyFieldSize As Long 'local copy
Private mvarMyFieldNull As Boolean 'local copy
Private mvarMyFieldDefault As String 'local copy
Private mvarMyFieldAutoIncrement As Boolean 'local copy
Private mvarMyFieldValue As Variant 'local copy

Public Enum API_refresh_options
 API_REFRESH_GRANT = 1
 API_REFRESH_LOG = 2
 API_REFRESH_TABLES = 4
 API_REFRESH_HOSTS = 8
 API_REFRESH_STATUS = 16
 API_REFRESH_THREADS = 32
 API_REFRESH_SLAVE = 64
 API_REFRESH_MASTER = 128
 API_REFRESH_READ_LOCK = 256
 API_REFRESH_FAST = 32768
End Enum
Public Enum API_enum_field_types
 FIELD_TYPE_DECIMAL = 0   ' adDecimal
 FIELD_TYPE_TINY = 1      ' adTinyInt
 FIELD_TYPE_SHORT = 2     ' adInteger
 FIELD_TYPE_LONG = 3      ' adBigInt
 FIELD_TYPE_FLOAT = 4
 FIELD_TYPE_DOUBLE = 5    ' adDouble
 FIELD_TYPE_NULL = 6      ' adUserDefined
 FIELD_TYPE_TIMESTAMP = 7 ' adDBTimeStamp
 FIELD_TYPE_LONGLONG = 8  ' adDouble
 FIELD_TYPE_INT24 = 9     ' adInt
 FIELD_TYPE_DATE = 10     ' adDate
 FIELD_TYPE_TIME = 11     ' adTime
 FIELD_TYPE_DATETIME = 12 ' adDateTime
 FIELD_TYPE_YEAR = 13
 FIELD_TYPE_NEWDATE = 14
 FIELD_TYPE_ENUM = 247
 FIELD_TYPE_SET = 248
 FIELD_TYPE_TINY_BLOB = 249
 FIELD_TYPE_MEDIUM_BLOB = 250
 FIELD_TYPE_LONG_BLOB = 251
 FIELD_TYPE_BLOB = 252
 FIELD_TYPE_VAR_STRING = 253 ' adVarChar
 FIELD_TYPE_STRING = 254     ' adBSTR
End Enum
'local variable(s) to hold property value(s)
Private mvarMyZeroFill As Boolean 'local copy
Private mvarMyUnsigned As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarMyFieldDecimals As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarMyFieldEnumDef As String 'local copy
Public Property Let MyFieldEnumDef(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldEnumDef = 5
    mvarMyFieldEnumDef = vData
End Property


Public Property Get MyFieldEnumDef() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldEnumDef
    MyFieldEnumDef = mvarMyFieldEnumDef
End Property



Public Property Let MyFieldDecimals(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldDecimals = 5
    mvarMyFieldDecimals = vData
End Property


Public Property Get MyFieldDecimals() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldDecimals
    MyFieldDecimals = mvarMyFieldDecimals
End Property



Public Property Let MyUnsigned(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyUnsigned = 5
    mvarMyUnsigned = vData
End Property


Public Property Get MyUnsigned() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyUnsigned
    MyUnsigned = mvarMyUnsigned
End Property



Public Property Let MyZeroFill(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyZeroFill = 5
    mvarMyZeroFill = vData
End Property


Public Property Get MyZeroFill() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyZeroFill
    MyZeroFill = mvarMyZeroFill
End Property




Public Property Let MyFieldValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldValue = 5
    mvarMyFieldValue = vData
End Property


Public Property Set MyFieldValue(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyFieldValue = Form1
    Set mvarMyFieldValue = vData
End Property


Public Property Get MyFieldValue() As Variant
Attribute MyFieldValue.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldValue
    If IsObject(mvarMyFieldValue) Then
        Set MyFieldValue = mvarMyFieldValue
    Else
        MyFieldValue = mvarMyFieldValue
    End If
End Property



Public Property Let MyFieldAutoIncrement(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldAutoIncrement = 5
    mvarMyFieldAutoIncrement = vData
End Property


Public Property Get MyFieldAutoIncrement() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldAutoIncrement
    MyFieldAutoIncrement = mvarMyFieldAutoIncrement
End Property



Public Property Let MyFieldDefault(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldDefault = 5
    mvarMyFieldDefault = vData
End Property


Public Property Get MyFieldDefault() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldDefault
    MyFieldDefault = mvarMyFieldDefault
End Property



Public Property Let MyFieldNull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldNull = 5
    mvarMyFieldNull = vData
End Property


Public Property Get MyFieldNull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldNull
    MyFieldNull = mvarMyFieldNull
End Property



Public Property Let MyFieldSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldSize = 5
    mvarMyFieldSize = vData
End Property


Public Property Get MyFieldSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldSize
    MyFieldSize = mvarMyFieldSize
End Property



Public Property Let MyFieldType(ByVal vData As API_enum_field_types)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyFieldType = Form1
    mvarMyFieldType = vData
End Property


Public Property Get MyFieldType() As API_enum_field_types
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldType
    MyFieldType = mvarMyFieldType
End Property



Public Property Let MyFieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyFieldName = 5
    mvarMyFieldName = vData
End Property


Public Property Get MyFieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFieldName
    MyFieldName = mvarMyFieldName
End Property



