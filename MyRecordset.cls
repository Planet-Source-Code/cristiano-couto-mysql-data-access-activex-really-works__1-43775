VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyRecordset"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MyFields"
Attribute VB_Ext_KEY = "Member1" ,"MyRows"
'local variable(s) to hold property value(s)
Private mvarMyEOF As Boolean 'local copy
Private mvarMyBOF As Boolean 'local copy
Private mvarMyRowID As Long 'local copy
Private mvarMyPercentPosition As Single 'local copy
Private mvarMyNumRows As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarMyRsFields As MyFields 'local copy
'local variable(s) to hold property value(s)
Private mvarMyActiveConnection As MyConnection 'local copy

Dim ResultArray()
'local variable(s) to hold property value(s)
Private mvarMySQLScript As String 'local copy
Private Sub fn_update_values()
Dim lFld As Long

For lFld = 1 To MyRsFields.Count
    MyRsFields(lFld).MyFieldValue = ResultArray(MyRowID, lFld)
Next

End Sub

Public Property Let MySQLScript(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MySQLScript = 5
    mvarMySQLScript = vData
End Property


Public Property Get MySQLScript() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MySQLScript
    MySQLScript = mvarMySQLScript
End Property




Public Property Set MyActiveConnection(ByVal vData As MyConnection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyActiveConnection = Form1
    Set mvarMyActiveConnection = vData
End Property


Public Property Get MyActiveConnection() As MyConnection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyActiveConnection
    Set MyActiveConnection = mvarMyActiveConnection
End Property



Public Property Set MyRsFields(ByVal vData As MyFields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyRsFields = Form1
    Set mvarMyRsFields = vData
End Property


Public Property Get MyRsFields() As MyFields
Attribute MyRsFields.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyRsFields
    Set MyRsFields = mvarMyRsFields
End Property



Sub MyStoreResult(pResArray())
Dim NRows

ResultArray = pResArray

On Error Resume Next

NRows = UBound(ResultArray)

On Error GoTo 0

MyNumRows = NRows
MyRowID = 1

End Sub

Public Property Let MyNumRows(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyNumRows = 5
    mvarMyNumRows = vData
End Property


Public Property Get MyNumRows() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyNumRows
    MyNumRows = mvarMyNumRows
End Property



Public Property Let MyPercentPosition(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyPercentPosition = 5
    mvarMyPercentPosition = vData
End Property


Public Property Get MyPercentPosition() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyPercentPosition
    MyPercentPosition = mvarMyPercentPosition
End Property



Public Property Let MyRowID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyRowID = 5
    mvarMyRowID = vData

    If MyRowID > MyNumRows Then
        MyPercentPosition = 100
    Else
        MyPercentPosition = (MyRowID / MyNumRows) * 100
    End If

    Select Case MyRowID
        Case MyNumRows + 1
            MyEOF = True
            MyBOF = False
        Case Is > MyNumRows + 1
            MyEOF = True
            MyBOF = False
            def_err vbObjectError & Len("MyRowID"), "MySQLDTA:MyRowID", "Try to past end of Recordset"
        Case 1
            MyBOF = True
            MyEOF = False
            fn_update_values
        Case Is < 1
            MyBOF = True
            MyEOF = False
            def_err vbObjectError & Len("MyRowID"), "MySQLDTA:MyRowID", "Try to past start of Recordset"
        Case Else
            MyBOF = False
            MyEOF = False
            fn_update_values
    End Select

End Property


Public Property Get MyRowID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyRowID
    MyRowID = mvarMyRowID
End Property



Public Property Let MyBOF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyBOF = 5
    mvarMyBOF = vData
End Property


Public Property Get MyBOF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyBOF
    MyBOF = mvarMyBOF
End Property



Public Property Let MyEOF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyEOF = 5
    mvarMyEOF = vData
End Property


Public Property Get MyEOF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyEOF
    MyEOF = mvarMyEOF
End Property



Public Sub MyRefresh()
Dim pResultArray()
Dim TmpRs As New MyRecordset
Dim pMyRec As API_MYSQL

MyActiveConnection.MyExecute MySQLScript, pResultArray
MyStoreResult pResultArray

End Sub

Public Sub MyGotoRow(MyGotoRowID As Long)

If MyGotoRowID < 1 Or MyGotoRowID > MyNumRows Then
    def_err vbObjectError & Len("MyGotoRow"), "MySQLDTA:MyGotoRow", "Try to past out of Recordset size"
Else
    MyRowID = MyGotoRowID
End If

End Sub

Public Sub MyMovePrevious()
MyRowID = MyRowID - 1
End Sub

Public Sub MyMoveLast()
MyRowID = MyNumRows
End Sub

Public Sub MyMoveNext()
MyRowID = MyRowID + 1
End Sub

Public Sub MyMoveFirst()
MyRowID = 1
End Sub

Private Sub Class_Initialize()
Set MyRsFields = New MyFields
End Sub


