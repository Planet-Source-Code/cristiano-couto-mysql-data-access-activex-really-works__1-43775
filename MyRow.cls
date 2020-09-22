VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyRow"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"MyFields"
Public Key As String
'local variable(s) to hold property value(s)
Private mvarMyFields As New MyFields 'local copy
'local variable(s) to hold property value(s)
Private mvarMyEOF As Boolean 'local copy
Private mvarMyBOF As Boolean 'local copy
Private mvarMyNumRows As Long 'local copy
Private mvarMyRowID As Long 'local copy

Public Property Let MyRowID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyRowID = 5
    mvarMyRowID = vData
End Property


Public Property Get MyRowID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyRowID
    MyRowID = mvarMyRowID
End Property



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



Public Sub MyAddNew()
End Sub

Public Property Set RowMyFields(ByVal vData As MyFields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyFields = Form1
    Set mvarMyFields = vData
End Property


Public Property Get RowMyFields() As MyFields
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyFields
    Set RowMyFields = mvarMyFields
End Property



Public Sub MyMoveFirst()
End Sub

Public Sub MyMoveLast()
End Sub

Public Sub MyMoveNext()
End Sub

Public Sub MyMovePrevious()
End Sub


Public Sub MyUpdate()
End Sub

