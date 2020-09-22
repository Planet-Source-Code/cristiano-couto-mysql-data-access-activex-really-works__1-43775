VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyIndex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMyIndexName As String 'local copy
Private mvarMyIndexPrimary As Boolean 'local copy
Private mvarMyIndexUnique As Boolean 'local copy
Private mvarMyIndexSequence As Long 'local copy
Private mvarMyIndexFields As MyFields 'local copy
Private mvarMyIndexCollation As String 'local copy
Private mvarMyIndexCardinality As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarMyIndexFullText As Boolean 'local copy
Public Property Let MyIndexFullText(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexFullText = 5
    mvarMyIndexFullText = vData
End Property


Public Property Get MyIndexFullText() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexFullText
    MyIndexFullText = mvarMyIndexFullText
End Property



Public Property Let MyIndexCardinality(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexCardinality = 5
    mvarMyIndexCardinality = vData
End Property


Public Property Get MyIndexCardinality() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexCardinality
    MyIndexCardinality = mvarMyIndexCardinality
End Property



Public Property Let MyIndexCollation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexCollation = 5
    mvarMyIndexCollation = vData
End Property


Public Property Get MyIndexCollation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexCollation
    MyIndexCollation = mvarMyIndexCollation
End Property



Public Property Set MyIndexFields(ByVal vData As MyFields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyIndexFields = Form1
    Set mvarMyIndexFields = vData
End Property


Public Property Get MyIndexFields() As MyFields
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexFields
    If mvarMyIndexFields Is Nothing Then
        Set mvarMyIndexFields = New MyFields
    End If
    
    Set MyIndexFields = mvarMyIndexFields
End Property



Public Property Let MyIndexSequence(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexSequence = 5
    mvarMyIndexSequence = vData
End Property


Public Property Get MyIndexSequence() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexSequence
    MyIndexSequence = mvarMyIndexSequence
End Property



Public Property Let MyIndexUnique(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexUnique = 5
    mvarMyIndexUnique = vData
End Property


Public Property Get MyIndexUnique() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexUnique
    MyIndexUnique = mvarMyIndexUnique
End Property



Public Property Let MyIndexPrimary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexPrimary = 5
    mvarMyIndexPrimary = vData
End Property


Public Property Get MyIndexPrimary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexPrimary
    MyIndexPrimary = mvarMyIndexPrimary
End Property



Public Property Let MyIndexName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyIndexName = 5
    mvarMyIndexName = vData
End Property


Public Property Get MyIndexName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyIndexName
    MyIndexName = mvarMyIndexName
End Property



