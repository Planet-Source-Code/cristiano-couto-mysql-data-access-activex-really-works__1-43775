VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyTables"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"MyRecordset"
Attribute VB_Ext_KEY = "Member0" ,"MyRecordset"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarMyTableFields As MyFields 'local copy
'local variable(s) to hold property value(s)
Private mvarMyTableIndexes As MyIndexes 'local copy
Public Property Let MyTableIndexes(ByVal vData As MyIndexes)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MyTableIndexes = 5
    Set mvarMyTableIndexes = vData
End Property


Public Property Set MyTableIndexes(ByVal vData As MyIndexes)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyTableIndexes = Form1
    Set mvarMyTableIndexes = vData
End Property


Public Property Get MyTableIndexes() As MyIndexes
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyTableIndexes
    If IsObject(mvarMyTableIndexes) Then
        Set MyTableIndexes = mvarMyTableIndexes
    Else
        Set MyTableIndexes = mvarMyTableIndexes
    End If
End Property



Public Property Set MyTableFields(ByVal vData As MyFields)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MyTableFields = Form1
    Set mvarMyTableFields = vData
End Property


Public Property Get MyTableFields() As MyFields
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MyTableFields
    Set MyTableFields = mvarMyTableFields
End Property




Public Function Add(Optional sKey As String) As MyRecordset
    'create a new object
    Dim objNewMember As MyRecordset
    Set objNewMember = New MyRecordset


    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MyRecordset
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

