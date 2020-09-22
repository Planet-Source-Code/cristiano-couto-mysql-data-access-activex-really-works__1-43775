VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   2970
   ClientTop       =   2280
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim pMy As New MySQLDTA.MyConnection
Dim pRs As MyRecordset

Dim pFlds As New MyFields
Dim pIdxs As New MyIndexes

pMy.MyConnect "jacouto", "laj42290", "10.4.41.152"

pMy.MyCreateDatabase "banco_teste", True
pMy.MySelectDatabase "banco_teste"

With pFlds.Add
    .MyFieldName = "codigo"
    .MyFieldType = FIELD_TYPE_INT24
    .MyFieldSize = 11
    .MyFieldNull = False
    .MyFieldDefault = 0
    .MyFieldAutoIncrement = True
End With

With pFlds.Add
    .MyFieldName = "data"
    .MyFieldType = FIELD_TYPE_DATETIME
    .MyFieldNull = False
End With

With pFlds.Add
    .MyFieldName = "valor"
    .MyFieldType = FIELD_TYPE_DECIMAL
    .MyFieldSize = 11
    .MyFieldDecimals = 2
    .MyFieldNull = False
    .MyFieldDefault = 0
End With

With pIdxs.Add
    .MyIndexName = "pk_codigo"
    .MyIndexPrimary = True
    With .MyIndexFields.Add
        .MyFieldName = "codigo"
    End With
End With

With pIdxs.Add
    .MyIndexName = "ix_data"
    With .MyIndexFields.Add
        .MyFieldName = "data"
    End With
End With

pMy.MyCreateTable "tabela_teste", pFlds, pIdxs, False, MYISAM, MyOperationCreateNew

MsgBox pMy.MyDumpStructure("tabela_teste")

pMy.MyDropTable "tabela_teste"
pMy.MyDropDatabase "banco_teste"

pMy.MyCloseConnection

End Sub


