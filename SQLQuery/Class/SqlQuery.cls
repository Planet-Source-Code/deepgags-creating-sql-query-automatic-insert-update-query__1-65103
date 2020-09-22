VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class for Creating SQL Data Source"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim SQLField As String
Dim SqlValues As String
Dim SqlUpdateFieldsWhere
Public GetSQL_String As String
Public Enum FieldType
    DateField = 0
    CharacterField = 1
    NumericField = 2
End Enum
Public Enum SQlType
InsertData = 0
UpdateData = 1
End Enum

Private tpe As FieldType
Public Function InsertSQL(FieldName As String, FieldType As FieldType, DataValue As Variant) As String
Select Case (FieldType)
Case Is = 0
SQLField = SQLField & FieldName & ","
SqlValues = SqlValues & "'" & Format(DataValue, "mm/dd/yyyy") & "',"

Case Is = 1
SQLField = SQLField & FieldName & ","
SqlValues = SqlValues & "'" & DataValue & "',"

Case Is = 2
SQLField = SQLField & FieldName & ","
SqlValues = SqlValues & DataValue & ","
End Select

End Function
Function SQLBuild(TableName As String, SQlType As SQlType)
Select Case UCase(SQlType)
Case Is = 0
GetSQL_String = "Insert into " & TableName & " (" & Left(SQLField, Len(SQLField) - 1) & ")" & " values(" & Left(SqlValues, Len(SqlValues) - 1) & ")"
Case Is = 1
GetSQL_String = "update " & TableName & " set " & Left(SQLField, Len(SQLField) - 1) & " where " & Left(SqlUpdateFieldsWhere, Len(SqlUpdateFieldsWhere) - 5)
End Select
End Function

Public Function UpdateSQL(FieldName As String, FieldType As FieldType, DataValue As Variant) As String
Select Case (FieldType)
Case Is = 0
SQLField = SQLField & FieldName & "='" & Format(DataValue, "mm/dd/yyyy") & "',"
Case Is = 1
SQLField = SQLField & FieldName & "='" & DataValue & "',"
Case Is = 2
SQLField = SQLField & FieldName & "=" & DataValue & ","
End Select

End Function

Public Function UpdateSQLWhere(FieldName As String, FieldType As FieldType, DataValue As Variant) As String
Select Case (FieldType)
Case Is = 0
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & FieldName & "="
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & "'" & Format(DataValue, "mm/dd/yyyy") & "' and "

Case Is = 1
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & FieldName & "="
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & "'" & DataValue & "' and "

Case Is = 2
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & FieldName & "="
SqlUpdateFieldsWhere = SqlUpdateFieldsWhere & DataValue & " and "
End Select

End Function
Public Function ClearString()
SqlUpdateFieldsWhere = ""
SQLField = ""
SqlValues = ""
End Function

