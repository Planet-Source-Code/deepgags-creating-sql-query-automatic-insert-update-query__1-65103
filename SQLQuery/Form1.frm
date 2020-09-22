VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Query"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert Query"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLInsert As New SqlQuery
Private Sub Command1_Click()
SQLInsert.ClearString
SQLInsert.InsertSQL "Name", CharacterField, "My Textbox Value"
SQLInsert.InsertSQL "Address", CharacterField, "My Textbox Value"
SQLInsert.InsertSQL "PIN", CharacterField, "My Textbox Value"
SQLInsert.InsertSQL "Age", NumericField, 25
SQLInsert.InsertSQL "DateofBirth", DateField, Date

SQLInsert.SQLBuild "Customer", InsertData
Text1.Text = SQLInsert.GetSQL_String

End Sub

Private Sub Command2_Click()

SQLInsert.ClearString
SQLInsert.UpdateSQL "Name", CharacterField, "My Textbox Value"
SQLInsert.UpdateSQL "Address", CharacterField, "My Textbox Value"
SQLInsert.UpdateSQL "PIN", CharacterField, "My Textbox Value"
SQLInsert.UpdateSQL "Age", NumericField, 25
SQLInsert.UpdateSQL "DateofBirth", DateField, Date
SQLInsert.UpdateSQLWhere "Age", NumericField, 20
SQLInsert.SQLBuild "Customer", UpdateData
Text1.Text = SQLInsert.GetSQL_String
End Sub

