Attribute VB_Name = "Declaraciones"

Global BD As New ADODB.Connection   'declaro variable tipo ADO de connection

'Variables recorset

Global Rscolor As New ADODB.Recordset
Global Rstela As New ADODB.Recordset
Global Rsprenda As New ADODB.Recordset
Global Rsmateriap As New ADODB.Recordset
Global Rsinventario As New ADODB.Recordset

'Variables de usuario
Global wcodigo As Integer
Global mcodigo As String

Global id_tela As Integer
Global id_color As Integer

