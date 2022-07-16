Attribute VB_Name = "Sentencias"
Sub main()
'conectar la BD
With BD
.CursorLocation = adUseClient 'Clente de la BaseDatos
.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BDsispro;Data Source=DESKTOP-BNRF03Q\SQLEXPRESS"
MDIForm1.Show
End With
End Sub


Sub Abrirtbcolor()
With Rscolor
    If .State = 1 Then .Close
    .Open "select *from color", BD, adOpenStatic, adLockOptimistic
End With
End Sub

Sub Abrirtbtela()
With Rstela
    If .State = 1 Then .Close
    .Open "select *from tela", BD, adOpenStatic, adLockOptimistic
End With
End Sub

Sub Abrirtbprenda()
With Rsprenda
    If .State = 1 Then .Close
    .Open "select *from prenda", BD, adOpenStatic, adLockOptimistic
End With
End Sub

Sub Abrirtbmateriaprima()
With Rsmateriap
    If .State = 1 Then .Close
    .Open "select *from materiaprima", BD, adOpenStatic, adLockOptimistic
End With
End Sub

Sub Abrirtbinventario()
With Rsinventario
    If .State = 1 Then .Close
    .Open "select *from inventario", BD, adOpenStatic, adLockOptimistic
End With
End Sub

Sub Abrirtbordenp()
With Rsordenp
    If .State = 1 Then .Close
    .Open "select *from ordenp", BD, adOpenStatic, adLockOptimistic
End With
End Sub

