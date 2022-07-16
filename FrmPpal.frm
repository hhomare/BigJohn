VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "PRODUCCION PRENDAS"
   ClientHeight    =   6330
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12915
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mn_color 
      Caption         =   "Colores"
      Index           =   1
   End
   Begin VB.Menu nm_tela 
      Caption         =   "TipoTela"
      Index           =   2
   End
   Begin VB.Menu nm_materip 
      Caption         =   "MateriaPrima"
      Index           =   3
   End
   Begin VB.Menu nm_prenda 
      Caption         =   "Prendas"
      Index           =   4
   End
   Begin VB.Menu nm_inventario 
      Caption         =   "Inventario"
      Index           =   5
   End
   Begin VB.Menu nm_ordep 
      Caption         =   "OrdeProduccion"
      Index           =   6
   End
   Begin VB.Menu nm_cprenda 
      Caption         =   "ConsultaPrendas"
   End
   Begin VB.Menu nmsalir 
      Caption         =   "Salir"
      Index           =   7
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mn_color_Click(Index As Integer)
FrmColor.Show
End Sub

Private Sub nm_cprenda_Click()
Frmcprenda.Show
End Sub

Private Sub nm_inventario_Click(Index As Integer)
FrmInventario.Show
End Sub

Private Sub nm_materip_Click(Index As Integer)
Frmprueba.Show
End Sub

Private Sub nm_ordep_Click(Index As Integer)
FrmOrdenP.Show
End Sub

Private Sub nm_prenda_Click(Index As Integer)
FrmPrenda.Show
End Sub

Private Sub nm_tela_Click(Index As Integer)
FrmtELA.Show
End Sub
