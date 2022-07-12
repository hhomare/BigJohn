VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mn_color 
      Caption         =   "Colores"
      Index           =   1
   End
   Begin VB.Menu mn_tela 
      Caption         =   "TipoTela"
      Index           =   2
   End
   Begin VB.Menu mn_materiap 
      Caption         =   "MateriaPrima"
      Index           =   3
   End
   Begin VB.Menu mn_prenda 
      Caption         =   "Prendas"
      Index           =   4
   End
   Begin VB.Menu nm_inventario 
      Caption         =   "Inventario"
      Index           =   5
   End
   Begin VB.Menu nm_ordenp 
      Caption         =   "OrdenProduccion"
      Index           =   6
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mn_color_Click(Index As Integer)
FrmTela.Show vbModal
End Sub
