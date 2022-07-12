VERSION 5.00
Begin VB.Form FrmOrdenP 
   Caption         =   "INVENTARIO"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   7020
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Text            =   "TxtPrenda"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Text            =   "TxtCantidad"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnModifica 
      Caption         =   "MODIFICAR"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Text            =   "TxMateriaP"
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Text            =   "Txtconsec"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Prenda"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Materia Prima"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Consecutivo"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmOrdenP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
