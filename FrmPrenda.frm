VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPrenda 
   Caption         =   "PRENDAS"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   9735
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Gridpj 
      Height          =   1935
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1920,189
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtcant 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "INTERFAZ PRENDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad Tela Requerida (mts)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Prenda"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPrenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnElimina_Click(Index As Integer)
wcodigo = Gridpj.Columns(0).Text
MsgBox (wcodigo)
With Rsprenda
.Requery
.Find "id=" & Val(wcodigo)
.Delete
.Requery
EstiloGrid
wcodigo = 0
End With

End Sub

Private Sub BtnSalir_Click(Index As Integer)
Unload Me

End Sub

Private Sub Form_Load()
Abrirtbprenda

'LLenar Grid
Set Gridpj.DataSource = Rsprenda
EstiloGrid
txtnombre.Text = ""
txtcant.Text = 0
End Sub

Private Sub BtnGrabar_Click(Index As Integer)
With Rsprenda
.AddNew
    !Nombre = txtnombre.Text
    !ConsumoInvUnd = Val(txtcant.Text)
.Update
Set Gridpj.DataSource = Rsprenda
End With
End Sub

Private Sub BtnNuevo_Click(Index As Integer)
txtnombre.Enabled = True
txtcant.Enabled = True
txtnombre.Text = ""
txtcant.Text = 0
txtnombre.SetFocus
End Sub



Sub EstiloGrid()
Gridpj.Columns(0).Width = 1000
Gridpj.Columns(1).Width = 4000
Gridpj.Columns(2).Width = 2000
Gridpj.Columns(0).Caption = "Codigo"
Gridpj.Columns(1).Caption = "Nombre"
Gridpj.Columns(2).Caption = "Cantidad Mts"

Gridpj.HeadFont.Bold = True

End Sub


