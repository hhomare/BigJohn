VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmInventario 
   Caption         =   "INVENTARIO"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid gridinventario 
      Height          =   1575
      Left            =   6840
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid gridmateriap 
      Height          =   1575
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2778
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
      ColumnCount     =   2
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtcant 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "INVENTARIO ACTUAL"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "MATERIA PRIMA EXISTENTE"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad Requerida"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Materia Prima"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGrabar_Click(Index As Integer)
With Rsinventario
.AddNew
    !Idmateriaprima = mcodigo
    !cantidad = txtcant.Text
    
.Update
Set gridinventario.DataSource = Rsinventario
End With
End Sub

Private Sub BtnNuevo_Click(Index As Integer)
txtnombre.Enabled = True
txtnombre.Text = ""
txtcant.Text = 0
txtnombre.SetFocus
End Sub

Private Sub BtnSalir_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()

Abrirtbmateriaprima
Set gridmateriap.DataSource = Rsmateriap
Abrirtbinventario
Set gridinventario.DataSource = Rsinventario
txtnombre.Text = ""
txtcant.Text = 0
End Sub

Private Sub gridmateriap_Click()
id_inv = 0
mnombre = ""
mcodigo = (gridmateriap.Columns(0).Text)
mnombre = (gridmateriap.Columns(1).Text)
txtnombre.Text = mnombre
txtcant.SetFocus
'EstiloGrid

'mnombre = ""
'mcodigo = 0

End Sub
