VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmOrdenP 
   Caption         =   "ORDEN DE PRODUCCION"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   13560
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   10
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtmateriap 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtcant 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtprenda 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   0
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid Gridordenp 
      Height          =   1935
      Left            =   3120
      TabIndex        =   2
      Top             =   4080
      Width           =   6375
      _ExtentX        =   11245
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
      ColumnCount     =   4
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
      BeginProperty Column03 
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
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Gridmateriap 
      Height          =   1935
      Left            =   5760
      TabIndex        =   6
      Top             =   1800
      Width           =   6255
      _ExtentX        =   11033
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
      ColumnCount     =   4
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
      BeginProperty Column03 
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
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Gridprenda 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Prenda"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccionar PRENDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Materia Prima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   12
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Materia Prima"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FrmOrdenP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGrabar_Click(Index As Integer)
With Rsordenp
.AddNew
    !materiaprima = id_materiap
    !prenda = id_prenda
    !unidades = txtcant.Text
.Update
Set Gridordenp.DataSource = Rsordenp
End With
End Sub

Private Sub BtnNuevo_Click(Index As Integer)
txtprenda.Enabled = False
txtprenda.Text = ""
txtmateriap.Enabled = False
txtmateriap.Text = ""
txtcant.Text = 0
txtcant.SetFocus
End Sub

Private Sub BtnSalir_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Abrirtbprenda
Set Gridprenda.DataSource = Rsprenda
Abrirtbordenp
Set Gridordenp.DataSource = Rsordenp
Abrirtbmateriaprima
Set Gridmateriap.DataSource = Rsmateriap
txtprenda.Text = ""
txtmateriap.Text = ""
txtcant.Text = 0
End Sub

Private Sub Gridmateriap_Click()
id_materiap = 0
id_materiap = (Gridmateriap.Columns(0).Text)
txtmateriap.Text = (Gridmateriap.Columns(1).Text)
End Sub

Private Sub Gridprenda_Click()
id_prenda = 0
id_prenda = (Gridprenda.Columns(0).Text)
txtprenda.Text = (Gridprenda.Columns(1).Text)

'EstiloGrid
End Sub
