VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmprueba 
   Caption         =   "MATERIAS PRIMAS"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid GRIDMP 
      Height          =   1335
      Left            =   3000
      TabIndex        =   14
      Top             =   4200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2355
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
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox txttipocolor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txttipotela 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid Gridcolor 
      Height          =   1935
      Left            =   6480
      TabIndex        =   10
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSDataGridLib.DataGrid Gridtela 
      Height          =   1935
      Left            =   600
      TabIndex        =   9
      Top             =   2040
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
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccionar Color"
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
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccionar Telas"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Color"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Tela"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Frmprueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGrabar_Click(Index As Integer)
With Rsmateriap
.AddNew
    !Nombre = txtnombre.Text
    !idtela = id_tela
    !idcolor = id_color
.Update
Set GRIDMP.DataSource = Rsmateriap
End With
End Sub

Private Sub BtnSalir_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Abrirtbtela
Set Gridtela.DataSource = Rstela
Abrirtbcolor
Set Gridcolor.DataSource = Rscolor
Abrirtbmateriaprima
Set GRIDMP.DataSource = Rsmateriap
txtnombre.Text = ""
txttipotela.Text = ""
txttipocolor.Text = ""
'With Rstela

'End With


End Sub


Sub EstiloGrid()
Gridpj.Columns(0).Width = 1000
Gridpj.Columns(1).Width = 4000
Gridpj.Columns(0).Caption = "Codigo"
Gridpj.Columns(0).Locked = True
Gridpj.Columns(1).Caption = "Nombre"
Gridpj.Columns(1).Locked = True
Gridpj.EditActive = False
Gridpj.HeadFont.Bold = True

End Sub









Private Sub Gridcolor_Click()
id_color = 0
mcodigo = (Gridcolor.Columns(1).Text)
txttipocolor.Text = mcodigo
id_color = (Gridcolor.Columns(0).Text)
'EstiloGrid

mcodigo = ""


End Sub

Private Sub Gridtela_Click()
id_tela = 0
mcodigo = (Gridtela.Columns(1).Text)
txttipotela.Text = mcodigo
id_tela = (Gridtela.Columns(0).Text)
'EstiloGrid

mcodigo = ""

End Sub

