VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmcprenda 
   Caption         =   "CONSULTA TIPO PRENDA"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Gridprenda 
      Height          =   1935
      Left            =   5280
      TabIndex        =   7
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
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
   Begin VB.CommandButton Btnbuscar 
      Caption         =   "BUSCAR PRENDA"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txttipotela 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid Gridpj 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   1
      Top             =   3960
      Width           =   1250
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "PRENDAS - UNIDADES"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Tela Seleccionada"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "SELECCIONE TIPO DE TELA"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Frmcprenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnElimina_Click(Index As Integer)
wcodigo = Gridpj.Columns(0).Text
MsgBox (wcodigo)
With Rstela
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
Abrirtbtela

'LLenar Grid
Set Gridpj.DataSource = Rstela
EstiloGrid
txttipotela.Text = ""
End Sub

Private Sub BtnGrabar_Click(Index As Integer)
With Rstela
.AddNew
    !Nombre = txtnombre.Text
.Update
Set Gridpj.DataSource = Rstela
End With
End Sub

Private Sub BtnNuevo_Click(Index As Integer)
txtnombre.Enabled = True
txtnombre.Text = ""
txtnombre.SetFocus
End Sub



Sub EstiloGrid()
Gridpj.Columns(0).Width = 1000
Gridpj.Columns(1).Width = 4000
Gridpj.Columns(0).Caption = "Codigo"
Gridpj.Columns(1).Caption = "Nombre"

Gridpj.HeadFont.Bold = True

End Sub

Private Sub Gridpj_Click()
id_tela = 0
id_tela = (Gridpj.Columns(0).Text)
txttipotela.Text = (Gridpj.Columns(1).Text)

End Sub
