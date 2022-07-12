VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmColor 
   Caption         =   "COLORES"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9210
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   960
      Width           =   4575
   End
   Begin MSDataGridLib.DataGrid Gridpj 
      Height          =   1935
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   5535
      _ExtentX        =   9763
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
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "INTERFAZ COLORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Color"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FrmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnElimina_Click(Index As Integer)
wcodigo = Gridpj.Columns(0).Text
MsgBox (wcodigo)
With Rscolor
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
Abrirtbcolor

'LLenar Grid
Set Gridpj.DataSource = Rscolor
EstiloGrid
txtnombre.Text = ""
End Sub

Private Sub BtnGrabar_Click(Index As Integer)
With Rscolor
.AddNew
    !Nombre = txtnombre.Text
.Update
Set Gridpj.DataSource = Rscolor
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


