VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMateriaP1 
   Caption         =   "MATERIAS PRIMAS"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   Begin VB.ComboBox Cbotela 
      Height          =   315
      ItemData        =   "FrmMateriaP1.frx":0000
      Left            =   2160
      List            =   "FrmMateriaP1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Text            =   "TxtColor"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnSalir 
      Caption         =   "SALIR"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnElimina 
      Caption         =   "ELIMINAR"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BtnNuevo 
      Caption         =   "NUEVO"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   0
      Text            =   "Txtnombre"
      Top             =   960
      Width           =   4575
   End
   Begin MSDataGridLib.DataGrid Gridpj 
      Height          =   1935
      Left            =   480
      TabIndex        =   10
      Top             =   3840
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
   Begin VB.Label Label1 
      Caption         =   "Color"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Tela"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FrmMateriaP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Abrirtbtela
Set Gridpj.DataSource = Rstela
EstiloGrid

With Rstela
MoveFirst
Do While Not .EOF

MsgBox (Nombre)

Loop

End With

'LLenar Grid

End Sub

Private Sub BtnSalir_Click(Index As Integer)
Unload Me

End Sub

Private Sub Cbotela_Change()
With Rstela
.Open
While Not .EOF
MsgBox (Nombre)
Loop

End With

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
Gridpj.Columns(0).Caption = "Codigo"
Gridpj.Columns(1).Caption = "Nombre"

Gridpj.HeadFont.Bold = True

End Sub




