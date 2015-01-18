VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InvoiceAndQuotation 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9345
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   15210
   Begin VB.Frame FrmOperative 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   14895
      Begin MSFlexGridLib.MSFlexGrid GridList_Operative 
         Height          =   1215
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   8388608
         BackColorFixed  =   0
         ForeColorFixed  =   65280
         BackColorSel    =   16776960
         ForeColorSel    =   -2147483630
         BackColorBkg    =   0
         GridColorFixed  =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblValue_Winner 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   6000
         TabIndex        =   47
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label LblWinner 
         BackStyle       =   0  'Transparent
         Caption         =   "Ganancia"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4560
         TabIndex        =   46
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label LdlDouble 
         BackStyle       =   0  'Transparent
         Caption         =   "Mano de Obra"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label LdValue_lDouble 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2280
         TabIndex        =   44
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label LblV_Neto 
         BackStyle       =   0  'Transparent
         Caption         =   "V. materiales"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label LblValue_Neto 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2280
         TabIndex        =   42
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label LblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10920
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblValue_Date 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12360
         TabIndex        =   40
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LblValue_Iva 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12360
         TabIndex        =   36
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label LblIva 
         BackStyle       =   0  'Transparent
         Caption         =   "IVA %"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10200
         TabIndex        =   35
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label LblValue_Total 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12360
         TabIndex        =   34
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label LblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10200
         TabIndex        =   33
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label LblValue_Subtotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   12360
         TabIndex        =   32
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label LblSubTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL"
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10200
         TabIndex        =   31
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label LblNumber 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   12360
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Lbltitle_in 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "@Microsoft JhengHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   9960
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame FrmDescription 
      BackColor       =   &H80000007&
      Caption         =   "Descripción del Proyecto"
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   14895
      Begin VB.TextBox TxtDescripProject 
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   14415
      End
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   13320
      TabIndex        =   21
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame FrmCapture 
      BackColor       =   &H80000012&
      Caption         =   "Escoja materiales para la "
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   14895
      Begin VB.TextBox TxtQuanty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   525
         Left            =   11760
         TabIndex        =   48
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OpNot 
         BackColor       =   &H80000012&
         Caption         =   "NO"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   13440
         TabIndex        =   39
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton OpYes 
         BackColor       =   &H80000012&
         Caption         =   "SI"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   12360
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox CbnMeasure 
         Height          =   315
         Left            =   5760
         TabIndex        =   23
         Text            =   "Seleccione..."
         Top             =   360
         Width           =   4455
      End
      Begin VB.ComboBox CbnImputs 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Text            =   "Seleccione..."
         Top             =   360
         Width           =   3615
      End
      Begin MSFlexGridLib.MSFlexGrid GridList_Input 
         Height          =   1215
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   8388608
         BackColorFixed  =   0
         ForeColorFixed  =   65280
         BackColorSel    =   16776960
         ForeColorSel    =   -2147483630
         BackColorBkg    =   0
         GridColorFixed  =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblQuanty 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10680
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblRequest 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¿Requiere IVA?"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   11880
         TabIndex        =   37
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label LblMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "Medida"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblMaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame FrmBody 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   14895
      Begin VB.TextBox TxtObservations 
         Height          =   375
         Left            =   6240
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   8415
      End
      Begin VB.TextBox TxtTypeDocument 
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtEmail 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox TxtAddress 
         Height          =   375
         Left            =   10920
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox TxtPhone 
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtDocument 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label LblObservations 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacion"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblDocumentNumber 
         BackStyle       =   0  'Transparent
         Caption         =   " Documento"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electronico"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   10080
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblPhone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   7080
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton BtnCreateClient 
      Caption         =   "CREAR CLIENTE"
      Height          =   495
      Left            =   11160
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame FrmClient 
      BackColor       =   &H80000007&
      Caption         =   "Escoja cliente para "
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.OptionButton Opdoc 
         BackColor       =   &H80000012&
         Caption         =   "Documento"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OpName 
         BackColor       =   &H80000012&
         Caption         =   "Nombre"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CbnSearch 
         Height          =   315
         ItemData        =   "InvoiceAndQuotation.frx":0000
         Left            =   2760
         List            =   "InvoiceAndQuotation.frx":0002
         TabIndex        =   2
         Text            =   "Seleccione..."
         Top             =   360
         Width           =   5865
      End
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
         Height          =   495
         Left            =   8880
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "InvoiceAndQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Crear_client As Integer
Public Count_GOperative As Long
Public GValueNeto As Long
Public GValueDouble As Long
Public GValueWinner As Long
Public GValueSubTotal As Long


Private Sub Form_Load()

    Dim Inputs() As Variant
    Dim C_Proc As New C_General_Procedures
    
    Crear_client = 0
    FrmBody.Visible = False
    FrmCapture.Visible = False
    FrmDescription.Visible = False
    FrmOperative.Visible = False
    LblValue_Date.Caption = Date
    
    CbnMeasure.Enabled = False
    Count_GOperative = 1
    
    'cargamos consulta datos insumos en BD
    Inputs = C_Proc.Datos_Charge("TC_Inputs", "Charge")
    
    'traemos la cantidad de insumos en BD
    Q_Inputs = C_Proc.Q_Combo("TC_Inputs")
    Q_Inputs = Q_Inputs - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_Inputs
           CbnImputs.AddItem Inputs(1, I)
    Next

    
    'dimencionamos el numero de columnas del grid
    GridList_Input.Cols = GridList_Input.Cols + 4
    
    'cargo titulos del grid
    GridList_Input.TextMatrix(0, 0) = "Proveedor"
    GridList_Input.TextMatrix(0, 1) = "Material"
    GridList_Input.TextMatrix(0, 2) = "Medida"
    GridList_Input.TextMatrix(0, 3) = "Descripción"
    GridList_Input.TextMatrix(0, 4) = "Valor"
    GridList_Input.TextMatrix(0, 5) = "Seleccionar"
    
    GridList_Input.Visible = False
    
    GridList_Operative.Cols = GridList_Operative.Cols + 2
    
    GridList_Operative.TextMatrix(0, 0) = "Descripción"
    GridList_Operative.TextMatrix(0, 1) = "Valor Unidad"
    GridList_Operative.TextMatrix(0, 2) = "Valor Total"
    GridList_Operative.TextMatrix(0, 3) = "Eliminar"
    
End Sub

Private Sub BtnExit_Click()
     Unload InvoiceAndQuotation
End Sub

Private Sub BtnCreateClient_Click()
    
    Crear_client = 1
    Load Client_Crud
    Client_Crud.Left = (MenuCarpenter.ScaleWidth - Client_Crud.Width) / 2
    Client_Crud.Top = (MenuCarpenter.ScaleHeight - Client_Crud.Height) / 2
    Client_Crud.Caption = "Nuevo Cliente"
    Client_Crud.FrmClient.Visible = False
    Client_Crud.BtnCreate.Caption = "CREAR CLIENTE"
    Client_Crud.Show

End Sub

Private Sub CbnSearch_GotFocus()
    
    If Crear_client = 1 Then
         If OpName.Value = True Then
            Call OpName_Click
        Else
            Call Opdoc_Click
        End If
    End If
    
End Sub

'para desbloquear el combo de medidas
Private Sub CbnImputs_GotFocus()
    CbnMeasure.Enabled = True
End Sub

'para cargar el combo de medidas segun el insumo
Private Sub CbnImputs_lostFocus()

   Dim C_ListInt As New C_List_Price
   Dim IndexInputs As Integer
   Dim ListInt() As Variant
    
   CbnMeasure.clear
   
   IndexInputs = CbnImputs.ListIndex
   
   'cargamos consulta datos medidas según el insumo solicitado en BD
   ListInt = C_ListInt.Measure(IndexInputs)
   
   'traemos la cantidad de medidas según el insumo solicitado en BD
   Q_Measure = C_ListInt.Q_Measure(IndexInputs)
   Q_Measure = Q_Measure - 1
   
   'cargamos el combo con los datos seleccionados
   For I = 0 To Q_Measure
          CbnMeasure.AddItem ListInt(1, I)
   Next
   
   GridList_Input.Visible = True
End Sub

'para cargar el combo de medidas segun el insumo
Private Sub CbnMeasure_lostFocus()

   Dim C_Proc As New C_General_Procedures
   Dim C_ListInt As New C_List_Price
   
   Dim IndexInputs As String
   Dim IndexMeasure As String
   
   Dim id As Integer
     
   Dim ListInt() As Variant
     
   IndexInputs = CbnImputs.ListIndex
   
   On Error GoTo ctrlerr
    
   'capturamos el id de la medida
   id = C_Proc.Recover_Id_Detail("TC_Measure", CbnMeasure.Text, IndexInputs)
   
   'cargamos consulta datos medidas según el insumo solicitado en BD
   ListInt = C_ListInt.Charge_List_View_sale(IndexInputs, id)
   
   'traemos la cantidad de medidas según el insumo solicitado en BD
   Q_Charge_List_View = C_ListInt.Q_Charge_List_View_sale(IndexInputs, id)
   Q_Charge_List_View = Q_Charge_List_View - 1
   
   'dimencionamos el grid
   GridList_Input.Rows = Q_Charge_List_View + 2
   
   'inicializamos variables
   IFF = 1
   Columnas = 4
   
   'cargamos el array
   For I = 0 To Q_Charge_List_View
      For IC = 0 To Columnas
          GridList_Input.TextMatrix(IFF, IC) = ListInt(IC, I)
      Next
      IFF = IFF + 1
   Next
   
   IFF = 1
    
   'CREAR BOTON SELECCIONAR
   For I = 0 To Q_Charge_List_View
        GridList_Input.TextMatrix(IFF, 5) = "INGRESAR"
        IFF = IFF + 1
   Next
     
   'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList_Input.Rows - 1
        For Col = 0 To GridList_Input.Cols - 1
            GridList_Input.ColWidth(Col) = IIf(Me.TextWidth(GridList_Input.TextMatrix(Row, Col)) + 400 > GridList_Input.ColWidth(Col), Me.TextWidth(GridList_Input.TextMatrix(Row, Col)) + 400, GridList_Input.ColWidth(Col))
        Next
    Next
    
   Exit Sub
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case 13
    Exit Sub
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    

   
End Sub

Private Sub BtnSearch_Click()

 Dim C_Client As New C_CRUD_client
    Dim Traer_Datos() As Variant
    Dim op_Search As String
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
      
    FrmBody.Visible = True
          
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(0)
      
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el cliente debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        ' traer los datos del registro seleccionado
        Traer_Datos = C_Client.Charge_List_Client(op_Search, CbnSearch.Text)
            
        'cargar los datos capturados en los campos
        
        TxtDocument.Text = Traer_Datos(2, 0)
        TxtTypeDocument.Text = Traer_Datos(1, 0)
        TxtAddress.Text = Traer_Datos(3, 0)
        TxtPhone.Text = Traer_Datos(4, 0)
        TxtEmail.Text = Traer_Datos(5, 0)
        TxtObservations.Text = Traer_Datos(6, 0)
        G_IdDocument = Traer_Datos(7, 0)
        FrmBody.Caption = "Información principal de " & Traer_Datos(0, 0)
        FrmBody.Visible = True
        FrmCapture.Visible = True
        FrmDescription.Visible = True
        block
        
    End If

End Sub

Private Sub GridList_Input_Click()
    
    Dim PosGrid() As Variant
    Dim id_GInput As String
    
    Dim Price As Long
    Dim Price_Total As String
    Dim Description As String
    Dim Rows_Q As Integer
    Dim Q_Inputs As Long
    
    
    If TxtQuanty.Text = "" Then
        MsgBox "El campo cantidad debe estar diligenciado!", vbExclamation + vbOKOnly, "Información!"
        TxtQuanty.SetFocus
        
    Else
         Q_Inputs = TxtQuanty.Text
         id_GInput = GridList_Input.Row
         
         Description = GridList_Input.TextMatrix(id_GInput, 3)
         Price = GridList_Input.TextMatrix(id_GInput, 4)
            
         Price_Total = Price * Q_Inputs
            
         Columnas = 2
          
         Rows_Q = GridList_Operative.Rows
         Rows_Q = Rows_Q + 1
         
         GridList_Operative.Rows = Rows_Q
         
         
        'cargamos el array
        GridList_Operative.TextMatrix(Count_GOperative, 0) = Description
        GridList_Operative.TextMatrix(Count_GOperative, 1) = Price
        GridList_Operative.TextMatrix(Count_GOperative, 2) = Price_Total
        GridList_Operative.TextMatrix(Count_GOperative, 3) = "Eliminar"
        
        Count_GOperative = Count_GOperative + 1
        
        'redimencionamos el tamaño de las columnas a los datos digitados
         For Row = 0 To GridList_Operative.Rows - 1
             For Col = 0 To GridList_Operative.Cols - 1
                 GridList_Operative.ColWidth(Col) = IIf(Me.TextWidth(GridList_Operative.TextMatrix(Row, Col)) + 400 > GridList_Operative.ColWidth(Col), Me.TextWidth(GridList_Operative.TextMatrix(Row, Col)) + 400, GridList_Operative.ColWidth(Col))
             Next
         Next
             
         FrmOperative.Visible = True
         Sum_Values (Price_Total)
    
    End If
    
    

End Sub


'OPCION PARA BUSCAR POR DOCUMENTO
Private Sub Opdoc_Click()
    
    Crear_client = 0
    
    Dim C_Proc As New C_General_Procedures
    Dim cargar_datos() As Variant
    
    CbnSearch.clear
    CbnSearch.Text = "Seleccione..."
    CbnSearch.Width = 2500
    
    'cargamos consulta datos documentos en BD
    cargar_datos = C_Proc.Datos_Charge("Client", "Add_Doc")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Client")
    Q_cargar_datos = Q_cargar_datos - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_cargar_datos
        CbnSearch.AddItem cargar_datos(1, I)
    Next

End Sub
'OPCION PARA BUSCAR POR NOMBRE
Private Sub OpName_Click()
      
    Crear_client = 0
      
    Dim C_Proc As New C_General_Procedures
    Dim cargar_datos() As Variant
    
    CbnSearch.clear
    CbnSearch.Text = "Seleccione..."
    CbnSearch.Width = 5865
    
    'cargamos consulta datos documentos en BD
    cargar_datos = C_Proc.Datos_Charge("Client", "Add_Name")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Client")
    Q_cargar_datos = Q_cargar_datos - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_cargar_datos
        CbnSearch.AddItem cargar_datos(1, I)
    Next

End Sub


Private Sub TxtQuanty_Change()
    
    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    initial = TxtQuanty.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtQuanty.Text = final

End Sub


''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos(verificar As Integer) As Integer
    
    'instanciamos variables
    Dim validate As Integer
    Dim valideedit As Integer
    
    'inicializamos en 0
    validate = 0
    valideedit = 0
    
    If valideedit = 1 Then
        validate = 1
        ValidateCampos = validate
        LblEditCrud.Visible = True
        Exit Function
    Else
    End If
    
    ValidateCampos = validate

End Function

'BLOQUEAR CAMPOS
Function block()

    TxtDocument.Enabled = False
    TxtTypeDocument.Enabled = False
    TxtAddress.Enabled = False
    TxtPhone.Enabled = False
    TxtEmail.Enabled = False
    TxtObservations.Enabled = False
  
End Function

 Function Sum_Values(Value As Long)
   
    'averiguamos si es el primer valor
    If LblValue_Neto.Caption = "" Then
      GValueNeto = 0
    Else
      GValueNeto = LblValue_Neto.Caption
    End If
     
    'sumamos y asignamos el valor de los materiales
    GValueNeto = GValueNeto + Value
    LblValue_Neto.Caption = GValueNeto
    
    'multiplicamos y asignamos el valor de la mano de obra
    GValueDouble = GValueNeto * 1.4
    LdValue_lDouble.Caption = GValueDouble
    
    'multiplicamos y asignamos el valor de la mano de obra
    GValueWinner = GValueNeto * 0.4
    LblValue_Winner.Caption = GValueWinner
    
    'sumamos subtotales
    GValueSubTotal = GValueNeto + GValueDouble + GValueWinner
    LblValue_Subtotal.Caption = GValueSubTotal
    
    TxtQuanty.Text = ""
 End Function

