VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PaymentsCredits 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   16305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "PaymentsCredits.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   16305
   Begin VB.PictureBox FrmDescription 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   120
      Picture         =   "PaymentsCredits.frx":26EAE
      ScaleHeight     =   5295
      ScaleWidth      =   15855
      TabIndex        =   3
      Top             =   2400
      Width           =   15855
      Begin VB.CommandButton BtnCreate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FINALIZAR PROYECTO"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         Picture         =   "PaymentsCredits.frx":4DD5C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox TxtFactPro 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         TabIndex        =   21
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton BtnIn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INGRESAR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13920
         Picture         =   "PaymentsCredits.frx":4F44F
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox TxtReal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         TabIndex        =   19
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox TxtDescripProject 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   10815
      End
      Begin MSFlexGridLib.MSFlexGrid GridList_Detail 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   4194304
         GridColorFixed  =   16777215
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
      Begin VB.Label LblDif 
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   29
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label LblVDif 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   9000
         TabIndex        =   28
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label LblTC 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Comprado "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   4080
         TabIndex        =   27
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label LblTVC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label LblTVM 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Facturado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label LblValTVM 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label LblhelpGeneral 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   4560
         Width           =   11775
      End
      Begin VB.Label LblFactPro 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Fact Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   12000
         TabIndex        =   22
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label LblReal 
         BackStyle       =   0  'Transparent
         Caption         =   "Vr Real"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   12000
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label LblFact 
         BackStyle       =   0  'Transparent
         Caption         =   "Vr Facturado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   12000
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label LblFactVal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   13560
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label LblMaterialsVal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   13560
         TabIndex        =   15
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Lblmaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   12000
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label LblNumberFact 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   13560
         TabIndex        =   13
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Lblfature 
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11880
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label LblClient 
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label LlbDetailProject 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle del proyecto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Lblsale 
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO PENDIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   13560
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label LblValueSale 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   13560
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Lbltitle_in 
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11520
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label LblVallueTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   11160
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      Picture         =   "PaymentsCredits.frx":50B42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid GridList_Project 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      ForeColor       =   4194304
      BackColorFixed  =   4194304
      ForeColorFixed  =   16777215
      BackColorSel    =   16776960
      ForeColorSel    =   4194304
      BackColorBkg    =   4194304
      GridColor       =   4194304
      GridColorFixed  =   16777215
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
   Begin VB.Label GridList_ 
      BackStyle       =   0  'Transparent
      Caption         =   "Proyectos  en proceso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "PaymentsCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public G_PosDetail As Integer
Public CountItems As Long
Public G_IdProject As Integer

Private Sub Form_Load()
    
    Dim C_Project As New C_Project
    Dim ListProject() As Variant
    Dim C_Proc As New C_General_Procedures
      
    CountItems = 0
    FrmDescription.Visible = False
    LblhelpGeneral.Visible = False
    
    'dimencionamos el numero de columnas del grid
    GridList_Project.Cols = GridList_Project.Cols + 2
    
    'cargo titulos del grid
    GridList_Project.TextMatrix(0, 0) = "N° de Factura"
    GridList_Project.TextMatrix(0, 1) = "Cliente"
    GridList_Project.TextMatrix(0, 2) = "Valor Total"
    GridList_Project.TextMatrix(0, 3) = "Acción"
    
    'cargamos consulta datos de proyectos pendientes por pago final
    ListProject = C_Project.Charge_GProject
    
    Q_Project = C_Project.Q_Charge_GProject
    Q_Project = Q_Project - 1
    
    'dimencionamos el grid
   GridList_Project.Rows = Q_Project + 2
   
   'inicializamos variables
   IFF = 1
   Columnas = 2
   
   'cargamos el array
   For I = 0 To Q_Project
      For IC = 0 To Columnas
          GridList_Project.TextMatrix(IFF, IC) = ListProject(IC, I)
      Next
      IFF = IFF + 1
   Next
   
   IFF = 1
    
   'CREAR BOTON SELECCIONAR
   For I = 0 To Q_Project
        GridList_Project.TextMatrix(IFF, 3) = "BUSCAR"
        IFF = IFF + 1
   Next
     
   'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList_Project.Rows - 1
        For Col = 0 To GridList_Project.Cols - 1
            GridList_Project.ColWidth(Col) = IIf(Me.TextWidth(GridList_Project.TextMatrix(Row, Col)) + 400 > GridList_Project.ColWidth(Col), Me.TextWidth(GridList_Project.TextMatrix(Row, Col)) + 400, GridList_Project.ColWidth(Col))
        Next
    Next
    
    Call C_Proc.pvSetColors(GridList_Project, RGB(233, 233, 233), RGB(209, 222, 253))
    Call C_Proc.pvSetColorsColumns(GridList_Project, 3, RGB(46, 46, 46))
    Call C_Proc.PaintText(GridList_Project, 3, RGB(255, 255, 255), "BUSCAR")
End Sub


'boton salir
Private Sub BtnExit_Click()
     Unload PaymentsCredits
End Sub

Private Sub GridList_Project_Click()
    
    Dim C_Project As New C_Project
    Dim ListProject() As Variant
    Dim ListProjectDetails() As Variant
    
    Dim id_GInput As String
    Dim IdProject As Integer
    
    Dim C_Proc As New C_General_Procedures
    
     
    id_GInput = GridList_Project.Row
    FrmDescription.Visible = True
      
    LblClient.Caption = "CLIENTE: " & GridList_Project.TextMatrix(id_GInput, 1)
    LblNumberFact.Caption = GridList_Project.TextMatrix(id_GInput, 0)
      
    ListProject = C_Project.DateProjectDetail(id_GInput)
    
    TxtDescripProject.Text = ListProject(0, 0)
    TxtDescripProject.Enabled = False
    LblVallueTotal.Caption = Format(ListProject(1, 0), "####,####")
    LblValueSale.Caption = Format(ListProject(2, 0), "####,####")
    
    'capturamos el id del proyecto de la operacion
    IdProject = C_Proc.Recover_Id("project", "project", id_GInput)
    'capturamos el detalle del proyecto
    ListProjectDetails = C_Project.SearchDetailsProject(IdProject)
    
    'dimencionamos el numero de columnas del grid
    GridList_Detail.Cols = GridList_Detail.Cols + 7
    'cargo titulos del grid
    GridList_Detail.TextMatrix(0, 0) = "Material"
    GridList_Detail.TextMatrix(0, 1) = "Medida"
    GridList_Detail.TextMatrix(0, 2) = "Cantidad"
    GridList_Detail.TextMatrix(0, 3) = "Valor Unitario"
    GridList_Detail.TextMatrix(0, 4) = "Valor Total"
    GridList_Detail.TextMatrix(0, 5) = "Acción"
    GridList_Detail.TextMatrix(0, 6) = "N° Fac. Proveedor"
    GridList_Detail.TextMatrix(0, 7) = "Valor Factura"
    GridList_Detail.TextMatrix(0, 8) = "valor Conciliado"
    
    G_IdProject = IdProject
    
    Q_ProjectDetails = C_Project.Q_SearchDetailsProject(IdProject)
    Q_ProjectDetails = Q_ProjectDetails - 1
    
    'dimencionamos el grid
    GridList_Detail.Rows = Q_ProjectDetails + 2
    
    'inicializamos variables
    IFF = 1
    Columnas = 4
    
    'cargamos el array
    For I = 0 To Q_ProjectDetails
       For IC = 0 To Columnas
           GridList_Detail.TextMatrix(IFF, IC) = ListProjectDetails(IC, I)
       Next
       IFF = IFF + 1
    Next
    
   
   IFF = 1
    
   'CREAR BOTON SELECCIONAR
   For I = 0 To Q_ProjectDetails
        GridList_Detail.TextMatrix(IFF, 5) = "CONCILIAR"
        IFF = IFF + 1
   Next
    
   
    'redimencionamos el tamaño de las columnas a los datos digitados
     For Row = 0 To GridList_Detail.Rows - 1
         For Col = 0 To GridList_Detail.Cols - 1
             GridList_Detail.ColWidth(Col) = IIf(Me.TextWidth(GridList_Detail.TextMatrix(Row, Col)) + 400 > GridList_Detail.ColWidth(Col), Me.TextWidth(GridList_Detail.TextMatrix(Row, Col)) + 400, GridList_Detail.ColWidth(Col))
         Next
     Next


   Call C_Proc.ColorsColumns(GridList_Detail, RGB(233, 233, 233), RGB(209, 222, 253))
   
   Call C_Proc.pvSetColorsColumns(GridList_Detail, 4, RGB(184, 183, 153))
   Call C_Proc.pvSetColorsColumns(GridList_Detail, 7, RGB(184, 183, 153))
   Call C_Proc.pvSetColorsColumns(GridList_Detail, 8, RGB(132, 195, 190))
   
   Call C_Proc.pvSetColorsColumns(GridList_Detail, 5, RGB(46, 46, 46))
   Call C_Proc.PaintText(GridList_Detail, 5, RGB(255, 255, 255), "CONCILIAR")
   
   Suma_total = C_Proc.Sum_Columns(GridList_Detail, 4)
   LblValTVM.Caption = Format(Suma_total, "####,####")
        
End Sub
    
Private Sub GridList_Detail_Click()
 
    Dim id_GInput As String
    
    id_GInput = GridList_Detail.Row
    G_PosDetail = id_GInput
    LblMaterialsVal.Caption = GridList_Detail.TextMatrix(id_GInput, 0)
    LblFactVal.Caption = Format(GridList_Detail.TextMatrix(id_GInput, 4), "####,####")
    TxtFactPro.Text = GridList_Detail.TextMatrix(id_GInput, 6)
    TxtReal.Text = GridList_Detail.TextMatrix(id_GInput, 7)

End Sub


Private Sub BtnIn_Click()

    Dim C_Proc As New C_General_Procedures
    
    Dim valfact As String
    Dim valreal As String
    Dim valextra As String
    
    If TxtReal.Text = "" Or TxtFactPro.Text = "" Then
         
        If TxtReal.Text = "" Then
          TxtReal.BackColor = &H40&
        Else
          TxtReal.BackColor = &HFFFFFF
        End If
        If TxtFactPro.Text = "" Then
           TxtFactPro.BackColor = &H40&
        Else
          TxtFactPro.BackColor = &HFFFFFF
        End If
          
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para ingresar la conciliación debe llenar los campos en Rojo!!!"
        LblhelpGeneral.ForeColor = &H80&
          
    Else
    
        valfact = Format(LblFactVal.Caption, "##")
        valreal = TxtReal.Text
        valextra = valfact - valreal
        
        GridList_Detail.TextMatrix(G_PosDetail, 6) = TxtFactPro.Text
        GridList_Detail.TextMatrix(G_PosDetail, 7) = valreal
        GridList_Detail.TextMatrix(G_PosDetail, 8) = valextra
         
        Suma_Real = C_Proc.Sum_Columns(GridList_Detail, 7)
        LblTVC.Caption = Format(Suma_Real, "####,####")
        Suma_Conciliar = C_Proc.Sum_Columns(GridList_Detail, 8)
        LblVDif.Caption = Format(Suma_Conciliar, "####,####")
        
        LblMaterialsVal.Caption = ""
        LblFactVal.Caption = ""
        TxtReal.Text = ""
        TxtFactPro.Text = ""
        
        CountItems = CountItems + 1
        
        verifica = C_Proc.NumbersNegative(GridList_Detail, 8, RGB(138, 8, 8))

    End If
    
       
End Sub

Private Sub BtnCreate_Click()

    Dim C_Project As New C_Project
    Dim Id_LProjectDetails() As Variant
    
    Q_ProjectDetails = C_Project.Q_SearchDetailsProject(G_IdProject)
    
    If Q_ProjectDetails <= CountItems Then
        LblhelpGeneral.Caption = ""
        'capturos los id de detalles de proyecto
        Id_LProjectDetails = C_Project.SearchDetailsProject_ID(G_IdProject)
        
        Dim Q_GL_Operative As Integer
        Dim Id_Details As Integer
        Dim NumberProviderFact As String
        Dim Vr_FactProvider As String
        Dim Winner_Loser As String
        Dim GUARDAR_PD As String
        Dim GUARDAR_P As String
        
        Dim IC As Integer
        
        Q_GL_Operative = GridList_Detail.Rows
        Q_GL_Operative = Q_GL_Operative - 1
        IC = 0
        
        For I = 1 To Q_GL_Operative
            
            Id_Details = Id_LProjectDetails(0, IC)
            NumberProviderFact = GridList_Detail.TextMatrix(I, 6)
            Vr_FactProvider = GridList_Detail.TextMatrix(I, 7)
            Winner_Loser = GridList_Detail.TextMatrix(I, 8)
            
            'guardamos los detalles de la factura
            GUARDAR_PD = C_Project.Update_DatailProject(Id_Details, NumberProviderFact, Vr_FactProvider, Winner_Loser)
            IC = IC + 1
        Next
        
        GUARDAR_P = C_Project.Update_Project(G_IdProject)
        
         'validamos el resultado de la operacion anterior
        If GUARDAR_P = "OK" And GUARDAR_PD = "OK" Then
        
            LblhelpGeneral.Visible = True
            LblhelpGeneral.Caption = " conciliación del proyecto realizada con exito!"
            LblhelpGeneral.ForeColor = &H8000&
            
        Else
            
            LblhelpGeneral.Visible = True
            LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
            LblhelpGeneral.ForeColor = &H80&
        
        End If

        
    Else
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "debe diligenciar todos materiales!!!"
        LblhelpGeneral.ForeColor = &H80&
    End If
    
    
End Sub


'VALIDAR CAMPO NUMERICO
Private Sub TxtReal_Change()
    
    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    TxtReal.BackColor = &HFFFFFF
    
    initial = TxtReal.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtReal.Text = final

End Sub

'VALIDAR CAMPO NUMERICO
Private Sub TxtFactPro_Change()
    
    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    TxtFactPro.BackColor = &HFFFFFF
   
    initial = TxtFactPro.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtFactPro.Text = final

End Sub
