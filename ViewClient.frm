VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ViewClient 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ViewClient.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   17625
   Begin VB.PictureBox FrmDescription 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      Picture         =   "ViewClient.frx":26EAE
      ScaleHeight     =   3735
      ScaleWidth      =   17295
      TabIndex        =   3
      Top             =   3480
      Width           =   17295
      Begin MSFlexGridLib.MSFlexGrid GListVentas 
         Height          =   2055
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   16777215
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
      Begin MSFlexGridLib.MSFlexGrid GListCotizaciones 
         Height          =   2055
         Left            =   4800
         TabIndex        =   7
         Top             =   1200
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   16777215
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
      Begin MSFlexGridLib.MSFlexGrid GListProyecto 
         Height          =   2055
         Left            =   10920
         TabIndex        =   12
         Top             =   1200
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   4194304
         BackColorFixed  =   4194304
         ForeColorFixed  =   16777215
         BackColorSel    =   16776960
         ForeColorSel    =   4194304
         BackColorBkg    =   4194304
         GridColor       =   16777215
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
      Begin VB.Label LblFact 
         BackStyle       =   0  'Transparent
         Caption         =   "Proyectos"
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
         Left            =   11040
         TabIndex        =   14
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label LblHelpProyecto 
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
         Height          =   1455
         Left            =   10920
         TabIndex        =   13
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label LblHelpCotizaciones 
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
         Height          =   1455
         Left            =   4800
         TabIndex        =   11
         Top             =   1080
         Width           =   5895
      End
      Begin VB.Label Lblhelpminimas 
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
         Height          =   1455
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label LblCotizaciones 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizaciones"
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
         Left            =   4920
         TabIndex        =   8
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label LblMinimas 
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas Minimas"
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
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Lbltittledescrip 
         BackStyle       =   0  'Transparent
         Caption         =   "Historial del Cliente:  "
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10695
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
      Left            =   15960
      Picture         =   "ViewClient.frx":4DD5C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid GridList 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   4194304
      BackColorFixed  =   4194304
      ForeColorFixed  =   16777215
      BackColorSel    =   16776960
      ForeColorSel    =   4194304
      BackColorBkg    =   4194304
      GridColor       =   16777215
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
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   15015
   End
   Begin VB.Label lblTitleGL 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuestros Clientes"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "ViewClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    Dim C_Client As New C_CRUD_client
    Dim C_Proc As New C_General_Procedures
    
    Dim ListClient() As Variant
    Dim GUARDAR As String
    LblhelpGeneral.Visible = False
    FrmDescription.Visible = False
    
    'dimencionamos el numero de columnas del grid
    GridList.Cols = GridList.Cols + 7
    
    'cargo titulos del grid
    GridList.TextMatrix(0, 0) = "Cliente"
    GridList.TextMatrix(0, 1) = "Tipo de Documento"
    GridList.TextMatrix(0, 2) = "N° Documento"
    GridList.TextMatrix(0, 3) = "Telefono"
    GridList.TextMatrix(0, 4) = "Dirección"
    GridList.TextMatrix(0, 5) = "Correo"
    GridList.TextMatrix(0, 6) = "Observaciones"
    GridList.TextMatrix(0, 7) = "Fecha de Creación"
    GridList.TextMatrix(0, 8) = "Revisar"
    
    Q_CLIENT = C_Client.Q_ChargeClient
    
    If Q_CLIENT = 0 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "La Base de Datos Actualmente esta vacia"
        Exit Sub
    Else
        Q_CLIENT = Q_CLIENT - 1
        
    End If
   
   'cargamos consulta datos de proyectos pendientes por pago final
    ListClient = C_Client.ChargeClient
   
   'dimencionamos el grid
   GridList.Rows = Q_CLIENT + 2
   
   'inicializamos variables
   IFF = 1
   Columnas = 7
   
   'cargamos el array
   For I = 0 To Q_CLIENT
      For IC = 0 To Columnas
          GridList.TextMatrix(IFF, IC) = ListClient(IC, I)
      Next
      IFF = IFF + 1
   Next
   
   IFF = 1
    
   'CREAR BOTON SELECCIONAR
   For I = 0 To Q_CLIENT
        GridList.TextMatrix(IFF, 8) = "HISTORICO"
        IFF = IFF + 1
   Next
     
   'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList.Rows - 1
        For Col = 0 To GridList.Cols - 1
            GridList.ColWidth(Col) = IIf(Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400 > GridList.ColWidth(Col), Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400, GridList.ColWidth(Col))
        Next
    Next
   
    Call C_Proc.pvSetColors(GridList, RGB(233, 233, 233), RGB(209, 222, 253))
    Call C_Proc.pvSetColorsColumns(GridList, 8, RGB(46, 46, 46))
    Call C_Proc.PaintText(GridList, 8, RGB(255, 255, 255), "HISTORICO")
    
End Sub

'salir del formulario
Private Sub BtnExit_Click()
    Unload ViewClient
End Sub

Private Sub GridList_Click()
    
    Dim C_Proc As New C_General_Procedures
    Dim C_Project As New C_Project
    
    Dim List_GVentas() As Variant
    Dim List_GCot() As Variant
    Dim List_GFact() As Variant
    
    Dim Id_GInput As String
    Dim DocClient As String
    Dim Id_Client As String
    
    GListVentas.Visible = True
    GListCotizaciones.Visible = True
    GListProyecto.Visible = True
    Lblhelpminimas.Visible = False
    LblHelpCotizaciones.Visible = False
    LblHelpProyecto.Visible = False
    
    FrmDescription.Visible = True
    Id_GInput = GridList.Row
    
    'dimencionamos el grid
    GListVentas.Cols = 2
    GListCotizaciones.Cols = 2
    GListProyecto.Cols = 2
     
    DocClient = GridList.TextMatrix(Id_GInput, 2)
    Lbltittledescrip.Caption = "Historial del Cliente:   " & GridList.TextMatrix(Id_GInput, 0)
    Id_Client = C_Proc.Recover_Id("Doc", "Client", DocClient)
        
    Q_ventas = C_Project.Q_SearchProjectClient(Id_Client, "Venta")
    Q_cotizacion = C_Project.Q_SearchProjectClient(Id_Client, "Cotización")
    Q_fact = C_Project.Q_SearchProjectClient(Id_Client, "Factura")
    
    If Q_ventas = 0 Then
        Lblhelpminimas.Visible = True
        GListVentas.Visible = False
        Lblhelpminimas.Caption = "El cliente actualmente no tiene ventas minimas"
    Else
        List_GVentas = C_Project.SearchProjectClient(Id_Client, "Venta")
        
        Q_ventas = Q_ventas - 1
        'dimencionamos el numero de columnas del grid
        GListVentas.Cols = GListVentas.Cols + 1
        
        'cargo titulos del grid
        GListVentas.TextMatrix(0, 0) = "Descripción"
        GListVentas.TextMatrix(0, 1) = "Fecha"
        GListVentas.TextMatrix(0, 2) = "Valor"
        
        'dimencionamos el grid
        GListVentas.Rows = Q_ventas + 2
        
        'inicializamos variables
        IFF = 1
        Columnas = 2
        
        'cargamos el array
        For I = 0 To Q_ventas
           For IC = 0 To Columnas
               GListVentas.TextMatrix(IFF, IC) = List_GVentas(IC, I)
           Next
           IFF = IFF + 1
        Next
          
        'redimencionamos el tamaño de las columnas a los datos digitados
        For Row = 0 To GListVentas.Rows - 1
            For Col = 0 To GListVentas.Cols - 1
                GListVentas.ColWidth(Col) = IIf(Me.TextWidth(GListVentas.TextMatrix(Row, Col)) + 400 > GListVentas.ColWidth(Col), Me.TextWidth(GListVentas.TextMatrix(Row, Col)) + 400, GListVentas.ColWidth(Col))
            Next
        Next
        
        Call C_Proc.pvSetColors(GListVentas, RGB(233, 233, 233), RGB(209, 222, 253))
   
    End If
    
    If Q_cotizacion = 0 Then
        LblHelpCotizaciones.Visible = True
        GListCotizaciones.Visible = False
        LblHelpCotizaciones.Caption = "El cliente actualmente no tiene cotizaciones"
    Else
        List_GCot = C_Project.SearchProjectClient(Id_Client, "Cotización")
        
        Q_cotizacion = Q_cotizacion - 1
        'dimencionamos el numero de columnas del grid
        GListCotizaciones.Cols = GListCotizaciones.Cols + 2
       
        'cargo titulos del grid
        GListCotizaciones.TextMatrix(0, 0) = "Descripción"
        GListCotizaciones.TextMatrix(0, 1) = "Fecha"
        GListCotizaciones.TextMatrix(0, 2) = "Valor"
        GListCotizaciones.TextMatrix(0, 3) = "Estado"
        
        'dimencionamos el grid
        GListCotizaciones.Rows = Q_cotizacion + 2
        
        'inicializamos variables
        IFFC = 1
        ColumnasC = 3
        
        'cargamos el array
        For I = 0 To Q_cotizacion
           For IC = 0 To ColumnasC
               GListCotizaciones.TextMatrix(IFFC, IC) = List_GCot(IC, I)
           Next
           IFFC = IFFC + 1
        Next
          
        'redimencionamos el tamaño de las columnas a los datos digitados
        For Row = 0 To GListCotizaciones.Rows - 1
            For Col = 0 To GListCotizaciones.Cols - 1
                GListCotizaciones.ColWidth(Col) = IIf(Me.TextWidth(GListCotizaciones.TextMatrix(Row, Col)) + 400 > GListCotizaciones.ColWidth(Col), Me.TextWidth(GListCotizaciones.TextMatrix(Row, Col)) + 400, GListCotizaciones.ColWidth(Col))
            Next
        Next
        
        Call C_Proc.pvSetColors(GListCotizaciones, RGB(233, 233, 233), RGB(209, 222, 253))
        Call C_Proc.PaintText(GListCotizaciones, 3, RGB(4, 180, 4), "valida")
        Call C_Proc.PaintText(GListCotizaciones, 3, RGB(223, 1, 1), "Expiro")
   
        
    End If
    
    If Q_fact = 0 Then
        LblHelpProyecto.Visible = True
        GListProyecto.Visible = False
        LblHelpProyecto.Caption = "El cliente actualmente no tiene proyectos"
    Else
        List_GFact = C_Project.SearchProjectClient(Id_Client, "Factura")
        
        Q_fact = Q_fact - 1
        'dimencionamos el numero de columnas del grid
        GListProyecto.Cols = GListProyecto.Cols + 2
              
        'cargo titulos del grid
        GListProyecto.TextMatrix(0, 0) = "Descripción"
        GListProyecto.TextMatrix(0, 1) = "Fecha"
        GListProyecto.TextMatrix(0, 2) = "Valor"
        GListProyecto.TextMatrix(0, 3) = "Estado"
        
        'dimencionamos el grid
        GListProyecto.Rows = Q_fact + 2
        
        'inicializamos variables
        IFFP = 1
        ColumnasP = 3
        
        'cargamos el array
        For I = 0 To Q_fact
           For IC = 0 To ColumnasP
               GListProyecto.TextMatrix(IFFP, IC) = List_GFact(IC, I)
           Next
           IFFP = IFFP + 1
        Next
          
        'redimencionamos el tamaño de las columnas a los datos digitados
        For Row = 0 To GListProyecto.Rows - 1
            For Col = 0 To GListProyecto.Cols - 1
                GListProyecto.ColWidth(Col) = IIf(Me.TextWidth(GListProyecto.TextMatrix(Row, Col)) + 400 > GListProyecto.ColWidth(Col), Me.TextWidth(GListProyecto.TextMatrix(Row, Col)) + 400, GListProyecto.ColWidth(Col))
            Next
        Next
        
        Call C_Proc.pvSetColors(GListProyecto, RGB(233, 233, 233), RGB(209, 222, 253))
        Call C_Proc.PaintText(GListProyecto, 3, RGB(4, 180, 4), "Cancelado")
        Call C_Proc.PaintText(GListProyecto, 3, RGB(223, 1, 1), "inicial")
        
        
        
    End If
    
End Sub
