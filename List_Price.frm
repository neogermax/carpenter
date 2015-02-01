VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form List_Price 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5760
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "List_Price.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   10245
   Begin VB.PictureBox FrmCapture 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      Picture         =   "List_Price.frx":26EAE
      ScaleHeight     =   1335
      ScaleWidth      =   9975
      TabIndex        =   15
      Top             =   1800
      Width           =   9975
      Begin VB.TextBox TxtDescription 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
      Begin VB.TextBox TxtValues 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label LblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LblValues 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblhelpDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   7440
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label LblhelpValue 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.PictureBox FrmBody 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      Picture         =   "List_Price.frx":4DD5C
      ScaleHeight     =   1575
      ScaleWidth      =   9975
      TabIndex        =   4
      Top             =   120
      Width           =   9975
      Begin VB.ComboBox CbnProvider 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Text            =   "Seleccione..."
         Top             =   120
         Width           =   4695
      End
      Begin VB.ComboBox CbnImputs 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "Seleccione..."
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox CbnMeasure 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Text            =   "Seleccione..."
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
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
         Left            =   8400
         Picture         =   "List_Price.frx":74C0A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LblProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label LblMaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LblMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "Medida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label LblhelpProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label LblhelpInput 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6000
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label LblhelpMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridList 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1931
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
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00808080&
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
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "List_Price.frx":762FD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
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
      Left            =   8520
      Picture         =   "List_Price.frx":779F0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   5775
   End
End
Attribute VB_Name = "List_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


''''''----------- REGION ENVENTOS
'INICIO DE FORM LISTADO
Private Sub Form_Load()

    Dim C_Proc As New C_General_Procedures
    Dim provider() As Variant
    Dim Inputs() As Variant
      
    List_Price.Height = 4665
    'ocultar controles
    LblhelpProvider.Visible = False
    LblhelpInput.Visible = False
    LblhelpMeasure.Visible = False
    LblhelpDescription.Visible = False
    LblhelpValue.Visible = False
    LblhelpGeneral.Visible = False
    GridList.Visible = False
          
    'cargamos consulta datos proveedores en BD
    provider = C_Proc.Datos_Charge("Provider", "Add_Name")
    'cargamos consulta datos insumos en BD
    Inputs = C_Proc.Datos_Charge("TC_Inputs", "Charge")
    
    'traemos la cantidad de proveedores en BD
    Q_Provider = C_Proc.Q_Combo("Provider")
    Q_Provider = Q_Provider - 1
    'traemos la cantidad de insumos en BD
    Q_Inputs = C_Proc.Q_Combo("TC_Inputs")
    Q_Inputs = Q_Inputs - 1

    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_Provider
           CbnProvider.AddItem provider(1, I)
    Next
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_Inputs
           CbnImputs.AddItem Inputs(1, I)
    Next
        
    CbnMeasure.Enabled = False
    
    'dimencionamos el numero de columnas del grid
    GridList.Cols = GridList.Cols + 3
    
    'cargo titulos del grid
    GridList.TextMatrix(0, 0) = "Proveedor"
    GridList.TextMatrix(0, 1) = "Insumo"
    GridList.TextMatrix(0, 2) = "Medida"
    GridList.TextMatrix(0, 3) = "Descripción"
    GridList.TextMatrix(0, 4) = "Valor"
    
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
   
   On Error GoTo ctrlerr
    
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

'validar campo numerico
Private Sub TxtValues_Change()

    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    initial = TxtValues.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtValues.Text = final

End Sub

'salir del formulario
Private Sub BtnExit_Click()
    Unload List_Price
End Sub

Private Sub BtnCreate_Click()

    Dim C_Proc As New C_General_Procedures
    Dim C_List As New C_List_Price
    
    Dim validate As Integer
    Dim SL_Prices() As Variant
    
    'validamos campos de diligenciamiento
    validate = ValidateCampos
    
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
     
        Dim Proccess As String
        Proccess = BtnCreate.Caption
            
        If Proccess = "CREAR INSUMO" Then
        
            'traemos la cantidad de clientes en BD
            Q_SL_Prices = C_Proc.Q_Combo("Suppliers_List_Prices")
            Q_SL_Prices = Q_SL_Prices - 1
                   
            'validamos si la tabla es ta vacia o no
            If Q_SL_Prices <> -1 Then
                
                'cargamos consulta datos clientes en BD
                SL_Prices = C_List.Datos_Charge_S()
                
                'recorremos el arreglo para verifica si esta repetido
                For I = 0 To Q_SL_Prices
                    'validamos el campo de cliente o documento
                    If SL_Prices(1, I) = UCase(CbnImputs.Text) And SL_Prices(2, I) = UCase(CbnMeasure.Text) And SL_Prices(3, I) = UCase(CbnProvider.Text) And SL_Prices(4, I) = UCase(TxtDescription.Text) Then
                        MsgBox "el Insumo " & TxtDescription.Text & " del proveedor " & CbnProvider.Text & " ya existe en la base de datos!!!", vbInformation + vbOKOnly, "Información!"
                        LblhelpGeneral.Visible = True
                        LblhelpGeneral.Caption = "el Insumo " & TxtDescription.Text & " del proveedor " & CbnProvider.Text & " ya existe en la base de datos!!!"
                        LblhelpGeneral.ForeColor = &H80&
                        Exit Sub
                    End If
                Next
            
            End If
        
        End If
        
         Select Case Proccess
            Case "CREAR INSUMO"
                'llamar la funcion de insertar en la BD
                Call Insert
                                   
                      
            Case Else
        
        End Select
            
        
    End If
    
End Sub
''''''-----------END_REGION ENVENTOS

''''''-----------REGION FUNCIONES BD
'LLAMAR METODO CREAR CLIENTE
Function Insert()

   Dim C_Proc As New C_General_Procedures
   Dim C_ListInt As New C_List_Price
   Dim id As Integer
   Dim Id_User As Integer
   
   Dim GUARDAR As String
   Dim IndexMeasure As Integer
          
    IndexMeasure = C_Proc.Recover_Id_Detail("TC_Measure", CbnMeasure.Text, CbnImputs.ListIndex)
    'capturamos el proveedor del registro
    id = C_Proc.Recover_Id("Name", "Provider", CbnProvider.Text)
    Id_User = C_Proc.Recover_Id("User", "Users", MenuCarpenter.Lbl_Value_User.Caption)
       
    'llamamos la funcion crear cliente
    GUARDAR = C_ListInt.Add_ListInt(id, CbnImputs.ListIndex, IndexMeasure, UCase(TxtDescription.Text), UCase(TxtValues.Text), Id_User)

    'validamos el resultado de la operacion anterior
    If GUARDAR = "OK" Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Insumo creado con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        InsertGrid
        
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If

End Function
''''''-----------END_REGION FUNCIONES BD

''''''-----------REGION FUNCIONES
'validar campos obligatorios
Function ValidateCampos()

    'instanciamos variables
    Dim valProvider As Integer
    Dim valInput As Integer
    Dim valmeasure As Integer
    Dim valDescription As Integer
    Dim valValues As Integer
    Dim Valide_Total As Integer
    
    'inicializamos en 0
    valProvider = 0
    valInput = 0
    valmeasure = 0
    valDescription = 0
    valValues = 0
    Valide_Total = 0
    
    'validamos campos obligatorios
    If CbnProvider.Text = "Seleccione..." Then
          valProvider = 1
    End If
    If CbnImputs.Text = "Seleccione..." Then
          valInput = 1
    End If
    If CbnMeasure.Text = "Seleccione..." Then
          valmeasure = 1
    End If
    If TxtDescription.Text = "" Then
          valDescription = 1
    End If
    If TxtValues.Text = "" Then
          valValues = 1
    End If

    'verificamos validacion anterior
    If valProvider = 1 Or valInput = 1 Or valmeasure = 1 Or valDescription = 1 Or valValues = 1 Then
    
        Valide_Total = 1
        
        'validamos para mostrar campos sin diligenciar
        If valProvider = 1 Then
            LblhelpProvider.Visible = True
        Else
            LblhelpProvider.Visible = False
        End If
        If valInput = 1 Then
            LblhelpInput.Visible = True
        Else
            LblhelpInput.Visible = False
        End If
        If valmeasure = 1 Then
            LblhelpMeasure.Visible = True
        Else
            LblhelpMeasure.Visible = False
        End If
        If valDescription = 1 Then
            LblhelpDescription.Visible = True
        Else
            LblhelpDescription.Visible = False
        End If
        If valValues = 1 Then
            LblhelpValue.Visible = True
        Else
            LblhelpValue.Visible = False
        End If
    Else
    
        LblhelpProvider.Visible = False
        LblhelpInput.Visible = False
        LblhelpMeasure.Visible = False
        LblhelpDescription.Visible = False
        LblhelpValue.Visible = False
    
    End If

    ValidateCampos = Valide_Total
    
End Function

Function InsertGrid()

    Dim Pos_Grid As Integer
    
    List_Price.Height = 6250
    
    'habilitamos el grid
    GridList.Visible = True
    'agregamos una nueva fila
    GridList.Rows = GridList.Rows + 1
    
    'asignamos la posicion del grid
    Pos_Grid = GridList.Rows - 1
    
    'cargamos los datos
    GridList.TextMatrix(Pos_Grid, 0) = UCase(CbnProvider.Text)
    GridList.TextMatrix(Pos_Grid, 1) = UCase(CbnImputs.Text)
    GridList.TextMatrix(Pos_Grid, 2) = UCase(CbnMeasure.Text)
    GridList.TextMatrix(Pos_Grid, 3) = UCase(TxtDescription.Text)
    GridList.TextMatrix(Pos_Grid, 4) = TxtValues.Text
        
    'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList.Rows - 1
        For Col = 0 To GridList.Cols - 1
            GridList.ColWidth(Col) = IIf(Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400 > GridList.ColWidth(Col), Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400, GridList.ColWidth(Col))
        Next
    Next
    
    'alineamos columnas disparejas
    GridList.ColAlignment(2) = 1
    GridList.ColAlignment(4) = 4
    'llamamos funcion limpiar
    clear
    
End Function

Function clear()
    
    CbnProvider.Text = "Seleccione..."
    CbnImputs.Text = "Seleccione..."
    CbnMeasure.Text = "Seleccione..."
    TxtDescription.Text = ""
    TxtValues.Text = ""
    
End Function
''''''-----------END REGION FUNCIONES
