VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form List_Price 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10245
   Begin MSFlexGridLib.MSFlexGrid GridList 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1931
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
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame FrmCapture 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   9975
      Begin VB.TextBox TxtValues 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtDescription 
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label LblhelpValue 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label LblhelpDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   7200
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label LblValues 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame FrmBody 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
         Height          =   495
         Left            =   8400
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox CbnMeasure 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "Seleccione..."
         Top             =   1200
         Width           =   4695
      End
      Begin VB.ComboBox CbnImputs 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "Seleccione..."
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox CbnProvider 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label LblhelpMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label LblhelpInput 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label LblhelpProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Campo Obligatorio!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label LblMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "Medida"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblMaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label LblhelpGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   3840
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
                                   
            
            Case "MODIFICAR INSUMO"
                 'llamar la funcion de modificar en la BD
                 Call Update
                
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
   Dim guardar As String
   
       
    'capturamos el proveedor del registro
    id = C_Proc.Recover_Id("Name", "Provider", CbnProvider.Text)
       
    'llamamos la funcion crear cliente
    guardar = C_ListInt.Add_ListInt(id, CbnImputs.ListIndex, CbnMeasure.ListIndex, UCase(TxtDescription.Text), UCase(TxtValues.Text))

    'validamos el resultado de la operacion anterior
    If guardar = "OK" Then
    
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

'LLAMAR METODO MODIFICAR CLIENTE
Function Update()

    Dim C_ListInt As New C_List_Price
    Dim C_Proc As New C_General_Procedures
    Dim guardar As String
    Dim op_Search As String
    Dim id As Integer
    
    'averiguar por que metodo buscar el id
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
   
    'capturamos el id del registro
    id = C_Proc.Recover_Id(op_Search, "Client", CbnSearch.Text)
    
    'actualizar el registro seleccionado
    guardar = C_Client.Update_Client(id, UCase(TxtName.Text), G_IdDocument, UCase(TxtDocument.Text), UCase(TxtAddress.Text), UCase(TxtPhone.Text), UCase(TxtEmail.Text), UCase(TxtObservations.Text))

    'validamos el resultado de la operacion anterior
    If guardar = "OK" Then
        
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Cliente ha sido modificado con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        InsertGrid
        
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No actualizo revisar insercion a la BD!"
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
