VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form List_PriceEdit 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10260
   Begin VB.Frame FrmBody 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9975
      Begin VB.ComboBox CbnDescription 
         Height          =   315
         Left            =   1320
         TabIndex        =   22
         Text            =   "Seleccione..."
         Top             =   1680
         Width           =   4695
      End
      Begin VB.ComboBox CbnProvider 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox CbnImputs 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Text            =   "Seleccione..."
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox CbnMeasure 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Text            =   "Seleccione..."
         Top             =   1200
         Width           =   4695
      End
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
         Height          =   495
         Left            =   8400
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LblDescriptionC 
         BackStyle       =   0  'Transparent
         Caption         =   "insumos a modificar"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label LblHelpNewdescription 
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
         TabIndex        =   23
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label LblProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblMaterials 
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblMeasure 
         BackStyle       =   0  'Transparent
         Caption         =   "Medida"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   17
         Top             =   240
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
         TabIndex        =   16
         Top             =   720
         Width           =   2415
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
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.Frame FrmCapture 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   9975
      Begin VB.TextBox TxtDescription 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox TxtValues 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label LblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblValues 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   855
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
         TabIndex        =   7
         Top             =   480
         Width           =   2415
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
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid GridList 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   4800
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
      TabIndex        =   21
      Top             =   3960
      Width           =   5775
   End
End
Attribute VB_Name = "List_PriceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''' REGION GLOBALES
Public G_Id_Provider As String
'''''' END REGION GLOBALES

''''''----------- REGION ENVENTOS
'INICIO DE FORM LISTADO
Private Sub Form_Load()

    Dim C_Proc As New C_General_Procedures
    Dim provider() As Variant
    Dim Inputs() As Variant
      
    List_PriceEdit.Height = 4755
    'ocultar controles
    LblhelpProvider.Visible = False
    LblhelpInput.Visible = False
    LblhelpMeasure.Visible = False
    LblhelpDescription.Visible = False
    LblhelpValue.Visible = False
    LblhelpGeneral.Visible = False
    LblHelpNewdescription.Visible = False
    GridList.Visible = False
    TxtDescription.Enabled = False
    TxtValues.Enabled = False
          
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
    CbnDescription.Enabled = False
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
   
   On Error GoTo ctrlerr
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

'para desbloquear el combo de descripción
Private Sub CbnMeasure_GotFocus()
    CbnDescription.Enabled = True
End Sub

'para cargar el combo de medidas segun el insumo
Private Sub CbnMeasure_lostFocus()

   Dim C_Proc As New C_General_Procedures
   Dim C_ListInt As New C_List_Price
   
   Dim IndexInputs As String
   Dim IndexMeasure As Integer
   Dim TextDescription As String
   
   Dim id As Integer
     
   List_PriceEdit.Height = 6500
   GridList.Visible = True
     
   On Error GoTo ctrlerr
     
   Dim ListInt() As Variant
    
   CbnDescription.clear
   
   IndexInputs = CbnImputs.ListIndex
   IndexMeasure = CbnMeasure.ListIndex
   
    'capturamos el id del registro
   id = C_Proc.Recover_Id("Name", "Provider", CbnProvider.Text)
    
   'capturamos el id de la medida
   IndexMeasure = C_Proc.Recover_Id_Detail("TC_Measure", CbnMeasure.Text, IndexInputs)
  
    
   'cargamos consulta datos medidas según el insumo solicitado en BD
   ListInt = C_ListInt.Charge_List_View(id, IndexInputs, IndexMeasure, TextDescription, "General")
   
   'traemos la cantidad de medidas según el insumo solicitado en BD
   Q_Charge_List_View = C_ListInt.Q_Charge_List_View(id, IndexInputs, IndexMeasure)
   Q_Charge_List_View = Q_Charge_List_View - 1
   
   'cargamos el combo con los datos seleccionados
   For I = 0 To Q_Charge_List_View
          CbnDescription.AddItem ListInt(3, I)
   Next
    
   'dimencionamos el grid
   GridList.Rows = Q_Charge_List_View + 2
   
   'inicializamos variables
   IFF = 1
   Columnas = 4
   
   'cargamos el array
   For I = 0 To Q_Charge_List_View
      For IC = 0 To Columnas
          GridList.TextMatrix(IFF, IC) = ListInt(IC, I)
      Next
      IFF = IFF + 1
   Next
   
    'redimencionamos el tamaño de las columnas a los datos digitados
    For Row = 0 To GridList.Rows - 1
        For Col = 0 To GridList.Cols - 1
            GridList.ColWidth(Col) = IIf(Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400 > GridList.ColWidth(Col), Me.TextWidth(GridList.TextMatrix(Row, Col)) + 400, GridList.ColWidth(Col))
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
    Unload List_PriceEdit
End Sub

'buscar los datos seleccionados
Private Sub BtnSearch_Click()

    Dim C_Proc As New C_General_Procedures
    Dim C_ListInt As New C_List_Price
   
    Dim IndexInputs As String
    Dim IndexMeasure As Integer
    Dim TextDescription As String
   
    Dim id As Integer
    Dim ListInt() As Variant

    On Error GoTo ctrlerr

    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(0)
      
    
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el cliente debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        LblhelpGeneral.Caption = ""
        
        IndexInputs = CbnImputs.ListIndex
        TextDescription = CbnDescription
        
        TxtDescription.Enabled = True
        TxtValues.Enabled = True
    
        'capturamos el id del registro
        id = C_Proc.Recover_Id("Name", "Provider", CbnProvider.Text)
    
        IndexMeasure = C_Proc.Recover_Id_Detail("TC_Measure", CbnMeasure.Text, IndexInputs)
  
        ListInt = C_ListInt.Charge_List_View(id, IndexInputs, IndexMeasure, TextDescription, "Detallado")
        
        TxtDescription.Text = ListInt(3, 0)
        TxtValues.Text = ListInt(4, 0)
        G_Id_Provider = ListInt(8, 0)
       
    End If
    
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
'actualizar el registro del listado
Private Sub BtnCreate_Click()
    Dim C_Proc As New C_General_Procedures
    Dim C_ListInt As New C_List_Price
   
    Dim id As Integer
    Dim ListInt() As Variant

    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(1)
        
    If validate = 1 Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el cliente debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        LblhelpGeneral.Caption = ""
       
        Dim guardar As String
       
        'actualizar el registro seleccionado
        guardar = C_ListInt.Update_List(G_Id_Provider, UCase(TxtDescription.Text), UCase(TxtValues.Text))
    
        'validamos el resultado de la operacion anterior
        If guardar = "OK" Then
            
            LblhelpGeneral.Visible = True
            LblhelpGeneral.Caption = "Cliente ha sido modificado con exito!"
            LblhelpGeneral.ForeColor = &H8000&
            clear
            
        Else
        
            LblhelpGeneral.Visible = True
            LblhelpGeneral.Caption = "No actualizo revisar insercion a la BD!"
            LblhelpGeneral.ForeColor = &H80&
        
        End If
   
    End If

End Sub
''''''-----------END_REGION ENVENTOS

''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos(verificar As Integer) As Integer
    
    'instanciamos variables
    Dim valProvider As Integer
    Dim valInput As Integer
    Dim valmeasure As Integer
    Dim valDescription As Integer
    Dim valNewDescription As Integer
    Dim valValues As Integer
      
    Dim validate As Integer
    
    'inicializamos en 0
    valProvider = 0
    valInput = 0
    valmeasure = 0
    valDescription = 0
    valNewDescription = 0
    valValues = 0
      
    If CbnProvider.Text = "Seleccione..." Then
         valProvider = 1
    End If
    If CbnImputs.Text = "Seleccione..." Then
         valInput = 1
    End If
    If CbnMeasure.Text = "Seleccione..." Then
         valmeasure = 1
    End If
    If CbnDescription.Text = "Seleccione..." Then
         valDescription = 1
    End If
    
    'verificamos validacion anterior
    If valProvider = 1 Or valInput = 1 Or valmeasure = 1 Or valDescription = 1 Then
    
        validate = 1
        
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
            LblHelpNewdescription.Visible = True
        Else
            LblHelpNewdescription.Visible = False
        End If
        
    Else
        
        LblhelpProvider.Visible = False
        LblhelpInput.Visible = False
        LblhelpMeasure.Visible = False
        LblHelpNewdescription.Visible = False
        
    End If

    
    If verificar = 1 Then
         
         'validamos campos obligatorios
         If TxtDescription.Text = "" Then
              valNewDescription = 1
         End If
         If TxtValues.Text = "" Then
              valValues = 1
         End If
               
         'verificamos validacion anterior
         If valNewDescription = 1 Or valValues = 1 Then
         
             validate = 1
             
             'validamos para mostrar campos sin diligenciar
             If valNewDescription = 1 Then
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
             'ocultar labels de mensajes
             LblhelpDescription.Visible = False
             LblhelpValue.Visible = False
         End If
    
    End If
    
    ValidateCampos = validate

End Function

'limpiar campos
Function clear()
    
    CbnProvider.Text = "Seleccione..."
    CbnImputs.Text = "Seleccione..."
    CbnMeasure.Text = "Seleccione..."
    CbnDescription.Text = "Seleccione..."
    TxtDescription.Text = ""
    TxtValues.Text = ""
    TxtDescription.Enabled = False
    TxtValues.Enabled = False
End Function
''''''----------- END_REGION FUNCIONES
