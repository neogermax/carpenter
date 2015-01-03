VERSION 5.00
Begin VB.Form Client_Crud 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12780
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   12780
   Begin VB.Frame FrmBody 
      BackColor       =   &H80000007&
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   12495
      Begin VB.TextBox TxtName 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   7095
      End
      Begin VB.ComboBox CbnTypeDocument 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Text            =   "Seleccione..."
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtDocument 
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtPhone 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox TxtAddress 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TxtEmail 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox TxtObservations 
         Height          =   855
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   7095
      End
      Begin VB.Label LblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblDocument 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo De Documento"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblPhone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label LblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electronico"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label LblObservations 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacion"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label LblDocumentNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "N� Documento"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblhelpName 
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
         Left            =   9360
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label LblhelpDoc 
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
         Left            =   9360
         TabIndex        =   17
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LlbhelpPhone 
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
         Left            =   5520
         TabIndex        =   16
         Top             =   1440
         Width           =   2415
      End
   End
   Begin VB.Frame FrmClient 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   12495
      Begin VB.CommandButton BtnSearch 
         Caption         =   "BUSCAR"
         Height          =   495
         Left            =   8880
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CbnSearch 
         Height          =   315
         ItemData        =   "Client_Create.frx":0000
         Left            =   2760
         List            =   "Client_Create.frx":0002
         TabIndex        =   0
         Text            =   "Seleccione..."
         Top             =   360
         Width           =   5865
      End
      Begin VB.OptionButton OpName 
         BackColor       =   &H80000012&
         Caption         =   "Nombre"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Opdoc 
         BackColor       =   &H80000012&
         Caption         =   "Documento"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblEditCrud 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Debe seleccionar una opci�n"
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
         Height          =   495
         Left            =   10560
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   11160
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton BtnCreate 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
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
      TabIndex        =   11
      Top             =   5280
      Width           =   7935
   End
End
Attribute VB_Name = "Client_Crud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''----------- REGION ENVENTOS
'INICIO DE FORM CLIENTE
Private Sub Form_Load()

    Dim C_Proc As New C_General_Procedures
   
    Dim Type_doc() As Variant
    
    'ocultar labels de mensajes
    LblhelpName.Visible = False
    LblhelpDoc.Visible = False
    LlbhelpPhone.Visible = False
    LblhelpGeneral.Visible = False
    LblEditCrud.Visible = False
    
    'cargamos consulta datos documentos en BD
    Type_doc = C_Proc.Datos_Charge("TC_Document", "Charge")
    
    'traemos la cantidad de doc en BD
    Q_typedoc = C_Proc.Q_Combo("TC_Document")
    Q_typedoc = Q_typedoc - 1

    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_typedoc
           CbnTypeDocument.AddItem Type_doc(1, I)
    Next
          
    Dim Proccess As String
    Proccess = BtnCreate.Caption
    
End Sub
'BOTON PARA CREAR MODIFICAR O ELIMINAR
Private Sub BtnCreate_Click()

    Dim C_Proc As New C_General_Procedures
    Dim Client() As Variant
    
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(1)
    
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el cliente debe registrar los campos obligatorios se�alados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        Dim Proccess As String
        Proccess = BtnCreate.Caption
        
        If Proccess = "CREAR CLIENTE" Then
        
            'traemos la cantidad de clientes en BD
            Q_Client = C_Proc.Q_Combo("Client")
            Q_Client = Q_Client - 1
                   
            'validamos si la tabla es ta vacia o no
            If Q_Client <> -1 Then
                
                'cargamos consulta datos clientes en BD
                Client = C_Proc.Datos_Charge("Client", "Add_Doc")
                
                'recorremos el arreglo para verifica si esta repetido
                For I = 0 To Q_Client
                    'validamos el campo de cliente o documento
                    If Client(1, I) = TxtDocument.Text Then
                        MsgBox "el Cliente  " & TxtName.Text & "   ya existe en la base de datos!!!", vbInformation + vbOKOnly, "Informaci�n!"
                        LblhelpGeneral.Visible = True
                        LblhelpGeneral.Caption = "el Cliente  " & TxtName.Text & "   ya existe en la base de datos!!!"
                        LblhelpGeneral.ForeColor = &H80&
                        Exit Sub
                    End If
                Next
            
            End If
        
        End If
        
        Select Case Proccess
            Case "CREAR CLIENTE"
                'llamar la funcion de insertar en la BD
                Call Insert
            
            Case "MODIFICAR CLIENTE"
                 'llamar la funcion de modificar en la BD
                 Call Update
                
            Case "ELIMINAR CLIENTE"
                 'llamar la funcion de eliminar en la BD
                 Call Delete
            Case Else
        
        End Select
        
    End If
    
End Sub
'BOTON BUSCAR DATOS
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
        LblhelpGeneral.Caption = "Para poder generar el cliente debe registrar los campos obligatorios se�alados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        ' traer los datos del registro seleccionado
        Traer_Datos = C_Client.Charge_List_Client(op_Search, CbnSearch.Text)
            
        'cargar los datos capturados en los campos
        TxtName.Text = Traer_Datos(0, 0)
        TxtDocument.Text = Traer_Datos(2, 0)
        TxtAddress.Text = Traer_Datos(3, 0)
        TxtPhone.Text = Traer_Datos(4, 0)
        TxtEmail.Text = Traer_Datos(5, 0)
        TxtObservations.Text = Traer_Datos(6, 0)
        
        'seleccionamos el combo de nit con los datos consultados
        For I = 0 To CbnTypeDocument.ListCount - 1
            If Traer_Datos(1, 0) = CbnTypeDocument.List(I) Then
                CbnTypeDocument.Text = CbnTypeDocument.List(I)
            Exit For
            End If
        Next

        Dim Proccess As String
        Proccess = BtnCreate.Caption
        
        If Proccess = "ELIMINAR CLIENTE" Then
            block
        End If
        
    End If
End Sub
'BOTON SALIR
Private Sub BtnExit_Click()
    Unload Client_Crud
End Sub
'OPCION PARA BUSCAR POR DOCUMENTO
Private Sub Opdoc_Click()
    
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
'VALIDAR CAMPO NUMERICO
Private Sub TxtDocument_Change()

    Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    initial = TxtDocument.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtDocument.Text = final

End Sub
''''''----------- END_REGION ENVENTOS

''''''----------- REGION FUNCIONES BD
'LLAMAR METODO CREAR CLIENTE
Function Insert()

    Dim C_Client As New C_CRUD_client
    Dim guardar As String
       
    'llamamos la funcion crear cliente
    guardar = C_Client.Add_Client(UCase(TxtName.Text), CbnTypeDocument.ListIndex, UCase(TxtDocument.Text), UCase(TxtAddress.Text), UCase(TxtPhone.Text), UCase(TxtEmail.Text), UCase(TxtObservations.Text))

    'validamos el resultado de la operacion anterior
    If guardar = "OK" Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Cliente creado con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        clear
        
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If

End Function

'LLAMAR METODO MODIFICAR CLIENTE
Function Update()

    Dim C_Client As New C_CRUD_client
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
    guardar = C_Client.Update_Client(id, UCase(TxtName.Text), CbnTypeDocument.ListIndex, UCase(TxtDocument.Text), UCase(TxtAddress.Text), UCase(TxtPhone.Text), UCase(TxtEmail.Text), UCase(TxtObservations.Text))

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

End Function

'LLAMAR METODO ELIMINAR CLIENTE
Function Delete()

    Dim C_Client As New C_CRUD_client
    Dim C_Proc As New C_General_Procedures
    
    Dim guardar As String
    Dim id As String
    Dim op_Search As String
    
    'averiguar por que metodo buscar el id
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
   
   
    'capturamos el id del registro
    id = C_Proc.Recover_Id(op_Search, "Client", CbnSearch.Text)
    guardar = C_Client.Delete_client(id)
    
    'validamos el resultado de la operacion anterior
    If guardar = "OK" Then
        
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Cliente ha sido eliminado con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        clear
        
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No elimino revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If

    
End Function
''''''----------- END REGION FUNCIONES BD

''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos(verificar As Integer) As Integer
    
    'instanciamos variables
    Dim valName As Integer
    Dim valtypedoc As Integer
    Dim valndoc As Integer
    Dim valphone As Integer
    Dim validate As Integer
    Dim valideedit As Integer
    
    'inicializamos en 0
    valName = 0
    valtypedoc = 0
    valndoc = 0
    valphone = 0
    validate = 0
    valideedit = 0
    
     Dim Proccess As String
       
     Proccess = BtnCreate.Caption
    'validamos combo de seleccion de edicion y eliminacion
        
    If Proccess <> "CREAR CLIENTE" Then
        If CbnSearch.Text = "Seleccione..." Then
             valideedit = 1
        End If
    End If
    
    If valideedit = 1 Then
        validate = 1
        ValidateCampos = validate
        LblEditCrud.Visible = True
        Exit Function
    Else
        LblEditCrud.Visible = False
    End If
   
    If verificar = 1 Then
         
         'validamos campos obligatorios
         If TxtName.Text = "" Then
              valName = 1
         End If
         If CbnTypeDocument.Text = "Seleccione..." Then
              valtypedoc = 1
         End If
         If TxtDocument.Text = "" Then
              valndoc = 1
         End If
         If TxtPhone.Text = "" Then
              valphone = 1
         End If
        
         'verificamos validacion anterior
         If valName = 1 Or valtypedoc = 1 Or valndoc = 1 Or valphone = 1 Then
         
             validate = 1
             
             'validamos para mostrar campos sin diligenciar
             If valName = 1 Then
                 LblhelpName.Visible = True
             Else
                 LblhelpName.Visible = False
             End If
             
             If valtypedoc = 1 Or valndoc = 1 Then
                 LblhelpDoc.Visible = True
             Else
                 LblhelpDoc.Visible = False
             End If
                     
             If valphone = 1 Then
                 LlbhelpPhone.Visible = True
             Else
                 LlbhelpPhone.Visible = False
             End If
         Else
             'ocultar labels de mensajes
             LblhelpName.Visible = False
             LblhelpDoc.Visible = False
             LlbhelpPhone.Visible = False
         End If
    End If
    
    
    
    
    ValidateCampos = validate

End Function

'LIMPIAR CAMPOS
Function clear()

    TxtName.Text = ""
    TxtDocument.Text = ""
    TxtAddress.Text = ""
    TxtPhone.Text = ""
    TxtEmail.Text = ""
    TxtObservations.Text = ""
    CbnSearch.Text = "Seleccione..."
    CbnTypeDocument.Text = "Seleccione..."

End Function

'BLOQUERA CAMPOS
Function block()

    TxtName.Enabled = False
    TxtDocument.Enabled = False
    TxtAddress.Enabled = False
    TxtPhone.Enabled = False
    TxtEmail.Enabled = False
    TxtObservations.Enabled = False
    CbnTypeDocument.Enabled = False

End Function
''''''----------- REGION FUNCIONES

