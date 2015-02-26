VERSION 5.00
Begin VB.Form Sale_Easy 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4395
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Sale_Easy.frx":0000
   ScaleHeight     =   4395
   ScaleWidth      =   11535
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
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Sale_Easy.frx":26EAE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox FrmValue 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Sale_Easy.frx":285A1
      ScaleHeight     =   735
      ScaleWidth      =   14775
      TabIndex        =   12
      Top             =   2520
      Width           =   14775
      Begin VB.TextBox TxtValue 
         Height          =   375
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Lbltotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL  --->>>"
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
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Diguite el valor"
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LblValueView 
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
         Height          =   495
         Left            =   8520
         TabIndex        =   14
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.CommandButton BtnCreateClient 
      Caption         =   "CREAR CLIENTE"
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
      Left            =   9600
      Picture         =   "Sale_Easy.frx":4F44F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1815
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
      Left            =   9840
      Picture         =   "Sale_Easy.frx":50B42
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox FrmDescription 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      Picture         =   "Sale_Easy.frx":52235
      ScaleHeight     =   1335
      ScaleWidth      =   14775
      TabIndex        =   6
      Top             =   1200
      Width           =   14775
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
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   8655
      End
      Begin VB.Label Lbltitle_in 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label LblNumber 
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
         Left            =   9360
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.PictureBox FrmClient 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Sale_Easy.frx":790E3
      ScaleHeight     =   735
      ScaleWidth      =   11415
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.OptionButton Opdoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Picture         =   "Sale_Easy.frx":9FF91
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OpName 
         BackColor       =   &H00400000&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1560
         Picture         =   "Sale_Easy.frx":A1684
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CbnSearch 
         Height          =   315
         ItemData        =   "Sale_Easy.frx":A2D77
         Left            =   2880
         List            =   "Sale_Easy.frx":A2D79
         TabIndex        =   2
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   5865
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
         Left            =   9600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sale_Easy.frx":A2D7B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label LbltitleDate 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblValue_Date 
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
      Left            =   9360
      TabIndex        =   20
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label LbltittleInfo 
      BackStyle       =   0  'Transparent
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
      Left            =   3360
      TabIndex        =   18
      Top             =   840
      Width           =   6975
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
      TabIndex        =   17
      Top             =   3360
      Width           =   7815
   End
   Begin VB.Label Lbltittledescrip 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCIÓN DE LA VENTA"
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
      TabIndex        =   5
      Top             =   840
      Width           =   10215
   End
End
Attribute VB_Name = "Sale_Easy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GType_operation As String
Public Crear_client As String



'INICIO DE FORM operativo
Private Sub Form_Load()

    Dim Inputs() As Variant
    Dim C_Proc As New C_General_Procedures
    FrmDescription.Visible = False
    FrmValue.Visible = False
    Crear_client = 0
    LblhelpGeneral.Visible = False
    LblValue_Date.Caption = Date
    
    Count_GOperative = 1
   
End Sub
'FIN DE FORM operativo

'boton crear cliente
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

'boton salir
Private Sub BtnExit_Click()
     Unload Sale_Easy
End Sub

'para volver a cargar el combo de clientes si se creo uno nuevo
Private Sub CbnSearch_GotFocus()
    
    If Crear_client = 1 Then
         If OpName.Value = True Then
            Call OpName_Click
        Else
            Call Opdoc_Click
        End If
    End If
    
End Sub

Private Sub BtnSearch_Click()

    Dim C_Client As New C_CRUD_client
    Dim C_Project As New C_Project
    Dim Traer_Datos() As Variant
    Dim op_Search As String
    
    FrmDescription.Visible = True
    FrmValue.Visible = True
    
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
                
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(1)
      
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar la Venta Debe almenos seleccionar un cliente por favor!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        ' traer los datos del registro seleccionado
        Traer_Datos = C_Client.Charge_List_Client(op_Search, CbnSearch.Text)
            
        'cargar los datos capturados en los campos
        LbltittleInfo.Caption = "CLIENTE: " & Traer_Datos(0, 0)
         
        'capturamos si hay factura o no
        Q_Fact = C_Project.Q_Project(GType_operation)
        
        If Q_Fact = 0 Then
         LblNumber.Caption = 1
        Else
          Q_Fact = Q_Fact + 1
          LblNumber.Caption = Q_Fact
        End If
        
    End If
End Sub

Private Sub BtnCreate_Click()

   Dim validate As Integer
    
    'validamos campos de diligenciamiento
    validate = ValidateCampos(2)
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar la " & GType_operation & " Debe llenar los campos en Rojo!!!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
       G_Venta
      
    End If


End Sub

''''''----------- REGION FUNCIONES BD
'GENERAR FACTURA
Function G_Venta()
    
    Dim op_Search As String
    Dim id As Integer
    Dim Id_User As Integer
    Dim Id_Project As Integer
    
    Dim GUARDAR_P  As String
    Dim GUARDAR_PD  As String
    
    Dim C_Proc As New C_General_Procedures
    Dim C_Project As New C_Project
    Dim val_Iva As String
    
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Doc"
    End If
    
       
    'capturamos el cliente de la operacion
    id = C_Proc.Recover_Id(op_Search, "Client", CbnSearch.Text)
    'capturamos el usuario de la operacion
    Id_User = C_Proc.Recover_Id("User", "Users", MenuCarpenter.Lbl_Value_User.Caption)
    'guardamos la factura
    GUARDAR_P = C_Project.Add_Project(GType_operation, LblNumber.Caption, id, TxtDescripProject.Text, LblValue_Date.Caption, "0", "0", "0", "0", "0", "0", TxtValue.Text, "0", "0", Id_User)
    'capturamos el numeo de proyecto recien creado de la operacion
    Id_Project = C_Project.Recover_IDProject(GType_operation)
    
    'guardamos los detalles de la factura
    GUARDAR_PD = C_Project.Add_ProjectDetail(Id_Project, "N/A", GType_operation, "N/A", "0", "0", "0")
    
    
    'validamos el resultado de la operacion anterior
    If GUARDAR_P = "OK" And GUARDAR_PD = "OK" Then
         
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = GType_operation & " realizada con exito!"
        LblhelpGeneral.ForeColor = &H8000&
        
    Else
        
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If
    
    
End Function


''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos(verificar As Integer) As Integer
    
    'instanciamos variables
    Dim validate As Integer
      
    'inicializamos en 0
    validate = 0
    
    Select Case verificar
    
        Case 1
        
            If CbnSearch.Text = "Seleccione..." Then
                
                MsgBox "no se ha seleccionado Cliente!!!", vbExclamation + vbOKOnly, "Información!"
                validate = 1
              
            End If
          
        Case 2
        
            If CbnSearch.Text = "Seleccione..." Or TxtDescripProject = "" Or TxtValue.Text = "" Then
               
               If CbnSearch.Text = "Seleccione..." Then
                   CbnSearch.BackColor = &H80&
               Else
                   CbnSearch.BackColor = &H80000005
               End If
               If TxtDescripProject.Text = "" Then
                   TxtDescripProject.BackColor = &H80&
               Else
                   TxtDescripProject.BackColor = &H80000005
               End If
                If TxtValue.Text = "" Then
                   TxtValue.BackColor = &H80&
               Else
                   TxtValue.BackColor = &H80000005
               End If
        
               validate = 1
            
            End If
            
        Case Else
        
    End Select
    
      ValidateCampos = validate
          

End Function

Private Sub TxtValue_Change()

 Dim C_Proc As New C_General_Procedures
    
    Dim initial As String
    Dim final As String
    
    initial = TxtValue.Text
    final = C_Proc.Validate_Numeric(initial)
    TxtValue.Text = final
    
    LblValueView.Caption = Format(final, "####,####")
    
End Sub
