VERSION 5.00
Begin VB.Form Admin_Users 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Admin_Users.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   12705
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
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin_Users.frx":26EAE
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
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
      Left            =   11160
      Picture         =   "Admin_Users.frx":285A1
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5400
      Width           =   1455
   End
   Begin VB.PictureBox FrmBody 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      Picture         =   "Admin_Users.frx":29C94
      ScaleHeight     =   3855
      ScaleWidth      =   12615
      TabIndex        =   6
      Top             =   1080
      Width           =   12615
      Begin VB.TextBox TxtNick 
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtName 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   7095
      End
      Begin VB.ComboBox CbnRoll 
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Text            =   "Seleccione..."
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox TxtDocument 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox TxtPhone 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox TxtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox TxtPassword2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label LblHelpNick 
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
         Left            =   5640
         TabIndex        =   30
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label LblNick 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LblHelpPassword2 
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
         Left            =   5640
         TabIndex        =   24
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label LblHelpRoll 
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
         Left            =   4680
         TabIndex        =   23
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label LblHelpPassword 
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
         Left            =   5640
         TabIndex        =   22
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label LblPassword2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label LlbhelpPhone 
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
         Left            =   5640
         TabIndex        =   20
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label LblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblRoll 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label LblPhone 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
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
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label LblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label LblDocumentNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Documento"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label LblhelpName 
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
         Left            =   9480
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label LblhelpDoc 
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
         Left            =   5640
         TabIndex        =   13
         Top             =   2040
         Width           =   2415
      End
   End
   Begin VB.PictureBox FrmUser 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      Picture         =   "Admin_Users.frx":50B42
      ScaleHeight     =   975
      ScaleWidth      =   12615
      TabIndex        =   0
      Top             =   0
      Width           =   12615
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
         Left            =   8880
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Admin_Users.frx":779F0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CbnSearch 
         Height          =   315
         ItemData        =   "Admin_Users.frx":790E3
         Left            =   2760
         List            =   "Admin_Users.frx":790E5
         TabIndex        =   3
         Text            =   "Seleccione..."
         Top             =   240
         Width           =   5865
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
         Left            =   1440
         Picture         =   "Admin_Users.frx":790E7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OpNick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alias"
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
         Left            =   120
         Picture         =   "Admin_Users.frx":7A7DA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblEditCrud 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Debe seleccionar una opción"
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
         Height          =   615
         Left            =   10320
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
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
      TabIndex        =   27
      Top             =   5280
      Width           =   8415
   End
End
Attribute VB_Name = "Admin_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public G_IdRolls As Integer

Private Sub Form_Load()
  
    Dim C_Proc As New C_General_Procedures
    Dim Type_rolls() As Variant
    
    'ocultar labels de mensajes
    LblhelpName.Visible = False
    LblhelpDoc.Visible = False
    LlbhelpPhone.Visible = False
    LblhelpGeneral.Visible = False
    LblEditCrud.Visible = False
    LblHelpRoll.Visible = False
    LblHelpPassword.Visible = False
    LblHelpPassword2.Visible = False
    LblHelpNick.Visible = False
    CbnSearch.Visible = True
        
    'cargamos consulta datos documentos en BD
    Type_rolls = C_Proc.Datos_Charge("TC_Rolls", "Charge")
    
    'traemos la cantidad de doc en BD
    Q_rolls = C_Proc.Q_Combo("TC_Rolls")
    Q_rolls = Q_rolls - 1

    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_rolls
           CbnRoll.AddItem Type_rolls(1, I)
    Next
          
End Sub

'BOTON BUSCAR DATOS
Private Sub BtnSearch_Click()
    
    Dim C_User As New C_Users
    Dim Traer_Datos() As Variant
    Dim op_Search As String
    
    'revisamos la opcion de busqueda
    If OpName.Value = True Then
        op_Search = "Name"
    Else
        op_Search = "Nick"
    End If
      
    FrmBody.Visible = True
          
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(2)
      
    'comprobamos validacion
    If validate = 1 Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el usuario debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    
    Else
        
        ' traer los datos del registro seleccionado
        Traer_Datos = C_User.Charge_List_User(op_Search, CbnSearch.Text)
            
        'cargar los datos capturados en los campos
        TxtName.Text = Traer_Datos(0, 0)
        TxtDocument.Text = Traer_Datos(1, 0)
        TxtNick.Text = Traer_Datos(2, 0)
        TxtPhone.Text = Traer_Datos(3, 0)
        TxtPassword.Text = Traer_Datos(4, 0)
        TxtPassword2.Text = Traer_Datos(4, 0)
        G_IdRolls = Traer_Datos(6, 0)
        
        'seleccionamos el combo de nit con los datos consultados
        For I = 0 To CbnRoll.ListCount - 1
            If Traer_Datos(5, 0) = CbnRoll.List(I) Then
                CbnRoll.Text = CbnRoll.List(I)
            Exit For
            End If
        Next
     End If
End Sub

'BOTON SALIR
Private Sub BtnExit_Click()
    Unload Admin_Users
End Sub

Private Sub BtnCreate_Click()

    Dim C_Proc As New C_General_Procedures
    Dim Users() As Variant
    
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos(1)
    
    'comprobamos validacion
    If validate = 1 Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Para poder generar el Usuario debe registrar los campos obligatorios señalados en la parte superior!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        
        Dim Proccess As String
        Proccess = BtnCreate.Caption
        
        If Proccess = "CREAR USUARIO" Then
            'traemos la cantidad de usuarios en BD
            Q_User = C_Proc.Q_Combo("Users")
            Q_User = Q_User - 1
                   
            'validamos si la tabla esta vacia o no
            If Q_User <> -1 Then
                'cargamos consulta datos clientes en BD
                Users = C_Proc.Datos_Charge("Users", "Add_Doc")
                
                'recorremos el arreglo para verifica si esta repetido
                For I = 0 To Q_Client
                    'validamos el campo de cliente o documento
                    If Users(1, I) = TxtDocument.Text Then
                        MsgBox "El usuario  " & TxtName.Text & " ya existe en la base de datos!!!", vbInformation + vbOKOnly, "Información!"
                        LblhelpGeneral.Visible = True
                        LblhelpGeneral.Caption = "El usuario  " & TxtName.Text & "   ya existe en la base de datos!!!"
                        LblhelpGeneral.ForeColor = &H80&
                        Exit Sub
                    End If
                Next
            End If
        End If
                
        Select Case Proccess
            Case "CREAR USUARIO"
                'llamar la funcion de insertar en la BD
                Call Insert
            
            Case "MODIFICAR USUARIO"
                 'llamar la funcion de modificar en la BD
                 Call Update
           
        End Select
    
    End If
    
End Sub

'en edicion guardar el index
Private Sub CbnRoll_lostFocus()
    G_IdRolls = CbnRoll.ListIndex
End Sub

'OPCION PARA BUSCAR POR DOCUMENTO
Private Sub OpNick_Click()
    
    Dim C_Proc As New C_General_Procedures
    Dim cargar_datos() As Variant
    
    CbnSearch.clear
    CbnSearch.Text = "Seleccione..."
    CbnSearch.Width = 2500
    
    'cargamos consulta datos documentos en BD
    cargar_datos = C_Proc.Datos_Charge("Users", "User")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Users")
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
    cargar_datos = C_Proc.Datos_Charge("Users", "Add_Name")
    
    'traemos la cantidad de doc en BD
    Q_cargar_datos = C_Proc.Q_Combo("Users")
    Q_cargar_datos = Q_cargar_datos - 1
    
    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_cargar_datos
        CbnSearch.AddItem cargar_datos(1, I)
    Next

End Sub

Private Sub TxtPassword2_LostFocus()

    If TxtPassword.Text <> TxtPassword2.Text Then
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "La contraseña no coinside!!"
        LblhelpGeneral.ForeColor = &H80&
    Else
        LblhelpGeneral.Visible = False
    End If

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

Function ValidateCampos(validar As Integer)

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
    
    Select Case validar
    
        Case 1
           
            If TxtName.Text = "" Then
                LblhelpName.Visible = True
                validate = 1
            Else
                LblhelpName.Visible = False
                validate = 0
            End If
            
            If TxtNick.Text = "" Then
                LblHelpNick.Visible = True
                validate = 1
            Else
                LblHelpNick.Visible = False
                validate = 0
            End If
            
            If CbnRoll.Text = "Seleccione..." Then
                LblHelpRoll.Visible = True
                validate = 1
            Else
                LblHelpRoll.Visible = False
                validate = 0
            End If
            
            If TxtDocument.Text = "" Then
                LblhelpDoc.Visible = True
                validate = 1
            Else
                LblhelpDoc.Visible = False
                validate = 0
            End If
            
            If TxtPhone.Text = "" Then
                LlbhelpPhone.Visible = True
                validate = 1
            Else
                LlbhelpPhone.Visible = False
                validate = 0
            End If
            
            If TxtPassword.Text = "" Then
                LblHelpPassword.Visible = True
                validate = 1
            Else
                LblHelpPassword.Visible = False
                validate = 0
            End If
            
            If TxtPassword2.Text = "" Then
                LblHelpPassword2.Visible = True
                validate = 1
            Else
                LblHelpPassword2.Visible = False
                validate = 0
            End If
        
        Case 2
            If CbnSearch.Text = "Seleccione..." Then
                LblEditCrud.Visible = True
                validate = 1
            Else
                LblEditCrud.Visible = False
                validate = 0
            End If

        Case Else
    
    End Select
    
    ValidateCampos = validate

End Function

''''''----------- REGION FUNCIONES BD
'LLAMAR METODO CREAR CLIENTE
Function Insert()

    Dim C_User As New C_Users
    Dim GUARDAR As String
  
    Dim C_Proc As New C_General_Procedures
       
    'llamamos la funcion crear cliente
    GUARDAR = C_User.Add_User(CbnRoll.ListIndex, UCase(TxtName.Text), UCase(TxtDocument.Text), UCase(TxtNick.Text), UCase(TxtPhone.Text), TxtPassword2.Text)

    'validamos el resultado de la operacion anterior
    If GUARDAR = "OK" Then
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "Usuario creado con exito!"
        LblhelpGeneral.ForeColor = &H8000&
    
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If

End Function

Function Update()

    Dim C_Proc As New C_General_Procedures
    Dim C_User As New C_Users
    
    Dim op_Search As String
    Dim id As Integer
    Dim GUARDAR As String
    Dim TextVarible As String
    
    'averiguar por que metodo buscar el id
    If OpName.Value = True Then
        op_Search = "Name"
        TextVarible = CbnSearch.Text
    Else
        op_Search = "Doc"
        TextVarible = TxtDocument.Text
    End If
      
   
    'capturamos el id del registro
    id = C_Proc.Recover_Id(op_Search, "Users", TextVarible)
    'llamamos la funcion crear cliente
    GUARDAR = C_User.Update_User(id, G_IdRolls, UCase(TxtName.Text), UCase(TxtDocument.Text), UCase(TxtNick.Text), UCase(TxtPhone.Text), TxtPassword2.Text)
    
    'validamos el resultado de la operacion anterior
    If GUARDAR = "OK" Then
        
         LblhelpGeneral.Visible = True
         LblhelpGeneral.Caption = "Usuario modificado con exito!"
         LblhelpGeneral.ForeColor = &H8000&
    
    Else
    
        LblhelpGeneral.Visible = True
        LblhelpGeneral.Caption = "No guardo revisar insercion a la BD!"
        LblhelpGeneral.ForeColor = &H80&
    
    End If
End Function
