VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1710
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1010.324
   ScaleMode       =   0  'User
   ScaleWidth      =   4098.499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtUserName 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Text            =   "Seleccione..."
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      Picture         =   "frmLogin.frx":26EAE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      Picture         =   "frmLogin.frx":285A1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre de usuario:"
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
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
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
      Height          =   270
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public LoginSucceeded As Boolean

Private Sub Form_Load()
    
    Dim C_Proc As New C_General_Procedures
    Dim Users() As Variant
    
    
     'cargamos consulta datos documentos en BD
    Users = C_Proc.Datos_Charge("Users", "User")
    
    'traemos la cantidad de doc en BD
    Q_User = C_Proc.Q_Combo("Users")
    Q_User = Q_User - 1

    'cargamos el combo con los datos seleccionados
    For I = 0 To Q_User
           txtUserName.AddItem Users(1, I)
    Next
    
End Sub

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
     
    Dim C_User As New C_Users
    Dim Users() As Variant
    Dim validate As Integer
    'validamos campos de diligenciamiento
    validate = ValidateCampos

    If validate = 1 Then
        MsgBox "Algunos de los campos no esta diligeciado. Vuelva a intentarlo", , "Atención!"
        
    Else
        Users = C_User.User_Compary(txtUserName.Text)
        'comprobar si la contraseña es correcta
        If TxtPassword = Users(1, 0) Then
            Load MenuCarpenter
            MenuCarpenter.Show
            MenuCarpenter.Lbl_Value_User.Caption = Users(0, 0)
            MenuCarpenter.Lbl_Value_Roll.Caption = Users(2, 0)
            LoginSucceeded = True
            Me.Hide
        Else
            MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
            TxtPassword.SetFocus
            
        End If
    End If
End Sub

''''''----------- REGION FUNCIONES
'VALIDA CAMPOS OBLIGATORIOS
Function ValidateCampos() As Integer
    
    'instanciamos variables
    Dim valuecampos As Integer
    
    'inicializamos en 0
    valuecampos = 0
    
    'validamos campos obligatorios
    If txtUserName.Text = "Seleccione..." Then
        valuecampos = 1
    End If
    If TxtPassword.Text = "" Then
         valuecampos = 1
    End If
 
    ValidateCampos = valuecampos

End Function
