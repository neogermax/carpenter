VERSION 5.00
Begin VB.MDIForm MenuCarpenter 
   BackColor       =   &H00FFFFFF&
   Caption         =   "POS Carpinteria 2015"
   ClientHeight    =   4710
   ClientLeft      =   225
   ClientTop       =   540
   ClientWidth     =   12480
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Status_Bar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "MenuCarpenter.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   12480
      TabIndex        =   1
      Top             =   4215
      Width           =   12480
      Begin VB.Timer Tm_In 
         Interval        =   1000
         Left            =   17160
         Top             =   120
      End
      Begin VB.Label LblHours 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   17760
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Lbl_Value_Roll 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LblRoll 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Lbl_Value_User 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label LblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox PicStretch 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "MenuCarpenter.frx":259F6
      ScaleHeight     =   495
      ScaleWidth      =   12480
      TabIndex        =   0
      Top             =   0
      Width           =   12480
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   12480
      TabIndex        =   3
      Top             =   0
      Width           =   12480
   End
   Begin VB.Menu Client 
      Caption         =   "&Clientes"
      WindowList      =   -1  'True
      Begin VB.Menu ClientCreate 
         Caption         =   "&Crear"
      End
      Begin VB.Menu ClientChange 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu ClientDelete 
         Caption         =   "&Eliminar"
      End
   End
   Begin VB.Menu provider 
      Caption         =   "&Proveedores"
      Begin VB.Menu CreateProvider 
         Caption         =   "&Crear"
      End
      Begin VB.Menu UpdateProvider 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu DeleteProvider 
         Caption         =   "&Eliminar"
      End
   End
   Begin VB.Menu ListPrices 
      Caption         =   "&Listado de Materiales"
      Begin VB.Menu InputsCreate 
         Caption         =   "&Crear Insumo"
      End
      Begin VB.Menu InputsUpdate 
         Caption         =   "&Editar Insumo"
      End
   End
   Begin VB.Menu Module_sale 
      Caption         =   "&Modulo de Venta"
      Begin VB.Menu quotation 
         Caption         =   "&Cotización"
      End
      Begin VB.Menu SaleEasy 
         Caption         =   "&Venta Minima"
      End
      Begin VB.Menu invoice 
         Caption         =   "&Facturación"
      End
      Begin VB.Menu EndProyect 
         Caption         =   "&Pagos ó Abonos"
      End
   End
   Begin VB.Menu AdminUsers 
      Caption         =   "&Administración de Usuarios"
      Begin VB.Menu CreateUsers 
         Caption         =   "Crear"
      End
      Begin VB.Menu UpdateUsers 
         Caption         =   "&Modificar"
      End
   End
   Begin VB.Menu Querys 
      Caption         =   "&Consultas"
      Begin VB.Menu ViewClients 
         Caption         =   "&Ver Clientes"
      End
   End
   Begin VB.Menu OperationalProcesses 
      Caption         =   "&Procesos Operacionales"
      Begin VB.Menu TurnoverSecuritiesAdministrators 
         Caption         =   "&Administración de Valores Facturación"
      End
      Begin VB.Menu AdminDiscount 
         Caption         =   "&Administración de Descuentos en Facturación"
      End
   End
End
Attribute VB_Name = "MenuCarpenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMAGESIZE = 0.566893424036281 'BY MAXeXTREME


Private Sub MDIForm_Load()

   On Error GoTo ctrlerr
   Dim imagen As String
    
   Dim C_Proc As New C_General_Procedures
   Dim Config() As Variant
   
   Config = C_Proc.Config
   imagen = Config(0, 0)
   
   MenuCarpenter.WindowState = 2
   PicStretch.Picture = LoadPicture(imagen)
    Exit Sub
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Sub

'tamaño de la imagen
Private Sub MDIForm_Resize()

    On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    PicStretch.Visible = False
    PicStretch.AutoRedraw = True
    PicStretch.Height = Me.ScaleHeight
    ImageWidth = PicStretch.Picture.Width * IMAGESIZE
    ImageHeight = PicStretch.Picture.Height * IMAGESIZE
    PicStretch.PaintPicture PicStretch.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, ImageWidth, ImageHeight
    Set Me.Picture = PicStretch.Image
    PicStretch.Visible = False

End Sub



'hora en linea
Private Sub Tm_In_Timer()
    LblHours.Caption = Time
End Sub

'---------------------------------- modulo clientes ------------------------------------------------
Private Sub ClientCreate_Click()
    
    Load Client_Crud
    Client_Crud.Left = (MenuCarpenter.ScaleWidth - Client_Crud.Width) / 2
    Client_Crud.Top = (MenuCarpenter.ScaleHeight - Client_Crud.Height) / 2
    Client_Crud.Caption = "Nuevo Cliente"
    Client_Crud.FrmClient.Visible = False
    Client_Crud.BtnCreate.Caption = "CREAR CLIENTE"
    Client_Crud.Show

End Sub

Private Sub ClientChange_Click()

    Load Client_Crud
    Client_Crud.Left = (MenuCarpenter.ScaleWidth - Client_Crud.Width) / 2
    Client_Crud.Top = (MenuCarpenter.ScaleHeight - Client_Crud.Height) / 2
    Client_Crud.Caption = "Modificar Datos del Cliente"
    Client_Crud.FrmClient.Visible = True
    Client_Crud.BtnCreate.Caption = "MODIFICAR CLIENTE"
    Client_Crud.Opdoc.Value = True
    Client_Crud.FrmBody.Visible = False
    Client_Crud.Show

End Sub

Private Sub ClientDelete_Click()

    Load Client_Crud
    Client_Crud.Left = (MenuCarpenter.ScaleWidth - Client_Crud.Width) / 2
    Client_Crud.Top = (MenuCarpenter.ScaleHeight - Client_Crud.Height) / 2
    Client_Crud.Caption = "Eliminar Datos del Cliente"
    Client_Crud.FrmClient.Visible = True
    Client_Crud.BtnCreate.Caption = "ELIMINAR CLIENTE"
    Client_Crud.Opdoc.Value = True
    Client_Crud.FrmBody.Visible = False
    Client_Crud.Show
    
End Sub

'---------------------------------- modulo proveedores ------------------------------------------------
Private Sub CreateProvider_Click()

    Load Provider_Crud
    Provider_Crud.Left = (MenuCarpenter.ScaleWidth - Provider_Crud.Width) / 2
    Provider_Crud.Top = (MenuCarpenter.ScaleHeight - Provider_Crud.Height) / 2
    Provider_Crud.Caption = "Nuevo Proveedor"
    Provider_Crud.FrmClient.Visible = False
    Provider_Crud.BtnCreate.Caption = "CREAR PROVEEDOR"
    Provider_Crud.Show

End Sub

Private Sub UpdateProvider_Click()

    Load Provider_Crud
    Provider_Crud.Left = (MenuCarpenter.ScaleWidth - Provider_Crud.Width) / 2
    Provider_Crud.Top = (MenuCarpenter.ScaleHeight - Provider_Crud.Height) / 2
    Provider_Crud.Caption = "Modificar Datos del Proveedor"
    Provider_Crud.FrmClient.Visible = True
    Provider_Crud.BtnCreate.Caption = "MODIFICAR PROVEEDOR"
    Provider_Crud.Opdoc.Value = True
    Provider_Crud.FrmBody.Visible = False
    Provider_Crud.Show

End Sub

Private Sub DeleteProvider_Click()

    Load Provider_Crud
    Provider_Crud.Left = (MenuCarpenter.ScaleWidth - Provider_Crud.Width) / 2
    Provider_Crud.Top = (MenuCarpenter.ScaleHeight - Provider_Crud.Height) / 2
    Provider_Crud.Caption = "Modificar Datos del Proveedor"
    Provider_Crud.FrmClient.Visible = True
    Provider_Crud.BtnCreate.Caption = "ELIMINAR PROVEEDOR"
    Provider_Crud.Opdoc.Value = True
    Provider_Crud.FrmBody.Visible = False
    Provider_Crud.Show
End Sub

'---------------------------------- modulo listado insumos ------------------------------------------------
Private Sub InputsCreate_Click()

    Load List_Price
    List_Price.Left = (MenuCarpenter.ScaleWidth - List_Price.Width) / 2
    List_Price.Top = (MenuCarpenter.ScaleHeight - List_Price.Height) / 2
    List_Price.Caption = "Nuevo Insumo"
    List_Price.BtnCreate.Caption = "CREAR INSUMO"
    List_Price.BtnSearch.Visible = False
    List_Price.Show

End Sub

Private Sub InputsUpdate_Click()

    Load List_PriceEdit
    List_PriceEdit.Left = (MenuCarpenter.ScaleWidth - List_PriceEdit.Width) / 2
    List_PriceEdit.Top = (MenuCarpenter.ScaleHeight - List_PriceEdit.Height) / 2
    List_PriceEdit.Caption = "Nuevo Insumo"
    List_PriceEdit.BtnCreate.Caption = "MODIFICAR INSUMO"
    List_PriceEdit.BtnSearch.Visible = True
    List_PriceEdit.Show
    
End Sub

'---------------------------------- modulo venta------------------------------------------------
Private Sub invoice_Click()

    Load InvoiceAndQuotation
    InvoiceAndQuotation.Left = (MenuCarpenter.ScaleWidth - InvoiceAndQuotation.Width) / 2
    InvoiceAndQuotation.Top = (MenuCarpenter.ScaleHeight - InvoiceAndQuotation.Height) / 2
    InvoiceAndQuotation.Caption = "Facturación el ebanista"
    InvoiceAndQuotation.LblTittleCapture.Caption = "ESCOJA MATERIALES PARA FACTURACIÓN"
    InvoiceAndQuotation.Lbltitle_in.Caption = "FACTURA N°"
    InvoiceAndQuotation.Opdoc.Value = True
    InvoiceAndQuotation.OpNot.Value = True
    InvoiceAndQuotation.Show
    InvoiceAndQuotation.GType_operation = "Factura"
    InvoiceAndQuotation.BtnCreate.Caption = "FACTURAR"
End Sub

Private Sub quotation_Click()

    Load InvoiceAndQuotation
    InvoiceAndQuotation.Left = (MenuCarpenter.ScaleWidth - InvoiceAndQuotation.Width) / 2
    InvoiceAndQuotation.Top = (MenuCarpenter.ScaleHeight - InvoiceAndQuotation.Height) / 2
    InvoiceAndQuotation.Caption = "Cotización el ebanista"
    InvoiceAndQuotation.LblTittleCapture.Caption = "ESCOJA MATERIALES PARA COTIZACIÓN"
    InvoiceAndQuotation.Lbltitle_in.Caption = "COTIZACIÓN N°"
    InvoiceAndQuotation.Opdoc.Value = True
    InvoiceAndQuotation.OpNot.Value = True
    InvoiceAndQuotation.frmPagos.Visible = False
    InvoiceAndQuotation.Show
    InvoiceAndQuotation.GType_operation = "Cotización"
    InvoiceAndQuotation.BtnCreate.Caption = "COTIZAR"
    
End Sub

Private Sub SaleEasy_Click()

    Load Sale_Easy
    Sale_Easy.Left = (MenuCarpenter.ScaleWidth - Sale_Easy.Width) / 2
    Sale_Easy.Top = (MenuCarpenter.ScaleHeight - Sale_Easy.Height) / 2
    Sale_Easy.Caption = "Venta y chicharrones el ebanista"
    Sale_Easy.Lbltitle_in.Caption = "VENTA N°"
    Sale_Easy.Opdoc.Value = True
    Sale_Easy.Show
    Sale_Easy.GType_operation = "Venta"
    Sale_Easy.BtnCreate.Caption = "VENTA"

End Sub

Private Sub EndProyect_Click()
    
    Load PaymentsCredits
    PaymentsCredits.Left = (MenuCarpenter.ScaleWidth - PaymentsCredits.Width) / 2
    PaymentsCredits.Top = (MenuCarpenter.ScaleHeight - PaymentsCredits.Height) / 2
    PaymentsCredits.Caption = "Pagos y abonos el ebanista"
    PaymentsCredits.Show
   
End Sub

'---------------------------------- modulo administracion de usuarios------------------------------------------------

Private Sub CreateUsers_Click()

    Load Admin_Users
    Admin_Users.Left = (MenuCarpenter.ScaleWidth - Admin_Users.Width) / 2
    Admin_Users.Top = (MenuCarpenter.ScaleHeight - Admin_Users.Height) / 2
    Admin_Users.Caption = "Nuevo Usuario"
    Admin_Users.FrmUser.Visible = False
    Admin_Users.BtnCreate.Caption = "CREAR USUARIO"
    Admin_Users.Show
    
End Sub

Private Sub UpdateUsers_Click()
    
    Load Admin_Users
    Admin_Users.Left = (MenuCarpenter.ScaleWidth - Admin_Users.Width) / 2
    Admin_Users.Top = (MenuCarpenter.ScaleHeight - Admin_Users.Height) / 2
    Admin_Users.Caption = "Modificar Usuario"
    Admin_Users.OpName.Value = True
    Admin_Users.FrmUser.Visible = True
    Admin_Users.BtnCreate.Caption = "MODIFICAR USUARIO"
    Admin_Users.Show
    
End Sub

'---------------------------------- modulo administracion operaciones------------------------------------------------

Private Sub TurnoverSecuritiesAdministrators_Click()
 
    Load Admin_ValuesFact
    Admin_ValuesFact.Left = (MenuCarpenter.ScaleWidth - Admin_ValuesFact.Width) / 2
    Admin_ValuesFact.Top = (MenuCarpenter.ScaleHeight - Admin_ValuesFact.Height) / 2
    Admin_ValuesFact.Caption = "Administrador de valores operativos en facturación"
    Admin_ValuesFact.BtnCreate.Caption = "ACTUALIZAR"
    Admin_ValuesFact.Show

End Sub
Private Sub AdminDiscount_Click()

    Load Admin_ValuesDiscounts
    Admin_ValuesDiscounts.Left = (MenuCarpenter.ScaleWidth - Admin_ValuesDiscounts.Width) / 2
    Admin_ValuesDiscounts.Top = (MenuCarpenter.ScaleHeight - Admin_ValuesDiscounts.Height) / 2
    Admin_ValuesDiscounts.Caption = "Administrador de descuentos en facturación"
    Admin_ValuesDiscounts.BtnCreate.Caption = "ACTUALIZAR"
    Admin_ValuesDiscounts.Show

End Sub



'---------------------------------- modulo consultas------------------------------------------------

Private Sub ViewClients_Click()

    Load ViewClient
    ViewClient.Left = (MenuCarpenter.ScaleWidth - ViewClient.Width) / 2
    ViewClient.Top = (MenuCarpenter.ScaleHeight - ViewClient.Height) / 2
    ViewClient.Caption = "Nuestro Clientes"
    ViewClient.Show

End Sub
