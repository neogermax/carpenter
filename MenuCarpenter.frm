VERSION 5.00
Begin VB.MDIForm MenuCarpenter 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000007&
   Caption         =   "POS Carpinteria 2015"
   ClientHeight    =   2355
   ClientLeft      =   225
   ClientTop       =   540
   ClientWidth     =   9570
   LinkTopic       =   "MDIForm1"
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
End
Attribute VB_Name = "MenuCarpenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()

    On Error GoTo ctrlerr
    
    MenuCarpenter.WindowState = 2
   'PicStretch.Picture = LoadPicture(imagen)
    Exit Sub
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Sub

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
    Client_Crud.FrmClient.Caption = "Escoja La opcion:"
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
    Client_Crud.FrmClient.Caption = "Escoja La opcion:"
    Client_Crud.BtnCreate.Caption = "ELIMINAR CLIENTE"
    Client_Crud.Opdoc.Value = True
    Client_Crud.FrmBody.Visible = False
    Client_Crud.Show
    
End Sub

'---------------------------------- modulo clientes ------------------------------------------------
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
    Provider_Crud.FrmClient.Caption = "Escoja La opcion:"
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
    Provider_Crud.FrmClient.Caption = "Escoja La opcion:"
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

    Load List_Price
    List_Price.Left = (MenuCarpenter.ScaleWidth - List_Price.Width) / 2
    List_Price.Top = (MenuCarpenter.ScaleHeight - List_Price.Height) / 2
    List_Price.Caption = "Nuevo Insumo"
    List_Price.BtnCreate.Caption = "MODIFICAR INSUMO"
    List_Price.BtnSearch.Visible = True
    List_Price.Show
    
End Sub

