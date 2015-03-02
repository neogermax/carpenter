VERSION 5.00
Begin VB.Form Admin_ValuesDiscounts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Admin_ValuesDiscounts.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   8040
   Begin VB.TextBox TxtValDes4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox TxtValDes1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtValDes2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TxtValDes3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
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
      Left            =   5880
      Picture         =   "Admin_ValuesDiscounts.frx":26EAE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1935
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
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin_ValuesDiscounts.frx":285A1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento de 10.000.001 en adelante"
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
      TabIndex        =   14
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label LblValIVA 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento de 0 a 2.000.000 millones"
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
      TabIndex        =   11
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label LblValLabor 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento de 2.000.001 a 5.000.000 millones"
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
      TabIndex        =   10
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label LblValWinner 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento de 5.000.001 a 10.000.000 millones"
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
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
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
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5400
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "Admin_ValuesDiscounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'carga de formulario
Private Sub Form_Load()

    Dim C_OVal As New C_OperativeVal
    Dim listOpe() As Variant
    
    Dim ValDes1 As String
    Dim ValDes2 As String
    Dim ValDes3 As String
    Dim ValDes4 As String
      
    listOpe = C_OVal.SearchValuesDiscounts
      
    ValDes1 = listOpe(0, 0)
    ValDes2 = listOpe(1, 0)
    ValDes3 = listOpe(2, 0)
    ValDes4 = listOpe(3, 0)
    
    TxtValDes1.Text = ValDes1
    TxtValDes2.Text = ValDes2
    TxtValDes3.Text = ValDes3
    TxtValDes4.Text = ValDes4
    
End Sub

'actualizar valores dela bd
Private Sub BtnCreate_Click()
  
    Dim C_OVal As New C_OperativeVal
    Dim listOpe() As Variant
    Dim GUARDAR As String
    
    If TxtValDes1.Text = "" Or TxtValDes2.Text = "" Or TxtValDes3.Text = "" Or TxtValDes4.Text = "" Then
                
        If TxtValDes1.Text = "" Then
            TxtValDes1.BackColor = &H40&
        Else
            TxtValDes1.BackColor = &HFFFFFF
        End If
        
        If TxtValDes2.Text = "" Then
            TxtValDes2.BackColor = &H40&
        Else
            TxtValDes2.BackColor = &HFFFFFF
        End If
        
        If TxtValDes3.Text = "" Then
            TxtValDes3.BackColor = &H40&
        Else
            TxtValDes3.BackColor = &HFFFFFF
        End If
        
         If TxtValDes4.Text = "" Then
            TxtValDes4.BackColor = &H40&
        Else
            TxtValDes4.BackColor = &HFFFFFF
        End If
        
    Else
        GUARDAR = C_OVal.Update_ValuesDiscounts(TxtValDes1.Text, TxtValDes2.Text, TxtValDes3.Text, TxtValDes4.Text)
        
         If GUARDAR = "OK" Then
        
            LblhelpGeneral.Caption = "Valores de facturacion actualizados con exito!"
            LblhelpGeneral.ForeColor = &H8000&
            
        Else
        
            LblhelpGeneral.Caption = "No actualizo revisar insercion a la BD!"
            LblhelpGeneral.ForeColor = &H80&
        
        End If
    End If
  
End Sub
'salir de formulario
Private Sub BtnExit_Click()
    Unload Admin_ValuesDiscounts
End Sub

'advertencia descuento 1
Private Sub TxtValDes1_GotFocus()
    TxtValDes1.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es el porcentaje descuento entre 0 a 2.000.000 en facturación,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

'advertencia  descuento 2
Private Sub TxtValDes2_GotFocus()
    TxtValDes2.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es el porcentaje descuento entre 2.000.001 a 5.000.000 en facturación,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

'advertencia  descuento 3
Private Sub TxtValDes3_GotFocus()
    TxtValDes3.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es el porcentaje descuento entre 5.000.001 a 10.000.000 en facturación,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

'advertencia  descuento 4
Private Sub TxtValDes4_GotFocus()
    TxtValDes4.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es el porcentaje descuento entre 10.000.001 en adelante en facturación,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub
