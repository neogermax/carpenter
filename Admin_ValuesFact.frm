VERSION 5.00
Begin VB.Form Admin_ValuesFact 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Admin_ValuesFact.frx":0000
   ScaleHeight     =   3180
   ScaleWidth      =   6975
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
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin_ValuesFact.frx":26EAE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1935
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
      Left            =   4920
      Picture         =   "Admin_ValuesFact.frx":285A1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox TxtValWinner 
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtValLabor 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TxtValIVA 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1575
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
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
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
      Left            =   4440
      TabIndex        =   10
      Top             =   840
      Width           =   255
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
      Left            =   4440
      TabIndex        =   9
      Top             =   360
      Width           =   255
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
      Top             =   1800
      Width           =   6735
   End
   Begin VB.Label LblValWinner 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor de la Ganancia"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LblValLabor 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor de la Mano de Obra"
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
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label LblValIVA 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor del IVA"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Admin_ValuesFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'carga de formulario
Private Sub Form_Load()

    Dim C_OVal As New C_OperativeVal
    Dim listOpe() As Variant
    
    Dim OperateIva As String
    Dim ValLabor As String
    Dim ValWinner As String
      
    listOpe = C_OVal.SearchValuesOperatives
      
    OperateIva = listOpe(1, 0)
    ValLabor = listOpe(2, 0)
    ValWinner = listOpe(3, 0)
    
    TxtValIVA.Text = OperateIva
    TxtValLabor.Text = ValLabor
    TxtValWinner.Text = ValWinner
    
End Sub

'actualizar valores dela bd
Private Sub BtnCreate_Click()
  
    Dim C_OVal As New C_OperativeVal
    Dim listOpe() As Variant
    Dim GUARDAR As String
    
    If TxtValIVA.Text = "" Or TxtValLabor.Text = "" Or TxtValWinner.Text = "" Then
                
        If TxtValIVA.Text = "" Then
            TxtValIVA.BackColor = &H40&
        Else
            TxtValIVA.BackColor = &HFFFFFF
        End If
        
        If TxtValLabor.Text = "" Then
            TxtValLabor.BackColor = &H40&
        Else
            TxtValLabor.BackColor = &HFFFFFF
        End If
        
        If TxtValWinner.Text = "" Then
            TxtValWinner.BackColor = &H40&
        Else
            TxtValWinner.BackColor = &HFFFFFF
        End If
        
    Else
        GUARDAR = C_OVal.Update_AmdValues(TxtValIVA.Text, TxtValLabor.Text, TxtValWinner.Text)
        
         If GUARDAR = "OK" Then
        
            LblhelpGeneral.Caption = "Valores de facturacion actualizados con exito!"
            LblhelpGeneral.ForeColor = &H8000&
            
        Else
        
            LblhelpGeneral.Caption = "No actualizo revisar insercion a la BD!"
            LblhelpGeneral.ForeColor = &H80&
        
        End If
    End If
  
End Sub

'salir del formulario
Private Sub BtnExit_Click()
    Unload Admin_ValuesFact
End Sub

'advertencia iva
Private Sub TxtValIVA_GotFocus()
    TxtValIVA.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es el porcentaje para el iva en la facturación,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

'advertencia obra de mano
Private Sub TxtValLabor_GotFocus()
    TxtValLabor.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es para multiplicar por los materiales para obtener la mano de obra,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

'advertencia ganacia
Private Sub TxtValWinner_GotFocus()
    TxtValWinner.BackColor = &HFFFFFF
    LblhelpGeneral.Caption = " El valor registrado es para multiplicar por los materiales para obtener la ganacia,(Si ingresa valores decimales registrarlos con (punto).)"
    LblhelpGeneral.ForeColor = &H800000
End Sub

