VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'consultar numeracion de la factura o cotizacion
Function Q_Project(Type_operator As String)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT COUNT(Project.Id_Operator) AS Expr1" & _
              " FROM Project " & _
              " WHERE (((Project.CampoType_Operator)='" & Type_operator & "'));", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_Project = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function

'crear Proyecto en la BD
Function Add_Project(CampoType_Operator As String, Id_Operator As String, Id_Client As Integer, Description As String, CreateDate As String, TimeOfManufacture As String, VR_Inputs As String, VR_Labor As String, VR_Winner As String, VR_Subtotal As String, VR_Iva As String, VR_Total As String, VR_Abono As String, VR_Saldo As String, Id_User As Integer)
    
    On Error GoTo ctrlerr
    
    Dim C_conect As New C_Connections
    Dim STRINGCONECT As String
    
    'traemos el tipo de coneccion a la BD
    STRINGCONECT = C_conect.StringConnect_Access
    
    'abrimos la coneccion a la BD
    Set dbcon = New ADODB.Connection
        dbcon.ConnectionString = STRINGCONECT
        dbcon.Open
    
    'ejecutamos la sentecia en la BD
    Set rstusernames = New ADODB.Recordset
        rstusernames.CursorLocation = adUseClient
        rstusernames.Open "INSERT INTO project (CampoType_Operator, Id_Operator, Id_Client, Description, CreateDate, TimeOfManufacture , VR_Inputs, VR_Labor, VR_Winner, VR_Subtotal, VR_Iva, VR_Total, VR_Abono, VR_Saldo,Id_User) VALUES( '" & _
                            CampoType_Operator & _
                            "'," & Id_Operator & _
                            "," & Id_Client & _
                            ",'" & Description & _
                            "','" & CreateDate & _
                            "'," & TimeOfManufacture & _
                            "," & VR_Inputs & _
                            "," & VR_Labor & _
                            "," & VR_Winner & _
                            "," & VR_Subtotal & _
                            "," & VR_Iva & _
                            "," & VR_Total & _
                            "," & VR_Abono & _
                            "," & VR_Saldo & _
                            "," & Id_User & ")", STRINGCONECT, adOpenStatic, adLockOptimistic

    dbcon.Close
    
    Add_Project = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El usuario no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicación y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "Código del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

    End Select

End Function

