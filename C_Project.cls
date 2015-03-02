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
'---------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------  detalle de proyecto ----------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------

'actualizar proyecto en la BD
Function Update_Project(id As Integer)
   
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
        rstusernames.Open " update Project set Project.State_Project ='Cancelado' where (((Project.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_Project = "OK"
    
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

'actualizar detalles de proyecto en la BD
Function Update_DatailProject(id As Integer, NumberProviderFact As String, Vr_FactProvider As String, Winner_Loser As String)
    
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
        rstusernames.Open " update Project_Detail set Project_Detail.NumberProviderFact ='" & NumberProviderFact & "', Project_Detail.Vr_FactProvider ='" & Vr_FactProvider & "', Project_Detail.Winner_Loser ='" & Winner_Loser & "' where (((Project_Detail.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_DatailProject = "OK"
    
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

'funcion que  nos trae la cantidad de facturas pendientes
Function Q_SearchDetailsProject(Id_Project As Integer)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT Count(Project_Detail.id) AS Expr1 " & _
              " FROM Project_Detail  " & _
              " WHERE (((Project_Detail.Id_Project)= " & Id_Project & ")); ", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_SearchDetailsProject = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function

Function SearchDetailsProject_ID(Id_Project As Integer)


   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Project_Detail.id " & _
                " FROM Project_Detail " & _
                " WHERE (((Project_Detail.Id_Project)= " & Id_Project & ")); "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    SearchDetailsProject_ID = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
End Function

Function SearchDetailsProject(Id_Project As Integer)

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Project_Detail.Description, Project_Detail.Measure, Project_Detail.Quantity, Project_Detail.Vr_Unit, Project_Detail.Vr_Total " & _
                " FROM Project_Detail " & _
                " WHERE (((Project_Detail.Id_Project)= " & Id_Project & ")); "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    SearchDetailsProject = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------- fin detalle de proyecto -------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- proyecto -------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
'funcion que  nos trae la cantidad de proyecto por cliente
Function Q_SearchProjectClient(Id_Client As String, TypeOperative As String)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    
    If TypeOperative = "All" Then
        
         Myrs.Open " SELECT count(Project.Id) " & _
                   " FROM Project " & _
                   " WHERE (((Project.Id_Client)=" & Id_Client & "));  ", MyCon, adOpenDynamic, adLockPessimistic
    
    Else
            
         Myrs.Open " SELECT count(Project.Id) " & _
                   " FROM Project " & _
                   " WHERE (((Project.Id_Client)=" & Id_Client & ")AND ((Project.CampoType_Operator)='" & TypeOperative & "'));  ", MyCon, adOpenDynamic, adLockPessimistic
        
    End If
        
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_SearchProjectClient = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function

'captura de proyectos por cliente
Function SearchProjectClient(Id_Client As String, TypeOperative As String)

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        If TypeOperative = "All" Then
        
             MYSQL = " SELECT Project.Description, Project.CreateDate, Project.VR_Total, Project.State_Project " & _
                     " FROM Project " & _
                     " WHERE (((Project.Id_Client)=" & Id_Client & "));  "
        
        Else
            
             MYSQL = " SELECT Project.Description, Project.CreateDate, Project.VR_Total, Project.State_Project " & _
                     " FROM Project " & _
                     " WHERE (((Project.Id_Client)=" & Id_Client & ")AND ((Project.CampoType_Operator)='" & TypeOperative & "'));  "
        
        End If
         
         
       
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    SearchProjectClient = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function


'captura detalle del proyecto
Function DateProjectDetail(Id_Number As String)

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Project.Description, Project.VR_Total, Project.VR_Saldo " & _
                " FROM Client INNER JOIN Project ON Client.Id = Project.Id_Client " & _
                " WHERE (((Project.Id_Operator)= " & Id_Number & ") AND ((Project.CampoType_Operator)='Factura')) " & _
                " order by Project.Id_Operator asc; "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    DateProjectDetail = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function

'funcion que  nos trae la cantidad de facturas pendientes
Function Q_Charge_GProject()

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT count(Project.Id)" & _
              " FROM Client INNER JOIN Project ON Client.Id = Project.Id_Client  " & _
              " WHERE (((Project.CampoType_Operator)='Factura') AND ((Project.State_Project)='inicial'));", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_Charge_GProject = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function

'funcion que trae los proyectos pendientes por pago
Function Charge_GProject()

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Project.Id_Operator, Client.Name, Project.VR_Total " & _
                " FROM Client INNER JOIN Project ON Client.Id = Project.Id_Client " & _
                " WHERE (((Project.CampoType_Operator)='Factura') AND ((Project.State_Project)='inicial')) " & _
                " order by Project.Id_Operator asc; "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    Charge_GProject = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    
End Function

'funcion de captura el id para el detalle del proyecto
Function Recover_IDProject(Type_operator As String)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT Max(Project.Id) AS Expr1" & _
              " FROM Project " & _
              " WHERE (((Project.CampoType_Operator)='" & Type_operator & "'));", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Recover_IDProject = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function



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
    Dim State As String
    
    
    Select Case CampoType_Operator
    
        Case "Factura"
            State = "inicial"
        Case "Cotización"
            State = "valida"
        Case "Venta"
            State = "Cancelado"
        Case Else
            State = ""
    End Select
    
    
    'traemos el tipo de coneccion a la BD
    STRINGCONECT = C_conect.StringConnect_Access
    
    'abrimos la coneccion a la BD
    Set dbcon = New ADODB.Connection
        dbcon.ConnectionString = STRINGCONECT
        dbcon.Open
    
    'ejecutamos la sentecia en la BD
    Set rstusernames = New ADODB.Recordset
        rstusernames.CursorLocation = adUseClient
        rstusernames.Open "INSERT INTO project (CampoType_Operator, State_Project, Id_Operator, Id_Client, Description, CreateDate, TimeOfManufacture , VR_Inputs, VR_Labor, VR_Winner, VR_Subtotal, VR_Iva, VR_Total, VR_Abono, VR_Saldo, Id_User) VALUES( '" & _
                            CampoType_Operator & _
                            "','" & State & _
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

'crear los detalles del Proyecto en la BD
Function Add_ProjectDetail(Id_Project As Integer, Measure As String, CampoType_Operator As String, Description As String, Quantity As String, VR_Unit As String, VR_Total As String)
    
    On Error GoTo ctrlerr
    
    Dim C_conect As New C_Connections
    Dim STRINGCONECT As String
    Dim State As String
    
    'traemos el tipo de coneccion a la BD
    STRINGCONECT = C_conect.StringConnect_Access
    
    'abrimos la coneccion a la BD
    Set dbcon = New ADODB.Connection
        dbcon.ConnectionString = STRINGCONECT
        dbcon.Open
    
    'ejecutamos la sentecia en la BD
    Set rstusernames = New ADODB.Recordset
        rstusernames.CursorLocation = adUseClient
        rstusernames.Open "INSERT INTO Project_Detail (Id_Project, CampoType, Description, Measure, Quantity, VR_Unit, VR_Total ) VALUES( '" & _
                            Id_Project & _
                            "','" & CampoType_Operator & _
                            "','" & Description & _
                            "','" & Measure & _
                            "','" & Quantity & _
                            "','" & VR_Unit & _
                            "','" & VR_Total & "')", STRINGCONECT, adOpenStatic, adLockOptimistic

    dbcon.Close
    
    Add_ProjectDetail = "OK"
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


