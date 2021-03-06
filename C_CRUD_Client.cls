VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_CRUD_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'crear cliente en la BD
Function Add_Client(Name As String, IdDocument As String, Documents As String, StreetAddress As String, Phone As String, Email As String, Observations As String, Id_User As Integer)
    
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
        rstusernames.Open "INSERT INTO CLIENT (Name , IdDocument , Documents , StreetAddress , Phone , Email , Observations, Id_user, CreateDate) VALUES( '" & Name & "'," & IdDocument & ",'" & Documents & "','" & StreetAddress & "','" & Phone & "','" & Email & "','" & Observations & "'," & Id_User & ",'" & Date & "')", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Add_Client = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El usuario no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicaci�n y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "C�digo del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

    End Select

End Function

'actualizar cliente en la BD
Function Update_Client(id As Integer, Name As String, IdDocument As String, Documents As String, StreetAddress As String, Phone As String, Email As String, Observations As String, Id_User As Integer)
    
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
        rstusernames.Open " update client set client.Name ='" & Name & _
                          "', client.IdDocument ='" & IdDocument & _
                          "', client.Documents ='" & Documents & _
                          "', client.StreetAddress = '" & StreetAddress & _
                          "', client.Phone ='" & Phone & _
                          "', client.Email ='" & Email & _
                          "', client.Observations ='" & Observations & _
                          "', client.Id_user =" & Id_User & _
                          ", client.CreateDate ='" & Date & _
                          "' where (((client.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_Client = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El usuario no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicaci�n y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "C�digo del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

End Select

End Function

'eliminar cliente en la BD
Function Delete_client(id As String)

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
        rstusernames.Open "delete from Client  WHERE (((Client.[Id])= " & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Delete_client = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El usuario no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicaci�n y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "C�digo del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

End Select
  

End Function

'metodo de coneccion para consultar datos
Function Charge_List_Client(op_Search As String, MODE As String)

    Dim C_conect As New C_Connections
    Dim NameBD As String
    Dim Datos() As Variant
    Dim StringBD As String
    Dim BDNAME As String
  
    'traemos el tipo de coneccion a la BD
    StringBD = C_conect.StringConnect_Access_consult
    
     'metodo de busqueda de datos
    Datos = TRAER_ARRAY_List_Client(op_Search, MODE, CStr(StringBD))
    
    Charge_List_Client = Datos

End Function

Function TRAER_ARRAY_List_Client(op_Search As String, DateSearch As String, StringBD As String) As Variant

    Dim MYRECORDSET As ADODB.Recordset
    Dim MYSQL As String
    Dim I%
    Dim List_datos() As Variant
        
    'creamos recorset
    Set MYRECORDSET = New ADODB.Recordset
          
    On Error GoTo ctrlerr
    
    'validamos el tipo de busqueda
    If op_Search = "Name" Then
    
          'ejecutamos instruccion
          MYSQL = " SELECT Client.Name, TC_Document.Description, Client.Documents, Client.StreetAddress, Client.Phone, Client.Email, Client.Observations, Client.IdDocument " & _
                  " FROM TC_Document INNER JOIN Client ON TC_Document.Id = Client.IdDocument " & _
                  " WHERE (((Client.Name)= '" & DateSearch & "')); "
    Else
    
          'ejecutamos instruccion
          MYSQL = " SELECT Client.Name, TC_Document.Description, Client.Documents, Client.StreetAddress, Client.Phone, Client.Email, Client.Observations, Client.IdDocument " & _
                  " FROM TC_Document INNER JOIN Client ON TC_Document.Id = Client.IdDocument " & _
                  " WHERE (((Client.Documents)= '" & DateSearch & "'));"
    End If
        
    'abrimos connecion y capturamos los datos
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    List_datos = MYRECORDSET.GetRows
    
    'retornamos la consulta realizada
    TRAER_ARRAY_List_Client = List_datos
   
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "La consulta realizada NO existe!!!", vbExclamation + vbOKOnly, "Informaci�n!"
    
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function

'funcion que trae los proyectos pendientes por pago
Function ChargeClient()

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Client.Name, TC_Document.Description, Client.Documents, Client.StreetAddress, Client.Phone, Client.Email, Client.Observations, Client.CreateDate " & _
                " FROM TC_Document INNER JOIN Client ON TC_Document.Id = Client.IdDocument " & _
                " ORDER BY Client.Name; "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    ChargeClient = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Informaci�n!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    
End Function

'funcion que  nos trae la cantidad de facturas pendientes
Function Q_ChargeClient()

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT count(Client.id) " & _
              " FROM TC_Document INNER JOIN Client ON TC_Document.Id = Client.IdDocument;", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_ChargeClient = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "La Base de Datos notiene clientes en este momento!!!", vbExclamation + vbOKOnly, "Informaci�n!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function


