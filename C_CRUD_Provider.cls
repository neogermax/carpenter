VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_CRUD_Provider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'crear proveedor en la BD
Function Add_Provider(Name As String, IdDocument As String, Documents As String, StreetAddress As String, Phone As String, Email As String, Observations As String, Id_User As Integer)
    
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
        rstusernames.Open "INSERT INTO PROVIDER (Name , IdDocument , Documents , StreetAddress , Phone , Email , Observations,Id_User,CreateDate ) VALUES( '" & Name & "'," & IdDocument & ",'" & Documents & "','" & StreetAddress & "','" & Phone & "','" & Email & "','" & Observations & "'," & Id_User & ",'" & Date & "')", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Add_Provider = "OK"
    
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

'actualizar proveedor en la BD
Function Update_Provider(id As Integer, Name As String, IdDocument As String, Documents As String, StreetAddress As String, Phone As String, Email As String, Observations As String, Id_User As Integer)
    
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
        rstusernames.Open " update Provider set Provider.Name ='" & Name & _
                          "', Provider.IdDocument ='" & IdDocument & _
                          "', Provider.Documents ='" & Documents & _
                          "', Provider.StreetAddress = '" & StreetAddress & _
                          "', Provider.Phone ='" & Phone & _
                          "', Provider.Email ='" & Email & _
                          "', Provider.Observations ='" & Observations & _
                          "', Provider.CreateDate='" & Date & _
                          "', Provider.Id_user ='" & Id_User & _
                          "' where (((Provider.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_Provider = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El proveedor no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicación y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "Código del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

End Select

End Function

'eliminar cliente en la BD
Function Delete_Provider(id As String)

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
        rstusernames.Open "delete from Provider  WHERE (((Provider.[Id])= " & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Delete_Provider = "OK"
    
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

'metodo de coneccion para consultar datos
Function Charge_List_Provider(op_Search As String, MODE As String)

    Dim C_conect As New C_Connections
    Dim NameBD As String
    Dim Datos() As Variant
    Dim StringBD As String
    Dim BDNAME As String
  
    'traemos el tipo de coneccion a la BD
    StringBD = C_conect.StringConnect_Access_consult
    
     'metodo de busqueda de datos
    Datos = TRAER_ARRAY_List_Provider(op_Search, MODE, CStr(StringBD))
    
    Charge_List_Provider = Datos

End Function

Function TRAER_ARRAY_List_Provider(op_Search As String, DateSearch As String, StringBD As String) As Variant

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
          MYSQL = " SELECT Provider.Name, TC_Document.Description, Provider.Documents, Provider.StreetAddress, Provider.Phone, Provider.Email, Provider.Observations, Provider.IdDocument " & _
                  " FROM TC_Document INNER JOIN Provider ON TC_Document.Id = Provider.IdDocument " & _
                  " WHERE (((Provider.Name)= '" & DateSearch & "')); "
    Else
    
          'ejecutamos instruccion
          MYSQL = " SELECT Provider.Name, TC_Document.Description, Provider.Documents, Provider.StreetAddress, Provider.Phone, Provider.Email, Provider.Observations, Provider.IdDocument " & _
                  " FROM TC_Document INNER JOIN Provider ON TC_Document.Id = Provider.IdDocument " & _
                  " WHERE (((Provider.Documents)= '" & DateSearch & "'));"
    End If
        
    'abrimos connecion y capturamos los datos
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    List_datos = MYRECORDSET.GetRows
    
    'retornamos la consulta realizada
    TRAER_ARRAY_List_Provider = List_datos
   
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "La consulta realizada NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function






