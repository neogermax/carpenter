VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'traer todos los datos del usuario
Function Charge_List_User(op_Search As String, DateSearch As String) As Variant

    Dim C_conect As New C_Connections
    Dim StringBD As String
        'traemos el tipo de coneccion a la BD
    StringBD = C_conect.StringConnect_Access_consult
        
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
          MYSQL = " SELECT Users.Name, Users.Documents, Users.Nick, Users.Phone, Users.Password, TC_Rolls.Description, Users.Id_TC_Rolls " & _
                  " FROM TC_Rolls INNER JOIN Users ON TC_Rolls.Id = Users.Id_TC_Rolls " & _
                  " WHERE (((Users.Name)= '" & DateSearch & "')); "
    Else
    
          'ejecutamos instruccion
          MYSQL = " SELECT Users.Name, Users.Documents, Users.Nick, Users.Phone, Users.Password, TC_Rolls.Description, Users.Id_TC_Rolls " & _
                  " FROM TC_Rolls INNER JOIN Users ON TC_Rolls.Id = Users.Id_TC_Rolls " & _
                  " WHERE (((Users.Nick)= '" & DateSearch & "'));"
    End If
        
    'abrimos connecion y capturamos los datos
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    List_datos = MYRECORDSET.GetRows
    
    'retornamos la consulta realizada
    Charge_List_User = List_datos
   
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "La consulta realizada NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function


'traer las datos de usuario para el ingreso
Function User_Compary(Nick As String)
    
   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
    
   MYSQL = " SELECT Users.Nick, Users.Password, TC_Rolls.Alias  " & _
           " FROM TC_Rolls INNER JOIN Users ON TC_Rolls.Id = Users.Id_TC_Rolls " & _
           " WHERE (((Users.Nick)='" & Nick & "'));"
     
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    User_Compary = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El Usuario NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    

End Function

'crear usuario en la BD
Function Add_User(Id_Rolls As Integer, Name As String, Documents As String, Nick As String, Phone As String, Password As String)
    
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
        rstusernames.Open "INSERT INTO Users(Id_TC_Rolls, Name, Documents, Nick, Phone, [Password], CreateDate)VALUES(" & Id_Rolls & ",'" & Name & "','" & Documents & "','" & Nick & "','" & Phone & "','" & Password & "','" & Date & "')", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Add_User = "OK"
    
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

'actualizar cliente en la BD
Function Update_User(id As Integer, Id_Rolls As Integer, Name As String, Documents As String, Nick As String, Phone As String, Password As String)
    
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
        rstusernames.Open " update Users set Users.Id_TC_Rolls =" & Id_Rolls & _
                          ", Users.Name ='" & Name & _
                          "', Users.Documents ='" & Documents & _
                          "', Users.Nick = '" & Nick & _
                          "', Users.Phone ='" & Phone & _
                          "', Users.[Password] ='" & Password & _
                          "', Users.CreateDate ='" & Date & _
                          "' where (((Users.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_User = "OK"
    
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

