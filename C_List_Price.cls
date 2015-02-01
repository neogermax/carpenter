VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_List_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function Charge_List_View_sale(IndexInputs As String, IndexMeasure As Integer)

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT Provider.Name, TC_Inputs.Description, TC_Measure.Description, Suppliers_List_Prices.Description, Suppliers_List_Prices.Price, Suppliers_List_Prices.Id_Provider, Suppliers_List_Prices.Id_TC_Inputs, Suppliers_List_Prices.Id_TC_Measure " & _
                " FROM TC_Measure INNER JOIN (TC_Inputs INNER JOIN (Provider INNER JOIN Suppliers_List_Prices ON Provider.Id = Suppliers_List_Prices.Id_Provider) ON TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) ON TC_Measure.Id = Suppliers_List_Prices.Id_TC_Measure " & _
                " WHERE (((Suppliers_List_Prices.Id_TC_Inputs)=" & IndexInputs & ") AND ((Suppliers_List_Prices.Id_TC_Measure)=" & IndexMeasure & ")); "
   
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    Charge_List_View_sale = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    

End Function

Function Q_Charge_List_View_sale(IndexInputs As String, IndexMeasure As Integer)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT count(Suppliers_List_Prices.Id) AS Expr1" & _
              " FROM TC_Measure INNER JOIN (TC_Inputs INNER JOIN (Provider INNER JOIN Suppliers_List_Prices ON Provider.Id = Suppliers_List_Prices.Id_Provider) ON TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) ON TC_Measure.Id  " & _
              " = Suppliers_List_Prices.Id_TC_Measure " & _
              " WHERE (((Suppliers_List_Prices.Id_TC_Inputs)=" & IndexInputs & ") AND ((Suppliers_List_Prices.Id_TC_Measure)=" & IndexMeasure & "));", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_Charge_List_View_sale = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function


'traer los datos  segun el proveedor,insumo y medida
Function Charge_List_View(provider As Integer, IndexInputs As String, IndexMeasure As Integer, TextDescription As String, MODE As String)

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
    
   If MODE = "General" Then
         
        MYSQL = " SELECT Provider.Name, TC_Inputs.Description, TC_Measure.Description, Suppliers_List_Prices.Description, Suppliers_List_Prices.Price, Suppliers_List_Prices.Id_Provider, Suppliers_List_Prices.Id_TC_Inputs, Suppliers_List_Prices.Id_TC_Measure " & _
                " FROM TC_Measure INNER JOIN (TC_Inputs INNER JOIN (Provider INNER JOIN Suppliers_List_Prices ON Provider.Id = Suppliers_List_Prices.Id_Provider) ON TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) ON TC_Measure.Id = Suppliers_List_Prices.Id_TC_Measure " & _
                " WHERE (((Suppliers_List_Prices.Id_Provider)= " & provider & " ) AND ((Suppliers_List_Prices.Id_TC_Inputs)=" & IndexInputs & ") AND ((Suppliers_List_Prices.Id_TC_Measure)= " & IndexMeasure & ")); "
   Else
        MYSQL = " SELECT Provider.Name, TC_Inputs.Description, TC_Measure.Description, Suppliers_List_Prices.Description, Suppliers_List_Prices.Price, Suppliers_List_Prices.Id_Provider, Suppliers_List_Prices.Id_TC_Inputs, Suppliers_List_Prices.Id_TC_Measure, Suppliers_List_Prices.Id " & _
                " FROM TC_Inputs INNER JOIN (TC_Measure INNER JOIN (Provider INNER JOIN Suppliers_List_Prices ON Provider.Id = Suppliers_List_Prices.Id_Provider) ON TC_Measure.Id = Suppliers_List_Prices.Id_TC_Measure) ON (TC_Inputs.Id = TC_Measure.Id_TC_Inputs) AND (TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) " & _
                " WHERE (((Suppliers_List_Prices.Description)='" & TextDescription & "') AND ((Suppliers_List_Prices.Id_Provider)=" & provider & ") AND ((Suppliers_List_Prices.Id_TC_Inputs)=" & IndexInputs & ") AND ((Suppliers_List_Prices.Id_TC_Measure)=" & IndexMeasure & "));"
  
   End If
   
          
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    Charge_List_View = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    Charge_List_View = "Vacio"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    

End Function

Function Q_Charge_List_View(provider As Integer, IndexInputs As String, IndexMeasure As Integer)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open " SELECT Count(Suppliers_List_Prices.Id) AS Expr1" & _
              " FROM TC_Measure INNER JOIN (TC_Inputs INNER JOIN (Provider INNER JOIN Suppliers_List_Prices ON Provider.Id = Suppliers_List_Prices.Id_Provider) ON TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) ON TC_Measure.Id " & _
              " = Suppliers_List_Prices.Id_TC_Measure " & _
              " WHERE (((Suppliers_List_Prices.Id_Provider)= " & provider & ") AND ((Suppliers_List_Prices.Id_TC_Inputs)=" & IndexInputs & ") AND ((Suppliers_List_Prices.Id_TC_Measure)=" & IndexMeasure & "));", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_Charge_List_View = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function
'traer las medidas segun el insumo
Function Measure(IndexInputs As Integer)
    
   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
    
   MYSQL = " SELECT TC_Measure.Id, TC_Measure.Description " & _
           " FROM TC_Measure " & _
           " WHERE(((TC_Measure.Id_TC_Inputs)=" & IndexInputs & "));"
     
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    Measure = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
    

End Function
' para saber la cantidad de items segun la seleccion
Function Q_Measure(IndexInputs As Integer)

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open "select count(TC_Measure.id) from  TC_Measure where (((TC_Measure.Id_TC_Inputs) =" & IndexInputs & "));", MyCon, adOpenDynamic, adLockPessimistic
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_Measure = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select


End Function

Function Add_ListInt(CbnProvider As Integer, CbnImputs As Integer, CbnMeasure As Integer, TxtDescription As String, TxtValues As String, Id_User As Integer)

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
        rstusernames.Open "INSERT INTO Suppliers_List_Prices(Id_TC_Inputs, Id_TC_Measure, Id_Provider , Description, Price, Id_User, CreateDate)VALUES( " & CbnImputs & "," & CbnMeasure & "," & CbnProvider & ",'" & TxtDescription & "','" & TxtValues & "'," & Id_User & ",'" & Date & "')", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Add_ListInt = "OK"
    
    Exit Function

ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo no existe en la base de datos, intentelo de nuevo!", "Error"
    
    Case -2147467259
    MsgBox "El archivo de la base de datos no existe! Revise su ubicación y vuelva a intentarlo.", "Error"
    End
    
    Case Else
    MsgBox "Se ha producido un error desconocido." & vbCrLf & "Código del error: " & Err.Number & vbCrLf & " " & Err.Description, "Error"

    End Select



End Function

Function Update_List(id As String, Description As String, VALUES As String, Id_User As Integer)

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
        rstusernames.Open " update Suppliers_List_Prices set Suppliers_List_Prices.Description ='" & Description & _
                          "', Suppliers_List_Prices.Price ='" & VALUES & _
                          "', Suppliers_List_Prices.Id_User =" & Id_User & _
                          ", Suppliers_List_Prices.CreateDate ='" & Date & _
                          "' where (((Suppliers_List_Prices.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
    dbcon.Close
    
    Update_List = "OK"
    
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

Function Datos_Charge_S()

    Dim NameBD As String
    Dim Datos() As Variant
    Dim StringBD As String
    Dim C_conect As New C_Connections
  
    
    StringBD = C_conect.StringConnect_Access_consult
   
     '** CARGAMOS LOS COMBOS **
    Datos = TRAER_ARRAY_Datos_Charge_S(CStr(StringBD))
    
    Datos_Charge_S = Datos

End Function

Function TRAER_ARRAY_Datos_Charge_S(STRINGDECONECCION As String) As Variant

    Dim MYRECORDSET As ADODB.Recordset
    Dim MYSQL As String
    Dim I%
    Dim Datos() As Variant
    
    '** DIMENSIONAMOS LA BASE DE DATOS **
    
    Set MYRECORDSET = New ADODB.Recordset
    
    MYSQL = " SELECT Suppliers_List_Prices.Id, TC_Inputs.Description, TC_Measure.Description, Provider.Name, Suppliers_List_Prices.Description, Suppliers_List_Prices.Id_TC_Inputs, Suppliers_List_Prices.Id_TC_Measure, Suppliers_List_Prices.Id_Provider " & _
            " FROM ((Suppliers_List_Prices INNER JOIN TC_Inputs ON TC_Inputs.Id = Suppliers_List_Prices.Id_TC_Inputs) " & _
            " INNER JOIN TC_Measure ON TC_Measure.Id = Suppliers_List_Prices.Id_TC_Measure) " & _
            " INNER JOIN Provider ON Provider.id = Suppliers_List_Prices.Id_Provider "
    
    On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, STRINGDECONECCION
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
SALTO:
    TRAER_ARRAY_Datos_Charge_S = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El cliente ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case -2147217871
    GoTo SALTO
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select
              
End Function


