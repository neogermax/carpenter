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

Function Add_ListInt(CbnProvider As Integer, CbnImputs As Integer, CbnMeasure As Integer, TxtDescription As String, TxtValues As String)

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
        rstusernames.Open "INSERT INTO Suppliers_List_Prices(Id_TC_Inputs, Id_TC_Measure, Id_Provider , Description, Price)VALUES( " & CbnImputs & "," & CbnMeasure & "," & CbnProvider & ",'" & TxtDescription & "','" & TxtValues & "')", STRINGCONECT, adOpenStatic, adLockOptimistic
    
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

