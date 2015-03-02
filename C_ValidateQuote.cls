VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_ValidateQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'funcion que  nos trae la cantidad de proyecto por cliente
Function Q_SearchQuote()

    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    
    Myrs.Open " SELECT count(Project.Id) " & _
              " FROM Project " & _
              " WHERE (((Project.CampoType_Operator)='Cotización'));  ", MyCon, adOpenDynamic, adLockPessimistic
        
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Q_SearchQuote = TOTAL
    
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
Function SearchQuote()

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
   MYSQL = " SELECT Project.id, Project.CreateDate " & _
           " FROM Project " & _
           " WHERE (((Project.CampoType_Operator)='Cotización')); "
    
            
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    SearchQuote = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function

Function UpdateStateQuote()

   Dim Q_Quote As Integer
   Dim ListQuote() As Variant
   Dim DateInitial As Date
   Dim id As Integer
   
    Q_Quote = Q_SearchQuote
    
    If Q_Quote = 0 Then
        Exit Function
    End If
    
    Q_Quote = Q_Quote - 1
    ListQuote = SearchQuote
    
    For I = 1 To Q_Quote
        id = ListQuote(0, I)
        DateInitial = CDate(ListQuote(1, I))
        DateFinal = DateAdd("D", 15, DateInitial)
        
        If Date > DateFinal Then
            Update_Project (id)
        End If
    Next

End Function


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
        rstusernames.Open " update Project set Project.State_Project ='Expiro' where (((Project.id)=" & id & "));", STRINGCONECT, adOpenStatic, adLockOptimistic
    
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

