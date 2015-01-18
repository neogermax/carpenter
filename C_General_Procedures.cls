VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_General_Procedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'funcion para comprobar el formato de email
Function Validate_Emails(Validate_S_Email As String) As Integer

    Dim DataSearch_Arroba As String
    Dim DataSearch_Exten
    DataSearch_Exten = Array(".gov", ".com", ".net", ".org", ".biz", ".info", ".name", ".us", ".tv", ".in", ".cc")
    
    Dim TestArroba As Integer
    Dim TestExt As Integer
    Dim Validate_arroba As Integer
    Dim Validate_ext As Integer
    Dim validate_general As Integer
    
    Validate_arroba = 0
    Validate_ext = 0
    validate_general = 0
    
    DataSearch_Arroba = "@"
    
    TestArroba = InStr(1, Validate_S_Email, DataSearch_Arroba, vbTextCompare)
    
    If TestArroba = 0 Then
        Validate_arroba = 1
    End If

    
    If Validate_arroba = 0 Then
    
        For I = 0 To 9
            TestExt = InStr(1, Validate_S_Email, DataSearch_Exten(I), vbTextCompare)
            
            If TestExt = 0 Then
               Validate_ext = 1
            Else
               Validate_ext = 0
               Exit For
            End If
        Next
    
    End If

    If Validate_arroba = 1 Or Validate_ext = 1 Then
        validate_general = 1
    End If
    
    Validate_Emails = validate_general

End Function
'funcion para saber la cantidad de elementos de una tabla
Function Q_Combo(TABLA As String) As Integer
    
    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
   
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    Myrs.Open "select count(id) from " & TABLA, MyCon, adOpenDynamic, adLockPessimistic
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

SALTO:
    Q_Combo = TOTAL
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El recibo ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case -2147217871
    GoTo SALTO
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.description
    End Select

End Function

Function Datos_Charge(TABLA As String, MODE As String)

    Dim NameBD As String
    Dim Datos() As Variant
    Dim StringBD As String
    Dim C_conect As New C_Connections
  
    
    StringBD = C_conect.StringConnect_Access_consult
   
     '** CARGAMOS LOS COMBOS **
    Datos = TRAER_ARRAY_Datos_Charge(TABLA, MODE, CStr(StringBD))
    
    Datos_Charge = Datos

End Function

Function TRAER_ARRAY_Datos_Charge(TABLA As String, MODE As String, STRINGDECONECCION As String) As Variant

    Dim MYRECORDSET As ADODB.Recordset
    Dim MYSQL As String
    Dim I%
    Dim Datos() As Variant
    
    '** DIMENSIONAMOS LA BASE DE DATOS **
    
    Set MYRECORDSET = New ADODB.Recordset
    
    Select Case MODE
    
        Case "Charge"
             MYSQL = "SELECT " & TABLA & ".Id, " & TABLA & ".Description FROM " & TABLA
        
        Case "Add_Doc"
             MYSQL = "SELECT " & TABLA & ".Id, " & TABLA & ".Documents FROM " & TABLA
        
        Case "Add_Name"
             MYSQL = "SELECT " & TABLA & ".Id, " & TABLA & ".Name From " & TABLA
             
        Case "Add_Descrip"
             MYSQL = "SELECT " & TABLA & ".Id, " & TABLA & ".Description From " & TABLA
             
        Case Else
        
    End Select
    
    
    On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, STRINGDECONECCION
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
SALTO:
    TRAER_ARRAY_Datos_Charge = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El cliente ingresado NO existe!!!", vbExclamation + vbOKOnly, "Información!"
    
    Case -2147217871
    GoTo SALTO
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.description
    End Select
              
End Function


Function Recover_Id(op_Search As String, TABLA As String, DateSearch As String) As Integer
    
    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
  
     
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    
    If op_Search = "Name" Then
        Myrs.Open "select " & TABLA & ".id from " & TABLA & " where  (((" & TABLA & ".Name)= '" & DateSearch & "')); ", MyCon, adOpenDynamic, adLockPessimistic
    Else
        Myrs.Open "select " & TABLA & ".id from " & TABLA & " where  (((" & TABLA & ".Documents)= '" & DateSearch & "')); ", MyCon, adOpenDynamic, adLockPessimistic
    End If
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Recover_Id = TOTAL
  
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "No se encontro id seleccionado!!!", vbExclamation + vbOKOnly, "Información!"
    
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.description
    End Select

End Function

Function Recover_Id_Detail(TABLA As String, DateSearch As String, IdInput As String) As Integer
    
    Dim MyCon As New ADODB.Connection
    Dim Myrs As New ADODB.Recordset
    Dim StringBD As String
    Dim TOTAL As Integer
    Dim C_conect As New C_Connections
       
    TOTAL = 0
    StringBD = C_conect.StringConnect_Access_consult
   
    On Error GoTo ctrlerr
    MyCon.Open StringBD
    
        Myrs.Open "select " & TABLA & ".id from " & TABLA & " where  (((" & TABLA & ".Description)= '" & DateSearch & "') AND ((" & TABLA & ".Id_TC_Inputs)= " & IdInput & ")); ", MyCon, adOpenDynamic, adLockPessimistic
    
    TOTAL = Myrs.Fields(0)
    
    Myrs.Close
    MyCon.Close

    Recover_Id_Detail = TOTAL
  
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "No se encontro id seleccionado!!!", vbExclamation + vbOKOnly, "Información!"
    
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.description
    End Select

End Function

'VALIDAR CAMPO NUMERICO
Function Validate_Numeric(text_value As String)
    
    Dim Q_lent As Integer
    
    'capturamos el tamaño de la cadena
    Q_lent = Len(text_value)
    
    'validamos si no tiene nada
    If Q_lent <> 0 Then
        'validamos si tiene letras
        If Not IsNumeric(text_value) Then
                
                MsgBox "Solo acepta numero  ", vbExclamation + vbOKOnly, "Error!"
                
                Dim Value_Correct As String
                
                Q_lent = Q_lent - 1
                'borramos el ultimo caracter de la cadena
                Value_Correct = Left$(text_value, Q_lent)
                Validate_Numeric = Value_Correct
        Else
            Validate_Numeric = text_value
        End If
    Else
         Validate_Numeric = text_value
    End If

End Function

