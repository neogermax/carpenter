VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_OperativeVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'captura detalle del proyecto
Function SearchValuesOperatives()

   Dim StringBD As String
   Dim C_conect As New C_Connections
      
   StringBD = C_conect.StringConnect_Access
     
   Dim MYRECORDSET As ADODB.Recordset
   Dim MYSQL As String
   Dim I%
   Dim Datos() As Variant
    
   Set MYRECORDSET = New ADODB.Recordset
         
        MYSQL = " SELECT TC_OperatingValues.Id, TC_OperatingValues.IVA, TC_OperatingValues.Winner, TC_OperatingValues.Labor" & _
                " FROM TC_OperatingValues " & _
                " WHERE (((TC_OperatingValues.Id)=1)); "
   
     On Error GoTo ctrlerr
    MYRECORDSET.Open MYSQL, StringBD
    I = MYRECORDSET.RecordCount
    Datos = MYRECORDSET.GetRows
    SearchValuesOperatives = Datos
    
    Exit Function
    
ctrlerr:
    
    Select Case Err.Number
    
    Case 3021
    MsgBox "El insumo NO existe!!!", vbExclamation + vbOKOnly, "Información!"
  
    
    Case Else
    MsgBox "Ha ocurrido un error inesperado!" & Chr(13) & "Error " & Err.Number & ": " & Err.Description
    End Select

End Function


