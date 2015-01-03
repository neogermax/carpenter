VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Connections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function StringConnect() As String
    
    Dim StringBD As String
    StringBD = " driver=sql server; database=Carpinteria; server=GERMAX-PC; User ID=sa; Password=pikachu; "
    StringConnect = StringBD
    
End Function
Function CADENACONECTION()

    Dim StringBD As String
    StringBD = App.Path & "\CARPINTERIA_2003.mdb"

    CADENACONECTION = StringBD
     
End Function
 

   
Function StringConnect_Access_consult() As String
    
    Dim StringBD As String
    StringBD = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & App.Path & "\CARPINTERIA_2003.mdb"
    StringConnect_Access_consult = StringBD
    
End Function

Function StringConnect_Access() As String
    
    Dim StringBD As String
    
    Database_Path = App.Path & "\CARPINTERIA_2003.mdb"
    Set dbcon = New ADODB.Connection
    dbcon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database_Path & _
                             ";Persist Security Info=False;Jet OLEDB:Database Password="
    StringConnect_Access = dbcon
    
End Function


