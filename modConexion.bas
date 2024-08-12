Attribute VB_Name = "modConexion"
' modConexion.bas

Option Explicit

' Declarar una variable global para la conexi�n
Public conn As ADODB.Connection

' Funci�n para abrir la conexi�n a la base de datos
Public Sub AbrirConexion()
    ' Verificar si la conexi�n ya est� abierta
    If conn Is Nothing Then
        Set conn = New ADODB.Connection
    End If
    
    ' Si la conexi�n ya est� abierta, no hacer nada
    If conn.State = adStateOpen Then Exit Sub
    
    ' Configurar la cadena de conexi�n (ajusta seg�n tu base de datos)
    
    
    Dim sConnString As String

    
    'Conexion BD Aqui colocar las contrase�as de acceso
    
    sConnString = "Provider=SQLOLEDB;Data Source=;Initial Catalog=;User ID=;Password=;"
    
    
    ' Abrir la conexi�n
    conn.ConnectionString = sConnString
    conn.Open
End Sub

' Funci�n para cerrar la conexi�n a la base de datos
Public Sub CerrarConexion()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

