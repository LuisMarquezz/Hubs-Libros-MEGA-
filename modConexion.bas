Attribute VB_Name = "modConexion"
' modConexion.bas

Option Explicit

' Declarar una variable global para la conexión
Public conn As ADODB.Connection

' Función para abrir la conexión a la base de datos
Public Sub AbrirConexion()
    ' Verificar si la conexión ya está abierta
    If conn Is Nothing Then
        Set conn = New ADODB.Connection
    End If
    
    ' Si la conexión ya está abierta, no hacer nada
    If conn.State = adStateOpen Then Exit Sub
    
    ' Configurar la cadena de conexión (ajusta según tu base de datos)
    
    
    Dim sConnString As String

    
    'Conexion BD Aqui colocar las contraseñas de acceso
    
    sConnString = "Provider=SQLOLEDB;Data Source=;Initial Catalog=;User ID=;Password=;"
    
    
    ' Abrir la conexión
    conn.ConnectionString = sConnString
    conn.Open
End Sub

' Función para cerrar la conexión a la base de datos
Public Sub CerrarConexion()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

