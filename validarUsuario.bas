Attribute VB_Name = "validarUsuario"
Option Explicit
Public Function validacionUsuario(userName As String, password As String) As Boolean
    On Error GoTo ManejarError
    
    ' Abrir la conexi�n
    AbrirConexion
    
    ' SQL para validar el usuario y la contrase�a
    Dim sql As String
    sql = "SELECT COUNT(*) AS Conteo " & _
          "FROM Usuarios " & _
          "WHERE NombreDeUsuario = '" & userName & "' " & _
          "AND Password = '" & password & "'"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Validar si el usuario y la contrase�a coinciden
    If rs.Fields("Conteo").Value > 0 Then
        validacionUsuario = True
    Else
        validacionUsuario = False
    End If
    
    rs.Close
    Exit Function

ManejarError:
    MsgBox "Ocurri� un error: " & Err.Description, vbCritical
    validacionUsuario = False
End Function


