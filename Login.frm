VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000005&
   Caption         =   "Login"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton buttonRegister 
      Caption         =   "Register"
      Height          =   1000
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   2000
   End
   Begin VB.TextBox textPassword 
      Height          =   500
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox textUsername 
      Height          =   500
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   1000
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   2000
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Usuario"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buttonRegister_Click()
    Me.Hide
    
    Load Register
    Register.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
    ' Capturar los datos desde los TextBox
    Dim userName As String
    Dim password As String
    
    userName = textUsername.Text
    password = textPassword.Text
    
    ' Llamar a la función de validación
    If validacionUsuario(userName, password) Then
        ' Cerrar el formulario de inicio de sesión y abrir el formulario Home
        Unload Login
        Home.Show
    Else
        ' Mostrar mensaje de error
        MsgBox "Nombre de usuario o contraseña incorrectos.", vbExclamation
    End If
    
    ' Limpiar los TextBox
    textUsername.Text = ""
    textPassword.Text = ""
End Sub
