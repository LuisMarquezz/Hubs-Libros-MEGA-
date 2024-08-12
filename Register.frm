VERSION 5.00
Begin VB.Form Register 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton buttonCancel 
      Caption         =   "Cancel"
      Height          =   975
      Left            =   5400
      TabIndex        =   11
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CommandButton buttonRegister 
      Caption         =   "Register"
      Height          =   975
      Left            =   2640
      TabIndex        =   10
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox textPassword 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox textEmail 
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox textUsername 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox textLastName 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox textName 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Email"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Nombre de Usuario"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Apellido"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Nombre"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Text1_Change()

End Sub

Private Sub buttonCancel_Click()
    Me.Hide
    
    Login.Show
End Sub

Private Sub buttonRegister_Click()
    'Aqui va la logica para mandar los datos a la BD'
    Dim name As String
    Dim lastname As String
    Dim userName As String
    Dim email As String
    Dim password As String
    Dim sql As String
    
    
    
    name = textName.Text
    lastname = textLastName.Text
    userName = textUsername.Text
    email = textEmail.Text
    password = textPassword.Text
    
    
    AbrirConexion
    
    sql = "INSERT INTO Usuarios (Nombre, Apellido, Email, Password, NombreDeUsuario) VALUES ('" & name & "', '" & lastname & "', '" & email & "', '" & password & "', '" & userName & "')"
    
    conn.Execute sql
    
    MsgBox "Usuario insertado con éxito."
    
    textName.Text = ""
    textLastName.Text = ""
    textUsername.Text = ""
    textEmail.Text = ""
    textPassword.Text = ""
    
    CerrarConexion
    
End Sub

