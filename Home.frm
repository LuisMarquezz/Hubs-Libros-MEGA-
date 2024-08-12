VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H80000005&
   ClientHeight    =   12945
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   24450
   LinkTopic       =   "Form1"
   ScaleHeight     =   12945
   ScaleWidth      =   24450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton buttonSearch 
      Caption         =   "Buscar Libro"
      Height          =   735
      Left            =   15000
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox textSearch 
      Height          =   855
      Left            =   8520
      TabIndex        =   8
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   5
      Left            =   4680
      TabIndex        =   7
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   4
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   3
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label labelLink 
      Height          =   495
      Index           =   0
      Left            =   8520
      TabIndex        =   10
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MEGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   24495
   End
   Begin VB.Menu Usuario 
      Caption         =   "Usuario"
      Begin VB.Menu Usuario_1 
         Caption         =   "Cerrar Sesion"
      End
   End
   Begin VB.Menu Favoritos 
      Caption         =   "Libros Favoritos"
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub buttonSearch_Click()
    
    Dim bookName As String
    bookName = textSearch.Text
    
    GetBookData bookName
End Sub

Private Sub Label2_Click()
    Dim name As String
    Dim lastname As String
    
    AbrirConexion()
    
    CerrarConexion()
    
    
End Sub




Private Sub Usuario_1_Click()
    Unload Me
    Load Login
    Login.Show
End Sub

Private Sub GetBookData(bookName As String)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = "https://www.googleapis.com/books/v1/volumes?q=" & bookName
    
    http.Open "Get", url, False
    http.send
    
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText
        
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)
        
        'MsgBox response
        labelLink.Caption = json("items")(1)("kind")
        
    End If
    
End Sub


