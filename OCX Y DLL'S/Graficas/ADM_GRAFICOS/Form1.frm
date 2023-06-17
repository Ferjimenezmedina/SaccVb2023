VERSION 5.00
Object = "{2B26B39A-53D1-4401-B64E-1B727C1D2B68}#9.0#0"; "ADMGráficos.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6327.974
   ScaleMode       =   0  'User
   ScaleWidth      =   10594.66
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar Gráfico"
         Height          =   465
         Left            =   180
         TabIndex        =   2
         Top             =   315
         Width           =   1560
      End
   End
   Begin ADMGráficos.ADMGraf Graf 
      Height          =   5745
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   10134
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mostrar_Leyenda =   0   'False
      Color_Fondo     =   16777215
      Color_Barra1    =   0
      Color_Barra2    =   8421504
      Gráfico_Barras  =   -1  'True
      Color_Texto     =   0
      Mostrar_Media   =   0   'False
      Color_Media     =   4210752
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private n As Integer
Private Sub Command1_Click()
    n = n + 1
    If n Mod 2 = 0 Then
        Graf.Gráfico_Barras = False
    Else
        Graf.Gráfico_Barras = True
    End If
    Graf.Limpiar
    
    Graf.Título = "GRÁFICO DE PRUEBAS"
    Graf.Introducir "ENE", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "FEB", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "MAR", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "ABR", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "MAY", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "JUN", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "JUL", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "AGO", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "SEP", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "OCT", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "NOV", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Introducir "DIC", CSng(Rnd * 1000) + 1, CLng(Rnd * 16000000), QBColor(15)
    Graf.Dibujar
End Sub

Private Sub Form_Load()
    Randomize Timer
    Command1_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Graf.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
    Frame1.Move 0, Me.ScaleHeight - Frame1.Height, Me.ScaleWidth, Frame1.Height
    On Error GoTo 0
End Sub
