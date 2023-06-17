VERSION 5.00
Object = "{2FE3662E-0169-4252-8869-49150227B9EC}#2.0#0"; "Grafico_Circular.ocx"
Begin VB.Form FRM_Grafico_Circular 
   Caption         =   "Form1"
   ClientHeight    =   7692
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   12576
   LinkTopic       =   "Form1"
   ScaleHeight     =   7692
   ScaleWidth      =   12576
   StartUpPosition =   3  'Windows Default
   Begin Grafico_Circular.ADMPorc ADMPorc1 
      Height          =   5832
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   10332
      _ExtentX        =   18225
      _ExtentY        =   10287
      Valor_Total     =   0
      Mostrar_Leyenda =   -1  'True
      BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   12632256
      Separación_Filas=   10
   End
End
Attribute VB_Name = "FRM_Grafico_Circular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ADMPorc1.Limpiar
    ADMPorc1.Valor_Total = 2400
    ADMPorc1.Añadir_Sector "F.C. Barcelona", 1, QBColor(9), 700
    ADMPorc1.Añadir_Sector "Bayern de Munich", 2, QBColor(10), 600
    ADMPorc1.Añadir_Sector "Chelsea", 3, QBColor(11), 500
    ADMPorc1.Añadir_Sector "Manchester United", 4, QBColor(12), 300
    ADMPorc1.Añadir_Sector "Milan", 5, QBColor(13), 200
    ADMPorc1.Añadir_Sector "Borussia de Dormunt", 6, QBColor(14), 100
    ADMPorc1.Dibujar
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim sep As Single
    sep = 100
    ADMPorc1.Move sep, sep, Me.ScaleWidth - (sep * 2), Me.ScaleHeight - (sep * 2)
    On Error GoTo 0
End Sub
