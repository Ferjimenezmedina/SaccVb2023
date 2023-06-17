VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form RecioMod 
   Caption         =   "Modificar precio del producto."
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form11"
   ScaleHeight     =   3315
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Clave"
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Nombre"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   7095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Precio base para venta :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "RecioMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
