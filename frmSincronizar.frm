VERSION 5.00
Begin VB.Form frmSincronizar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESPERE..."
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   1410
         Left            =   0
         Picture         =   "frmSincronizar.frx":0000
         Top             =   0
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmSincronizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    deAPTONER.TRAER_HORA_FECHA_SISTEMA
    With deAPTONER.rsTRAER_HORA_FECHA_SISTEMA
        Time = TimeValue(!FECHAHORA)
        Date = DateValue(!FECHAHORA)
        .Close
    End With
    
    Unload Me
    
End Sub
