VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCancelCompCartVac 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar Compra de Cartuchos Vacios"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCancelCompCartVac.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCancelCompCartVac.frx":030A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   0
      Top             =   840
      Width           =   975
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmCancelCompCartVac.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCancelCompCartVac.frx":26F6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cancelar Compra de Cartuchos Vacios"
      TabPicture(0)   =   "FrmCancelCompCartVac.frx":41A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Compra :"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmCancelCompCartVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image6_Click()
    If Text1.Text <> "" Then
        If MsgBox("ESTA SEGURO QUE DESEA CANCELAR LA COMPRA NO. " & Text1.Text & "?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            
             sBuscar = "SELECT * FROM REV_COMPRA_ALMACEN1 WHERE GRUPO = " & Text1.Text
            Set tRs = cnn.Execute(sBuscar)
            
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "DELETE FROM REV_COMPRA_ALMACEN1 WHERE GRUPO = " & Text1.Text
                cnn.Execute (sBuscar)
                MsgBox "LA COMPRA HA SIDO CANCELADA", vbInformation, "SACC"
                Text1.Text = ""
            Else
                MsgBox "EL NUMERO DE COMPRA NO EXISTE", vbInformation, "SACC"
            End If
        End If
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

