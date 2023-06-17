VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRepComProd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Comiciones de Producción"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmRepComProd.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepComProd.frx":030A
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Reporte"
      TabPicture(0)   =   "FrmRepComProd.frx":1F0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Rango de Fechas"
         Height          =   1335
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   15925249
            CurrentDate     =   39493
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   15925249
            CurrentDate     =   39493
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Toner :"
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
         Left            =   480
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Tinta :"
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
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3840
      TabIndex        =   0
      Top             =   1440
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmRepComProd.frx":1F28
         MousePointer    =   99  'Custom
         Picture         =   "FrmRepComProd.frx":2232
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmRepComProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
    DTPicker1.Value = Format(Date - 15, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image14_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT SUM(CANTIDAD - CANTIDAD_NO_SIRVIO) AS Total From COMANDAS_DETALLES_2 WHERE (ESTADO_ACTUAL = 'N' OR ESTADO_ACTUAL = 'L' OR  ESTADO_ACTUAL = 'I') AND (FECHA_FIN BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (TIPO = 'I')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("Total")) Then
            Label5.Caption = tRs.Fields("Total")
        Else
            Label5.Caption = "0"
        End If
    Else
        Label5.Caption = "0"
    End If
    sBuscar = "SELECT SUM(CANTIDAD - CANTIDAD_NO_SIRVIO) AS Total From COMANDAS_DETALLES_2 WHERE (ESTADO_ACTUAL = 'N' OR ESTADO_ACTUAL = 'L' OR  ESTADO_ACTUAL = 'I') AND (FECHA_FIN BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') AND (TIPO = 'T')"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If Not IsNull(tRs.Fields("Total")) Then
            Label6.Caption = tRs.Fields("Total")
        Else
            Label6.Caption = "0"
        End If
    Else
        Label6.Caption = "0"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
