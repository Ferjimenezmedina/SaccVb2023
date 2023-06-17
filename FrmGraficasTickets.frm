VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2B26B39A-53D1-4401-B64E-1B727C1D2B68}#9.0#0"; "ADMGráficos.ocx"
Object = "{2FE3662E-0169-4252-8869-49150227B9EC}#2.0#0"; "Grafico_Circular.ocx"
Begin VB.Form FrmGraficasTickets 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gràficas de Tickets"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   10
      Top             =   6360
      Width           =   975
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmGraficasTickets.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmGraficasTickets.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " De Barra"
      TabPicture(0)   =   "FrmGraficasTickets.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ADMGraf4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ADMGraf3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ADMGraf2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ADMGraf1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTPicker5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Porcentajes"
      TabPicture(1)   =   "FrmGraficasTickets.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ADMPorc1"
      Tab(1).ControlCount=   1
      Begin Grafico_Circular.ADMPorc ADMPorc1 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   11668
         Valor_Total     =   0
         Mostrar_Leyenda =   0   'False
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin VB.CommandButton Command6 
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
         Height          =   375
         Left            =   3960
         Picture         =   "FrmGraficasTickets.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   44847
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   44847
      End
      Begin ADMGráficos.ADMGraf ADMGraf1 
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf2 
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf3 
         Height          =   1455
         Left            =   240
         TabIndex        =   6
         Top             =   4200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin ADMGráficos.ADMGraf ADMGraf4 
         Height          =   1455
         Left            =   240
         TabIndex        =   7
         Top             =   5760
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mostrar_Leyenda =   0   'False
         Color_Fondo     =   -2147483643
         Color_Barra1    =   8388608
         Color_Barra2    =   0
         Gráfico_Barras  =   -1  'True
         Color_Texto     =   -2147483640
         Mostrar_Media   =   0   'False
         Color_Media     =   0
      End
      Begin VB.Label Label16 
         Caption         =   "Del :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Al :"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmGraficasTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
'Dim tRs As ADODB.Recordset
Private Sub Command6_Click()
    FunGraficas
End Sub
Private Sub FunGraficas()
    Dim n As Integer
    Dim sBuscar As String
    Dim PROMEDIO As Double
    Dim tRs As ADODB.Recordset
    n = n + 1
    ADMGraf1.Gráfico_Barras = True
    ADMGraf1.Limpiar
    sBuscar = "SELECT DEPARTAMENTO_DESTINO, SUM(DATEDIFF(HOUR, FECHA, FECHA_CIERRE)) AS HORAS, COUNT(DEPARTAMENTO_DESTINO) AS VECES From TICKETS WHERE (ESTATUS = 'F') AND (FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY DEPARTAMENTO_DESTINO"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf1.Título = "Promedio de horas por ticket por Dpto."
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("HORAS")) Then
                PROMEDIO = CDbl(tRs.Fields("HORAS")) / CDbl(tRs.Fields("VECES"))
                ADMGraf1.Introducir Mid(tRs.Fields("DEPARTAMENTO_DESTINO"), 1, 8), CSng(PROMEDIO), CLng(PROMEDIO * 1600), QBColor(15)
            End If
            tRs.MoveNext
        Loop
        ADMGraf1.Dibujar
    End If
    ADMGraf2.Gráfico_Barras = True
    ADMGraf2.Limpiar
    sBuscar = "SELECT DEPARTAMENTO_DESTINO, SUM(DATEDIFF(HOUR, FECHA, FECHA_CIERRE)) AS HORAS, COUNT(DEPARTAMENTO_DESTINO) AS VECES From TICKETS WHERE (ESTATUS = 'F') AND (FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY DEPARTAMENTO_DESTINO"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf2.Título = "Número de ticket cerrados por Dpto."
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("HORAS")) Then
                ADMGraf2.Introducir Mid(tRs.Fields("DEPARTAMENTO_DESTINO"), 1, 8), CSng(CDbl(tRs.Fields("VECES"))), CLng(CDbl(tRs.Fields("VECES")) * 1600), QBColor(15)
            End If
            tRs.MoveNext
        Loop
        ADMGraf2.Dibujar
    End If
    ADMGraf3.Gráfico_Barras = True
    ADMGraf3.Limpiar
    sBuscar = "SELECT DEPARTAMENTO_DESTINO, COUNT(DEPARTAMENTO_DESTINO) AS VECES From TICKETS WHERE (FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY DEPARTAMENTO_DESTINO"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf3.Título = "Número de ticket levantados al Dpto."
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("VECES")) Then
                ADMGraf3.Introducir Mid(tRs.Fields("DEPARTAMENTO_DESTINO"), 1, 8), CSng(CDbl(tRs.Fields("VECES"))), CLng(CDbl(tRs.Fields("VECES")) * 1600), QBColor(15)
            End If
            tRs.MoveNext
        Loop
        ADMGraf3.Dibujar
    End If
    ADMGraf4.Gráfico_Barras = True
    ADMGraf4.Limpiar
    sBuscar = "SELECT DEPARTAMENTO_DESTINO, COUNT(DEPARTAMENTO_DESTINO) AS VECES From TICKETS WHERE (FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') AND (ESTATUS <> 'F') GROUP BY DEPARTAMENTO_DESTINO"
    Set tRs = cnn.Execute(sBuscar)
    ADMGraf4.Título = "Número de ticket abiertos por Dpto."
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("VECES")) Then
                ADMGraf4.Introducir Mid(tRs.Fields("DEPARTAMENTO_DESTINO"), 1, 8), CSng(CDbl(tRs.Fields("VECES"))), CLng(CDbl(tRs.Fields("VECES")) * 1600), QBColor(15)
            End If
            tRs.MoveNext
        Loop
        ADMGraf4.Dibujar
    End If
    
    

    
    Dim Cont As Long
    Dim SUMA As Integer
    Cont = 1
    sBuscar = "SELECT DEPARTAMENTO_DESTINO, COUNT(DEPARTAMENTO_DESTINO) AS VECES From TICKETS WHERE (FECHA BETWEEN '" & DTPicker5.Value & "' AND '" & DTPicker6.Value & "') GROUP BY DEPARTAMENTO_DESTINO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("VECES")) Then
                SUMA = CDbl(tRs.Fields("VECES")) + SUMA
            End If
            tRs.MoveNext
        Loop
    End If
    tRs.MoveFirst
    If Not (tRs.EOF And tRs.BOF) Then
        ADMPorc1.Limpiar
        ADMPorc1.Valor_Total = SUMA
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("VECES")) Then
                ADMPorc1.Añadir_Sector tRs.Fields("DEPARTAMENTO_DESTINO"), Cont, QBColor(Cont), CDbl(tRs.Fields("VECES"))
                Cont = Cont + 1
            End If
            tRs.MoveNext
        Loop
        ADMPorc1.Dibujar
    End If
    
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    DTPicker5.Value = Date - 30
    DTPicker6.Value = Date
    FunGraficas
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub

