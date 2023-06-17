VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmProvConsumibles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores de Consumibles"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   32
      Top             =   3480
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
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmProvConsumibles.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmProvConsumibles.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   30
      Top             =   2280
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar"
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
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmProvConsumibles.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmProvConsumibles.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "FrmProvConsumibles.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Información"
      TabPicture(1)   =   "FrmProvConsumibles.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text14"
      Tab(1).Control(1)=   "cmdEnviar"
      Tab(1).Control(2)=   "Text13"
      Tab(1).Control(3)=   "Text12"
      Tab(1).Control(4)=   "Text11"
      Tab(1).Control(5)=   "Text10"
      Tab(1).Control(6)=   "Text9"
      Tab(1).Control(7)=   "Text8"
      Tab(1).Control(8)=   "Text7"
      Tab(1).Control(9)=   "Text6"
      Tab(1).Control(10)=   "Text5"
      Tab(1).Control(11)=   "Text4"
      Tab(1).Control(12)=   "Text3"
      Tab(1).Control(13)=   "Label14"
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(15)=   "Label12"
      Tab(1).Control(16)=   "Label11"
      Tab(1).Control(17)=   "Label10"
      Tab(1).Control(18)=   "Label9"
      Tab(1).Control(19)=   "Label8"
      Tab(1).Control(20)=   "Label7"
      Tab(1).Control(21)=   "Label6"
      Tab(1).Control(22)=   "Label5"
      Tab(1).Control(23)=   "Label4"
      Tab(1).Control(24)=   "Label3"
      Tab(1).ControlCount=   25
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   -73920
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2400
         Width           =   5895
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Limpiar"
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
         Left            =   -72120
         Picture         =   "FrmProvConsumibles.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   975
         Left            =   -74880
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3000
         Width           =   6855
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -69360
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -71640
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -73920
         MaxLength       =   30
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -71400
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -73920
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -69000
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -70680
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -73920
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73920
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4200
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   6015
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label14 
         Caption         =   "* E-mail"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "País"
         Height          =   255
         Left            =   -71760
         TabIndex        =   29
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Estado"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Nota"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono 2"
         Height          =   255
         Left            =   -70200
         TabIndex        =   26
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   -72360
         TabIndex        =   25
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "* RFC"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "* C.P."
         Height          =   255
         Left            =   -69480
         TabIndex        =   23
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   -71280
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "* Colonia"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "* Dirección"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "* Nombre"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionado :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmProvConsumibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim IdProv As String
Private Sub cmdEnviar_Click()
    IdProv = ""
    Text2.Text = ""
    Text3.Text = ""
    Text11.Text = ""
    Text10.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text7.Text = ""
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Id Proveedor", 0
        .ColumnHeaders.Add , , "Nombre", 5500
        .ColumnHeaders.Add , , "Telefono", 1500
        .ColumnHeaders.Add , , "RFC", 1500
        .ColumnHeaders.Add , , "Direccion", 0
        .ColumnHeaders.Add , , "Colonia", 0
        .ColumnHeaders.Add , , "Ciudad", 0
        .ColumnHeaders.Add , , "Estado", 0
        .ColumnHeaders.Add , , "Pais", 0
        .ColumnHeaders.Add , , "Telefono 2", 0
        .ColumnHeaders.Add , , "Notas", 0
        .ColumnHeaders.Add , , "CP", 0
        .ColumnHeaders.Add , , "EMAIL", 0
    End With
End Sub
Private Sub Image8_Click()
    Text2.Text = Replace(Text2.Text, ",", "")
    Text3.Text = Replace(Text3.Text, ",", "")
    Text11.Text = Replace(Text11.Text, ",", "")
    Text10.Text = Replace(Text10.Text, ",", "")
    Text4.Text = Replace(Text4.Text, ",", "")
    Text5.Text = Replace(Text5.Text, ",", "")
    Text6.Text = Replace(Text6.Text, ",", "")
    Text8.Text = Replace(Text8.Text, ",", "")
    Text9.Text = Replace(Text9.Text, ",", "")
    Text12.Text = Replace(Text12.Text, ",", "")
    Text13.Text = Replace(Text13.Text, ",", "")
    Text7.Text = Replace(Text7.Text, ",", "")
    If Text7.Text = "" Then
        Text7.Text = "0"
    End If
    Dim sBuscar As String
    If IdProv <> "" Then
        If MsgBox("TIENE UN PROVEEDOR SELECCIONADO, DESEA REEMPLAZAR LA INFORMACIÓN DEL PROVEEDOR CON LA NUEVA INFORMACIÓN?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            Dim tRs As ADODB.Recordset
            sBuscar = "UPDATE PROVEEDOR_CONSUMO SET NOMBRE = '" & Text3.Text & "', TELEFONO1 = '" & Text11.Text & "', RFC = '" & Text10.Text & "', DIRECCION = '" & Text4.Text & "', COLONIA = '" & Text5.Text & "', CIUDAD = '" & Text6.Text & "', ESTADO = '" & Text8.Text & "', PAIS = '" & Text9.Text & "', TELEFONO2 = '" & Text12.Text & "', NOTAS = '" & Text13.Text & "', CP = " & Text7.Text & ", EMAIL = '" & Text14.Text & "' WHERE ID_PROVEEDOR = " & IdProv
            Set tRs = cnn.Execute(sBuscar)
            MsgBox "LA INFORMACION HA SIDO CAMBIADA!", vbInformation, "SACC"
        Else
            sBuscar = "INSERT INTO PROVEEDOR_CONSUMO (NOMBRE, TELEFONO1, RFC, DIRECCION, COLONIA, CIUDAD, ESTADO, PAIS, TELEFONO2, CP, NOTAS, EMAIL) VALUES ('" & Text3.Text & "', '" & Text11.Text & "', '" & Text10.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text12.Text & "', '" & Text7.Text & "', '" & Text13.Text & "', '" & Text14.Text & "');"
            cnn.Execute (sBuscar)
            MsgBox "SE HA AGREGADO UN NUEVO PROVEEDOR!", vbInformation, "SACC"
        End If
    Else
        sBuscar = "INSERT INTO PROVEEDOR_CONSUMO (NOMBRE, TELEFONO1, RFC, DIRECCION, COLONIA, CIUDAD, ESTADO, PAIS, TELEFONO2, CP, NOTAS, EMAIL) VALUES ('" & Text3.Text & "', '" & Text11.Text & "', '" & Text10.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & Text12.Text & "', '" & Text7.Text & "', '" & Text13.Text & "', '" & Text14.Text & "');"
        cnn.Execute (sBuscar)
    End If
    IdProv = ""
    Text2.Text = ""
    Text3.Text = ""
    Text11.Text = ""
    Text10.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text7.Text = ""
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdProv = Item
    Text2.Text = Item.SubItems(1)
    Text3.Text = Item.SubItems(1)
    Text11.Text = Item.SubItems(2)
    Text10.Text = Item.SubItems(3)
    Text4.Text = Item.SubItems(4)
    Text5.Text = Item.SubItems(5)
    Text6.Text = Item.SubItems(6)
    Text8.Text = Item.SubItems(7)
    Text9.Text = Item.SubItems(8)
    Text12.Text = Item.SubItems(9)
    Text13.Text = Item.SubItems(10)
    Text7.Text = Item.SubItems(11)
    Text14.Text = Item.SubItems(12)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1.Text <> "" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT * FROM PROVEEDOR_CONSUMO WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                    Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
                    If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    If Not IsNull(tRs.Fields("TELEFONO1")) Then tLi.SubItems(2) = tRs.Fields("TELEFONO1")
                    If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(3) = tRs.Fields("RFC")
                    If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(4) = tRs.Fields("DIRECCION")
                    If Not IsNull(tRs.Fields("COLONIA")) Then tLi.SubItems(5) = tRs.Fields("COLONIA")
                    If Not IsNull(tRs.Fields("CIUDAD")) Then tLi.SubItems(6) = tRs.Fields("CIUDAD")
                    If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(7) = tRs.Fields("ESTADO")
                    If Not IsNull(tRs.Fields("PAIS")) Then tLi.SubItems(8) = tRs.Fields("PAIS")
                    If Not IsNull(tRs.Fields("TELEFONO2")) Then tLi.SubItems(9) = tRs.Fields("TELEFONO2")
                    If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(10) = tRs.Fields("NOTAS")
                    If Not IsNull(tRs.Fields("CP")) Then tLi.SubItems(11) = tRs.Fields("CP")
                    If Not IsNull(tRs.Fields("EMAIL")) Then tLi.SubItems(12) = tRs.Fields("EMAIL")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890-()"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890-()"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.-_@abcdefghijklmnopqrstuvwxyz;"
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
