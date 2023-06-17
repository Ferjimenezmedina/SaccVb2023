VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form listadodecomandas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Garantias"
   ClientHeight    =   7110
   ClientLeft      =   2205
   ClientTop       =   1680
   ClientWidth     =   11655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   18
      Top             =   4560
      Width           =   975
      Begin VB.Image cmdEnviar 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmatgarantias.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmatgarantias.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label8 
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10560
      TabIndex        =   0
      Top             =   5760
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmatgarantias.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "frmatgarantias.frx":1FD6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmatgarantias.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDesc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtID"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Juego De Reparacion"
         Height          =   5895
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   5175
         Begin MSComctlLib.ListView ListView1 
            Height          =   5535
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   9763
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
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Frame Frame1 
         Caption         =   "Materiales"
         Height          =   5775
         Left            =   5400
         TabIndex        =   5
         Top             =   840
         Width           =   4815
         Begin VB.CommandButton Command1 
            Caption         =   "Quitar"
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
            Left            =   3600
            Picture         =   "frmatgarantias.frx":40D4
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5280
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   4815
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   8493
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblID3 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   6000
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Clave:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "------------------------------------------------"
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   5280
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4320
         Picture         =   "frmatgarantias.frx":6AA6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDesc 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Juego de Reparacion"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "listadodecomandas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim StrRep As String
Dim StrRep2 As String
Dim StrRep3 As String
Dim StrRep4 As String
Dim DelIndex As Integer
Private Sub cmdEnviar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim Con As Integer
    Dim NumeroRegistros As Integer
    Dim produc As String
    Dim cant As Integer
    If MsgBox("¿DESEA CONTINUAR,¡LOS PRODUCTOS SON LOS CORRECTOS?  ", vbYesNo + vbCritical + vbDefaultButton1) = vbYes Then
        NumeroRegistros = ListView3.ListItems.Count
        For Con = 1 To NumeroRegistros
            If Me.ListView3.ListItems.Item(Con).Checked = True Then
                sBuscar = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Me.ListView3.ListItems(Con).SubItems(1) & "' AND SUCURSAL = 'BODEGA'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.BOF And tRs.EOF) Then
                    Do While Not tRs.EOF
                        produc = tRs.Fields("ID_PRODUCTO")
                        sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - ('" & Me.ListView3.ListItems(Con).SubItems(3) & "') WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "' AND SUCURSAL = 'BODEGA'"
                        cnn.Execute (sBuscar)
                        tRs.MoveNext
                    Loop
                End If
             End If
        Next Con
        ListView3.ListItems.Clear
        ListView1.ListItems.Clear
        Text1.Text = ""
    End If
End Sub
Private Sub Command1_Click()
    ListView3.ListItems.Remove DelIndex
End Sub
Private Sub Command3_Click()
    Buscar
End Sub
Private Sub Command6_Click()
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Text2.Text
    sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION LIKE '%" & sBuscar & "%'"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("Descripcion") & ""
                     tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                .MoveNext
            Loop
        End If
    End With
    If ListView1.ListItems.Count = 0 Then
        MsgBox "El producto no existe en el almacen"
    End If
    '///////////////////////////////////////////////////////////////////////////////
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
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
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 3200
        .ColumnHeaders.Add , , "Descripcion", 3200
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Juego", 1200
        .ColumnHeaders.Add , , "Clave del Producto", 1200
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Cantidad que se uso", 1000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = Text1.Text
    txtId.Text = ""
    lblDesc.Caption = ""
    sBuscar = "SELECT ID_REPARACION FROM JUEGO_REPARACION  WHERE ID_REPARACION='" & Text1.Text & "' GROUP BY ID_REPARACION "
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                       .MoveNext
            Loop
        Else
            ListView1.ListItems.Clear
            MsgBox "El producto buscado no es juego de reparacion o no existe"
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    If ListView1.ListItems.Count > 0 Then
        txtId.Text = Item
        lblDesc.Caption = Item.SubItems(1)
        sBuscar = Item
        sBuscar = "SELECT * FROM JUEGO_REPARACION WHERE ID_REPARACION = '" & Item & "'"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                ListView3.ListItems.Clear
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView3.ListItems.Add(, , .Fields("ID_REPARACION") & "")
                    tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                    tLi.SubItems(2) = .Fields("CANTIDAD")
                    tLi.SubItems(3) = CDbl(.Fields("CANTIDAD")) * CDbl(Text3.Text)
                    .MoveNext
                Loop
                StrRep = sBuscar
            End If
        End With
        Frame2.Enabled = True
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView3.ListItems.Count > 0 Then
        lblID3.Caption = Item.Index
        Label6.Caption = Item.ListSubItems(1)
        DelIndex = 0
        DelIndex = Item.Index
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo ManejaError
    If KeyAscii = 13 Then
        If Text1.Text <> "" Then
            Buscar
        End If
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ%1234567890- "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
