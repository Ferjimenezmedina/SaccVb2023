VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmScrap 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scrap"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmScrap.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTPicker1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtIDCom"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.TextBox txtIDCom 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Material Seleccionado"
         Enabled         =   0   'False
         Height          =   2295
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   8055
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   7
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1935
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2160
            TabIndex        =   9
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            Height          =   885
            Left            =   4320
            TabIndex        =   10
            Top             =   1200
            Width           =   3615
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
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
            Left            =   1440
            Picture         =   "frmScrap.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtCantOriginal 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   1560
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtID_REP 
            Height          =   285
            Left            =   480
            TabIndex        =   28
            Top             =   1560
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   255
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Candidad Dañada"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            Height          =   255
            Left            =   2160
            TabIndex        =   31
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario"
            Height          =   255
            Left            =   4320
            TabIndex        =   30
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Material Dañado"
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   5640
         Width           =   8055
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
            Left            =   6720
            Picture         =   "frmScrap.frx":29EE
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtIndex 
            Height          =   285
            Left            =   6720
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   150
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2175
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   3836
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
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Clave"
            Height          =   255
            Left            =   6600
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1800
         Picture         =   "frmScrap.frx":53C0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   435
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50593793
         CurrentDate     =   39121
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   1935
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Comanda"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Juegos de Reparación"
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   135
         Left            =   5520
         TabIndex        =   35
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   21
      Top             =   7200
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmScrap.frx":7D92
         MousePointer    =   99  'Custom
         Picture         =   "frmScrap.frx":809C
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   19
      Top             =   4800
      Width           =   975
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         MouseIcon       =   "frmScrap.frx":A17E
         MousePointer    =   99  'Custom
         Picture         =   "frmScrap.frx":A488
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario"
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   9000
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   8760
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   0
      Top             =   6000
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmScrap.frx":BDC2
         MousePointer    =   99  'Custom
         Picture         =   "frmScrap.frx":C0CC
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label13 
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   735
      Left            =   8520
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
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
Attribute VB_Name = "frmScrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub cmdAgregar_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Combo1.Text <> "" Then
        If (Val(Replace(Text3.Text, ",", "")) <= Val(Replace(txtCantOriginal.Text, ",", ""))) Or (Val(Replace(Text3.Text, ",", "")) > 0) Then
            Dim tLi As ListItem
            Dim NR As Integer
            Dim C As Integer
            Dim Agregar As Boolean
            Agregar = True
            NR = ListView3.ListItems.Count
            For C = 1 To NR
                If ListView3.ListItems.Item(C).SubItems(1) = Text1.Text Then
                    Agregar = False
                End If
            Next C
            If Agregar Then
                Set tLi = ListView3.ListItems.Add(, , txtIDCom.Text)
                    tLi.SubItems(1) = Text1.Text
                    tLi.SubItems(2) = Text3.Text
                    tLi.SubItems(3) = Combo1.Text
                    tLi.SubItems(4) = Text4.Text
                    tLi.SubItems(5) = DTPicker1.Value
                    sBuscar = "SELECT Descripcion FROM ALMACEN2 WHERE ID_PRODUCTO = '" & Text1.Text & "'"
                    Set tRs = cnn.Execute(sBuscar)
                    With tRs
                        If Not .EOF And Not .BOF Then
                            tLi.SubItems(6) = "2"
                        Else
                            sBuscar = "SELECT Descripcion FROM ALMACEN1 WHERE ID_PRODUCTO = '" & Text1.Text & "'"
                            Set tRs = cnn.Execute(sBuscar)
                            If Not tRs.EOF And Not tRs.BOF Then
                                tLi.SubItems(6) = "1"
                            Else
                                tLi.SubItems(6) = "-"
                            End If
                        End If
                    End With
                    tLi.SubItems(7) = txtID_REP.Text
            Else
                MsgBox "NO SE PUEDE INSERTAR 2 VECES EL MISMO PRODUCTO", vbCritical, "SACC"
            End If
            txtIDCom.Enabled = False
            Command2.Enabled = False
            Frame1.Enabled = False
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text6.Text = ""
            txtIndex.Text = ""
            txtCantOriginal.Text = ""
            Combo1.Text = ""
        Else
            MsgBox "LA CANTIDAD NO PUEDE SER MAYOR A LA REQUERIDA O CERO", vbCritical, "SACC"
        End If
    Else
        MsgBox "SE DEBE INTRODUCIR UN MOTIVO", vbCritical, "SACC"
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    ListView3.ListItems.Remove (Val(txtIndex.Text))
    txtIndex.Text = ""
    Text6.Text = ""
    Command1.Enabled = False
    If ListView3.ListItems.Count = 0 Then
        txtIDCom.Enabled = True
        Command2.Enabled = True
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Limpiar
    If txtIDCom.Text <> "" Then
        sBuscar = "SELECT * FROM COMANDAS_DETALLES_2 WHERE ESTADO_ACTUAL IN ('P','M','R','S') AND ID_COMANDA = " & txtIDCom.Text
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_COMANDA"))
                        tLi.SubItems(1) = .Fields("ID_PRODUCTO")
                        tLi.SubItems(2) = .Fields("CANTIDAD")
                    .MoveNext
                Loop
            Else
                MsgBox "ESTA COMANDA NO PUEDE PROCESAR MATERIAL DAÑADO DEVIDO A SU ESTADO", vbInformation, "SACC"
            End If
        End With
    Else
        MsgBox "DEBE INTRODUCIR EL NUMERO DE COMANDA", vbCritical, "SACC"
    End If
    
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
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
        .ColumnHeaders.Add , , "No. Comanda", 1000
        .ColumnHeaders.Add , , "Clave", 1500
        .ColumnHeaders.Add , , "Cantidad", 6500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Comanda", 0
        .ColumnHeaders.Add , , "Clave", 1000
        .ColumnHeaders.Add , , "Descripcion", 2500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "ID_Rep", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Comanda", 0
        .ColumnHeaders.Add , , "Clave", 1000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Motivo", 1000
        .ColumnHeaders.Add , , "Comentario", 1000
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Almacen", 100
        .ColumnHeaders.Add , , "ID_Rep", 0
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Comanda", 0
        .ColumnHeaders.Add , , "Clave", 1000
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Motivo", 1000
        .ColumnHeaders.Add , , "Comentario", 1000
        .ColumnHeaders.Add , , "Fecha", 1000
        .ColumnHeaders.Add , , "Almacen", 100
    End With
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    sBuscar = "SELECT * FROM MOTIVOS_SCRAP"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            sBuscar = "SELECT * FROM MOTIVOS_SCRAP"
            Set tRs = cnn.Execute(sBuscar)
            Do While Not .EOF
                Combo1.AddItem (.Fields("MOTIVO"))
                .MoveNext
            Loop
        Else
            Do While Not .EOF
                Combo1.AddItem (.Fields("MOTIVO"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Limpiar()
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    Frame1.Enabled = False
    Command1.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    txtIndex.Text = ""
    txtCantOriginal.Text = ""
    Combo1.Text = ""
End Sub
Private Sub Image1_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo1.Clear
    sBuscar = "SELECT * FROM MOTIVOS_SCRAP"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                Combo1.AddItem (.Fields("MOTIVO"))
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Image8_Click()
    Dim NR As Integer
    Dim Cont As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    NR = ListView3.ListItems.Count
    ListView4.ListItems.Clear
    Cont = 1
    Do While Cont <= NR
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = 'BODEGA' AND ID_PRODUCTO = '" & ListView3.ListItems.Item(Cont).SubItems(1) & "' AND CANTIDAD >= " & ListView3.ListItems.Item(Cont).SubItems(2)
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If .EOF And .BOF Then
                If MsgBox("NO TIENE EXISTENCIA PARA SURTIR EL PRODUCTO, DESEA SUSTITUIRLO", vbYesNo, "SACC") = vbYes Then
                    Text5(0).Text = ""
                    Text5(1).Text = ""
                    frmJRAlterno.txtCantidad.Text = ListView3.ListItems.Item(Cont).SubItems(2)
                    frmJRAlterno.lblProducto.Caption = ListView3.ListItems.Item(Cont).SubItems(7)
                    frmJRAlterno.lblPieza.Caption = ListView3.ListItems.Item(Cont).SubItems(1)
                    frmJRAlterno.lblAlmacen.Caption = ListView3.ListItems.Item(Cont).SubItems(6)
                    frmJRAlterno.Caption = "JUEGO DE REPARACION ALTERNO DE PRODUCTO DAÑADO"
                    frmJRAlterno.Show vbModal, Me
                    If Text5(0).Text <> "" Then
                        ListView3.ListItems.Item(Cont).SubItems(1) = Text5(0).Text
                        ListView3.ListItems.Item(Cont).SubItems(2) = Text5(1).Text
                        Set tLi = ListView4.ListItems.Add(, , ListView3.ListItems.Item(Cont))
                        tLi.SubItems(1) = ListView3.ListItems.Item(Cont).SubItems(1)
                        tLi.SubItems(2) = ListView3.ListItems.Item(Cont).SubItems(2)
                        tLi.SubItems(3) = ListView3.ListItems.Item(Cont).SubItems(3)
                        tLi.SubItems(4) = ListView3.ListItems.Item(Cont).SubItems(4)
                        tLi.SubItems(5) = ListView3.ListItems.Item(Cont).SubItems(5)
                        tLi.SubItems(6) = ListView3.ListItems.Item(Cont).SubItems(6)
                    Else
                        Cont = NR + 1
                    End If
                Else
                    Cont = NR + 1
                End If
            Else
                Set tLi = ListView4.ListItems.Add(, , ListView3.ListItems.Item(Cont))
                tLi.SubItems(1) = ListView3.ListItems.Item(Cont).SubItems(1)
                tLi.SubItems(2) = ListView3.ListItems.Item(Cont).SubItems(2)
                tLi.SubItems(3) = ListView3.ListItems.Item(Cont).SubItems(3)
                tLi.SubItems(4) = ListView3.ListItems.Item(Cont).SubItems(4)
                tLi.SubItems(5) = ListView3.ListItems.Item(Cont).SubItems(5)
                tLi.SubItems(6) = ListView3.ListItems.Item(Cont).SubItems(6)
            End If
        End With
        Cont = Cont + 1
    Loop
    If ListView3.ListItems.Count = ListView4.ListItems.Count Then
        For Cont = 1 To NR
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & ListView3.ListItems.Item(Cont).SubItems(2) & " WHERE SUCURSAL = 'BODEGA' AND ID_PRODUCTO = '" & ListView3.ListItems.Item(Cont).SubItems(1) & "'"
            cnn.Execute (sBuscar)
            sBuscar = "INSERT INTO SCRAP (ID_COMANDA, ID_PRODUCTO, CANTIDAD, MOTIVO, COMENTARIO, FECHA) VALUES (" & ListView3.ListItems.Item(Cont) & ", '" & ListView3.ListItems.Item(Cont).SubItems(1) & "', " & ListView3.ListItems.Item(Cont).SubItems(2) & ", '" & ListView3.ListItems.Item(Cont).SubItems(3) & "', '" & ListView3.ListItems.Item(Cont).SubItems(4) & "', '" & ListView3.ListItems.Item(Cont).SubItems(5) & "');"
            cnn.Execute (sBuscar)
        Next Cont
        Imprimir
    End If
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    Limpiar
End Sub
Private Sub Image9_Click()
    If ListView3.ListItems.Count = 0 Then
        Unload Me
    Else
        MsgBox "ELIMINE TODOS LOS PRODUCTOS DEL MATERIAL DAÑADO PARA PODER SALIR", vbInformation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.ListItems.Count > 0 Then
        Dim Clave As String
        Dim IDC As String
        Clave = Item.SubItems(1)
        IDC = Item
        If Tiene_JR_Temporal(IDC, Clave) Then
            Llenar_Lista_JR_Temporal Clave, IDC, CDbl(Item.SubItems(2))
        Else
            Llenar_Lista_JR Clave, IDC, CDbl(Item.SubItems(2))
        End If
    End If
End Sub
Function Tiene_JR_Temporal(IDC As String, Clave As String) As Boolean
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    
    sqlQuery = "SELECT COUNT(TEMPORAL) TEMPORAL FROM JR_TEMPORALES WHERE ID_COMANDA = " & IDC & " AND ID_REPARACION = '" & Clave & "'"
    Set tRs = cnn.Execute(sqlQuery)
    
    If tRs.Fields("TEMPORAL") = 0 Then
        Tiene_JR_Temporal = False
    Else
        Tiene_JR_Temporal = True
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_JR_Temporal(ID As String, CID As String, cant As Double)
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sqlQuery = "SELECT J.ID_PRODUCTO, J.CANTIDAD, A.Descripcion FROM JR_TEMPORALES AS J JOIN ALMACEN2 AS A ON J.ID_PRODUCTO = A.ID_PRODUCTO WHERE J.ID_COMANDA = " & CID & " AND J.ID_REPARACION = '" & ID & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.EOF And tRs.BOF Then
        sqlQuery = "SELECT J.ID_PRODUCTO, J.CANTIDAD, A.Descripcion FROM JR_TEMPORALES AS J JOIN ALMACEN1 AS A ON J.ID_PRODUCTO = A.ID_PRODUCTO WHERE J.ID_COMANDA = " & CID & " AND J.ID_REPARACION = '" & ID & "'"
        Set tRs = cnn.Execute(sqlQuery)
    End If
    With tRs
        While Not .EOF
            Set tLi = ListView2.ListItems.Add(, , CID)
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO")
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = .Fields("Descripcion")
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = CDbl(.Fields("CANTIDAD")) * cant
            tLi.SubItems(4) = ID
            .MoveNext
        Wend
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_JR(ID As String, CID As String, cant As Double)
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sqlQuery = "SELECT J.ID_PRODUCTO, J.CANTIDAD, A.Descripcion FROM JUEGO_REPARACION AS J JOIN ALMACEN2 AS A ON J.ID_PRODUCTO = A.ID_PRODUCTO WHERE ID_REPARACION = '" & ID & "'"
    Set tRs = cnn.Execute(sqlQuery)
    If tRs.EOF And tRs.BOF Then
        sqlQuery = "SELECT J.ID_PRODUCTO, J.CANTIDAD, A.Descripcion FROM JUEGO_REPARACION AS J JOIN ALMACEN1 AS A ON J.ID_PRODUCTO = A.ID_PRODUCTO WHERE ID_REPARACION = '" & ID & "'"
        Set tRs = cnn.Execute(sqlQuery)
    End If
    With tRs
        While Not .EOF
            Set tLi = ListView2.ListItems.Add(, , CID)
            If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO")
            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = .Fields("Descripcion")
            If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(3) = CDbl(.Fields("CANTIDAD")) * cant
            tLi.SubItems(4) = ID
            .MoveNext
        Wend
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView2.ListItems.Count > 0 Then
        Frame1.Enabled = True
        Text1.Text = Item.SubItems(1)
        Text2.Text = Item.SubItems(2)
        Text3.Text = Item.SubItems(3)
        txtID_REP.Text = Item.SubItems(4)
        txtCantOriginal.Text = Item.SubItems(3)
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView3.ListItems.Count > 0 Then
        txtIndex.Text = Item.Index
        Text6.Text = Item
        Command1.Enabled = True
    End If
End Sub
Private Sub txtIDCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Imprimir()
    Dim POSY As Integer
    Dim NR As Integer
    Dim Cont As Integer
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "                                                                                          REPOSICION DE MATERIAL POR DAÑO"
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = 2200
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "PRODUCTO"
    Printer.CurrentY = POSY
    Printer.CurrentX = 2200
    Printer.Print "CANTIDAD"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3500
    Printer.Print "COMANDA"
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    POSY = POSY + 200
    NR = ListView4.ListItems.Count
    For Cont = 1 To NR
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView4.ListItems.Item(Cont).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print ListView4.ListItems.Item(Cont).SubItems(2)
        Printer.CurrentY = POSY
        Printer.CurrentX = 3500
        Printer.Print ListView4.ListItems.Item(Cont)
        If POSY >= 14200 Then
            Printer.NewPage
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
            Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "                                                                                          REPOSICION DE MATERIAL POR DAÑO"
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            POSY = 2200
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "PRODUCTO"
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print "CANTIDAD"
            Printer.CurrentY = POSY
            Printer.CurrentX = 3500
            Printer.Print "COMANDA"
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            POSY = POSY + 200
        End If
    Next Cont
    POSY = POSY + 200
    Printer.CurrentY = POSY
    Printer.CurrentX = 0
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.EndDoc
End Sub
