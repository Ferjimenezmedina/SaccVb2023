VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmtrassucursal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificar Traspaso a Sucusal"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   11
      Top             =   6240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmtrassucursal.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmtrassucursal.frx":030A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label34 
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
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "C"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Historial de Traspasos  entre Sucursales"
      TabPicture(0)   =   "frmtrassucursal.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Traspasos del Dia"
      TabPicture(1)   =   "frmtrassucursal.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "ListView1"
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(3)=   "Combo1"
      Tab(1).ControlCount=   4
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74640
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
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
         Left            =   -68400
         Picture         =   "frmtrassucursal.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   18
         Top             =   1560
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9551
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
         Left            =   4320
         TabIndex        =   17
         Top             =   4200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Rechazar"
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
         Left            =   6480
         Picture         =   "frmtrassucursal.frx":4DF6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aprobar"
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
         Left            =   5040
         Picture         =   "frmtrassucursal.frx":77C8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de Reporte"
         Height          =   1335
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   3015
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   6
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39885
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   960
            TabIndex        =   7
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39885
         End
         Begin VB.Label Label4 
            Caption         =   "Del:"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Al:"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
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
         Left            =   2760
         Picture         =   "frmtrassucursal.frx":A19A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   4800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   -74520
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Traspasos Detallados :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Traspasos :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmtrassucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim Venta As String
Dim costo As String
Dim utilidad As String
Dim Sucursal As String
Dim Cliente As String
Dim producto As String
Dim num As Integer
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Combo3.SetFocus
    ListView3.ListItems.Clear
    sBuscar = "SELECT * FROM TRASPASOS WHERE  SUCURSAL_AL = '" & Combo3.Text & "' AND FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & " ' ORDER BY FECHA DESC "
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView3.ListItems.Add(, , .Fields("ID_TRASPASO") & "")
                If Not IsNull(.Fields("SUCURSAL_DE")) Then tLi.SubItems(1) = .Fields("SUCURSAL_DE") & ""
                If Not IsNull(.Fields("SUCURSAL_AL")) Then tLi.SubItems(2) = .Fields("SUCURSAL_AL") & ""
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(3) = .Fields("FECHA") & ""
                tRs.MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim IDT As Integer
    Dim IDT1 As String
    For Cont = 1 To ListView4.ListItems.COUNT
        sBuscar = "UPDATE TRASPASO_DETALLE SET ESTADO ='A' WHERE ID_TRASPASO= '" & num & "'"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView4.ListItems.Item(Cont).SubItems(1) & "' AND SUCURSAL = '" & Combo3.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & ListView4.ListItems.Item(Cont).SubItems(2) & ", '" & ListView4.ListItems.Item(Cont).SubItems(1) & "', '" & Combo3.Text & "');"
        Else
           sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(Replace(tRs.Fields("CANTIDAD"), ",", "")) + CDbl(ListView4.ListItems.Item(Cont).SubItems(2)) & " WHERE ID_PRODUCTO = '" & ListView4.ListItems.Item(Cont).SubItems(1) & "' AND SUCURSAL = '" & Combo3.Text & "'"
        End If
        cnn.Execute (sBuscar)
    Next Cont
    ListView4.ListItems.Clear
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim IDT As Integer
    sBuscar = "UPDATE TRASPASO_DETALLE SET ESTADO ='R' WHERE ID_TRASPASO= '" & num & "'"
    cnn.Execute (sBuscar)
    ListView4.ListItems.Clear
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM VsTraspasos WHERE SUCURSAL_AL = '" & Combo1.Text & "'  and  fecha='" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("SUCURSAL_DE")) Then tLi.SubItems(1) = tRs.Fields("SUCURSAL_DE")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_TRASPASO_DETALLE")) Then tLi.SubItems(4) = tRs.Fields("ID_TRASPASO_DETALLE")
            tRs.MoveNext
        Loop
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
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Id Producto", 1500
        .ColumnHeaders.Add , , "Sucursal que envia", 4500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "ID", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "ID TRASPASO", 1000
        .ColumnHeaders.Add , , "DE SUCURSAL", 2000
        .ColumnHeaders.Add , , "A SUCURSAL", 2000
        .ColumnHeaders.Add , , "FECHA", 1500
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
         .ColumnHeaders.Add , , "TRASPASO", 2000
        .ColumnHeaders.Add , , "ID PRODUCTO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "ESTADO", 0
    End With
    Dim sBuscar As String
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        Do While Not tRs2.EOF
            Combo3.AddItem tRs2.Fields("NOMBRE")
            tRs2.MoveNext
        Loop
    End If
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs3 = cnn.Execute(sBuscar)
    If Not (tRs3.EOF And tRs3.BOF) Then
        Do While Not tRs3.EOF
            Combo1.AddItem tRs3.Fields("NOMBRE")
            tRs3.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
   Unload Me
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    num = Item
    Text1.Text = Item
    ListView4.ListItems.Clear
    sBuscar = "SELECT *  FROM TRASPASO_DETALLE WHERE ID_TRASPASO='" & num & "' "
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView4.ListItems.Add(, , .Fields("ID_TRASPASO") & "")
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = .Fields("ID_PRODUCTO") & ""
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                If tRs.Fields("ESTADO") = "A" Then
                    tLi.SubItems(3) = "ACEPTADO"
                End If
                If tRs.Fields("ESTADO") = "P" Then
                    tLi.SubItems(3) = "PENDIENTE"
                End If
                If tRs.Fields("ESTADO") = "R" Then
                    tLi.SubItems(3) = "RECHAZADO"
                End If
                tRs.MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub


