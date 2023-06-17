VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form PrestamosCartuchos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prestamos cartuchos vacios"
   ClientHeight    =   6615
   ClientLeft      =   2220
   ClientTop       =   1215
   ClientWidth     =   12255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11160
      TabIndex        =   33
      Top             =   5280
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "PrestamosCartuchos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "PrestamosCartuchos.frx":030A
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
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "Prestamo"
      TabPicture(0)   =   "PrestamosCartuchos.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ListView2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Cerrar"
      TabPicture(1)   =   "PrestamosCartuchos.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Regresar"
      Tab(1).Control(1)=   "txtIDPrestamo"
      Tab(1).Control(2)=   "txtSucursal"
      Tab(1).Control(3)=   "CantRegresar"
      Tab(1).Control(4)=   "ID"
      Tab(1).Control(5)=   "Text8"
      Tab(1).Control(6)=   "ListView5"
      Tab(1).Control(7)=   "ListView4"
      Tab(1).Control(8)=   "Label11"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label9"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   8400
         Picture         =   "PrestamosCartuchos.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   6960
         Picture         =   "PrestamosCartuchos.frx":4DF6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4320
         Picture         =   "PrestamosCartuchos.frx":77C8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Regresar 
         Caption         =   "Regresar"
         Enabled         =   0   'False
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
         Left            =   -65400
         Picture         =   "PrestamosCartuchos.frx":A19A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtIDPrestamo 
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         Top             =   5880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtSucursal 
         Height          =   285
         Left            =   -74400
         TabIndex        =   38
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox CantRegresar 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -68160
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox ID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   5880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   5760
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   14
         Top             =   3120
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
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
         Height          =   1575
         Left            =   5880
         TabIndex        =   10
         Top             =   4080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Usuario seleccionado"
         Height          =   1575
         Left            =   5880
         TabIndex        =   25
         Top             =   480
         Width           =   4935
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "No usuario"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Articulo seleccionado"
         Height          =   1335
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   5415
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   2880
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   3000
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clave"
         Height          =   195
         Left            =   4320
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   2520
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   570
         Width           =   4695
      End
      Begin VB.Label Label11 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   -68880
         TabIndex        =   37
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Folio"
         Height          =   255
         Left            =   -70080
         TabIndex        =   36
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Clave del producto"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Notas"
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "PrestamosCartuchos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CanExi As String
Dim CanExi2 As String
Dim IndItm As String
Dim NoElim As String
Dim Regresado As Boolean
Private cnn As ADODB.Connection
Private Sub Combo1_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        Combo1.Clear
        If (.EOF And .BOF) Then
            MsgBox ("NO EXISTEN SUCURSALES")
        Else
            Do While Not .EOF
                Combo1.AddItem (.Fields("NOMBRE"))
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Combo1_LostFocus()
    Combo1.Enabled = False
End Sub
Private Sub Command1_Click()
    Dim tLi As ListItem
    If Text5.Text <> "" And Text4.Text <> "" Then
        If Text5.Text <= CanExi Then
            Set tLi = ListView3.ListItems.Add(, , Text4.Text)
            tLi.SubItems(1) = Text5.Text
            tLi.SubItems(2) = CanExi
            Text4.Text = ""
            Text5.Text = ""
            Text7.Enabled = True
            Text7.SetFocus
            Text1.Enabled = False
            ListView1.Enabled = False
        Else
            MsgBox "Existencia insuficiente para surtir!", vbExclamation, "SACC"
        End If
    End If
End Sub
Private Sub Command2_Click()
    On Error GoTo ManejaError
    If ListView3.ListItems.Count <> 0 Then
        ListView3.ListItems.Remove (CDbl(IndItm))
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If Text6.Text <> "" And ListView3.ListItems.Count <> 0 Then
        Dim nIDPrestamoTmporal
        Dim NoReg As Integer
        Dim NueCanEx As String
        Dim Con As Integer
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim sBuscar3 As String
        Dim sBuscar4 As String
        Dim sBuscar5 As String
        Dim sBuscar6 As String
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim tRs3 As ADODB.Recordset
        Dim tRs4 As ADODB.Recordset
        Dim tRs5 As ADODB.Recordset
        Dim tRs6 As ADODB.Recordset
        sBuscar = "INSERT INTO PRESTAMOS_CARTUCHOS(USUARIO,SUCURSAL,FECHA,NOTAS,ESTADO) VALUES ('" & Text6.Text & "','" & Combo1.Text & "','" & Format(Date, "dd/mm/yyyy") & "','" & Text7.Text & "','P')"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar2 = "SELECT TOP 1 ID_PRESTAMO FROM PRESTAMOS_CARTUCHOS ORDER BY ID_PRESTAMO DESC"
        Set tRs2 = cnn.Execute(sBuscar2)
        nIDPrestamoTmporal = tRs2.Fields("ID_PRESTAMO")
        NoReg = ListView3.ListItems.Count
        For Con = 1 To NoReg
            sBuscar3 = "INSERT INTO PRESTAMOS_CARTUCHOS_DETALLE (ID_PRESTAMO, ID_PRODUCTO, CANTIDAD) VALUES ('" & nIDPrestamoTmporal & "', '" & ListView3.ListItems(Con).Text & "', '" & ListView3.ListItems(Con).SubItems(1) & "')"
            cnn.Execute (sBuscar3)
            NueCanEx = CDbl(ListView3.ListItems(Con).SubItems(2)) - CDbl(ListView3.ListItems(Con).SubItems(1))
            sBuscar4 = "UPDATE EXISTENCIAS SET CANTIDAD = " & NueCanEx & " WHERE ID_PRODUCTO = '" & ListView3.ListItems(Con).Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Set tRs4 = cnn.Execute(sBuscar4)
            sBuscar5 = "SELECT CANTIDAD,ID_PRODUCTO,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView3.ListItems(Con).Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Set tRs5 = cnn.Execute(sBuscar5)
            If Not (tRs5.BOF And tRs5.EOF) Then
                If tRs5.Fields("CANTIDAD") = 0 Then
                    sBuscar6 = "DELETE FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & tRs5.Fields("ID_PRODUCTO") & "' AND SUCURSAL = '" & tRs5.Fields("SUCURSAL") & "' "
                    Set tRs6 = cnn.Execute(sBuscar6)
                End If
            End If
        Next Con
        Limpiar
        Actualiza
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "DELETE PRESTAMOS_CARTUCHOS_DETALLE WHERE WHERE ID ="
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
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID_USUARIO", 0
        .ColumnHeaders.Add , , "NOMBRE", 5600
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID ARTICULO", 2000
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 200
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Existencia", 1000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HoverSelection = False
        .FullRowSelect = False
        .ColumnHeaders.Add , , "Folio Prestamo", 1200
        .ColumnHeaders.Add , , "Usuario", 1300
        .ColumnHeaders.Add , , "Fecha de Prestamo", 1300
        .ColumnHeaders.Add , , "Notas", 6750
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "ID Folio Detalle", 1000
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 4800
        .ColumnHeaders.Add , , "Cantidad", 800
        .ColumnHeaders.Add , , "Sucursal", 1400
        .ColumnHeaders.Add , , "FOLIO", 1000
    End With
    Actualiza
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text6.Text = Item
    Text2.Text = Item.SubItems(1)
    Combo1.Enabled = True
    Text3.Enabled = True
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text4.Text = Item
    CanExi = Item.SubItems(2)
    Text5.Enabled = True
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IndItm = Item.Index
End Sub
Private Sub ListView4_BeforeLabelEdit(Cancel As Integer)
   Actualiza
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoElim = Item
    Text8.Text = Item.SubItems(1)
    ID.Text = Item
    CantRegresar = Item.SubItems(3)
    txtSucursal = Item.SubItems(4)
    txtIDPrestamo = Item.SubItems(5)
    Regresar.Enabled = True
End Sub
Private Sub Option1_Click()
    Text3.SetFocus
End Sub
Private Sub Option2_Click()
    Text3.SetFocus
End Sub
Private Sub Regresar_Click()
On Error GoTo ManejaError
    If Text8.Text <> "" And ID.Text <> "" And CantRegresar.Text <> "" And txtIDPrestamo.Text <> "" And txtSucursal.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim tRs3 As ADODB.Recordset
        Dim tRs4 As ADODB.Recordset
        Dim tRs5 As ADODB.Recordset
        Dim tRs6 As ADODB.Recordset
        Dim sBuscar As String
        Dim sBuscar2 As String
        Dim sBuscar3 As String
        Dim sBuscar4 As String
        Dim sBuscar5 As String
        Dim sBuscar6 As String
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Dim POSY As Integer
        Dim sBuscar7 As String
        Dim tRs7 As ADODB.Recordset
        POSY = 2200
        sBuscar7 = "SELECT * FROM VSPRESTAMOSCARTUCHOSVACIOS WHERE ID = " & ID.Text
        Set tRs7 = cnn.Execute(sBuscar7)
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. "; VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL." & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. "; VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print ""
        Printer.Print "     PRESTAMO DE MERCANCIA A NOMBRE DE : " & tRs7.Fields("NOMBRE")
        Printer.Print "     FECHA DE PRESTAMO : " & tRs7.Fields("FECHA")
        If tRs7.Fields("NOTAS") <> "" Then
            Printer.Print "     NOTAS : " & tRs7.Fields("NOTAS")
        End If
        Printer.Print "     FECHA : " & Format(Date, "dd/mm/yyyy")
        Printer.Print ""
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print ""
        POSY = POSY + 1000
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "PRODUCTO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1700
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        Printer.Print "CANTIDAD"
        POSY = POSY + 400
        Do While Not tRs7.EOF
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs7.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1700
            Printer.Print tRs7.Fields("Descripcion")
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print tRs7.Fields("CANTIDAD")
            POSY = POSY + 200
            tRs7.MoveNext
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print ""
                Printer.Print "     PRESTAMO DE MERCANCIA A NOMBRE DE : " & tRs7.Fields("NOMBRE")
                Printer.Print "     FECHA DE PRESTAMO : " & tRs7.Fields("FECHA")
                If tRs7.Fields("NOTAS") <> "" Then
                    Printer.Print "     NOTAS : " & tRs7.Fields("NOTAS")
                End If
                Printer.Print "     FECHA : " & Format(Date, "dd/mm/yyyy")
                Printer.Print ""
                Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print ""
                POSY = POSY + 1000
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "PRODUCTO"
                Printer.CurrentY = POSY
                Printer.CurrentX = 1400
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 7700
                Printer.Print "CANTIDAD"
            End If
        Loop
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print " FIN DEL LISTADO"
        Printer.EndDoc
        CommonDialog1.Copies = 1
        sBuscar = "DELETE FROM PRESTAMOS_CARTUCHOS_DETALLE WHERE ID = '" & ID.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        sBuscar5 = "SELECT * FROM PRESTAMOS_CARTUCHOS_DETALLE WHERE ID_PRESTAMO ='" & txtIDPrestamo.Text & "'"
        Set tRs5 = cnn.Execute(sBuscar5)
        Regresado = True
        If (tRs5.BOF And tRs5.EOF) Then
            sBuscar6 = "DELETE FROM PRESTAMOS_CARTUCHOS WHERE ID_PRESTAMO = '" & txtIDPrestamo.Text & "'"
            Set tRs6 = cnn.Execute(sBuscar6)
        End If
        sBuscar2 = "SELECT ID_PRODUCTO,CANTIDAD,SUCURSAL FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text8.Text & "' AND SUCURSAL = '" & txtSucursal.Text & "'"
        Set tRs2 = cnn.Execute(sBuscar2)
        If (Not tRs2.BOF And Not tRs2.EOF) Then
            CanExi2 = tRs2.Fields("CANTIDAD") + CantRegresar
            sBuscar3 = "UPDATE EXISTENCIAS SET CANTIDAD = " & CanExi2 & " WHERE ID_PRODUCTO = '" & Text8.Text & "' AND SUCURSAL = '" & txtSucursal.Text & "'"
            Set tRs3 = cnn.Execute(sBuscar3)
        Else
           sBuscar4 = "INSERT INTO EXISTENCIAS(ID_PRODUCTO,CANTIDAD,SUCURSAL) VALUES('" & Text8.Text & "','" & CantRegresar & "','" & txtSucursal.Text & "')"
           Set tRs4 = cnn.Execute(sBuscar4)
        End If
   Else
        MsgBox "DEBE DE SELECCIONAR UN PRODUCTO!", vbInformation, "SACC"
   End If
    Limpiar2
    Actualiza
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Buscar
        ListView1.Enabled = True
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
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
Public Sub Buscar()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    sBuscar = "SELECT ID_USUARIO,NOMBRE FROM USUARIOS WHERE NOMBRE LIKE '%" & Text1.Text & "%' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    ListView1.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_USUARIO"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        BuscarProducto
        ListView2.Enabled = True
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Text5.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Sub BuscarProducto()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    If Option1.Value Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA1 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD FROM VSEXISALMA1 WHERE Descripcion LIKE '%" & Text3.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
    End If
    Set tRs = cnn.Execute(sBuscar)
    ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                tRs.MoveNext
            Loop
        End If
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Valido = "1234567890"
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
Public Sub Limpiar()
    Text1.Text = ""
    ListView1.ListItems.Clear
    Combo1.Clear
    Text3.Text = ""
    ListView2.ListItems.Clear
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text2.Text = ""
    Text7.Text = ""
    ListView3.ListItems.Clear
    Text1.Enabled = True
End Sub
Public Sub Limpiar2()
    Text8.Text = ""
    ID.Text = ""
    txtIDPrestamo.Text = ""
    txtSucursal.Text = ""
    CantRegresar.Text = ""
    Regresar.Enabled = False
End Sub
Public Sub Actualiza()
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tLi2 As ListItem
    sBuscar = "SELECT ID,ID_PRESTAMO,NOMBRE,FECHA,NOTAS,ESTADO,SUCURSAL,ID_PRODUCTO,CANTIDAD,Descripcion FROM VSPRESTAMOSCARTUCHOSVACIOS WHERE ESTADO ='P' ORDER BY FECHA"
    Set tRs = cnn.Execute(sBuscar)
    ListView4.ListItems.Clear
    ListView5.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRESTAMO"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = tRs.Fields("FECHA")
            tLi.SubItems(3) = tRs.Fields("NOTAS")
            Set tLi2 = ListView5.ListItems.Add(, , tRs.Fields("ID"))
            tLi2.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi2.SubItems(2) = tRs.Fields("Descripcion")
            tLi2.SubItems(3) = tRs.Fields("CANTIDAD")
            tLi2.SubItems(4) = tRs.Fields("SUCURSAL")
            tLi2.SubItems(5) = tRs.Fields("ID_PRESTAMO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
