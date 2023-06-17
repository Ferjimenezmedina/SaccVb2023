VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Transfe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspasos de Inventarios a Sucursales"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   37
      Top             =   4440
      Width           =   975
      Begin VB.Image Image13 
         Height          =   735
         Left            =   120
         MouseIcon       =   "Transfe.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Transfe.frx":030A
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Historial"
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   24
      Top             =   6840
      Width           =   975
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Transfe.frx":1D54
         MousePointer    =   99  'Custom
         Picture         =   "Transfe.frx":205E
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8160
      TabIndex        =   19
      Top             =   5640
      Width           =   975
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
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Transfe.frx":4140
         MousePointer    =   99  'Custom
         Picture         =   "Transfe.frx":444A
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Traspaso"
      TabPicture(0)   =   "Transfe.frx":5E0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text3(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Option1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCantE"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtIdProd"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtIndex"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Verificar Traspasos"
      TabPicture(1)   =   "Transfe.frx":5E28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "ListView4"
      Tab(1).Control(2)=   "ListView3"
      Tab(1).Control(3)=   "Combo3"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "Label9"
      Tab(1).ControlCount=   7
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
         Left            =   -72120
         Picture         =   "Transfe.frx":5E44
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   33
         Top             =   5280
         Width           =   7575
         _ExtentX        =   13361
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   32
         Top             =   2400
         Width           =   7575
         _ExtentX        =   13361
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
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74760
         TabIndex        =   31
         Top             =   960
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango de Reporte"
         Height          =   1455
         Left            =   -70320
         TabIndex        =   26
         Top             =   600
         Width           =   3015
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   27
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51118081
            CurrentDate     =   39885
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   960
            TabIndex        =   28
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51118081
            CurrentDate     =   39885
         End
         Begin VB.Label Label5 
            Caption         =   "Al:"
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Del:"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.TextBox txtIndex 
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtIdProd 
         Height          =   285
         Left            =   360
         TabIndex        =   22
         Top             =   7320
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtCantE 
         Height          =   285
         Left            =   3480
         TabIndex        =   21
         Top             =   4680
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   4815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5160
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Nombre"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Clave"
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   4320
         Width           =   5295
      End
      Begin VB.CommandButton Command2 
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
         Left            =   6480
         Picture         =   "Transfe.frx":8816
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   6
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
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
         Picture         =   "Transfe.frx":B1E8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7320
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   240
         TabIndex        =   8
         Top             =   5160
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
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
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Traspasos Detallados:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "De la Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "A la Sucursal"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "No. de Traspaso"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7680
         Y1              =   1560
         Y2              =   1560
      End
   End
End
Attribute VB_Name = "Transfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim elimi As Integer
Private Sub Combo1_DropDown()
    Combo1.Clear
    Buscarcbo
End Sub
Private Sub Buscarcbo(Optional ByVal Siguiente As Boolean = False)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscom(Optional ByVal Siguiente As Boolean = False)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo2.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_DropDown()
    Combo2.Clear
    Buscom
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HFFE1E1
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &H80000005
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Combo3.SetFocus
    ListView3.ListItems.Clear
    sBuscar = "SELECT *  FROM VSTRASPAS WHERE SUCURSAL_DE='" & Combo3.Text & "' OR SUCURSAL_AL='" & Combo3.Text & "' AND FECHA BETWEEN '" & DTPicker1.value & "' AND '" & DTPicker2.value & " ' ORDER BY FECHA DESC "
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
    Dim tRs As ADODB.Recordset
    If Text3(1).Text <> "" And Text4.Text <> "" And Text3(2).Text <> "" And Combo1.Text <> "" And Combo2.Text <> "" Then
        Dim tLi As ListItem
        If CDbl(Text3(2).Text) <= CDbl(txtCantE.Text) Then
            Set tLi = ListView2.ListItems.Add(, , Text3(1).Text & "")
            tLi.SubItems(1) = Text4.Text
            tLi.SubItems(2) = Text3(2).Text
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & Text3(2).Text & " WHERE ID_PRODUCTO = '" & Text3(1).Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            cnn.Execute (sBuscar)
            ListView1.ListItems.Clear
            If Option1.value = True Then
                sBuscar = "SELECT * FROM VSEXISALMA3 WHERE Descripcion LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            Else
                sBuscar = "SELECT * FROM VSEXISALMA3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
            End If
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                If Not (.BOF And .EOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                        If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                        .MoveNext
                    Loop
                End If
            End With
        Else
            MsgBox "NO TIENE SUFICIENTE EXISTENCIA PARA EL TRASPASO", vbInformation, "SACC"
        End If
        Text3(1).Text = ""
        Text4.Text = ""
        Text3(2).Text = ""
    Else
        If Text3(1).Text = "" Then
            MsgBox "DEBE SELECCIONAR UN PRODUCTO PARA SER AGREGADO!", vbInformation, "SACC"
        End If
        If Text3(2).Text = "" Then
            MsgBox "DEBE DAR UNA CANTIDAD!", vbInformation, "SACC"
        End If
        If Combo1.Text = "" Then
            MsgBox "DEBE SELECCIONAR UNA SUCURSAL DE ORIGEN", vbInformation, "SACC"
        End If
        If Combo2.Text = "" Then
            MsgBox "DEBE SELECCIONAR UNA SUCURSAL DE DESTINO!", vbInformation, "SACC"
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If txtIndex.Text <> "" Then
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & txtIDPROD.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If tRs.EOF And tRs.BOF Then
            sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & ListView2.ListItems.Item(Val(txtIndex.Text)).SubItems(2) & ", '" & txtIDPROD.Text & "', '" & Combo1.Text & "');"
        Else
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & Val(Replace(tRs.Fields("CANTIDAD"), ".", ",")) + Val(Replace(ListView2.ListItems.Item(Val(txtIndex.Text)).SubItems(2), ".", ",")) & " WHERE ID_PRODUCTO = '" & txtIDPROD.Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        cnn.Execute (sBuscar)
        ListView1.ListItems.Clear
        ListView2.ListItems.Remove Val(txtIndex.Text)
        txtIndex.Text = ""
        txtIDPROD.Text = ""
    Else
        MsgBox "DEBE SELECCIONAR UN PRODUCTO PARA QUITAR!", vbInformation, "SACC"
    End If
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
    DTPicker1.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker2.value = Format(Date, "dd/mm/yyyy")
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripcion", 5350
        .ColumnHeaders.Add , , "CANTIDAD", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Clave del Producto", 1850
        .ColumnHeaders.Add , , "Descripcion", 4000
        .ColumnHeaders.Add , , "CANTIDAD", 1500
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
        .ColumnHeaders.Add , , "ID PRODUCTO", 2000
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "ESTADO", 2000
    End With
    Dim sBuscar As String
    Dim tRs2 As ADODB.Recordset
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs2 = cnn.Execute(sBuscar)
    If Not (tRs2.EOF And tRs2.BOF) Then
        Do While Not tRs2.EOF
            Combo3.AddItem tRs2.Fields("NOMBRE")
            tRs2.MoveNext
        Loop
    End If
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_TRASPASO FROM TRASPASOS ORDER BY ID_TRASPASO DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text3(0).Text = CDbl(tRs.Fields("ID_TRASPASO")) + 1
    Else
        Text3(0).Text = "1"
    End If
    tRs.Close
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image13_Click()
    frmtrassucursal.Show vbModal
End Sub
Private Sub Image8_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim Cont As Integer
    Dim tRs As ADODB.Recordset
    Dim IDT As Integer
    Dim IDT1 As String
    If Combo1.Text <> "" And Combo2.Text <> "" And ListView2.ListItems.COUNT > 0 Then
        sBuscar = "INSERT INTO TRASPASOS (SUCURSAL_DE, SUCURSAL_AL, FECHA) VALUES ('" & Combo1.Text & "', '" & Combo2.Text & "', '" & Format(Date, "DD/MM/YYYY") & "');"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_TRASPASO FROM TRASPASOS ORDER BY ID_TRASPASO DESC"
        Set tRs = cnn.Execute(sBuscar)
        IDT = tRs.Fields("ID_TRASPASO")
        IDT1 = tRs.Fields("ID_TRASPASO")
        tRs.Close
        For Cont = 1 To ListView2.ListItems.COUNT
            sBuscar = "INSERT INTO TRASPASO_DETALLE (ID_TRASPASO, ID_PRODUCTO, CANTIDAD) VALUES ('" & IDT & "', '" & ListView2.ListItems.Item(Cont) & "', " & ListView2.ListItems.Item(Cont).SubItems(2) & ");"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "' AND SUCURSAL = '" & Combo2.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
               sBuscar = "INSERT INTO EXISTENCIAS (CANTIDAD, ID_PRODUCTO, SUCURSAL) VALUES(" & ListView2.ListItems.Item(Cont).SubItems(2) & ", '" & ListView2.ListItems.Item(Cont) & "', '" & Combo2.Text & "');"
            Else
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(Replace(tRs.Fields("CANTIDAD"), ",", "")) + CDbl(ListView2.ListItems.Item(Cont).SubItems(2)) & " WHERE ID_PRODUCTO = '" & ListView2.ListItems.Item(Cont) & "' AND SUCURSAL = '" & Combo2.Text & "'"
            End If
            cnn.Execute (sBuscar)
        Next Cont
        On Error GoTo ErrImp
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Impr_traspaso IDT1
        Impr_traspaso IDT1
        ListView2.ListItems.Clear
        ListView1.ListItems.Clear
        MsgBox "TRASPASO FINALIZADO CON EXITO!", vbInformation, "SACC"
        sBuscar = "SELECT ID_TRASPASO FROM TRASPASOS ORDER BY ID_TRASPASO DESC"
        Set tRs = cnn.Execute(sBuscar)
        Text3(0).Text = CDbl(tRs.Fields("ID_TRASPASO")) + 1
        tRs.Close
    End If
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
    Exit Sub
ErrImp:
    Err.Clear
    ListView2.ListItems.Clear
    ListView1.ListItems.Clear
    MsgBox "TRASPASO FINALIZADO EXITOSAMENTE CON ERROR DE IMPRESORA!", vbInformation, "SACC"
    sBuscar = "SELECT ID_TRASPASO FROM TRASPASOS ORDER BY ID_TRASPASO DESC"
    Set tRs = cnn.Execute(sBuscar)
    Text3(0).Text = CDbl(tRs.Fields("ID_TRASPASO")) + 1
    tRs.Close
End Sub
Private Sub Impr_traspaso(ID_TRASPASO As String)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim POSY As Integer
    sBuscar = "SELECT T.SUCURSAL_DE, T.SUCURSAL_AL, T.FECHA, TD.ID_PRODUCTO, TD.CANTIDAD FROM TRASPASOS AS T JOIN TRASPASO_DETALLE AS TD ON T.ID_TRASPASO = TD.ID_TRASPASO WHERE T.ID_TRASPASO = " & ID_TRASPASO
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
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
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        Printer.Print " FECHA DEL TRASPASO:  " & tRs.Fields("FECHA")
        Printer.Print ""
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print ""
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "PRODUCTO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2200
        Printer.Print "C. SURTIDA"
        Printer.CurrentY = POSY
        Printer.CurrentX = 3200
        Printer.Print "SUCURSAL DESTINO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 4200
        Printer.Print "SUCURSAL ORIGEN"
        POSY = POSY + 400
        Do While Not tRs.EOF
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2200
            Printer.Print tRs.Fields("CANTIDAD")
            Printer.CurrentY = POSY
            Printer.CurrentX = 3200
            Printer.Print tRs.Fields("SUCURSAL_AL")
            Printer.CurrentY = POSY
            Printer.CurrentX = 4200
            Printer.Print tRs.Fields("SUCURSAL_DE")
            tRs.MoveNext
            POSY = POSY + 200
            If POSY >= 14200 Then
                Printer.NewPage
                POSY = 2600
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
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.Print " FECHA DEL TRASPASO:  " & tRs.Fields("FECHA")
                Printer.Print ""
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print ""
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "PRODUCTO"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2200
                Printer.Print "C. SURTIDA"
                Printer.CurrentY = POSY
                Printer.CurrentX = 3200
                Printer.Print "SUCURSAL DESTINO"
                Printer.CurrentY = POSY
                Printer.CurrentX = 4200
                Printer.Print "SUCURSAL ORIGEN"
                POSY = POSY + 400
            End If
        Loop
        Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "FIN DEL LISTADO"
        Printer.EndDoc
        tRs.Close
    End If
End Sub
Private Sub Image9_Click()
    If ListView2.ListItems.COUNT > 0 Then
        MsgBox "NO PUEDE SALIR SIN TERMINAR EL MOVIEMINTO PENDIENTE", vbCritical, "SACC"
    Else
        Unload Me
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text3(1).Text = Item
    Text4.Text = Item.SubItems(1)
    txtCantE.Text = Item.SubItems(2)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView2.ListItems.COUNT > 0 Then
        txtIDPROD.Text = Item
        txtIndex.Text = Item.Index
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim num As Integer
    num = Item
    ListView4.ListItems.Clear
    sBuscar = "SELECT *  FROM TRASPASO_DETALLE WHERE ID_TRASPASO='" & num & "' "
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = ListView4.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                If tRs.Fields("ESTADO") = "A" Then
                    tLi.SubItems(2) = "ACEPTADO"
                End If
                If tRs.Fields("ESTADO") = "P" Then
                    tLi.SubItems(2) = "PENDIENTE"
                End If
                If tRs.Fields("ESTADO") = "R" Then
                    tLi.SubItems(2) = "RECHAZADO"
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        ListView1.ListItems.Clear
        'almacen 1
        If Option1.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA1 WHERE Descripcion LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        If Option2.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA1 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
        ' almacen 2
        If Option1.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA2 WHERE Descripcion LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        If Option2.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA2 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
        ' almacen 3
        If Option1.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA3 WHERE Descripcion LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        If Option2.value = True Then
            sBuscar = "SELECT * FROM VSEXISALMA3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                    If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = .Fields("CANTIDAD") & ""
                    .MoveNext
                Loop
            End If
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text3_Change(Index As Integer)
    If Text3(1).Text = "" Then
        Me.Command2.Enabled = False
    Else
        Me.Command2.Enabled = True
    End If
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).BackColor = &HFFE1E1
    Text3(Index).SetFocus
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890."
    If Index = 2 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).BackColor = &H80000005
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
