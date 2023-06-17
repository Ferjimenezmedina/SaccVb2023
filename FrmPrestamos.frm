VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPrestamos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prestamos de Inventario a Clientes"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   44
      Top             =   5520
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
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmPrestamos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmPrestamos.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9360
      TabIndex        =   42
      Top             =   4320
      Width           =   975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
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
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FrmPrestamos.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmPrestamos.frx":26F6
         Top             =   120
         Width           =   705
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Capturar"
      TabPicture(0)   =   "FrmPrestamos.frx":4478
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListView2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Option1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Cerrar"
      TabPicture(1)   =   "FrmPrestamos.frx":4494
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "Text11"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "Command4"
      Tab(1).Control(4)=   "ListView5"
      Tab(1).Control(5)=   "ListView4"
      Tab(1).Control(6)=   "Label12"
      Tab(1).ControlCount=   7
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
         Left            =   -67200
         Picture         =   "FrmPrestamos.frx":44B0
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -74160
         TabIndex        =   47
         Top             =   720
         Width           =   6615
      End
      Begin VB.Frame Frame4 
         Caption         =   "Devolición Parcial"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   36
         Top             =   5280
         Width           =   4935
         Begin VB.TextBox Text10 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cerrar"
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
            Left            =   3480
            Picture         =   "FrmPrestamos.frx":6E82
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1800
            TabIndex        =   38
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Clave del Producto"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Cantidad Pendiente"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1455
         End
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
         Left            =   -67320
         Picture         =   "FrmPrestamos.frx":9854
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   34
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   33
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
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
         Left            =   7800
         Picture         =   "FrmPrestamos.frx":C226
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6120
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   195
         Left            =   3000
         TabIndex        =   31
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Clave"
         Height          =   195
         Left            =   3000
         TabIndex        =   30
         Top             =   2640
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Información del Prestamo"
         Height          =   2295
         Left            =   4680
         TabIndex        =   20
         Top             =   1800
         Width           =   4335
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50266113
            CurrentDate     =   39099
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   2040
            TabIndex        =   25
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox Text7 
            Height          =   855
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de entrega :"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Deposito por prestamo"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Notas :"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label7 
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
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
         Left            =   6360
         Picture         =   "FrmPrestamos.frx":EBF8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Articulo Seleccionado"
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   5280
         Width           =   4335
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Text            =   "1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   3255
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
            Left            =   3000
            Picture         =   "FrmPrestamos.frx":115CA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cliente Seleccionado"
         Height          =   1095
         Left            =   4680
         TabIndex        =   8
         Top             =   600
         Width           =   4335
         Begin VB.TextBox Text6 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label6 
            Caption         =   "No. Cliente"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   3120
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2990
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2990
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   1815
         Left            =   4680
         TabIndex        =   7
         Top             =   4200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
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
      Begin VB.Label Label12 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim CanExi As String
Dim PVent As String
Dim IndItm As String
Dim NoFolioElim As String
Dim CantPend As String
Dim NoElim As String
Private Sub Combo1_DropDown()
    Me.Combo1.Clear
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    sBus = "SELECT NOMBRE FROM SUCURSALES ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBus)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            If Not IsNull(tRs.Fields("NOMBRE")) Then Combo1.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Command1_Click()
    If Text4.Text <> "" And Text3.Text <> "" Then
        If CDbl(Text4.Text) <= CDbl(CanExi) Then
            Set tLi = ListView3.ListItems.Add(, , Text3.Text)
            tLi.SubItems(1) = Text4.Text
            tLi.SubItems(2) = PVent
            tLi.SubItems(3) = CanExi
        Else
            MsgBox "LA EXISTENCIA ES INSUFICIENTE PARA SURTIR!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Command2_Click()
    ListView3.ListItems.Remove (CDbl(IndItm))
End Sub
Private Sub Command3_Click()
    If Text6.Text <> "" And ListView3.ListItems.Count <> 0 Then
        Dim sBuscar As String
        Dim FechEnt As String
        Dim tRs As ADODB.Recordset
        Dim tRs2 As ADODB.Recordset
        Dim NoReg As Integer
        Dim Con As Integer
        Dim NueCanEx As String
        If DTPicker1.Value = Format(Date, "dd/mm/yyyy") Then
            FechEnt = ""
        Else
            FechEnt = Format(DTPicker1.Value, "DD/MM/YYYY")
        End If
        Text8.Text = Replace(Text8.Text, ",", "")
        Text7.Text = Replace(Text7.Text, ",", "")
        If Text8.Text = "" Then
            Text8.Text = "0"
        End If
        sBuscar = "INSERT INTO PRESTAMOS_CLIENTES (ID_CLIENTE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS, ESTADO) VALUES (" & Text6.Text & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & FechEnt & "', " & Text8.Text & ", '" & Text7.Text & "', 'P' );"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_PRESTAMO FROM PRESTAMOS_CLIENTES ORDER BY ID_PRESTAMO DESC"
        Set tRs = cnn.Execute(sBuscar)
        NoReg = ListView3.ListItems.Count
        For Con = 1 To NoReg
            ListView3.ListItems(Con).SubItems(1) = Replace(ListView3.ListItems(Con).SubItems(1), ",", "")
            ListView3.ListItems(Con).SubItems(2) = Replace(ListView3.ListItems(Con).SubItems(2), ",", "")
            sBuscar = "INSERT INTO PRESTAMOS_CLIENTES_DETALLE (ID_PRESTAMO, ID_PRODUCTO, CANTIDAD, PRECIO_VENTA) VALUES (" & tRs.Fields("ID_PRESTAMO") & ", '" & ListView3.ListItems(Con).Text & "', " & ListView3.ListItems(Con).SubItems(1) & ", " & ListView3.ListItems(Con).SubItems(2) & ");"
            cnn.Execute (sBuscar)
            NoFolioElim = tRs.Fields("ID_PRESTAMO")
            NueCanEx = CDbl(ListView3.ListItems(Con).SubItems(3)) - CDbl(ListView3.ListItems(Con).SubItems(1))
            NueCanEx = Replace(NueCanEx, ",", "")
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NueCanEx & " WHERE ID_PRODUCTO = '" & ListView3.ListItems(Con).Text & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Set tRs2 = cnn.Execute(sBuscar)
        Next Con
        FunImpPrestamo
        ListView1.ListItems.Clear
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
    Else
        If Text6.Text <> "" Then
            MsgBox "ES NECESARIO SELECCIONAR UN CLIENTE!", vbInformation, "SACC"
        Else
            MsgBox "ES NECESARIO SELECCIONAR UNO O VARIOS ARTICULOS!", vbInformation, "SACC"
        End If
    End If
End Sub
Private Sub Command4_Click()
    Dim NoCheq As Integer
    Dim Con As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim NuExPro As String
    NoCheq = ListView5.ListItems.Count
    For Con = 1 To NoCheq
        If ListView5.ListItems(Con).Checked Then
            sBuscar = "SELECT SUM(CANTIDAD) AS CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con).SubItems(1) & "' AND SUCURSAL = 'BODEGA'"
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & ListView5.ListItems(Con).SubItems(1) & "', " & ListView5.ListItems(Con).SubItems(3) & ", 'BODEGA' );"
                cnn.Execute (sBuscar)
            Else
                NuExPro = CDbl(tRs.Fields("CANTIDAD")) + CDbl(ListView5.ListItems(Con).SubItems(3))
                NuExPro = Replace(NuExPro, ",", "")
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NuExPro & " WHERE ID_PRODUCTO = '" & ListView5.ListItems(Con).SubItems(1) & "' AND SUCURSAL = 'BODEGA'"
                Set tRs = cnn.Execute(sBuscar)
            End If
            'sBuscar = "DELETE FROM PRESTAMOS_CLIENTES_DETALLE WHERE ID = " & ListView5.ListItems(Con).Text
            'Set tRs = cnn.Execute(sBuscar)
            sBuscar = "UPDATE PRESTAMOS_CLIENTES SET ESTADO = 'C' WHERE ID = " & ListView5.ListItems(Con).Text
            Set tRs = cnn.Execute(sBuscar)
        End If
    Next Con
    Actualiza
End Sub
Private Sub Command5_Click()
    If Text10.Text <> "" Then
        If CDbl(Text9.Text) < CDbl(CantPend) Then
            MsgBox "LA CANTIDAD PENDIENTE ES MAYOR QUE LA CANTIDAD PRESTADA!", vbInformation, "SACC"
        Else
            Dim sBuscar As String
            Dim Can As String
            Dim tRs As ADODB.Recordset
            Dim NuExPro As String
            Text9.Text = Replace(Text9.Text, ".", ",")
            CantPend = Replace(CantPend, ".", ",")
            Can = Format(CDbl(CantPend) - CDbl(Text9.Text), "###,###,##0.00")
            Can = Replace(Can, ",", "")
            If CDbl(Can) <> 0 Then
                sBuscar = "UPDATE PRESTAMOS_CLIENTES_DETALLE SET CANTIDAD = " & Can & " WHERE ID = " & NoElim
                Set tRs = cnn.Execute(sBuscar)
            Else
                sBuscar = "UPDATE PRESTAMOS_CLIENTES SET ESTADO = 'C' WHERE ID = " & NoElim
                Set tRs = cnn.Execute(sBuscar)
                'sBuscar = "DELETE FROM PRESTAMOS_CLIENTES_DETALLE WHERE ID = " & NoElim
                'Set tRs = cnn.Execute(sBuscar)
            End If
            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text10.Text & "' AND SUCURSAL = 'BODEGA'"
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & Text10.Text & "', " & Text9.Text & ", 'BODEGA' );"
                cnn.Execute (sBuscar)
            Else
                NuExPro = CDbl(tRs.Fields("CANTIDAD")) + CDbl(Text9.Text)
                NuExPro = Replace(NuExPro, ",", "")
                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & NuExPro & " WHERE ID_PRODUCTO = '" & Text10.Text & "' AND SUCURSAL = 'BODEGA'"
                Set tRs = cnn.Execute(sBuscar)
            End If
        End If
        Actualiza
    End If
    Text9.Text = ""
    Text10.Text = ""
    Can = ""
    CantPend = ""
    NoElim = ""
End Sub
Private Sub Command6_Click()
    BuscaPrestamo
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Combo1.Text = "BODEGA"
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
        .ColumnHeaders.Add , , "Id Cliente", 0
        .ColumnHeaders.Add , , "Nombre", 5950
        .ColumnHeaders.Add , , "RFC", 1500
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 6000
        .ColumnHeaders.Add , , "Existencia", 1500
        .ColumnHeaders.Add , , "Precio Venta", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Precio Venta", 1500
        .ColumnHeaders.Add , , "Existencia", 0
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Folio Prestamo", 1500
        .ColumnHeaders.Add , , "Cliente", 5500
        .ColumnHeaders.Add , , "Fecha de Prestamo", 1500
        .ColumnHeaders.Add , , "Fecha de Entrega", 1500
        .ColumnHeaders.Add , , "Deposito", 1500
        .ColumnHeaders.Add , , "Notas", 7500
    End With
    With ListView5
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "", 300
        .ColumnHeaders.Add , , "Clave Producto", 2500
        .ColumnHeaders.Add , , "Descripcion", 5000
        .ColumnHeaders.Add , , "Cantidad", 1500
        .ColumnHeaders.Add , , "Precio Venta", 1500
    End With
    BuscaPrestamo
End Sub
Private Sub Image3_Click()
    FunImpPrestamo
End Sub
Private Sub FunImpPrestamo()
On Error GoTo ManejaError
    If ListView5.ListItems.Count = 0 Then
        MsgBox "DEBE SELECCIONAR EL PRESTAMO A IMPRIMIR!", vbInformation, "SACC"
    Else
        CommonDialog1.Flags = 64
        CommonDialog1.CancelError = True
        CommonDialog1.ShowPrinter
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim POSY As Integer
        POSY = 2200
        sBuscar = "SELECT * FROM VsRepPrestamos WHERE ID_PRESTAMO = " & NoFolioElim
        Set tRs = cnn.Execute(sBuscar)
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
        Printer.Print "     PRESTAMO DE MERCANCIA A NOMBRE DE : " & tRs.Fields("NOMBRE")
        Printer.Print "     FECHA DE PRESTAMO : " & tRs.Fields("FECHA_PRESTAMO")
        If tRs.Fields("FECHA_ENTREGA") <> "01/01/1900" Then
            Printer.Print "     FECHA DE ENTREGA : " & tRs.Fields("FECHA_ENTREGA")
        End If
        If tRs.Fields("DEPOSITO") <> 0 Then
            Printer.Print "     DEPOSITO : " & tRs.Fields("DEPOSITO")
        End If
        If tRs.Fields("NOTAS") <> "" Then
            Printer.Print "     NOTAS : " & tRs.Fields("NOTAS")
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
        Printer.CurrentX = 1900
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 7700
        Printer.Print "CANTIDAD"
        Printer.CurrentY = POSY
        Printer.CurrentX = 8800
        Printer.Print "P/UNITARIO"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10000
        POSY = POSY + 400
        Printer.Print "IMPORTE"
        sBuscar = "SELECT * FROM VsRepPrestamoDet WHERE ID_PRESTAMO = " & NoFolioElim
        Set tRs = cnn.Execute(sBuscar)
        Do While Not tRs.EOF
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            If Len(tRs.Fields("Descripcion")) > 55 Then
                Printer.Print Mid(tRs.Fields("Descripcion"), 1, 55)
            Else
                Printer.Print tRs.Fields("Descripcion")
            End If
            Printer.CurrentY = POSY
            Printer.CurrentX = 8100
            Printer.Print tRs.Fields("CANTIDAD")
            Printer.CurrentY = POSY
            Printer.CurrentX = 8800
            Printer.Print tRs.Fields("PRECIO_VENTA")
            Printer.CurrentY = POSY
            Printer.CurrentX = 10000
            Printer.Print Format(CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANTIDAD")), "###,###,##0.00")
            POSY = POSY + 200
            tRs.MoveNext
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
                Printer.Print "     PRESTAMO DE MERCANCIA A NOMBRE DE : " & tRs.Fields("NOMBRE")
                Printer.Print "     FECHA DE PRESTAMO : " & tRs.Fields("FECHA_PRESTAMO")
                If tRs.Fields("FECHA_ENTREGA") <> "01/01/1900" Then
                    Printer.Print "     FECHA DE ENTREGA : " & tRs.Fields("FECHA_ENTREGA")
                End If
                If tRs.Fields("DEPOSITO") <> 0 Then
                    Printer.Print "     DEPOSITO : " & tRs.Fields("DEPOSITO")
                End If
                If tRs.Fields("NOTAS") <> "" Then
                    Printer.Print "     NOTAS : " & tRs.Fields("NOTAS")
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
                Printer.CurrentX = 1900
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 7700
                Printer.Print "CANTIDAD"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8800
                Printer.Print "P/UNITARIO"
                Printer.CurrentY = POSY
                Printer.CurrentX = 10000
                POSY = POSY + 400
                Printer.Print "IMPORTE"
            End If
        Loop
         Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
         Printer.Print "      FIN DEL LISTADO"
         Printer.EndDoc
         CommonDialog1.Copies = 1
    End If
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text6.Text = Item
    Text5.Text = Item.SubItems(1)
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text3.Text = Item
    CanExi = Item.SubItems(2)
    PVent = Item.SubItems(3)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IndItm = Item.Index
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoFolioElim = Item
    Actualiza
End Sub
Private Sub ListView5_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoElim = Item
    Text10.Text = Item.SubItems(1)
    Text9.Text = Item.SubItems(3)
    CantPend = Item.SubItems(3)
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    BuscaPrestamo
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT NOMBRE, ID_CLIENTE, RFC FROM CLIENTE WHERE NOMBRE LIKE '%" & Text1.Text & "%'"
        Set tRs = cnn.Execute(sBuscar)
        ListView1.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_CLIENTE"))
                If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(2) = tRs.Fields("RFC")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        If Option1.Value = True Then
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_COSTO, GANANCIA FROM VSEXISALMA3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        Else
            sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO_COSTO, GANANCIA FROM VSEXISALMA3 WHERE DESCRIPCION LIKE '%" & Text2.Text & "%' AND SUCURSAL = '" & Combo1.Text & "'"
        End If
        Set tRs = cnn.Execute(sBuscar)
        ListView2.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(1) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("PRECIO_COSTO")) And Not IsNull(tRs.Fields("GANANCIA")) Then tLi.SubItems(3) = Format(CDbl(tRs.Fields("PRECIO_COSTO")) * (CDbl(tRs.Fields("GANANCIA")) + 1), "###,###,##0.00")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub BuscaPrestamo()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_PRESTAMO, NOMBRE, FECHA_PRESTAMO, FECHA_ENTREGA, DEPOSITO, NOTAS FROM VSPRESTAMO WHERE ESTADO = 'P' and nombre like '%" & Text11.Text & "%' ORDER BY ID_PRESTAMO"
    Set tRs = cnn.Execute(sBuscar)
    ListView4.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView4.ListItems.Add(, , tRs.Fields("ID_PRESTAMO"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("FECHA_PRESTAMO")) Then tLi.SubItems(2) = tRs.Fields("FECHA_PRESTAMO")
            If Not IsNull(tRs.Fields("FECHA_ENTREGA")) Then tLi.SubItems(3) = tRs.Fields("FECHA_ENTREGA")
            If Not IsNull(tRs.Fields("DEPOSITO")) Then tLi.SubItems(4) = tRs.Fields("DEPOSITO")
            If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(5) = tRs.Fields("NOTAS")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT PRESTAMOS_CLIENTES_DETALLE.ID_PRODUCTO, PRODUCTOS_CONSUMIBLES.DESCRIPCION, PRESTAMOS_CLIENTES_DETALLE.CANTIDAD, PRESTAMOS_CLIENTES_DETALLE.PRECIO_VENTA, PRESTAMOS_CLIENTES_DETALLE.ID_PRESTAMO, PRESTAMOS_CLIENTES_DETALLE.ID, PRESTAMOS_CLIENTES_DETALLE.NO_SERIE, PRESTAMOS_CLIENTES_DETALLE.CONT_INICIAL FROM PRESTAMOS_CLIENTES_DETALLE INNER JOIN PRODUCTOS_CONSUMIBLES ON PRESTAMOS_CLIENTES_DETALLE.ID_PRODUCTO = PRODUCTOS_CONSUMIBLES.ID_PRODUCTO WHERE PRESTAMOS_CLIENTES_DETALLE.ID_PRESTAMO = " & NoFolioElim
    Set tRs = cnn.Execute(sBuscar)
    ListView5.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID"))
            If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(2) = tRs.Fields("Descripcion")
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
            tRs.MoveNext
        Loop
    Else
        sBuscar = "SELECT PRESTAMOS_CLIENTES_DETALLE.ID_PRODUCTO, PRESTAMOS_CLIENTES_DETALLE.CANTIDAD, PRESTAMOS_CLIENTES_DETALLE.PRECIO_VENTA, PRESTAMOS_CLIENTES_DETALLE.ID_PRESTAMO, PRESTAMOS_CLIENTES_DETALLE.ID, PRESTAMOS_CLIENTES_DETALLE.NO_SERIE, PRESTAMOS_CLIENTES_DETALLE.CONT_INICIAL, ALMACEN3.Descripcion FROM PRESTAMOS_CLIENTES_DETALLE INNER JOIN ALMACEN3 ON PRESTAMOS_CLIENTES_DETALLE.ID_PRODUCTO = ALMACEN3.ID_PRODUCTO WHERE ID_PRESTAMO = " & NoFolioElim
        Set tRs = cnn.Execute(sBuscar)
        ListView5.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not (tRs.EOF)
                Set tLi = ListView5.ListItems.Add(, , tRs.Fields("ID"))
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
                If Not IsNull(tRs.Fields("Descripcion")) Then tLi.SubItems(2) = tRs.Fields("Descripcion")
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
                If Not IsNull(tRs.Fields("PRECIO_VENTA")) Then tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
