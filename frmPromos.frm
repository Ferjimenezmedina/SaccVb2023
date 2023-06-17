VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPromos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Promociónes"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9000
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8880
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   14
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmPromos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmPromos.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab sstPromos 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "Precios de Oferta"
      TabPicture(0)   =   "frmPromos.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescripcion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpFechaFin"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtId_Producto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lvwProductos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBuscar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdGuardar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPrecioOferta"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Ofertas por Clave"
      TabPicture(1)   =   "frmPromos.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ListView1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DTPicker1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Borrar"
      TabPicture(2)   =   "frmPromos.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "cmdBorrar"
      Tab(2).Control(2)=   "txtId_Promo"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -69120
         TabIndex        =   37
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
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
         Left            =   -69000
         Picture         =   "frmPromos.frx":2440
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3000
         Width           =   1215
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
         Left            =   -71760
         Picture         =   "frmPromos.frx":4E12
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -74760
         TabIndex        =   34
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
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
         Left            =   -70320
         Picture         =   "frmPromos.frx":77E4
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1500
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
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
         Left            =   -70320
         Picture         =   "frmPromos.frx":A1B6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1980
         Width           =   255
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4800
         TabIndex        =   29
         Top             =   1020
         Width           =   3735
         Begin VB.OptionButton Option1 
            Caption         =   "Categoria"
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Producto"
            Height          =   255
            Left            =   1800
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4800
         TabIndex        =   26
         Top             =   600
         Width           =   3615
         Begin VB.OptionButton Option4 
            Caption         =   "Por Porcentaje"
            Height          =   255
            Left            =   1680
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Precio"
            Height          =   195
            Left            =   480
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6600
         TabIndex        =   25
         Top             =   2940
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   6600
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   7200
         TabIndex        =   20
         Top             =   2580
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command5 
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
         Left            =   6000
         Picture         =   "frmPromos.frx":CB88
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6000
         TabIndex        =   18
         Top             =   3540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   39679
      End
      Begin VB.TextBox txtId_Promo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -74640
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
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
         Left            =   -68160
         Picture         =   "frmPromos.frx":F55A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -74640
         TabIndex        =   10
         Top             =   600
         Width           =   7815
         Begin MSComctlLib.ListView lvwPromos 
            Height          =   2895
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   5106
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtPrecioOferta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   2220
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton cmdGuardar 
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
         Left            =   6000
         Picture         =   "frmPromos.frx":11F2C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
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
         Left            =   3240
         Picture         =   "frmPromos.frx":148FE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwProductos 
         Height          =   3135
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtId_Producto 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   840
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   30
         Left            =   5040
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   53
         _Version        =   393216
         Format          =   50528256
         CurrentDate     =   38925
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3195
         Left            =   -74760
         TabIndex        =   38
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5636
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -69120
         TabIndex        =   39
         Top             =   2340
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   38925
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Precio Ofertado"
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
         Left            =   -69840
         TabIndex        =   42
         Top             =   1260
         Width           =   3015
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Fecha final de la promoción"
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
         Left            =   -69840
         TabIndex        =   41
         Top             =   2100
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "GUARDANDO...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -69960
         TabIndex        =   40
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Categoría"
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
         Left            =   5280
         TabIndex        =   24
         Top             =   2940
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Porcentaje"
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
         Left            =   5400
         TabIndex        =   23
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Cantidad En Producto"
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
         Left            =   4920
         TabIndex        =   21
         Top             =   2580
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Precio de Oferta"
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
         Left            =   5160
         TabIndex        =   9
         Top             =   1980
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha final de la promoción"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   3300
         Width           =   3015
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmPromos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim Cont As Integer
Dim NoRe As Integer
Private Sub cmdBorrar_Click()
On Error GoTo ManejaError
    If Puede_Borrar Then
        sqlQuery = "DELETE PROMOCION WHERE ID_PROMOCION = '" & Me.txtId_Promo.Text & "'"
        Set tRs = cnn.Execute(sqlQuery)
        Me.Llenar_Lista_Promociones
        MsgBox "PROMOCIÓN ELIMINADA", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdBuscar_Click()
On Error GoTo ManejaError
    Llenar_Lista_Productos (Trim(Me.txtId_Producto.Text))
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub CmdGuardar_Click()
On Error GoTo ManejaError
    Text4.Text = Me.txtPrecioOferta.Text
    sqlQuery = "INSERT INTO PROMO_PORCE(ID_PRODUCTO,CANTIDAD,PORCE_DESC,FECHA_VENCE) VALUES ('" & txtId_Producto.Text & "', '" & Text5.Text & "', '" & Text6.Text & "','" & DTPicker2.Value & "');"
    cnn.Execute (sqlQuery)
    MsgBox "PROMOCIÓN GUARDADA", vbInformation, "SACC"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
    Dim Cont As Integer
    Dim Guardado As Boolean
    If Text1.Text <> "" Then
        Guardado = False
        Label8.Visible = False
        For Cont = 1 To ListView1.ListItems.Count
            If ListView1.ListItems.Item(Cont).Checked Then
                sqlQuery = "INSERT INTO PROMOCION (ID_PRODUCTO, PRECIO_OFERTA, FECHA_FIN, TIPO, ESTADO_ACTUAL) VALUES('" & ListView1.ListItems.Item(Cont) & "', " & Replace(Text1.Text, ",", "") & ", '" & DTPicker1.Value & "', 'PROMOCION', 'A')"
                Set tRs = cnn.Execute(sqlQuery)
                Guardado = True
            End If
        Next Cont
        If Guardado Then
            MsgBox "PROMOCIÓN GUARDADA", vbInformation, "SACC"
        End If
        Label8.Visible = False
    End If
End Sub
Private Sub Command2_Click()
    If Text2.Text <> "" Then
        ListView1.ListItems.Clear
        sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        Set tRs = cnn.Execute(sqlQuery)
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
                    tLi.SubItems(1) = tRs.Fields("Descripcion")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub Command3_Click()
    Dim Cont As Integer
    For Cont = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(Cont).Checked = True
    Next Cont
End Sub
Private Sub Command4_Click()
    Dim Cont As Integer
    For Cont = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(Cont).Checked = False
    Next Cont
End Sub
Private Sub Command5_Click()
On Error GoTo ManejaError
    Text4.Text = Me.txtPrecioOferta.Text
    If Option1.Value = True Then
        sqlQuery = "INSERT INTO PROMO_CATEGO(CATEGORIA, CANTIDAD, FECHA_VENCE, PORCE_DESC) VALUES ('" & Combo1.Text & "', '" & Text5.Text & "',  '" & DTPicker2.Value & "', '" & Text6.Text & "');"
        cnn.Execute (sqlQuery)
        MsgBox "PROMOCIÓN GUARDADA", vbInformation, "SACC"
    End If
    If Option2.Value = True Then
        sqlQuery = "INSERT INTO PROMOCION(ID_PRODUCTO, PRECIO_OFERTA, FECHA_FIN, ESTADO_ACTUAL, TIPO) VALUES ('" & Text3.Text & "', '" & Text4.Text & "',  '" & DTPicker2.Value & "','A','PROMOCION' );"
        cnn.Execute (sqlQuery)
        'sqlQuery = "INSERT INTO PROMO_PORCE(ID_PRODUCTO,CANTIDAD,PORCE_DESC,FECHA_VENCE) VALUES ('" & txtId_Producto.Text & "', '" & Text5.Text & "', '" & Text6.Text & "','" & DTPicker2.Value & "');"
        'cnn.Execute (sqlQuery)
        MsgBox "PROMOCIÓN GUARDADA", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Activate()
On Error GoTo ManejaError
    Me.dtpFechaFin.Value = Now + 1
    Llenar_Lista_Promociones
    DTPicker1.Value = Format(Date + 30, "dd/mm/yyyy")
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
    dtpFechaFin = Format(Date, "dd/mm/yyyy")
    DTPicker2 = Format(Date, "dd/mm/yyyy")
    With lvwProductos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "Clave del Producto", 3000, 2
        .ColumnHeaders.Add , , "Descripcion", 0
    End With
    With lvwPromos
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PROMOCION", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200, 2
        .ColumnHeaders.Add , , "OFERTA/DESCUENTO", 1100, 2
        .ColumnHeaders.Add , , "FECHA FIN", 1800, 2
        .ColumnHeaders.Add , , "TIPO", 2200, 2
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .Checkboxes = True
        .ColumnHeaders.Add , , "Clave del Producto", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
    End With
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT CLASIFICACION FROM CLATINTONER"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.EOF) Then
        Do While Not tRs.EOF
            Combo1.AddItem tRs.Fields("CLASIFICACION")
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Productos(nId_Producto As String)
On Error GoTo ManejaError
    Me.lvwProductos.ListItems.Clear
    sqlQuery = "SELECT ID_PRODUCTO, Descripcion FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & nId_Producto & "%'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwProductos.ListItems.Add(, , "")
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(2) = Trim(.Fields("Descripcion"))
                .MoveNext
            Loop
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
Private Sub lvwProductos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtId_Producto.Text = Item.SubItems(1)
    Me.lblDescripcion.Caption = Item.SubItems(2)
    Text3.Text = Item.SubItems(1)
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwPromos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Me.txtId_Promo.Text = Item
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Option1_Click()
    Combo1.Visible = True
        Label12.Visible = True
End Sub
Private Sub Option2_Click()
    Combo1.Visible = False
    Label12.Visible = False
End Sub
Private Sub Option3_Click()
    If Option3.Value Then
        Command5.Visible = True
        Option1.Visible = True
        Option2.Visible = True
        Text5.Visible = True
        Text6.Visible = False
        Combo1.Visible = True
        Label10.Visible = True
        Label11.Visible = False
        Label12.Visible = True
        Label2.Visible = True
        cmdGuardar.Visible = False
        txtPrecioOferta.Visible = True
        Option1.Value = True
    End If
End Sub
Private Sub Option4_Click()
    If Option4.Value Then
        cmdGuardar.Visible = False
        Command5.Visible = False
        Option1.Visible = False
        Option2.Visible = False
        Text5.Visible = False
        Text6.Visible = True
        Combo1.Visible = False
        Label10.Visible = False
        Label11.Visible = True
        Label12.Visible = False
        Label2.Visible = False
        txtPrecioOferta.Visible = False
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890."
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.Value = True
    End If
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ *%_"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtId_Producto_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Llenar_Lista_Productos (Trim(Me.txtId_Producto.Text))
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtPrecioOferta_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    Else
        Dim Valido As String
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Guardar() As Boolean
On Error GoTo ManejaError
    Puede_Guardar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Promociones()
On Error GoTo ManejaError
    sqlQuery = "SELECT ID_PROMOCION, ID_PRODUCTO, PRECIO_OFERTA, FECHA_FIN, TIPO FROM PROMOCION"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwPromos.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwPromos.ListItems.Add(, , .Fields("ID_PROMOCION"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("PRECIO_OFERTA")) Then tLi.SubItems(2) = Trim(.Fields("PRECIO_OFERTA"))
                If Not IsNull(.Fields("FECHA_FIN")) Then tLi.SubItems(3) = Trim(.Fields("FECHA_FIN"))
                If Not IsNull(.Fields("TIPO")) Then tLi.SubItems(4) = Trim(.Fields("TIPO"))
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Puede_Borrar() As Boolean
On Error GoTo ManejaError
    If Me.txtId_Promo.Text = "" Then
        MsgBox "SELECCIONE LA PROMOCION QUE VA A BORRAR", vbInformation, "SACC"
        Puede_Borrar = False
        Exit Function
    End If
    Puede_Borrar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
