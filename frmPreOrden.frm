VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPreOrden 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRE-ORDEN DE COMPRA"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   49
      Top             =   2160
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
         TabIndex        =   50
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmPreOrden.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmPreOrden.frx":030A
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   47
      Top             =   3360
      Width           =   975
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   48
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmPreOrden.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "frmPreOrden.frx":1FD6
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   42
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10560
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtotros 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   37
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10440
      Picture         =   "frmPreOrden.frx":3BA8
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtflete 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   31
      Text            =   "0"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9720
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   28
      Top             =   4560
      Width           =   975
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modificar"
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
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "frmPreOrden.frx":657A
         MousePointer    =   99  'Custom
         Picture         =   "frmPreOrden.frx":6884
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   24
      Top             =   5760
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmPreOrden.frx":89AE
         MousePointer    =   99  'Custom
         Picture         =   "frmPreOrden.frx":8CB8
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
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox txtCantOrg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9840
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cambiar"
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
      Left            =   6600
      Picture         =   "frmPreOrden.frx":AD9A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtCant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   21
      Top             =   6690
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8760
      Picture         =   "frmPreOrden.frx":D76C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   975
   End
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
      Left            =   7680
      Picture         =   "frmPreOrden.frx":1013E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtId_Proveedor 
      Height          =   285
      Left            =   10080
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   3000
         TabIndex        =   51
         Top             =   2400
         Width           =   1695
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmPreOrden.frx":12B10
            Left            =   120
            List            =   "frmPreOrden.frx":12B23
            TabIndex        =   54
            Text            =   "15"
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contado"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Credito"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   195
         Left            =   3000
         TabIndex        =   46
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox txtcomen 
         Height          =   285
         Left            =   1560
         TabIndex        =   40
         Text            =   "0"
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   195
         Left            =   3000
         TabIndex        =   35
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox txtMoneda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton opnIndirecta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Indirecta"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton opnInternacional 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Internacional"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton opnNacional 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nacional"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtImpuesto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "0"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtSubtotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblProveedor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DESCUENTO"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMENTARIO"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OTROS CARGOS"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FLETE"
         Height          =   375
         Left            =   840
         TabIndex        =   34
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "MONEDA"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblFolio 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "IMPUESTO"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "SUBTOTAL"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwProveedores 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwCotizaciones 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   9840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Label LblTipoOrden 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9840
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "FLETE"
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "FLETE"
      Height          =   255
      Left            =   5280
      TabIndex        =   32
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   9840
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9840
      TabIndex        =   27
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblidprod 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9840
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblIndex 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSelec 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6600
      Width           =   5415
   End
End
Attribute VB_Name = "frmPreOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim tRs As ADODB.Recordset
Dim tRs1 As ADODB.Recordset
Dim Cont As Integer
Dim NoRe As Integer
Dim CanMod As String
Dim ProdMod As String
Dim PresMod As String
Dim IdCotiza As String
Dim orden As Integer
Dim IdProv As String
Dim IdOrden As String
Dim idcoti As String
Dim ordennum As Integer
Dim totgen As Double
Dim totiva     As Double
Dim numordena As Integer
Private Sub SUM()
    Me.txtTotal.Text = CDbl(Me.txtSubtotal.Text) + CDbl(Me.txtFlete.Text) + CDbl(Me.txtImpuesto.Text) + CDbl(Me.txtotros.Text)
End Sub
Private Sub Command1_Click()
On Error GoTo MANEJAERR
    Dim IDC As String
    Dim sqlQuery As String
    If lblIndex.Caption <> "" Then
        IDC = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption))
        If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9) <> "" Then
            IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9)
        End If
        If lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(10) <> "" Then
            sqlQuery = "DELETE FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(3) & "' AND ID_ORDEN_COMPRA = " & lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(10)
            cnn.Execute (sqlQuery)
            sqlQuery = "UPDATE COTIZA_REQUI SET NUMOC = 0 WHERE ID_COTIZACION IN (" & IDC & ") AND ID_PRODUCTO = '" & lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(3) & "'"
            cnn.Execute (sqlQuery)
        End If
        lvwCotizaciones.ListItems.Remove (Val(lblIndex.Caption))
        lblIndex.Caption = ""
        lblSelec.Caption = ""
    End If
    Sumar_Importe
MANEJAERR:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim IDC As String
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    If lblIndex.Caption <> "" Then
        If MsgBox("SEGURO QUE DESEA ELIMINAR PERMANENTEMENTE EL PRODUCTO " & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(3) & " DE ESTA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
            IDC = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption))
            If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9) <> "" Then
                IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9)
            End If
            If IdOrden <> "" Then
                ' ESTE ES EL DELETE QUE MENCIONAS ABAJO?? DONDE ELIMINAS EL PRODUCTO DE LA TABLA??
                ' SI LO TENIA... SOLO FALTABA PONER UN RECALCULO DE LOS TOTALES!
                sqlQuery = "DELETE FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & lblidprod.Caption & "' AND ID_ORDEN_COMPRA = " & IdOrden
                cnn.Execute (sqlQuery)
                ' SE AGREGO PARA ELIMINAR DE LA COTIZACION EL PRODUCTO
                ' 26 JUNIO 2009
                sqlQuery = "DELETE FROM COTIZA_REQUI WHERE ID_PRODUCTO = '" & lblidprod.Caption & "' AND ID_PROVEEDOR = '" & txtId_Proveedor & " ' AND ESTADO_ACTUAL = 'X'"
                cnn.Execute (sqlQuery)
            End If
            sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'Z' WHERE ID_COTIZACION IN (" & IDC & ") AND ID_PRODUCTO = '" & lblidprod.Caption & "'"
            cnn.Execute (sqlQuery)
            ' ESTO ES LO QUE AGREGUE!! QUE RECALCULE EL TOTAL DE LA ORDEN DE COMPRA Y LO GUARDE DESPUES DE LA ELIMINACION
            ' LA ORDEN SACA EL TOTAL DE LA TABLA DE ORDEN DE COMPRA Y NO DE LA DE REQUIS...
            If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) <> "" Then
                sqlQuery = "SELECT SUM (CANTIDAD * PRECIO) AS TOTAL FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = '" & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) & "'"
                Set tRs = cnn.Execute(sqlQuery)
                sqlQuery = "UPDATE ORDEN_COMPRA SET TOTAL = " & tRs.Fields("TOTAL") & " WHERE ID_ORDEN_COMPRA = '" & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) & "'"
                cnn.Execute (sqlQuery)
            End If
            If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) <> "" Then
                lvwCotizaciones.ListItems.Remove (Val(lblIndex.Caption))
                lblIndex.Caption = ""
                lblSelec.Caption = ""
            End If
        End If
    End If
    Sumar_Importe
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function NumPalabras(Cade As String) As Integer
    Dim pos As Integer
    Dim Cont As Integer
    pos = 1
    Cont = 1
    While InStr(pos, Cade, ",") > 0
        pos = InStr(pos, Cade, ",") + 1
        Cont = Cont + 1
    Wend
    NumPalabras = Cont
End Function
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim Total As Double
    Dim TAX As Double
    Dim Almacen As String
    Dim Marca As String
    Dim IDC As String
    Dim sqlQuery As String
    Dim Cont As Integer
    Dim pos As Integer
    Dim cant As Double
    Dim Aux1 As String
    Dim Aux2 As String
    Dim IdRequi As String
    If CDbl(CanMod) < CDbl(txtCant.Text) Then
        MsgBox "NO ES POSIBLE COMPRAR CANTIDADES MAYORES A LAS SOLICITADAS, TOME PRODUCTOS DE REQUISICION PARA COMPLETAR SU PEDIDO", vbExclamation, "SACC"
    Else
        If lblIndex.Caption <> "" Then
            If Command4.Caption = "Cambiar" Then
                txtCant.Enabled = True
                Command4.Caption = "Guardar"
                If VarMen.Text1(77).Text = "N" Then
                    Command1.Enabled = False
                    Command2.Enabled = False
                End If
                Frame8.Enabled = False
                Frame2.Enabled = False
                lvwCotizaciones.Enabled = False
                lvwProveedores.Enabled = False
                txtCantOrg.Text = txtCant.Text
            ElseIf txtCant.Text <> "" Then
                txtCant.Enabled = False
                txtCant.Enabled = False
                Command4.Caption = "Cambiar"
                If VarMen.Text1(77).Text = "S" Then
                    Command1.Enabled = True
                    Command2.Enabled = True
                End If
                Frame8.Enabled = True
                Frame2.Enabled = True
                lvwCotizaciones.Enabled = True
                lvwProveedores.Enabled = True
                pos = 1
                Aux2 = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption))
                IdRequi = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(1)
                cant = CDbl(Replace(txtCantOrg.Text, ",", "")) - CDbl(Replace(txtCant.Text, ",", ""))
                For Cont = 1 To NumPalabras(lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)))
                    'If InStr(1, Aux2, ",") = 0 Then
                    Aux1 = Aux2
                    'Else
                    '    Aux1 = Mid(Aux2, 1, InStr(1, Aux2, ",") - 1)
                    'End If
                    pos = InStr(1, Aux2, ",") + 1
                    sqlQuery = "SELECT ID_COTIZACION, CANTIDAD FROM COTIZA_REQUI WHERE ID_COTIZACION IN (" & Aux1 & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                    Set tRs = cnn.Execute(sqlQuery)
                    If Not (tRs.BOF And tRs.EOF) Then
                        cant = CDbl(txtCant.Text)
                        Do While Not tRs.EOF
                            'If cant > CDbl(tRs.Fields("CANTIDAD")) Then
                            '    sqlQuery = "UPDATE COTIZA_REQUI SET CANTIDAD = 0 WHERE ID_COTIZACION IN (" & Aux1 & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                            '    cant = cant - CDbl(tRs.Fields("CANTIDAD"))
                            'Else
                                sqlQuery = "UPDATE COTIZA_REQUI SET CANTIDAD = " & cant & " WHERE ID_COTIZACION IN (" & tRs.Fields("ID_COTIZACION") & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                                cant = 0
                            'End If
                            cnn.Execute (sqlQuery)
                            'Aux2 = Right(Aux2, Len(Aux2) - pos)
                            tRs.MoveNext
                        Loop
                    End If
                Next Cont
                'If CDbl(txtCantOrg.Text) > CDbl(txtCant.Text) Then
                sBuscar = "SELECT ID_REQUISICION, URGENTE, COMENTARIO FROM REQUISICION WHERE ACTIVO = 0 AND ID_PRODUCTO = '" & ProdMod & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.BOF And tRs.EOF) Then
                    sBuscar = "SELECT Descripcion, MARCA FROM ALMACEN3 WHERE ID_PRODUCTO = '" & ProdMod & "'"
                    Set tRs1 = cnn.Execute(sBuscar)
                    Almacen = "A3"
                    If tRs1.BOF And tRs1.EOF Then
                        sBuscar = "SELECT Descripcion, MARCA FROM ALMACEN2 WHERE ID_PRODUCTO = '" & ProdMod & "'"
                        Set tRs1 = cnn.Execute(sBuscar)
                        Almacen = "A2"
                        If tRs1.BOF And tRs1.EOF Then
                            sBuscar = "SELECT Descripcion, MARCA FROM ALMACEN1 WHERE ID_PRODUCTO = '" & lblidprod.Caption & "'"
                            Set tRs1 = cnn.Execute(sBuscar)
                            Label10.Caption = tRs1.Fields("Descripcion")
                            Marca = tRs1.Fields("MARCA")
                            Almacen = "A1"
                        Else
                            Label10.Caption = tRs1.Fields("Descripcion")
                            Marca = tRs1.Fields("MARCA")
                        End If
                    Else
                        Label10.Caption = tRs1.Fields("Descripcion")
                        Marca = tRs1.Fields("MARCA")
                    End If
                    sBuscar = "INSERT INTO REQUISICION (FECHA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, CONTADOR, ALMACEN, MARCA, URGENTE, COMENTARIO) Values('" & Format(Date, "dd/mm/yyyy") & "', '" & lblidprod.Caption & "', '" & Label10.Caption & "', " & Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) & ", 0, '" & Almacen & "', '" & Marca & "', 'N', 'PEDIDO COMPLEMENTARIO POR MODIFICACION DE CANTIDADES') "
                    sBuscar2 = ""
                    If Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) = 0 Then
                        IDC = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption))
                        If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9) <> "" Then
                            IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9)
                        End If
                        If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) <> "" Then
                            sqlQuery = "DELETE FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(3) & "' AND ID_ORDEN_COMPRA = " & vwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10)
                            cnn.Execute (sqlQuery)
                        End If
                        sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'Z' WHERE ID_COTIZACION IN (" & IDC & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                        cnn.Execute (sqlQuery)
                        lvwCotizaciones.ListItems.Remove (Val(lblIndex.Caption))
                    End If
                Else
                    sBuscar = "UPDATE REQUISICION SET CANTIDAD = " & Val(Replace(txtCant.Text, ",", "")) & " WHERE ID_REQUISICION IN (" & IdRequi & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                    sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & Replace(IdRequi, ",", ", '" & VarMen.Text1(1).Text & "'  ," & txtCant & ", 'REQUISICION HECHA DESDE LA PRE-ORDEN POR DEVOLUCION DE MATERIAL','" & Format(Date, "dd/mm/yyyy") & "'), (") & ", '" & VarMen.Text1(1).Text & "'  ," & txtCant & ", 'REQUISICION HECHA DESDE LA PRE-ORDEN POR DEVOLUCION DE MATERIAL','" & Format(Date, "dd/mm/yyyy") & "')"
                    If Val(Replace(txtCantOrg.Text, ",", "")) - Val(Replace(txtCant.Text, ",", "")) = 0 Then
                        IDC = lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption))
                        If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9) <> "" Then
                            IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(9)
                        End If
                        If lvwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10) <> "" Then
                            sqlQuery = "DELETE FROM ORDEN_COMPRA_DETALLE WHERE ID_PRODUCTO = '" & lvwCotizaciones.ListItems.Item(CDbl(lblIndex.Caption)).SubItems(3) & "' AND ID_ORDEN_COMPRA = " & vwCotizaciones.ListItems.Item(Val(lblIndex.Caption)).SubItems(10)
                            cnn.Execute (sqlQuery)
                        End If
                        sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'Z' WHERE ID_COTIZACION IN (" & IDC & ") AND ID_PRODUCTO = '" & ProdMod & "'"
                        cnn.Execute (sqlQuery)
                        lvwCotizaciones.ListItems.Remove (Val(lblIndex.Caption))
                    End If
                End If
                tRs.Close
                cnn.Execute (sBuscar)
                If sBuscar2 = "" Then
                    sBuscar = "SELECT ID_REQUISICION FROM REQUISICION ORDER BY ID_REQUISICION DESC"
                    Set tRs = cnn.Execute(sBuscar)
                    sBuscar2 = "INSERT INTO RASTREOREQUI (ID_REQUI, SOLICITO, CANTIDAD, COMENTARIO,FECHA) Values(" & tRs.Fields("ID_REQUISICION") & ", '" & VarMen.Text1(1).Text & "'  ," & txtCant & ", 'REQUISICION HECHA DESDE LA PRE-ORDEN POR DEVOLUCION DE MATERIAL','" & Format(Date, "dd/mm/yyyy") & "')"
                    tRs.Close
                End If
                cnn.Execute (sBuscar2)
                'End If
                Llenar_Lista_Cotizaciones (Val(txtId_Proveedor.Text))
                Sumar_Importe
                Label10.Caption = ""
            Else
                MsgBox "DEBE PONER UNA CANTIDAD", vbCritical, "SACC"
            End If
        End If
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command5_Click()
    txtFlete.Enabled = True
    txtotros.Enabled = True
    Text3.Enabled = True
End Sub
Private Sub Command6_Click()
    Me.txtTotal.Text = CDbl(Me.txtSubtotal.Text) + CDbl(Me.txtFlete.Text) + CDbl(Me.txtImpuesto.Text) + CDbl(Me.txtotros.Text)
End Sub
Private Sub Command7_Click()
    txtImpuesto.Enabled = True
End Sub
Private Sub Form_Activate()
    If Hay_Proveedores Then
        Llenar_Lista_Proveedores
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
    With Me.lvwCotizaciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID COTIZACION", 0
        .ColumnHeaders.Add , , "ID REQUISICION", 0
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "DIAS ENTREGA", 1440, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "FECHA", 0, 2
        .ColumnHeaders.Add , , "IDS", 0
        .ColumnHeaders.Add , , "NUMOC", 100
        .ColumnHeaders.Add , , "MONEDA", 100
        .ColumnHeaders.Add , , "No PEDIDO", 100
    End With
    With Me.lvwProveedores
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID PROVEEDOR", 0
        .ColumnHeaders.Add , , "PROVEEDOR", 4500, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
        .ColumnHeaders.Add , , "", 0, 2
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Proveedores()
On Error GoTo ManejaError
    'Dim nId_Proveedor As Integer
    Dim sBuscar As String
    sBuscar = "SELECT PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, PROVEEDOR.DIRECCION, PROVEEDOR.COLONIA, PROVEEDOR.CIUDAD, PROVEEDOR.CP, PROVEEDOR.RFC, PROVEEDOR.TELEFONO1, PROVEEDOR.TELEFONO2, Proveedor.TELEFONO3 , Proveedor.NOTAS, Proveedor.Estado, Proveedor.PAIS FROM PROVEEDOR INNER JOIN COTIZA_REQUI ON PROVEEDOR.ID_PROVEEDOR = COTIZA_REQUI.ID_PROVEEDOR WHERE (PROVEEDOR.ELIMINADO = 'N') AND (COTIZA_REQUI.ESTADO_ACTUAL = 'X') AND (COTIZA_REQUI.ID_PRODUCTO <> '') GROUP BY PROVEEDOR.ID_PROVEEDOR, PROVEEDOR.NOMBRE, PROVEEDOR.DIRECCION, PROVEEDOR.COLONIA, PROVEEDOR.CIUDAD, PROVEEDOR.CP, PROVEEDOR.RFC, PROVEEDOR.TELEFONO1, PROVEEDOR.TELEFONO2, PROVEEDOR.TELEFONO3, PROVEEDOR.NOTAS, PROVEEDOR.ESTADO, PROVEEDOR.PAIS ORDER BY PROVEEDOR.NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.BOF And tRs.EOF) Then
        Me.lvwProveedores.ListItems.Clear
        Do While Not tRs.EOF
            'If nId_Proveedor <> tRs.Fields("ID_PROVEEDOR") Then
            Set tLi = lvwProveedores.ListItems.Add(, , tRs.Fields("ID_PROVEEDOR"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = Trim(tRs.Fields("NOMBRE"))
            If Not IsNull(tRs.Fields("DIRECCION")) Then tLi.SubItems(2) = Trim(tRs.Fields("DIRECCION"))
            If Not IsNull(tRs.Fields("COLONIA")) Then tLi.SubItems(3) = Trim(tRs.Fields("COLONIA"))
            If Not IsNull(tRs.Fields("CP")) Then tLi.SubItems(4) = Trim(tRs.Fields("CP"))
            If Not IsNull(tRs.Fields("RFC")) Then tLi.SubItems(5) = Trim(tRs.Fields("RFC"))
            If Not IsNull(tRs.Fields("TELEFONO1")) Then tLi.SubItems(6) = Trim(tRs.Fields("TELEFONO1"))
            If Not IsNull(tRs.Fields("TELEFONO2")) Then tLi.SubItems(7) = Trim(tRs.Fields("TELEFONO2"))
            If Not IsNull(tRs.Fields("TELEFONO3")) Then tLi.SubItems(8) = Trim(tRs.Fields("TELEFONO3"))
            If Not IsNull(tRs.Fields("NOTAS")) Then tLi.SubItems(9) = Trim(tRs.Fields("NOTAS"))
            If Not IsNull(tRs.Fields("ESTADO")) Then tLi.SubItems(10) = Trim(tRs.Fields("ESTADO"))
            If Not IsNull(tRs.Fields("PAIS")) Then tLi.SubItems(11) = Trim(tRs.Fields("PAIS"))
            'INICIO PARA NO REPETIR PROVEEDORES
            'nId_Proveedor = tRs.Fields("ID_PROVEEDOR")
            'End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Llenar_Lista_Cotizaciones(nId_Proveedor As Integer)
On Error GoTo ManejaError
    Dim Cont As Integer
    Dim CONT2 As Integer
    'V3
    sqlQuery = "SELECT NUMOC, STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_COTIZACION) FROM COTIZA_REQUI CR WHERE CR.FOLIO = RE.FOLIO  AND (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') FOR XML PATH('')), 1, 1, '') AS  ID_COTIZACION, STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM COTIZA_REQUI CR WHERE CR.FOLIO  = RE.FOLIO  AND (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') FOR XML PATH('')), 1, 1, '') AS  ID_REQUISICION,  ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, SUM(CANTIDAD) AS CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ISNULL(MONEDA, '') AS MONEDA, NO_PEDIDO FROM COTIZA_REQUI RE WHERE (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') AND CANTIDAD <> 0 GROUP BY NUMOC, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, DIAS_ENTREGA, PRECIO, FECHA, FOLIO, MONEDA, NO_PEDIDO"
    'V2
    'sqlQuery = "SELECT NUMOC, STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_COTIZACION) FROM COTIZA_REQUI CR WHERE CR.FOLIO = RE.FOLIO  AND (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') FOR XML PATH('')), 1, 1, '') AS  ID_COTIZACION, STUFF(( SELECT ',' + CONVERT(VARCHAR(20), ID_REQUISICION) FROM COTIZA_REQUI CR WHERE CR.FOLIO  = RE.FOLIO  AND (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') FOR XML PATH('')), 1, 1, '') AS  ID_REQUISICION,  ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ISNULL(MONEDA, '') AS MONEDA, NO_PEDIDO FROM COTIZA_REQUI RE WHERE (ESTADO_ACTUAL = 'X') AND (ID_PROVEEDOR = " & nId_Proveedor & ") AND (ID_PRODUCTO <> '') GROUP BY NUMOC, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, FOLIO, MONEDA, NO_PEDIDO"
    'V1
    'sqlQuery = "SELECT NUMOC, ID_COTIZACION, ID_REQUISICION, ID_PROVEEDOR, ID_PRODUCTO, DESCRIPCION, CANTIDAD, DIAS_ENTREGA, PRECIO, FECHA, ISNULL(MONEDA, '') AS MONEDA, NO_PEDIDO FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'X' AND ID_PROVEEDOR = " & nId_Proveedor & " AND (ID_PRODUCTO <> '')"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        Me.lvwCotizaciones.ListItems.Clear
        If Not (.BOF And .EOF) Then
            Do While Not .EOF
                Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_COTIZACION"))
                If Not IsNull(.Fields("ID_REQUISICION")) Then tLi.SubItems(1) = Trim(.Fields("ID_REQUISICION"))
                If Not IsNull(.Fields("ID_PROVEEDOR")) Then tLi.SubItems(2) = Trim(.Fields("ID_PROVEEDOR"))
                If Not IsNull(.Fields("ID_PRODUCTO")) Then tLi.SubItems(3) = Trim(.Fields("ID_PRODUCTO"))
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(4) = Trim(.Fields("Descripcion"))
                If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(5) = Trim(.Fields("CANTIDAD"))
                If Not IsNull(.Fields("DIAS_ENTREGA")) Then tLi.SubItems(6) = Trim(.Fields("DIAS_ENTREGA"))
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(7) = Trim(.Fields("PRECIO"))
                If Not IsNull(.Fields("FECHA")) Then tLi.SubItems(8) = Trim(.Fields("FECHA"))
                If Not IsNull(.Fields("NUMOC")) Then
                    If .Fields("NUMOC") <> "0" Then
                        tLi.SubItems(10) = Trim(.Fields("NUMOC"))
                    End If
                End If
                If Not IsNull(.Fields("MONEDA")) Then tLi.SubItems(11) = Trim(.Fields("MONEDA"))
                If Not IsNull(.Fields("NO_PEDIDO")) Then tLi.SubItems(12) = Trim(.Fields("NO_PEDIDO"))
                .MoveNext
            Loop
        End If
    End With
    If lvwCotizaciones.ListItems.Count > 0 Then
        txtMoneda.Text = lvwCotizaciones.ListItems.Item(1).SubItems(11)
    End If
    For Cont = 1 To lvwCotizaciones.ListItems.Count
        If lvwCotizaciones.ListItems.Item(Cont).SubItems(3) <> "" Then
            For CONT2 = Cont + 1 To lvwCotizaciones.ListItems.Count
                If lvwCotizaciones.ListItems.Item(Cont).SubItems(3) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(3) Then
                   lvwCotizaciones.ListItems.Item(Cont).SubItems(5) = Val(lvwCotizaciones.ListItems.Item(Cont).SubItems(5)) + Val(lvwCotizaciones.ListItems.Item(CONT2).SubItems(5))
                   If Val(lvwCotizaciones.ListItems.Item(Cont).SubItems(6)) < Val(lvwCotizaciones.ListItems.Item(CONT2).SubItems(6)) Then
                        lvwCotizaciones.ListItems.Item(Cont).SubItems(6) = lvwCotizaciones.ListItems.Item(CONT2).SubItems(6)
                   End If
                   lvwCotizaciones.ListItems.Item(CONT2).SubItems(3) = ""
                   If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) = "" Then
                        lvwCotizaciones.ListItems.Item(Cont).SubItems(9) = lvwCotizaciones.ListItems.Item(CONT2)
                   Else
                        lvwCotizaciones.ListItems.Item(Cont).SubItems(9) = lvwCotizaciones.ListItems.Item(Cont).SubItems(9) & "," & lvwCotizaciones.ListItems.Item(CONT2)
                   End If
                   lvwCotizaciones.ListItems.Item(Cont) = lvwCotizaciones.ListItems.Item(Cont) & ", " & lvwCotizaciones.ListItems.Item(CONT2)
                End If
            Next CONT2
        End If
    Next Cont
    Cont = 1
    Do While Cont <= lvwCotizaciones.ListItems.Count
        If lvwCotizaciones.ListItems.Item(Cont).SubItems(3) = "" Then
            lvwCotizaciones.ListItems.Remove (Cont)
        Else
            Cont = Cont + 1
        End If
    Loop
    If lvwCotizaciones.ListItems.Count > 0 Then
        If lvwCotizaciones.ListItems.Item(1).SubItems(10) <> "" Then
            sqlQuery = "SELECT NUM_ORDEN, TIPO, FORMA_PAGO FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & lvwCotizaciones.ListItems.Item(1).SubItems(10)
            Set tRs = cnn.Execute(sqlQuery)
            If Not (tRs.BOF And tRs.EOF) Then
                lblFolio.Caption = tRs.Fields("NUM_ORDEN")
                If tRs.Fields("Tipo") = "I" Then
                    opnInternacional.Value = True
                ElseIf tRs.Fields("Tipo") = "N" Then
                    opnNacional.Value = True
                Else
                    opnIndirecta.Value = True
                End If
                opnInternacional.Enabled = False
                opnNacional.Enabled = False
                opnIndirecta.Enabled = False
                If tRs.Fields("FORMA_PAGO") = "F" Then
                    Option1.Value = True
                    Combo1.Enabled = True
                Else
                    If tRs.Fields("FORMA_PAGO") = "C" Then
                        Option2.Value = True
                        Combo1.Enabled = False
                    Else
                        Option1.Value = False
                        Option2.Value = False
                        Combo1.Enabled = False
                    End If
                End If
            Else
                lblFolio.Caption = ""
                opnInternacional.Enabled = True
                opnNacional.Enabled = True
                opnIndirecta.Enabled = True
                opnInternacional.Value = False
                opnNacional.Value = False
                opnIndirecta.Value = False
                Option1.Value = False
                Option2.Value = False
                Combo1.Enabled = False
            End If
        Else
            lblFolio.Caption = ""
            opnInternacional.Enabled = True
            opnNacional.Enabled = True
            opnIndirecta.Enabled = True
            opnInternacional.Value = False
            opnNacional.Value = False
            opnIndirecta.Value = False
        End If
    End If
    
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Function Hay_Cotizaciones() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_COTIZACION) AS ID_COTIZACION FROM COTIZA_REQUI WHERE ESTADO_ACTUAL = 'X'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_COTIZACION") <> 0 Then
            Hay_Cotizaciones = True
        Else
            Hay_Cotizaciones = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Function Hay_Proveedores() As Boolean
On Error GoTo ManejaError
    sqlQuery = "SELECT COUNT(ID_PROVEEDOR)ID_PROVEEDOR FROM PROVEEDOR WHERE ELIMINADO = 'N'"
    Set tRs = cnn.Execute(sqlQuery)
    With tRs
        If .Fields("ID_PROVEEDOR") <> 0 Then
            Hay_Proveedores = True
        Else
            Hay_Proveedores = False
        End If
    End With
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub Image2_Click()
On Error GoTo ManejaError
If Puede_Guardar Then
    txtFlete.Enabled = False
    txtotros.Enabled = False
    Dim ID_PROVEEDOR As String
    Dim Total As String
    Dim TAX As String
    Dim OTROS As String
    Dim FREIGHT As String
    Dim DISCOUNT As String
    Dim COMENTARIO As String
    Dim Tipo As String
    Dim NUM_ORDENS As String
    Dim num_orden As Integer
    Dim ID_ORDEN_COMPRA As Integer
    Dim ID_PRODUCTO As String
    Dim Descripcion As String
    Dim CANTIDAD As String
    Dim Precio As String
    Dim DIAS_ENTREGA As Integer
    Dim IDC As String
    Dim nLvw As Integer
    Dim Moneda As String
    Dim FormaPago As String
    If (opnInternacional.Value = False) And (opnNacional.Value = False) And (opnIndirecta.Value = False) Then
        MsgBox "DEBE SELECCIONAR TIPO DE ORDEN PRIMERO"
    Else
        If Me.opnInternacional.Value = True Then
            Tipo = "I"
            nLvw = 1
            Moneda = "DOLARES"
        ElseIf Me.opnNacional.Value = True Then
            Tipo = "N"
            nLvw = 2
            Moneda = txtMoneda.Text
        Else
            Tipo = "X"
            nLvw = 3
            Moneda = txtMoneda.Text
        End If
            Cont = 1
            NUM_ORDENS = ""
            If lvwCotizaciones.ListItems.Count > 0 Then
                Do While (Cont <= lvwCotizaciones.ListItems.Count) And NUM_ORDENS = ""
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(10) <> "" Then
                        NUM_ORDENS = lvwCotizaciones.ListItems.Item(Cont).SubItems(10)
                    End If
                    Cont = Cont + 1
                Loop
            End If
            If NUM_ORDENS = "" Then 'NO EXISTE UNA O.C DE ESTA REQUI
                num_orden = 0
                sqlQuery = "SELECT TOP 1 NUM_ORDEN FROM ORDEN_COMPRA WHERE TIPO = '" & Tipo & "' ORDER BY NUM_ORDEN DESC"
                Set tRs = cnn.Execute(sqlQuery)
                With tRs
                    If Not (.BOF And .EOF) Then
                            If Not IsNull(.Fields("NUM_ORDEN")) Then num_orden = .Fields("NUM_ORDEN")
                        .Close
                    End If
                    num_orden = num_orden + 1
                    numordena = num_orden
                End With
                'FIN TRAER ULTIMA ORDEN DE COMPRA
                'INICIO TRER ULTIMO ID_ORDEN_COMPRA
                ID_ORDEN_COMPRA = 0
                'FIN
                If txtotros.Text = "" Then
                    txtotros.Text = "0"
                End If
                If txtFlete.Text = "" Then
                    txtFlete.Text = "0"
                End If
                Total = Replace(txtSubtotal.Text, ",", "")
                TAX = Replace(txtImpuesto.Text, ",", "")
                OTROS = Replace(txtFlete.Text, ",", "")
                FREIGHT = Replace(txtotros.Text, ",", "")
                If Option1.Value Then
                    FormaPago = "F"
                Else
                    FormaPago = "C"
                End If
                sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, TAX, CONFIRMADA, TIPO, NUM_ORDEN, MONEDA, COMENTARIO, ID_USUARIO, FORMA_PAGO, DIAS_CREDITO, FREIGHT, OTROS_CARGOS) VALUES ('" & txtId_Proveedor.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", " & TAX & ", 'N', '" & Tipo & "', " & num_orden & ", '" & Moneda & "','" & txtcomen.Text & "', " & VarMen.Text1(0).Text & ", '" & FormaPago & "', " & Combo1.Text & ", '" & txtFlete.Text & "', '" & txtotros.Text & "');"
                'If VarMen.TxtEmp(12).Text = "EXTENDIDO" Then
                '    sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, TAX, CONFIRMADA, TIPO, NUM_ORDEN, MONEDA, COMENTARIO, ID_USUARIO, FORMA_PAGO, DIAS_CREDITO, FREIGHT, OTROS_CARGOS) VALUES ('" & txtId_Proveedor.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", " & TAX & ", 'P', '" & Tipo & "', " & num_orden & ", '" & Moneda & "','" & txtcomen.Text & "', " & VarMen.Text1(0).Text & ", '" & FormaPago & "', " & Combo1.Text & ", " & txtFlete.Text & ", '" & txtotros.Text & "');"
                'Else
                '    sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, TAX, CONFIRMADA, TIPO, NUM_ORDEN, MONEDA, COMENTARIO, ID_USUARIO, FORMA_PAGO, DIAS_CREDITO, FREIGHT, OTROS_CARGOS) VALUES ('" & txtId_Proveedor.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", " & TAX & ", 'S', '" & Tipo & "', " & num_orden & ", '" & Moneda & "','" & txtcomen.Text & "', " & VarMen.Text1(0).Text & ", '" & FormaPago & "', " & Combo1.Text & ", " & txtFlete.Text & ", '" & txtotros.Text & "');"
                'End If
                cnn.Execute (sqlQuery)
                sqlQuery = "SELECT TOP 1 ID_ORDEN_COMPRA FROM ORDEN_COMPRA ORDER BY ID_ORDEN_COMPRA DESC"
                Set tRs = cnn.Execute(sqlQuery)
                With tRs
                    If Not (.BOF And .EOF) Then
                            If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ID_ORDEN_COMPRA = .Fields("ID_ORDEN_COMPRA")
                            orden = .Fields("ID_ORDEN_COMPRA")
                        .Close
                    End If
                End With
                NoRe = Me.lvwCotizaciones.ListItems.Count
                For Cont = 1 To NoRe
                    ID_PRODUCTO = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(3)
                    Descripcion = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(4)
                    CANTIDAD = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(5)
                    DIAS_ENTREGA = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(6)
                    Precio = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(7)
                    CANTIDAD = Replace(CANTIDAD, ",", "")
                    Precio = Replace(Precio, ",", "")
                    sqlQuery = "INSERT INTO ORDEN_COMPRA_DETALLE (ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA) VALUES (" & ID_ORDEN_COMPRA & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", " & Precio & ", " & DIAS_ENTREGA & ");"
                    cnn.Execute (sqlQuery)
                    IDC = lvwCotizaciones.ListItems.Item(Cont)
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
                        IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
                    End If
                    sqlQuery = "UPDATE COTIZA_REQUI SET NUMOC = " & ID_ORDEN_COMPRA & " WHERE ID_COTIZACION IN (" & IDC & ")"
                    cnn.Execute (sqlQuery)
                    num_orden = ID_ORDEN_COMPRA
                    orden = num_orden
                Next Cont
            Else
                ID_ORDEN_COMPRA = Val(NUM_ORDENS)
                Total = txtSubtotal.Text
                Total = Replace(Total, ",", "")
                TAX = txtImpuesto.Text
                TAX = Replace(TAX, ",", "")
                '''''
                FREIGHT = txtFlete.Text
                FREIGHT = Replace(FREIGHT, ",", "")
                OTROS = txtotros.Text
                OTROS = Replace(OTROS, ",", "")
                '''''
                COMENTARIO = txtcomen.Text
                COMENTARIO = Replace(COMENTARIO, ",", "")
                DISCOUNT = Text3.Text
                DISCOUNT = Replace(DISCOUNT, ",", "")
                '''''
                sqlQuery = "UPDATE ORDEN_COMPRA SET TOTAL = " & Replace(Total, ",", "") & " ,  TAX = " & Replace(TAX, ",", "") & ", FREIGHT=" & Replace(FREIGHT, ",", "") & " ,OTROS_CARGOS=" & Replace(OTROS, ",", "") & "  WHERE ID_ORDEN_COMPRA = " & NUM_ORDENS
                cnn.Execute (sqlQuery)
                sqlQuery = "UPDATE ORDEN_COMPRA SET COMENTARIO ='" & txtcomen.Text & "',DISCOUNT='" & Text3.Text & "' WHERE ID_ORDEN_COMPRA = " & NUM_ORDENS
                cnn.Execute (sqlQuery)
                NoRe = Me.lvwCotizaciones.ListItems.Count
                For Cont = 1 To NoRe
                    ID_PRODUCTO = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(3)
                    Descripcion = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(4)
                    CANTIDAD = lvwCotizaciones.ListItems.Item(Cont).SubItems(5)
                    CANTIDAD = Replace(CANTIDAD, ",", "")
                    DIAS_ENTREGA = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(6)
                    Precio = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(7)
                    Precio = Replace(Precio, ",", "")
                    NUM_ORDENS = lvwCotizaciones.ListItems.Item(Cont).SubItems(10)
                    If NUM_ORDENS = "" Then
                        sqlQuery = "INSERT INTO ORDEN_COMPRA_DETALLE (ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA) VALUES (" & ID_ORDEN_COMPRA & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", " & Precio & ", " & DIAS_ENTREGA & ")"
                        cnn.Execute (sqlQuery)
                        NUM_ORDENS = ID_ORDEN_COMPRA
                    Else
                        sqlQuery = "UPDATE ORDEN_COMPRA_DETALLE SET CANTIDAD = " & CANTIDAD & ",  PRECIO = " & Precio & ", DIAS_ENTREGA = " & DIAS_ENTREGA & "  WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND ID_ORDEN_COMPRA = " & NUM_ORDENS
                        cnn.Execute (sqlQuery)
                    End If
                    IDC = lvwCotizaciones.ListItems.Item(Cont)
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
                        IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
                    End If
                    sqlQuery = "UPDATE COTIZA_REQUI SET NUMOC = " & ID_ORDEN_COMPRA & " WHERE ID_COTIZACION IN (" & IDC & ")"
                    cnn.Execute (sqlQuery)
                Next Cont
                num_orden = Val(NUM_ORDENS)
            End If
            NUM_ORDENS = num_orden
            'CD.ShowPrinter
            Select Case nLvw
                Case 1:
                    FunImp1
                Case 2:
                    FunImp2
                Case 3:
                    FunImp2
            End Select
            'CAMBIAR ESTADO DE COTIZACION
            'If MsgBox("DESEA GUARDAR LA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
            '    NoRe = Me.lvwCotizaciones.ListItems.Count
            '    For Cont = 1 To NoRe
            '        IDC = lvwCotizaciones.ListItems.Item(Cont)
            '        If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
            '            IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
            '        End If
            '        sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'C' WHERE ID_COTIZACION IN (" & IDC & ")"
            '        cnn.Execute (sqlQuery)
            '    Next Cont
            '    sqlQuery = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'P' WHERE ID_ORDEN_COMPRA = " & num_orden
            '    cnn.Execute (sqlQuery)
            'End If
            lblFolio.Caption = ""
            txtSubtotal.Text = "0.00"
            txtImpuesto.Text = "0.00"
            txtTotal.Text = "0.00"
            txtotros.Text = "0.00"
            txtFlete = "0.00"
            Text3.Text = "0.00"
            txtcomen = "0"
            Me.lvwProveedores.ListItems.Clear
            Me.lvwCotizaciones.ListItems.Clear
            Llenar_Lista_Proveedores
            CD.Copies = 1
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image7_Click()
    If ProdMod <> "" Then
        FrmCambiaPrecioPreorden.Label1 = ProdMod
        FrmCambiaPrecioPreorden.Label2 = CanMod
        FrmCambiaPrecioPreorden.Label3 = PresMod
        FrmCambiaPrecioPreorden.Label10 = IdCotiza
        FrmCambiaPrecioPreorden.Label11 = IdProv
        FrmCambiaPrecioPreorden.Label13 = lblFolio
        If opnNacional.Value Then
            FrmCambiaPrecioPreorden.LblTipoOrden.Caption = "N"
        Else
            If opnInternacional.Value Then
                FrmCambiaPrecioPreorden.LblTipoOrden.Caption = "I"
            Else
                FrmCambiaPrecioPreorden.LblTipoOrden.Caption = "X"
            End If
        End If
        FrmCambiaPrecioPreorden.Show vbModal
    Else
        MsgBox "NO SE HA TOMADO NINGUN ARTICULO!", vbInformation, "SACC"
    End If
End Sub
Private Sub Image8_Click()
    On Error GoTo ManejaError
If Puede_Guardar Then
    txtFlete.Enabled = False
    txtotros.Enabled = False
    Dim ID_PROVEEDOR As String
    Dim Total As String
    Dim TAX As String
    Dim OTROS As String
    Dim FREIGHT As String
    Dim DISCOUNT As String
    Dim COMENTARIO As String
    Dim Tipo As String
    Dim NUM_ORDENS As String
    Dim num_orden As Integer
    Dim ID_ORDEN_COMPRA As Integer
    Dim ID_PRODUCTO As String
    Dim Descripcion As String
    Dim CANTIDAD As String
    Dim Precio As String
    Dim DIAS_ENTREGA As Integer
    Dim IDC As String
    Dim nLvw As Integer
    Dim Moneda As String
    Dim FormaPago As String
    If (opnInternacional.Value = False) And (opnNacional.Value = False) And (opnIndirecta.Value = False) Then
        MsgBox "DEBE SELECCIONAR TIPO DE ORDEN PRIMERO"
    Else
        If Me.opnInternacional.Value = True Then
            Tipo = "I"
            nLvw = 1
            Moneda = "DOLARES"
        ElseIf Me.opnNacional.Value = True Then
            Tipo = "N"
            nLvw = 2
            Moneda = txtMoneda.Text
        Else
            Tipo = "X"
            nLvw = 3
            Moneda = txtMoneda.Text
        End If
            Cont = 1
            NUM_ORDENS = ""
            If lvwCotizaciones.ListItems.Count > 0 Then
                Do While (Cont <= lvwCotizaciones.ListItems.Count) And NUM_ORDENS = ""
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(10) <> "" Then
                        NUM_ORDENS = lvwCotizaciones.ListItems.Item(Cont).SubItems(10)
                    End If
                    Cont = Cont + 1
                Loop
            End If
            If NUM_ORDENS = "" Then 'NO EXISTE UNA O.C DE ESTA REQUI
                num_orden = 0
                sqlQuery = "SELECT TOP 1 NUM_ORDEN FROM ORDEN_COMPRA WHERE TIPO = '" & Tipo & "' ORDER BY NUM_ORDEN DESC"
                Set tRs = cnn.Execute(sqlQuery)
                With tRs
                    If Not (.BOF And .EOF) Then
                            If Not IsNull(.Fields("NUM_ORDEN")) Then num_orden = .Fields("NUM_ORDEN")
                        .Close
                    End If
                    num_orden = num_orden + 1
                    numordena = num_orden
                End With
                'FIN TRAER ULTIMA ORDEN DE COMPRA
                'INICIO TRER ULTIMO ID_ORDEN_COMPRA
                ID_ORDEN_COMPRA = 0
                'FIN
                If txtotros.Text = "" Then
                    txtotros.Text = "0"
                End If
                If txtfletes.Text = "" Then
                    txtfletes.Text = "0"
                End If
                Total = Replace(txtSubtotal.Text, ",", "")
                TAX = Replace(txtImpuesto.Text, ",", "")
                OTROS = Replace(txtFlete.Text, ",", "")
                FREIGHT = Replace(txtotros.Text, ",", "")
                If Option1.Value Then
                    FormaPago = "F"
                Else
                    FormaPago = "C"
                End If
                'SE VA A AUTORIZACION O DIRECTO A ORDEN SEGUN EL METODO DE COMPRA DE LA EMPRESA
                If VarMen.TxtEmp(12).Text = "EXTENDIDO" Then
                    sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, TAX, CONFIRMADA, TIPO, NUM_ORDEN, MONEDA, COMENTARIO, ID_USUARIO, FORMA_PAGO, DIAS_CREDITO, FREIGHT, OTROS_CARGOS) VALUES ('" & txtId_Proveedor.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", " & TAX & ", 'P', '" & Tipo & "', " & num_orden & ", '" & Moneda & "','" & txtcomen.Text & "', " & VarMen.Text1(0).Text & ", '" & FormaPago & "', " & Combo1.Text & ", " & txtFlete.Text & ", '" & txtotros.Text & "');"
                Else
                    sqlQuery = "INSERT INTO ORDEN_COMPRA (ID_PROVEEDOR, FECHA, TOTAL, TAX, CONFIRMADA, TIPO, NUM_ORDEN, MONEDA, COMENTARIO, ID_USUARIO, FORMA_PAGO, DIAS_CREDITO, FREIGHT, OTROS_CARGOS) VALUES ('" & txtId_Proveedor.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & Total & ", " & TAX & ", 'S', '" & Tipo & "', " & num_orden & ", '" & Moneda & "','" & txtcomen.Text & "', " & VarMen.Text1(0).Text & ", '" & FormaPago & "', " & Combo1.Text & ", " & txtFlete.Text & ", '" & txtotros.Text & "');"
                End If
                cnn.Execute (sqlQuery)
                sqlQuery = "SELECT TOP 1 ID_ORDEN_COMPRA FROM ORDEN_COMPRA ORDER BY ID_ORDEN_COMPRA DESC"
                Set tRs = cnn.Execute(sqlQuery)
                With tRs
                    If Not (.BOF And .EOF) Then
                            If Not IsNull(.Fields("ID_ORDEN_COMPRA")) Then ID_ORDEN_COMPRA = .Fields("ID_ORDEN_COMPRA")
                            orden = .Fields("ID_ORDEN_COMPRA")
                        .Close
                    End If
                End With
                NoRe = Me.lvwCotizaciones.ListItems.Count
                For Cont = 1 To NoRe
                    ID_PRODUCTO = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(3)
                    Descripcion = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(4)
                    CANTIDAD = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(5)
                    DIAS_ENTREGA = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(6)
                    Precio = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(7)
                    CANTIDAD = Replace(CANTIDAD, ",", "")
                    Precio = Replace(Precio, ",", "")
                    sqlQuery = "INSERT INTO ORDEN_COMPRA_DETALLE (ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA) VALUES (" & ID_ORDEN_COMPRA & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", " & Precio & ", " & DIAS_ENTREGA & ");"
                    cnn.Execute (sqlQuery)
                    IDC = lvwCotizaciones.ListItems.Item(Cont)
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
                        IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
                    End If
                    sqlQuery = "UPDATE COTIZA_REQUI SET NUMOC = " & ID_ORDEN_COMPRA & " WHERE ID_COTIZACION IN (" & IDC & ")"
                    cnn.Execute (sqlQuery)
                    num_orden = ID_ORDEN_COMPRA
                    orden = num_orden
                Next Cont
            Else
                ID_ORDEN_COMPRA = Val(NUM_ORDENS)
                Total = txtSubtotal.Text
                Total = Replace(Total, ",", "")
                TAX = txtImpuesto.Text
                TAX = Replace(TAX, ",", "")
                '''''
                FREIGHT = txtFlete.Text
                FREIGHT = Replace(FREIGHT, ",", "")
                OTROS = txtotros.Text
                OTROS = Replace(OTROS, ",", "")
                '''''
                COMENTARIO = txtcomen.Text
                COMENTARIO = Replace(COMENTARIO, ",", "")
                DISCOUNT = Text3.Text
                DISCOUNT = Replace(DISCOUNT, ",", "")
                '''''
                sqlQuery = "UPDATE ORDEN_COMPRA SET TOTAL = " & Total & " ,  TAX = " & TAX & ", FREIGHT=" & FREIGHT & " ,OTROS_CARGOS=" & OTROS & "  WHERE ID_ORDEN_COMPRA = " & NUM_ORDENS
                cnn.Execute (sqlQuery)
                sqlQuery = "UPDATE ORDEN_COMPRA SET COMENTARIO ='" & txtcomen.Text & "',DISCOUNT='" & Text3.Text & "' WHERE ID_ORDEN_COMPRA = " & NUM_ORDENS
                cnn.Execute (sqlQuery)
                NoRe = Me.lvwCotizaciones.ListItems.Count
                For Cont = 1 To NoRe
                    ID_PRODUCTO = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(3)
                    Descripcion = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(4)
                    CANTIDAD = lvwCotizaciones.ListItems.Item(Cont).SubItems(5)
                    CANTIDAD = Replace(CANTIDAD, ",", "")
                    DIAS_ENTREGA = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(6)
                    Precio = Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(7)
                    Precio = Replace(Precio, ",", "")
                    NUM_ORDENS = lvwCotizaciones.ListItems.Item(Cont).SubItems(10)
                    If NUM_ORDENS = "" Then
                        sqlQuery = "INSERT INTO ORDEN_COMPRA_DETALLE (ID_ORDEN_COMPRA, ID_PRODUCTO, DESCRIPCION, CANTIDAD, PRECIO, DIAS_ENTREGA) VALUES (" & ID_ORDEN_COMPRA & ", '" & ID_PRODUCTO & "', '" & Descripcion & "', " & CANTIDAD & ", " & Precio & ", " & DIAS_ENTREGA & ")"
                        cnn.Execute (sqlQuery)
                        NUM_ORDENS = ID_ORDEN_COMPRA
                    Else
                        sqlQuery = "UPDATE ORDEN_COMPRA_DETALLE SET CANTIDAD = " & CANTIDAD & ",  PRECIO = " & Precio & ", DIAS_ENTREGA = " & DIAS_ENTREGA & "  WHERE ID_PRODUCTO = '" & ID_PRODUCTO & "' AND ID_ORDEN_COMPRA = " & NUM_ORDENS
                        cnn.Execute (sqlQuery)
                    End If
                    IDC = lvwCotizaciones.ListItems.Item(Cont)
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
                        IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
                    End If
                    sqlQuery = "UPDATE COTIZA_REQUI SET NUMOC = " & ID_ORDEN_COMPRA & " WHERE ID_COTIZACION IN (" & IDC & ")"
                    cnn.Execute (sqlQuery)
                Next Cont
                num_orden = Val(NUM_ORDENS)
            End If
            NUM_ORDENS = num_orden
            'CD.ShowPrinter
            'CAMBIAR ESTADO DE COTIZACION
            If MsgBox("DESEA GUARDAR LA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
                NoRe = Me.lvwCotizaciones.ListItems.Count
                For Cont = 1 To NoRe
                    IDC = lvwCotizaciones.ListItems.Item(Cont)
                    If lvwCotizaciones.ListItems.Item(Cont).SubItems(9) <> "" Then
                        IDC = IDC & "," & lvwCotizaciones.ListItems.Item(Cont).SubItems(9)
                    End If
                    sqlQuery = "UPDATE COTIZA_REQUI SET ESTADO_ACTUAL = 'C' WHERE ID_COTIZACION IN (" & IDC & ")"
                    cnn.Execute (sqlQuery)
                Next Cont
                If VarMen.TxtEmp(12).Text = "EXTENDIDO" Then
                    sqlQuery = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'P' WHERE ID_ORDEN_COMPRA = " & num_orden
                Else
                    sqlQuery = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'S' WHERE ID_ORDEN_COMPRA = " & num_orden
                End If
                'sqlQuery = "UPDATE ORDEN_COMPRA SET CONFIRMADA = 'P' WHERE ID_ORDEN_COMPRA = " & num_orden
                cnn.Execute (sqlQuery)
            End If
            lblFolio.Caption = ""
            txtSubtotal.Text = "0.00"
            txtImpuesto.Text = "0.00"
            txtTotal.Text = "0.00"
            txtotros.Text = "0.00"
            txtFlete = "0.00"
            Text3.Text = "0.00"
            txtcomen = "0"
            Me.lvwProveedores.ListItems.Clear
            Me.lvwCotizaciones.ListItems.Clear
            Llenar_Lista_Proveedores
            CD.Copies = 1
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
    Unload Me
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwCotizaciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwCotizaciones.SortKey = ColumnHeader.Index - 1
    lvwCotizaciones.Sorted = True
    lvwCotizaciones.SortOrder = 1 Xor lvwCotizaciones.SortOrder
End Sub
Private Sub lvwCotizaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwCotizaciones.ListItems.Count > 0 Then
        lblIndex.Caption = Item.Index
        lblSelec.Caption = Item.SubItems(3) & " Cantidad: "
        lblidprod.Caption = Item.SubItems(3)
        txtCant.Text = Item.SubItems(5)
        ProdMod = Item.SubItems(3)
        CanMod = Item.SubItems(5)
        PresMod = Item.SubItems(7)
        IdCotiza = Item
        IdOrden = Item.SubItems(10)
        IdProv = Item.SubItems(2)
        Text1.Text = Item.SubItems(3)
    End If
End Sub
Private Sub lvwProveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    If Hay_Cotizaciones Then
        Llenar_Lista_Cotizaciones (Item)
        Me.txtId_Proveedor.Text = Item
        Me.lblProveedor.Caption = Item.SubItems(1)
        Sumar_Importe
        If lblFolio.Caption <> "" Then
          Busca
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Busca()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If opnNacional.Value = True Then
        sBuscar = "SELECT * FROM VSORDENCOMPRA22 WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'N' "
    End If
    If opnInternacional.Value = True Then
        sBuscar = "SELECT * FROM VSORDENCOMPRA22 WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'I' "
    End If
    If opnIndirecta.Value Then
        sBuscar = "SELECT * FROM VSORDENCOMPRA22 WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'X' "
    End If
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Text3 = tRs.Fields("DISCOUNT")
        txtFlete = tRs.Fields("FREIGHT")
        txtotros = tRs.Fields("OTROS_CARGOS")
        txtcomen = tRs.Fields("COMENTARIO")
    End If
    If opnNacional.Value = True Then
        txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        txtTotal = Format((((CDbl(txtSubtotal) - CDbl(Text3.Text))) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    Else
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub RECALCULAR()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If opnNacional.Value = True Then
        txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    Else
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Sub Sumar_Importe()
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Me.txtSubtotal.Text = "0"
    NoRe = Me.lvwCotizaciones.ListItems.Count
    For Cont = 1 To NoRe
        Me.txtSubtotal.Text = Format(CDbl(Me.txtSubtotal.Text) + (Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(5) * Me.lvwCotizaciones.ListItems.Item(Cont).SubItems(7)), "###,###,##0.00")
    Next Cont
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub opnIndirecta_Click()
    txtImpuesto.Text = Format((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(CDbl(Me.txtFlete.Text)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    Me.txtTotal.Text = Format(CDbl(Val(Me.txtSubtotal.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtImpuesto.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtFlete.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtotros.Text)), "###,###,##0.00")
End Sub
Private Sub opnInternacional_Click()
    txtImpuesto.Text = "0.00"
    Me.txtTotal.Text = Format(CDbl(Val(Me.txtSubtotal.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtImpuesto.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtFlete.Text)), "###,###,##0.00") + Format(CDbl(Val(Me.txtotros.Text)), "###,###,##0.00")
End Sub
Private Sub opnNacional_Click()
    Dim totiva As Double
    txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
End Sub
Private Sub txtCargos_Change()
On Error GoTo ManejaError
    Dim totiva As Double
    txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCargos_GotFocus()
    txtCargos.BackColor = &HFFE1E1
End Sub
Private Sub txtCargos_LostFocus()
    txtCargos.BackColor = &H80000005
End Sub
Private Sub txtComentarios_GotFocus()
    txtComentarios.BackColor = &HFFE1E1
End Sub
Private Sub txtComentarios_LostFocus()
    txtComentarios.BackColor = &H80000005
End Sub
Private Sub txtEnviara_GotFocus()
    txtEnviara.BackColor = &HFFE1E1
End Sub
Private Sub txtEnviara_LostFocus()
    txtEnviara.BackColor = &H80000005
End Sub
Private Sub Option1_Click()
    If Option1.Value = True Then
        Combo1.Enabled = True
    Else
        Combo1.Enabled = False
    End If
End Sub
Private Sub Option2_Click()
    If Option1.Value = True Then
        Combo1.Enabled = True
    Else
        Combo1.Enabled = False
    End If
End Sub
Private Sub Text3_Change()
On Error GoTo ManejaError
    Dim totiva As Double
    If opnNacional.Value = True Then
        txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    Else
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And NvoMen.Text1(47).Text = "N" Then
        Command4.Value = True
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
Private Sub txtFlete_Change()
On Error GoTo ManejaError
    Dim totiva As Double
    If opnNacional.Value = True Then
        txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
        'totgen = txtTotal
    Else
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtImpuesto_Change()
On Error GoTo ManejaError
    'If opnInternacional.Value = False Then
    '    txtImpuesto.Text = Val(Replace(txtSubtotal.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100)
    'Else
    '    txtImpuesto.Text = Val("###,###,##0.00")
    'End If
    txtTotal.Text = Val(Replace(txtSubtotal.Text, ",", "")) - Val(Replace(Text3.Text, ",", "")) + Val(Replace(txtImpuesto.Text, ",", "")) + Val(Replace(txtotros.Text, ",", "")) + Val(Replace(txtFlete.Text, ",", ""))
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtotros_Change()
    Dim totiva As Double
    On Error GoTo ManejaError
    If txtotros.Text = "" Then
        txtotros.Text = "0.00"
    End If
    If opnNacional.Value = True Then
        txtImpuesto = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        totiva = Format(((CDbl(Me.txtSubtotal.Text) - CDbl(Text3.Text)) + CDbl(txtFlete)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
        'totgen = txtTotal
    Else
        txtTotal = Format(((CDbl(txtSubtotal) - CDbl(Text3.Text)) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)), "###,###,##0.00")
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtImpuesto_GotFocus()
    txtImpuesto.BackColor = &HFFE1E1
End Sub
Private Sub txtImpuesto_LostFocus()
    txtImpuesto.BackColor = &H80000005
End Sub
Private Sub txtSubtotal_Change()
On Error GoTo ManejaError
    If opnInternacional.Value = False Then
        txtImpuesto.Text = Val(Replace(txtSubtotal.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100)
    Else
        txtImpuesto.Text = Val("###,###,##0.00")
    End If
    txtTotal.Text = Val(Replace(txtSubtotal.Text, ",", "")) - Val(Replace(Text3.Text, ",", "")) + Val(Replace(txtImpuesto.Text, ",", "")) + Val(Replace(txtotros.Text, ",", "")) + Val(Replace(txtFlete.Text, ",", ""))
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub txtSubtotal_GotFocus()
    txtSubtotal.BackColor = &HFFE1E1
End Sub
Private Sub txtSubtotal_LostFocus()
      txtSubtotal.BackColor = &H80000005
End Sub
Function Puede_Guardar() As Boolean
On Error GoTo ManejaError
    If Me.txtId_Proveedor.Text = "" Then
        MsgBox "SELECCIONE EL PROVEEDOR", vbInformation, "SACC"
        Puede_Guardar = False
        Exit Function
    End If
    Puede_Guardar = True
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Private Sub txtTotal_GotFocus()
    txtTotal.BackColor = &HFFE1E1
End Sub
Private Sub txtTotal_LostFocus()
    txtTotal.BackColor = &H80000005
End Sub
Private Sub FunImp1()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If lblFolio.Caption <> "" Then
        If opnNacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'N' "
        End If
        If opnInternacional.Value = True = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'I' "
        End If
        If opnIndirecta.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'X' "
        End If
     Else
        If opnNacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'N' "
        End If
        If opnInternacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'I' "
        End If
        If opnIndirecta.Value Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'X' "
        End If
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\preorden.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        ' Encabezado del reporte
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 20, 38, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR where STATUS = 'A' "
        Set tRs5 = cnn.Execute(sBuscar)
        ''caja2
        oDoc.WTextBox 115, 205, 100, 175, tRs5.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 175, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 138, 205, 100, 175, tRs5.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 175, tRs5.Fields("TEL1"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 175, tRs5.Fields("TEL2"), "F3", 8, hCenter
        oDoc.WTextBox 40, 175, 100, 240, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 50, 224, 100, 175, tRs4.Fields("DIRECCION"), "F3", 8, hLeft
        oDoc.WTextBox 50, 328, 100, 175, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 60, 185, 100, 220, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 40, 400, 20, 250, "Date :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs5 = cnn.Execute(sBuscar)
            If Not (tRs5.EOF And tRs5.BOF) Then
                oDoc.WTextBox 50, 400, 20, 250, "Person in charge :" & tRs5.Fields("NOMBRE") & " " & tRs5.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        If lblFolio.Caption <> "" Then
            oDoc.WTextBox 80, 340, 20, 250, "PRE-ORDEN DE COMPRA INTERNACIONAL#: " & lblFolio.Caption, "F3", 10, hCenter
        Else
            oDoc.WTextBox 60, 340, 20, 250, "PRE-ORDEN DE COMPRA INTERNACIONAL#: " & numordena, "F3", 10, hCenter
        End If
        If tRs.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 100, 340, 20, 250, tRs1.Fields("DIAS_CREDITO") & " Credit Days", "F3", 10, hCenter
        Else
            oDoc.WTextBox 100, 340, 20, 250, "Cash Payment", "F3", 10, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 100, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 100, 175, "INVOICE TO: :", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
        End If
        'CAJA2
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUNT", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        'vsordenesrep
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 160, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 40, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 7, hRight
                oDoc.WTextBox Posi, 477, 20, 50, Format(CDbl(tRs3.Fields("PRECIO")) * CDbl(tRs3.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 7, hRight
                If Len(tRs3.Fields("Descripcion")) > 118 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 236 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 354 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 472 Then
                    Posi = Posi + 15
                End If
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 650 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    Posi = 50
                    oDoc.WImage 50, 40, 43, 161, "Logo"
                    oDoc.WTextBox 30, 340, 20, 250, "Orden de Compra :", "F3", 8, hCenter
                    oDoc.WTextBox 30, 380, 20, 250, Text2.Text, "F3", 8, hCenter
                    ' ENCABEZADO DEL DETALLE
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 580, Posi
                    oDoc.LineStroke
                    Posi = Posi + 10
                    oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 482, 20, 50, "AMOUNT", "F2", 8, hCenter
                    Posi = Posi + 10
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 580, Posi
                    oDoc.LineStroke
                    Posi = Posi + 15
                End If
            Loop
        End If
        ' Linea
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 10
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 20, 100, 275, "THIS PRE-PURCHASE ORDER IS VALID ONLY FOR INTERNAL USE IN THE COMPANY APTONER, strictly FORBIDDEN IS TRYING TO MAKE IT WITH ANY PURCHASE", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox Posi, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtSubtotal), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 20
        oDoc.WTextBox Posi, 400, 20, 70, "Less discount::", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(Text3.Text), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 20
        oDoc.WTextBox Posi, 400, 20, 70, "Other charges:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtotros), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 20
        oDoc.WTextBox Posi, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, txtFlete, "F3", 8, hRight
        oDoc.WTextBox Posi, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Sales Tax", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtImpuesto), "###,###,##0.00"), "F3", 8, hRight 'iva
        Posi = Posi + 20
        oDoc.WTextBox Posi, 400, 60, 100, "TOTAL AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(((CDbl(txtSubtotal) + CDbl(txtFlete) + CDbl(txtotros) + CDbl(txtImpuesto)) - CDbl(Text3.Text)), "###,###,##0.00"), "F3", 8, hRight
        'totales
        Dim SUBB As Double
        Posi = Posi + 10
        'oDoc.WTextBox Posi, 200, 20, 250, "Lic. Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox Posi, 15, 20, 250, "Prices expressed in " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        Posi = Posi + 10
        oDoc.WTextBox Posi, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        Posi = Posi + 10
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImp2()
'NACIONAL//////////////////error corregir
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If lblFolio.Caption <> "" Then
        If opnNacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'N' "
        End If
        If opnInternacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'I' "
        End If
        If opnIndirecta.Value Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & lblFolio.Caption & "' AND TIPO = 'X' "
        End If
     Else
       If opnNacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'N' "
        End If
        If opnInternacional.Value = True Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'I' "
        End If
        If opnIndirecta.Value Then
            sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = '" & numordena & "' AND TIPO = 'X' "
        End If
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\preorden.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        ' Encabezado del reporte
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 38, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 175, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 50, 224, 100, 175, tRs4.Fields("DIRECCION") & "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 50, 328, 100, 175, , "F3", 8, hLeft
        oDoc.WTextBox 70, 205, 100, 175, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 80, 205, 100, 175, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 50, 380, 20, 250, "Fecha :" & Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs5 = cnn.Execute(sBuscar)
            If Not (tRs5.EOF And tRs5.BOF) Then
                oDoc.WTextBox 40, 400, 20, 250, "Responsable :" & tRs5.Fields("NOMBRE") & " " & tRs5.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        If lblFolio.Caption <> "" Then
            oDoc.WTextBox 60, 340, 20, 250, "PRE-ORDEN DE COMPRA: " & lblFolio.Caption, "F3", 10, hCenter
        Else
            oDoc.WTextBox 60, 340, 20, 250, "PRE-ORDEN DE COMPRA: " & numordena, "F3", 10, hCenter
        End If
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 100, 340, 20, 250, tRs1.Fields("DIAS_CREDITO") & " dias de Credito", "F3", 10, hCenter
        Else
            oDoc.WTextBox 100, 340, 20, 250, "Contado", "F3", 10, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 100, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 100, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 170, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 170, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 170, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 170, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 170, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 170, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 170, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO3")) Then oDoc.WTextBox 195, 20, 100, 170, tRs2.Fields("TELEFONO3"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
        oDoc.WTextBox 148, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        'CAJA3
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "CLAVE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "CANTIDAD", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPCION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "PRECIO", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "TOTAL", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        'vsordenesrep
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 160, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 40, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hRight
                oDoc.WTextBox Posi, 477, 20, 50, Format(CDbl(tRs3.Fields("PRECIO")) * CDbl(tRs3.Fields("CANTIDAD")), "###,###,##0.00"), "F3", 8, hRight
                If Len(tRs3.Fields("Descripcion")) > 118 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 236 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 354 Then
                    Posi = Posi + 15
                End If
                If Len(tRs3.Fields("Descripcion")) > 472 Then
                    Posi = Posi + 15
                End If
                Posi = Posi + 15
                tRs3.MoveNext
                If Posi >= 650 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    oDoc.NewPage A4_Vertical
                    Posi = 50
                    oDoc.WImage 50, 40, 43, 161, "Logo"
                    oDoc.WTextBox 30, 340, 20, 250, "Orden de Compra :", "F3", 8, hCenter
                    oDoc.WTextBox 30, 380, 20, 250, Text2.Text, "F3", 8, hCenter
                    ' ENCABEZADO DEL DETALLE
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 580, Posi
                    oDoc.LineStroke
                    Posi = Posi + 10
                    oDoc.WTextBox Posi, 20, 20, 90, "CLAVE", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 112, 20, 50, "CANTIDAD", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPCION", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 418, 20, 50, "PRECIO", "F2", 8, hCenter
                    oDoc.WTextBox Posi, 477, 20, 50, "TOTAL", "F2", 8, hCenter
                    Posi = Posi + 10
                    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                    oDoc.MoveTo 10, Posi
                    oDoc.WLineTo 580, Posi
                    oDoc.LineStroke
                    Posi = Posi + 15
                End If
            Loop
        End If
        Posi = Posi + 10
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        Posi = Posi + 10
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 20, 100, 275, "ESTA PRE-ORDEN DE COMPRA TIENE SOLAMENTE VALIDEZ PARA USO INTERNO EN LA EMPRESA APTONER,QUEDA EXTRICTAMENTE PROHIBIDO TRATAR DE REALIZAR CUALQUIER COMPRA CON ELLA", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox Posi, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtSubtotal), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(Text3.Text), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtotros), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtFlete), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        Posi = Posi - 10
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtImpuesto), "###,###,##0.00"), "F3", 8, hRight 'iva
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(CDbl(txtTotal), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        'oDoc.WTextBox Posi, 200, 20, 250, "Lic. Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox Posi, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        Posi = Posi + 10
        oDoc.WTextBox Posi, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
