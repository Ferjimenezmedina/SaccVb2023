VERSION 5.00
Begin VB.Form FrmRastreoVentProg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rastreo de Producto de Venta programada en Compras"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7560
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8040
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "FrmRastreoVentProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Activate()
    If Text1.Text <> "" And Text2.Text <> "" Then
        Label2.Caption = Text2.Text
        Buscar
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
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim sTipo As String
    Dim sEstado As String
    Dim sSurtido As String
    'Busca si el producto ya esta en Orden de compra (por el campo NO_PEDIDO)
    sBuscar = "SELECT ORDEN_COMPRA.NUM_ORDEN, PROVEEDOR.NOMBRE, ORDEN_COMPRA.CONFIRMADA, ORDEN_COMPRA.TIPO, ORDEN_COMPRA_DETALLE.CANTIDAD, ORDEN_COMPRA_DETALLE.SURTIDO FROM ORDEN_COMPRA_DETALLE, ORDEN_COMPRA, PROVEEDOR WHERE ORDEN_COMPRA_DETALLE.ID_ORDEN_COMPRA = ORDEN_COMPRA.ID_ORDEN_COMPRA AND ORDEN_COMPRA.ID_PROVEEDOR  = PROVEEDOR.ID_PROVEEDOR AND ORDEN_COMPRA_DETALLE.NO_PEDIDO = " & Text1.Text & " AND ORDEN_COMPRA_DETALLE.ID_PRODUCTO = '" & Text2.Text & "'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        If tRs.Fields("TIPO") = "N" Then
            sTipo = "Nacional"
        Else
            If tRs.Fields("TIPO") = "I" Then
                sTipo = "Internacional"
            Else
                sTipo = "Indirecta"
            End If
        End If
        If tRs.Fields("CONFIRMADA") = "N" Then
            sEstado = "Preorden"
        Else
            If tRs.Fields("CONFIRMADA") = "P" Then
                sEstado = "Espera de aprobación"
            Else
                If tRs.Fields("CONFIRMADA") = "S" Then
                    sEstado = "Aprobada para compra"
                Else
                    If tRs.Fields("CONFIRMADA") = "X" Then
                        sEstado = "Cerrada en espera de pago"
                    Else
                        If tRs.Fields("CONFIRMADA") = "Y" Then
                            sEstado = "Pagada"
                        End If
                    End If
                End If
            End If
        End If
        If tRs.Fields("SURTIDO") > 0 Then
            sSurtido = " con entrada de " & tRs.Fields("SURTIDO") & " en Almacen"
        Else
            sSurtido = " sin entrada en Almacen"
        End If
        Label1.Caption = "Producto encontrado en la orden numero " & tRs.Fields("NUM_ORDEN") & " " & sTipo & " del proveedor " & tRs.Fields("NOMBRE") & ", por la cantidad de " & tRs.Fields("CANTIDAD") & " unidades, la cual se ecnuentra en estado de " & sEstado & sSurtido
    Else
        sBuscar = "SELECT COTIZA_REQUI.CANTIDAD, PROVEEDOR.NOMBRE, COTIZA_REQUI.ESTADO_ACTUAL FROM COTIZA_REQUI, PROVEEDOR WHERE COTIZA_REQUI.ID_PROVEEDOR = PROVEEDOR.ID_PROVEEDOR AND COTIZA_REQUI.NO_PEDIDO = " & Text1.Text & " AND COTIZA_REQUI.ID_PRODUCTO = '" & Text2.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Label1.Caption = "Producto encontrado en cotización con el proveedor " & tRs.Fields("NOMBRE") & ", por la cantidad de " & tRs.Fields("CANTIDAD") & " unidades"
        Else
            sBuscar = "SELECT FECHA FROM REQUISICION WHERE NO_PEDIDO = " & Text1.Text & " AND ID_PRODUCTO = '" & Text2.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Label1.Caption = "Producto encontrado en requisición, subido el " & tRs.Fields("FECHA")
            Else
                Label1.Caption = "El producto no se subio a pedido por cumplir con la cantidad necesaria para el surtido al momento de la captura de la venta programada"
            End If
        End If
    End If
End Sub
