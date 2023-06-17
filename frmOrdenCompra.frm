VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrdenCompra 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Orden de Compra"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   37
      Top             =   6000
      Width           =   975
      Begin VB.Label Label11 
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Image cmdImprimir 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmOrdenCompra.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmOrdenCompra.frx":030A
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   35
      Top             =   4800
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmOrdenCompra.frx":1EDC
         MousePointer    =   99  'Custom
         Picture         =   "frmOrdenCompra.frx":21E6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL"
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   33
      Top             =   7200
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmOrdenCompra.frx":3D28
         MousePointer    =   99  'Custom
         Picture         =   "frmOrdenCompra.frx":4032
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
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmOrdenCompra.frx":6114
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Nacionales"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwCotizaciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwOCIndirectas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvwOCNacionales"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvwOCInternacionales"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   5535
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         Begin VB.Frame Frame5 
            Caption         =   "REIMPRIMIR ORDENES CERRADAS"
            Height          =   975
            Left            =   120
            TabIndex        =   39
            Top             =   4200
            Width           =   4335
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   360
               TabIndex        =   42
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Internacional"
               Height          =   255
               Left            =   1920
               TabIndex        =   41
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Nacional"
               Height          =   255
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   1215
            End
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
            Left            =   3240
            Picture         =   "frmOrdenCompra.frx":6130
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1920
            Width           =   1215
         End
         Begin VB.OptionButton opnIndirecta 
            Caption         =   "Indirecta"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton opnInternacional 
            Caption         =   "Internacional"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton opnNacional 
            Caption         =   "Nacional"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtDescuento 
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtFlete 
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCargos 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Text            =   "0"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtImpuesto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtSubtotal 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtTotal 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Text            =   "0"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   1695
            Left            =   120
            TabIndex        =   2
            Top             =   2280
            Width           =   4335
            Begin VB.TextBox txtComentarios 
               Height          =   285
               Left            =   240
               TabIndex        =   4
               Top             =   1080
               Width           =   3975
            End
            Begin VB.TextBox txtEnviara 
               Height          =   285
               Left            =   240
               TabIndex        =   3
               Top             =   480
               Width           =   3975
            End
            Begin VB.Label Label8 
               Caption         =   "COMENTARIOS"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "ENVIAR  A"
               Height          =   255
               Left            =   240
               TabIndex        =   5
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Label lblFolio 
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
            Left            =   3000
            TabIndex        =   25
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "FOLIO"
            Height          =   255
            Left            =   3000
            TabIndex        =   24
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "DESCUENTO"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "OTROS CARGOS"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "FLETE"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "IMPUESTO"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "SUBTOTAL"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "TOTAL"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblID 
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   4320
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lvwOCInternacionales 
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwOCNacionales 
         Height          =   2415
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwOCIndirectas 
         Height          =   135
         Left            =   3720
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   238
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwCotizaciones 
         Height          =   2055
         Left            =   120
         TabIndex        =   29
         Top             =   5880
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Nacionales 
         Caption         =   "Internacionales :"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Nacionales :"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Indirectas :"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   10440
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   10080
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim nOrdenCompra As Integer
Dim nLvw As Byte
Dim NoOrden As String
Dim CompDolar As Double
Dim modif As Double
Dim VarTipo As String
Private Sub CmdGuardar_Click()
    Dim sBuscar As String
    Dim sBuscar2 As String
    Dim Subtotal As String
    Dim Desc As String
    Dim Flete As String
    Dim Cargos As String
    Dim IVA As String
    If lblID.Caption = "" Then
        MsgBox "PRIMERO SELECCIONE UNA ORDEN DE COMPRA", vbInformation, "SACC"
    Else
        Subtotal = Replace(txtSubtotal.Text, ",", "")
        Desc = Replace(txtDescuento.Text, ",", "")
        Flete = Replace(txtFlete.Text, ",", "")
        Cargos = Replace(txtCargos.Text, ",", "")
        IVA = Replace(txtImpuesto.Text, ",", "")
        sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = '" & Subtotal & "', ENVIARA = '" & txtEnviara.Text & "', DISCOUNT = " & Desc & ", FREIGHT = " & Flete & ", TAX = " & IVA & ", OTROS_CARGOS = '" & Cargos & "', COMENTARIO = '" & txtComentarios.Text & "', CONFIRMADA = 'X'  WHERE ID_ORDEN_COMPRA = " & lblID.Caption
        cnn.Execute (sBuscar)
        txtSubtotal.Text = ""
        txtDescuento.Text = ""
        txtFlete.Text = ""
        txtCargos.Text = ""
        txtImpuesto.Text = ""
        txtEnviara.Text = ""
        txtComentarios.Text = ""
        lblID.Caption = ""
        lblFolio.Caption = ""
        nLvw = 0
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
End Sub
Private Sub precioss()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim Subtotal As String
    sBuscar = "SELECT TOP 1(ID_DOLAR) FROM DOLAR ORDER BY ID_DOLAR DESC"
    Set tRs6 = cnn.Execute(sBuscar)
    sBuscar = "SELECT COMPRA FROM DOLAR where  ID_DOLAR ='" & tRs6.Fields("ID_DOLAR") & "'"
    Set tRs3 = cnn.Execute(sBuscar)
    If Not (tRs3.EOF And tRs3.BOF) Then
        CompDolar = tRs3.Fields("COMPRA")
    End If
End Sub
Private Sub cmdImprimir_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim Subtotal As String
    Dim modif   As Double
    Dim Desc As String
    Dim Flete As String
    Dim Cargos As String
    Dim IVA As String
    Dim NOC As Integer
    Dim sBuscar2 As String
    If lblID.Caption = "" Then
        If Text2.Text = "" Then
            MsgBox "PRIMERO SELECCIONE UNA ORDEN DE COMPRA", vbInformation, "SACC"
            Exit Sub
        End If
    Else
        Subtotal = Replace(txtSubtotal.Text, ",", "")
        Desc = Replace(txtDescuento.Text, ",", "")
        Flete = Replace(txtFlete.Text, ",", "")
        Cargos = Replace(txtCargos.Text, ",", "")
        IVA = Replace(txtImpuesto.Text, ",", "")
        sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = '" & Subtotal & "', ENVIARA = '" & txtEnviara.Text & "', DISCOUNT = " & Desc & ", FREIGHT = " & Flete & ", TAX = " & IVA & ", OTROS_CARGOS = '" & Cargos & "', COMENTARIO = '" & txtComentarios.Text & "'  WHERE ID_ORDEN_COMPRA = " & lblID.Caption
        cnn.Execute (sBuscar)
    End If
    If Text2.Text = "" Then
        Select Case nLvw
            Case 1:
                FunImpr2
                FunImpr3
                If MsgBox("DESEA CERRAR LA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
                    Subtotal = Replace(txtSubtotal.Text, ",", "")
                    Desc = Replace(txtDescuento.Text, ",", "")
                    Flete = Replace(txtFlete.Text, ",", "")
                    Cargos = Replace(txtCargos.Text, ",", "")
                    IVA = Replace(txtImpuesto.Text, ",", "")
                    sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = " & Subtotal & ", ENVIARA = '" & txtEnviara.Text & "', DISCOUNT = " & Desc & ", FREIGHT = " & Flete & ", TAX = " & IVA & ", OTROS_CARGOS = " & Cargos & ", COMENTARIO = '" & txtComentarios.Text & "', CONFIRMADA = 'X' WHERE ID_ORDEN_COMPRA = " & nOrdenCompra
                    cnn.Execute (sBuscar)
                End If
            Case 2:
                FunImp
                FunImpco
                If MsgBox("DESEA CERRAR LA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
                    Subtotal = Replace(txtSubtotal.Text, ",", "")
                    Desc = Replace(txtDescuento.Text, ",", "")
                    Flete = Replace(txtFlete.Text, ",", "")
                    Cargos = Replace(txtCargos.Text, ",", "")
                    IVA = Replace(txtImpuesto.Text, ",", "")
                    sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = " & Subtotal & ", ENVIARA = '" & txtEnviara.Text & "', DISCOUNT = " & Desc & ", FREIGHT = " & Flete & ", TAX = " & IVA & ", OTROS_CARGOS = " & Cargos & ", COMENTARIO = '" & txtComentarios.Text & "', CONFIRMADA = 'X' WHERE ID_ORDEN_COMPRA = " & nOrdenCompra
                    cnn.Execute (sBuscar)
               End If
            Case 3:
                FunImpr2
                FunImpr3
                If MsgBox("DESEA CERRAR LA ORDEN DE COMPRA", vbYesNo, "SACC") = vbYes Then
                    Subtotal = Replace(txtSubtotal.Text, ",", "")
                    Desc = Replace(txtDescuento.Text, ",", "")
                    Flete = Replace(txtFlete.Text, ",", "")
                    Cargos = Replace(txtCargos.Text, ",", "")
                    IVA = Replace(txtImpuesto.Text, ",", "")
                    sBuscar = "UPDATE ORDEN_COMPRA SET TOTAL = " & Subtotal & ", ENVIARA = '" & txtEnviara.Text & "', DISCOUNT = " & Desc & ", FREIGHT = " & Flete & ", TAX = " & IVA & ", OTROS_CARGOS = " & Cargos & ", COMENTARIO = '" & txtComentarios.Text & "', CONFIRMADA = 'X' WHERE ID_ORDEN_COMPRA = " & nOrdenCompra
                    cnn.Execute (sBuscar)
                End If
        End Select
        precioss
      ''''''''no  entra  aqui
       sBuscar = "SELECT ID_ORDEN_COMPRA FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = '" & VarTipo & "'"
       Set tRs = cnn.Execute(sBuscar)
       If Not (tRs.EOF And tRs.BOF) Then
           sBuscar = "SELECT ID_PRODUCTO, PRECIO FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs.Fields("ID_ORDEN_COMPRA")
           Set tRs = cnn.Execute(sBuscar)
           If Not (tRs.EOF And tRs.BOF) Then
               sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
               Set tRs1 = cnn.Execute(sBuscar)
               If Not (tRs1.EOF And tRs1.BOF) Then
                    ''''''la var modif  la hice  para cuando sea  inter  ya salga  el pesos convertido y pueda  compara  bien
                    If opnInternacional = True Then
                         modif = CDbl(tRs.Fields("PRECIO")) * CDbl(CompDolar)
                        If modif > tRs1.Fields("PRECIO_COSTO") Then
                            sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO =' " & modif & "' WHERE ID_PRODUCTO  = '" & tRs.Fields("ID_PRODUCTO") & "'"
                            cnn.Execute (sBuscar)
                        End If
                    Else
                        If tRs.Fields("PRECIO") > tRs1.Fields("PRECIO_COSTO") Then
                            sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & tRs.Fields("PRECIO") & " WHERE ID_PRODUCTO  = '" & tRs.Fields("ID_PRODUCTO") & "'"
                            cnn.Execute (sBuscar)
                        End If
                    End If
               Else
                   sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN2 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                   Set tRs1 = cnn.Execute(sBuscar)
                   If Not (tRs1.EOF And tRs1.BOF) Then
                       If tRs.Fields("PRECIO") > tRs1.Fields("PRECIO_COSTO") Then
                            sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = " & tRs.Fields("PRECIO") & " WHERE ID_PRODUCTO  = '" & tRs.Fields("ID_PRODUCTO") & "'"
                            cnn.Execute (sBuscar)
                       End If
                   Else
                       sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN1 WHERE ID_PRODUCTO = '" & tRs.Fields("ID_PRODUCTO") & "'"
                       Set tRs1 = cnn.Execute(sBuscar)
                       If Not (tRs1.EOF And tRs1.BOF) Then
                           If tRs.Fields("PRECIO") > tRs1.Fields("PRECIO_COSTO") Then
                               sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & tRs.Fields("PRECIO") & " WHERE ID_PRODUCTO  = '" & tRs.Fields("ID_PRODUCTO") & "'"
                               cnn.Execute (sBuscar)
                           End If
                       End If
                   End If
               End If
           End If
       End If
    Else
        If Option1.Value Then
            FunImprime
            FunImprimecopia
        End If
        If Option2.Value Then
            FunImprime2
            FunImprime2Copia
        End If
    End If
    sBuscar = "UPDATE ORDEN_COMPRA SET REVISION = REVISION + 1 WHERE ID_ORDEN_COMPRA = " & lblID.Caption
    txtSubtotal.Text = ""
    txtDescuento.Text = ""
    txtFlete.Text = ""
    txtCargos.Text = ""
    txtImpuesto.Text = ""
    txtEnviara.Text = ""
    txtComentarios.Text = ""
    lblID.Caption = ""
    lblFolio.Caption = ""
    Llenar_Lista_Compras "Internacionales"
    Llenar_Lista_Compras "Nacionales"
    Llenar_Lista_Compras "Indirectas"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub FunImprime()
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
    Dim tRs6 As ADODB.Recordset
    Dim ConPag As Integer
    ConPag = 1
    Dim sBuscar As String
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N' AND CONFIRMADA <> 'A'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, ), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : " & Text2.Text, "F3", 8, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credito", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Contado", "F3", 10, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "ENVIAR A:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Responsable :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'N'"
        Set tRs5 = cnn.Execute(sBuscar)
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
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
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
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
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("DISCOUNT")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("DISCOUNT"), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("OTROS_CARGOS"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("FREIGHT")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("FREIGHT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TAX")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("TAX"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("DISCOUNT")) Then
            Total = Format(CDbl(tRs1.Fields("TOTAL")) - (CDbl(tRs1.Fields("DISCOUNT"))), "###,###,##0.00")
        Else
            Total = Format(CDbl(tRs1.Fields("TOTAL")), "###,###,##0.00")
        End If
        If Not IsNull(tRs1.Fields("FREIGHT")) Then
            Total = Format(CDbl(Total) + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")), "###,###,##0.00")
        End If
        If Not IsNull(tRs1.Fields("TAX")) Then
            Total = Format(CDbl(Total) + CDbl(tRs1.Fields("TAX")), "###,###,##0.00")
        End If
        oDoc.WTextBox 720, 488, 20, 50, Format(Total, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 720, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 720, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        ' cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se enc ontro la orden de compra solicitada, puede ser que este cancelda o aun no se genere el folio", vbExclamation, "SACC"
    End If
End Sub
Private Sub FunImprimecopia()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim ConPag As Integer
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    ConPag = 1
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM vsordenesrep WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N' AND CONFIRMADA <> 'A'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompraCopia.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : " & Text2.Text, "F3", 8, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credito", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Contado", "F3", 10, hCenter
        End If
        If Not IsNull(tRs1.Fields("REVISION")) Then
            If tRs1.Fields("REVISION") <> 0 Then
                oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
            End If
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "ENVIAR A:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Reponsable  :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'N'"
        Set tRs5 = cnn.Execute(sBuscar)
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
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
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
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
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, tRs1.Fields("TOTAL"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("DISCOUNT")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("DISCOUNT"), "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("OTROS_CARGOS"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("FREIGHT")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("FREIGHT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TAX")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("TAX"), "###,###,##0.00"), "F3", 8, hRight
        discon = Format(CDbl(tRs1.Fields("TOTAL")) - (CDbl(tRs1.Fields("DISCOUNT"))), "###,###,##0.00")
        Total = Format(CDbl(discon) + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")) + CDbl(tRs1.Fields("TAX")), "###,###,##0.00")
        oDoc.WTextBox 720, 488, 20, 50, Format(Total, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 720, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        Else
            oDoc.WTextBox 750, 200, 25, 250, "COPIA", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 720, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        If Not IsNull(tRs1.Fields("COMENTARIO")) Then oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImp()
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
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If opnNacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
    End If
    If opnInternacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'I'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompranacioa.pdf") Then
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
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : ", "F3", 10, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credito", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Contado", "F3", 10, hCenter
        End If
        oDoc.WTextBox 60, 395, 20, 250, lblFolio.Caption, "F2", 10, hCenter
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "ENVIAR A:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Responsable :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'N'"
        Set tRs5 = cnn.Execute(sBuscar)
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
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
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 160, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : " & lblFolio.Caption, "F3", 10, hCenter
                        If tRs1.Fields("REVISION") <> 0 Then
                            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
                        End If
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
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(txtSubtotal.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 640, 488, 20, 50, Format(txtDescuento.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 660, 488, 20, 50, Format(txtCargos.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 680, 488, 20, 50, Format(txtFlete.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 700, 488, 20, 50, Format(txtImpuesto, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 720, 488, 20, 50, Format(txtTotal, "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 730, 200, 20, 250, "Lic. Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox 730, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 740, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 10, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImpco()
'''copia nacional
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
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If opnNacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
    End If
    If opnInternacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'I'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompranacopia.pdf") Then
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
        sBuscar = "SELECT * FROM EMPRESA"
        Set tRs4 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Fecha :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "Orden de Compra : ", "F3", 10, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credito", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Contado", "F3", 10, hCenter
        End If
        oDoc.WTextBox 60, 395, 20, 250, lblFolio.Caption, "F2", 10, hCenter
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDEDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "FACTURAR A :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "ENVIAR A:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Responsable :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'N'"
        Set tRs5 = cnn.Execute(sBuscar)
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
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
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 160, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
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
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Envíe por favor los artículos a la direccion de envío, si usted tiene alguna duda o pregunta sobre esta orden de compra, por favor contacte a  Dept Compras, al Tel. " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "SUBTOTAL:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Descuento:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Otros Cargos:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Flete:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "I.V.A:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(txtSubtotal.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 640, 488, 20, 50, Format(txtDescuento.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 660, 488, 20, 50, Format(txtCargos.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 680, 488, 20, 50, Format(txtFlete.Text, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 700, 488, 20, 50, Format(txtImpuesto, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox 720, 488, 20, 50, Format(txtTotal, "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 730, 200, 20, 250, "Lic. Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox 730, 15, 20, 250, "Precios expresados en " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 740, 200, 20, 250, "Firma de autorizado", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        Else
            oDoc.WTextBox 750, 200, 25, 250, "COPIA", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 10, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImprime2()
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
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N' AND CONFIRMADA <> 'A'"
    End If
    If Option2.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I' AND CONFIRMADA <> 'A'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'I'"
        Set tRs5 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Date :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "PURCHASE ORDER : " & Text2.Text, "F3", 8, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credit Days", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Cash Payment", "F3", 10, hCenter
        End If
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "INVOICE TO :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "SHIP TO:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Person in charge  :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUN", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 30, 340, 20, 250, "PURCHASE ORDER # :", "F3", 8, hCenter
                        oDoc.WTextBox 30, 380, 20, 250, Text2.Text, "F3", 8, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "ITEM#", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "QTY", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "AMOUN", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 15
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Please include country of origin for all items", "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 640, 20, 100, 275, "Please deliver to our shipping Address. Any questions, contact Purchasing Dept at " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Less Discount:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Other Charges:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "Sales Tax:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL AMOUNT:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, Format(tRs1.Fields("TOTAL"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("DISCOUNT")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("DISCOUNT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("OTROS_CARGOS")) Then oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("OTROS_CARGOS"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("FREIGHT")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("FREIGHT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TAX")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("TAX"), "###,###,##0.00"), "F3", 8, hRight
        discon = Format(CDbl(tRs1.Fields("TOTAL")) - (CDbl(tRs1.Fields("DISCOUNT"))), "###,###,##0.00")
        Total = Format(CDbl(discon) + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")) + CDbl(tRs1.Fields("TAX")), "###,###,##0.00")
        oDoc.WTextBox 720, 488, 20, 50, Format(Total, "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 730, 200, 20, 250, "Mr Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox 730, 15, 20, 250, "Prices expressed in " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 740, 200, 20, 250, "Autorized Signature", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 680, 20, 100, 275, "COMMENTARIES:", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "No se encontro la orden de compra solicitada, puede ser que este cancelda o aun no se genere el folio", vbExclamation, "SACC"
    End If
End Sub
Private Sub FunImprime2Copia()
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
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If Option1.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N' AND CONFIRMADA <> 'A'"
    End If
    If Option2.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I' AND CONFIRMADA <> 'A'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompraCopia.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'I'"
        Set tRs5 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "Date :" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "PURCHASE ORDER : " & Text2.Text, "F3", 8, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credit Days", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Cash Payment", "F3", 10, hCenter
        End If
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "INVOICE TO :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "SHIP TO:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Person in charge  :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUN", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 8, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 8, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 30, 340, 20, 250, "PURCHASE ORDER # :", "F3", 8, hCenter
                        oDoc.WTextBox 30, 380, 20, 250, Text2.Text, "F3", 8, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "ITEM#", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "QTY", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "AMOUN", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 15
                    End If
                End If
            Loop
        End If
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, 600
        oDoc.WLineTo 580, 600
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox 620, 20, 100, 275, "Please include country of origin for all items", "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 640, 20, 100, 275, "Please deliver to our shipping Address. Any questions, contact Purchasing Dept at " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox 620, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox 640, 400, 20, 70, "Less Discount:", "F2", 8, hRight
        oDoc.WTextBox 660, 400, 20, 70, "Other Charges:", "F2", 8, hRight
        oDoc.WTextBox 680, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox 700, 400, 20, 70, "Sales Tax:", "F2", 8, hRight
        oDoc.WTextBox 720, 400, 20, 70, "TOTAL AMOUNT:", "F2", 8, hRight
        'totales
        oDoc.WTextBox 620, 488, 20, 50, tRs1.Fields("TOTAL"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("DISCOUNT")) Then oDoc.WTextBox 640, 488, 20, 50, Format(tRs1.Fields("DISCOUNT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("OTROS_CARGOS")) Then oDoc.WTextBox 660, 488, 20, 50, Format(tRs1.Fields("OTROS_CARGOS"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("FREIGHT")) Then oDoc.WTextBox 680, 488, 20, 50, Format(tRs1.Fields("FREIGHT"), "###,###,##0.00"), "F3", 8, hRight
        If Not IsNull(tRs1.Fields("TAX")) Then oDoc.WTextBox 700, 488, 20, 50, Format(tRs1.Fields("TAX"), "###,###,##0.00"), "F3", 8, hRight
        discon = Format(CDbl(tRs1.Fields("TOTAL")) - (CDbl(tRs1.Fields("DISCOUNT"))), "###,###,##0.00")
        Total = Format(CDbl(discon) + CDbl(tRs1.Fields("OTROS_CARGOS")) + CDbl(tRs1.Fields("FREIGHT")) + CDbl(tRs1.Fields("TAX")), "###,###,##0.00")
        oDoc.WTextBox 720, 488, 20, 50, Format(Total, "###,###,##0.00"), "F3", 8, hRight
        'oDoc.WTextBox 730, 200, 20, 250, "Mr Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox 730, 15, 20, 250, "Prices expressed in " & tRs1.Fields("MONEDA"), "F3", 10, hCenter
        oDoc.WTextBox 740, 200, 20, 250, "Autorized Signature", "F3", 10, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPY CANCELED", "F3", 25, hCenter, , vbBlue
        Else
            oDoc.WTextBox 750, 200, 25, 250, "COPY", "F3", 25, hCenter, , vbBlue
        End If
        oDoc.WTextBox 680, 20, 100, 275, "COMENTARIOS:", "F3", 8, hLeft, , , 0, vbBlack
        oDoc.WTextBox 690, 20, 100, 300, tRs1.Fields("COMENTARIO"), "F3", 8, hLeft, , , 0, vbBlack
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImpr2()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim discon As Double
    Dim Total As Double
    Dim Posi As Integer
    Dim Moneda As String
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tRs5 As ADODB.Recordset
    Dim tRs6 As ADODB.Recordset
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    'text1.text = no_orden
    'nvomen.Text1(0).Text = id_usuario
    If opnNacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
    End If
    If opnInternacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'I'"
    End If
    If opnIndirecta.Value Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'X'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        If Not IsNull(tRs1.Fields("MONEDA")) Then Moneda = tRs1.Fields("MONEDA")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompra.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'I'"
        Set tRs5 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 30, 380, 20, 250, "Date:" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 60, 340, 20, 250, "PURCHASE ORDER : ", "F3", 8, hCenter
        oDoc.WTextBox 60, 390, 20, 250, lblFolio.Caption, "F2", 11, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credit Days", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Cash Payment", "F3", 10, hCenter
        End If
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "INVOICE TO :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "SHIP TO:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Person in charge :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUN", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 7, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 30, 340, 20, 250, "PURCHASE ORDER # :", "F3", 9, hCenter, , , 1, vbBlack
                        oDoc.WTextBox 30, 390, 20, 250, lblFolio.Caption, "F3", 11, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "ITEM#", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "QTY", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "AMOUN", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 15
                    End If
                End If
            Loop
        End If
        ' Linea
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 20, 100, 275, "Please include country of origin for all items", "F3", 10, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 100, 275, "Please deliver to our shipping Address. Any questions, contact Purchasing Dept at " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox Posi, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtSubtotal.Text), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Less Discount:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtDescuento, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Other Charges:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtCargos, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtFlete, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Sales Tax:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtImpuesto, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 20, 100, 275, "COMENTARIOS:", "F2", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox 690, 20, 100, 300, Format(txtComentarios, "###,###,##0.00"), "F3", 11, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox Posi, 400, 20, 70, "TOTAL AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtTotal), "###,###,##0.00"), "F3", 8, hRight
        'totales
        Posi = Posi + 6
        'oDoc.WTextBox Posi, 200, 20, 250, "Mr Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox Posi, 15, 20, 250, "Prices expressed in " & Moneda, "F3", 10, hCenter
        Posi = Posi + 10
        oDoc.WTextBox Posi, 200, 20, 250, "Autorized Signature", "F3", 8, hCenter
        'If tRs1.Fields("CONFIRMADA") = "E" Then
        '    oDoc.WTextBox 750, 200, 25, 250, "COPY CANCELED", "F3", 25, hCenter, , vbBlue
        'Else
        '    oDoc.WTextBox 750, 200, 25, 250, "COPY", "F3", 25, hCenter, , vbBlue
        'End If
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub FunImpr3()
''''''''''''''copia de internacional
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
    Dim tRs6 As ADODB.Recordset
    Dim Moneda As String
    Dim sBuscar As String
    Dim ConPag As Integer
    ConPag = 1
    If opnNacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'N'"
    End If
    If opnInternacional.Value = True Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'I'"
    End If
    If opnIndirecta.Value Then
        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & lblFolio.Caption & " AND TIPO = 'X'"
    End If
    Set tRs1 = cnn.Execute(sBuscar)
    If Not (tRs1.EOF And tRs1.BOF) Then
        If Not IsNull(tRs1.Fields("MONEDA")) Then Moneda = tRs1.Fields("MONEDA")
        Set oDoc = New cPDF
        If Not oDoc.PDFCreate(App.Path & "\OrdenCompracopia.pdf") Then
            Exit Sub
        End If
        oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
        oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
        oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
        oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
        ' Encabezado del reporte
        ' asi se agregan los logos... solo te falto poner un control IMAGE1 para cargar la imagen en el
        Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
        oDoc.LoadImage Image1, "Logo", False, False
        oDoc.NewPage A4_Vertical
        oDoc.WImage 70, 40, 43, 161, "Logo"
        sBuscar = "SELECT * FROM EMPRESA  "
        Set tRs4 = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM DIREIMPOR WHERE STATUS = 'A' AND TIPO = 'I'"
        Set tRs5 = cnn.Execute(sBuscar)
        oDoc.WTextBox 40, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 60, 224, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        'oDoc.WTextBox 60, 328, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hLeft
        oDoc.WTextBox 80, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 90, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 40, 380, 20, 250, "Date:" & Format(tRs1.Fields("FECHA"), "dd/mm/yyyy"), "F3", 8, hCenter
        oDoc.WTextBox 70, 340, 20, 250, "PURCHASE ORDER : ", "F3", 8, hCenter
        If tRs1.Fields("FORMA_PAGO") = "F" Then
            oDoc.WTextBox 80, 400, 20, 70, tRs1.Fields("DIAS_CREDITO") & " Credit Days", "F3", 10, hCenter
        Else
            oDoc.WTextBox 80, 400, 20, 70, "Cash Payment", "F3", 10, hCenter
        End If
        oDoc.WTextBox 70, 390, 20, 250, lblFolio.Caption, "F2", 11, hCenter
        If tRs1.Fields("REVISION") <> 0 Then
            oDoc.WTextBox 70, 340, 20, 250, "Revision : " & tRs1.Fields("REVISION"), "F3", 8, hCenter
        End If
        ' cuadros encabezado
        oDoc.WTextBox 100, 20, 105, 175, "VENDOR : ", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 205, 105, 175, "INVOICE TO :", "F2", 10, hCenter, , , 1, vbBlack
        oDoc.WTextBox 100, 390, 105, 175, "SHIP TO:", "F2", 10, hCenter, , , 1, vbBlack
        ' LLENADO DE LAS CAJAS
        If Not IsNull(tRs1.Fields("ID_USUARIO")) Then
            sBuscar = "SELECT  NOMBRE, APELLIDOS FROM USUARIOS WHERE ID_USUARIO = " & tRs1.Fields("ID_USUARIO")
            Set tRs6 = cnn.Execute(sBuscar)
            If Not (tRs6.EOF And tRs6.BOF) Then
                oDoc.WTextBox 50, 350, 20, 250, "Person in charge  :" & tRs6.Fields("NOMBRE") & " " & tRs6.Fields("APELLIDOS"), "F3", 8, hCenter
            End If
        End If
        'CAJA1
        sBuscar = "SELECT * FROM PROVEEDOR WHERE ID_PROVEEDOR = " & tRs1.Fields("ID_PROVEEDOR")
        Set tRs2 = cnn.Execute(sBuscar)
        If Not (tRs2.EOF And tRs2.BOF) Then
            If Not IsNull(tRs2.Fields("NOMBRE")) Then oDoc.WTextBox 115, 20, 100, 175, tRs2.Fields("NOMBRE"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("DIRECCION")) Then oDoc.WTextBox 135, 20, 100, 175, tRs2.Fields("DIRECCION"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("COLONIA")) Then oDoc.WTextBox 145, 20, 100, 175, tRs2.Fields("COLONIA"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CIUDAD")) Then oDoc.WTextBox 155, 20, 100, 175, tRs2.Fields("CIUDAD"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("CP")) Then oDoc.WTextBox 165, 20, 100, 175, tRs2.Fields("CP"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO1")) Then oDoc.WTextBox 175, 20, 100, 175, tRs2.Fields("TELEFONO1"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("TELEFONO2")) Then oDoc.WTextBox 185, 20, 100, 175, tRs2.Fields("TELEFONO2"), "F3", 8, hCenter
            If Not IsNull(tRs2.Fields("RFC")) Then oDoc.WTextBox 195, 20, 100, 175, tRs2.Fields("RFC"), "F3", 8, hCenter
        End If
        'CAJA2
        oDoc.WTextBox 115, 205, 100, 170, tRs4.Fields("NOMBRE"), "F3", 8, hCenter
        oDoc.WTextBox 130, 205, 100, 170, tRs4.Fields("DIRECCION") & " " & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        'oDoc.WTextBox 138, 205, 100, 170, "Col." & tRs4.Fields("COLONIA"), "F3", 8, hCenter
        oDoc.WTextBox 158, 205, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
        oDoc.WTextBox 168, 205, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
        oDoc.WTextBox 176, 205, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        oDoc.WTextBox 196, 205, 100, 170, tRs4.Fields("RFC"), "F3", 8, hCenter
        'CAJA3
        If Not (tRs5.EOF And tRs5.BOF) Then
            oDoc.WTextBox 125, 390, 100, 170, tRs5.Fields("DIRECCION"), "F3", 8, hCenter
            'oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs5.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, tRs5.Fields("TEL1"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs5.Fields("TEL2"), "F3", 8, hCenter
        Else
            oDoc.WTextBox 125, 390, 100, 170, tRs4.Fields("DIRECCION"), "F3", 8, hCenter
            oDoc.WTextBox 145, 390, 100, 170, tRs4.Fields("COLONIA"), "F3", 8, hCenter
            oDoc.WTextBox 155, 390, 100, 170, tRs4.Fields("ESTADO") & "," & tRs4.Fields("CD"), "F3", 8, hCenter
            oDoc.WTextBox 165, 390, 100, 170, "CP " & tRs4.Fields("CP"), "F3", 8, hCenter
            oDoc.WTextBox 175, 390, 100, 170, tRs4.Fields("TELEFONO"), "F3", 8, hCenter
        End If
        Posi = 210
        ' ENCABEZADO DEL DETALLE
        oDoc.WTextBox Posi, 5, 20, 90, "ITEM#", "F2", 8, hCenter
        oDoc.WTextBox Posi, 100, 20, 50, "QTY", "F2", 8, hCenter
        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
        oDoc.WTextBox Posi, 415, 20, 50, "UNIT PRICE", "F2", 8, hCenter
        oDoc.WTextBox Posi, 482, 20, 50, "AMOUN", "F2", 8, hCenter
        Posi = Posi + 15
        ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' DETALLE
        sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & tRs1.Fields("ID_ORDEN_COMPRA")
        Set tRs3 = cnn.Execute(sBuscar)
        If Not (tRs3.EOF And tRs3.BOF) Then
            Do While Not tRs3.EOF
                oDoc.WTextBox Posi, 20, 20, 90, Mid(tRs3.Fields("ID_PRODUCTO"), 1, 18), "F3", 7, hLeft
                oDoc.WTextBox Posi, 110, 20, 50, tRs3.Fields("CANTIDAD"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 164, 20, 260, tRs3.Fields("Descripcion"), "F3", 7, hLeft
                oDoc.WTextBox Posi, 430, 20, 50, Format(tRs3.Fields("PRECIO"), "###,###,##0.00"), "F3", 7, hLeft
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
                If Posi >= 620 Then
                    oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                    ConPag = ConPag + 1
                    If Option1.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'N'"
                    End If
                    If Option2.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'I'"
                    End If
                    If opnIndirecta.Value Then
                        sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE NUM_ORDEN = " & Text2.Text & " AND TIPO = 'X'"
                    End If
                    Set tRs1 = cnn.Execute(sBuscar)
                    If Not (tRs1.EOF And tRs1.BOF) Then
                        oDoc.NewPage A4_Vertical
                        Posi = 50
                        oDoc.WImage 50, 40, 43, 161, "Logo"
                        oDoc.WTextBox 30, 340, 20, 250, "PURCHASE ORDER # :", "F3", 9, hCenter, , , 1, vbBlack
                        oDoc.WTextBox 30, 390, 20, 250, lblFolio.Caption, "F3", 11, hCenter
                        ' ENCABEZADO DEL DETALLE
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 10
                        oDoc.WTextBox Posi, 20, 20, 90, "ITEM#", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 112, 20, 50, "QTY", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 164, 20, 260, "DESCRIPTION", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 418, 20, 50, "UNIT PRICE", "F2", 8, hCenter
                        oDoc.WTextBox Posi, 477, 20, 50, "AMOUN", "F2", 8, hCenter
                        Posi = Posi + 10
                        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                        oDoc.MoveTo 10, Posi
                        oDoc.WLineTo 580, Posi
                        oDoc.LineStroke
                        Posi = Posi + 15
                    End If
                End If
            Loop
        End If
        Posi = Posi + 6
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 10, Posi
        oDoc.WLineTo 580, Posi
        oDoc.LineStroke
        Posi = Posi + 6
        ' TEXTO ABAJO
        oDoc.WTextBox Posi, 20, 100, 275, "Please include country of origin for all items", "F3", 10, hLeft, , , 0, vbBlack
        Posi = Posi + 10
        oDoc.WTextBox Posi, 20, 100, 275, "Please deliver to our shipping Address. Any questions, contact Purchasing Dept at " & VarMen.TxtEmp(2).Text, "F3", 10, hLeft, , , 0, vbBlack
        oDoc.WTextBox Posi, 400, 20, 70, "NET AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtSubtotal.Text), "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Less Discount:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtDescuento, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Other Charges:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtCargos, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Freight:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtFlete, "###,###,##0.00"), "F3", 8, hRight
        Posi = Posi + 10
        oDoc.WTextBox Posi, 400, 20, 70, "Sales Tax:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format(txtImpuesto, "###,###,##0.00"), "F3", 8, hRight
        oDoc.WTextBox Posi, 20, 100, 275, "COMENTARIOS:", "F2", 8, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox 690, 20, 100, 300, Format(txtComentarios, "###,###,##0.00"), "F3", 11, hLeft, , , 0, vbBlack
        Posi = Posi + 5
        oDoc.WTextBox Posi, 400, 20, 70, "TOTAL AMOUNT:", "F2", 8, hRight
        oDoc.WTextBox Posi, 488, 20, 50, Format((txtTotal), "###,###,##0.00"), "F3", 8, hRight
        'totales
        Posi = Posi + 6
        'oDoc.WTextBox Posi, 200, 20, 250, "Mr Lorenzo Bujaidar", "F3", 10, hCenter
        oDoc.WTextBox Posi, 15, 20, 250, "Prices expressed in " & Moneda, "F3", 10, hCenter
        Posi = Posi + 10
        oDoc.WTextBox Posi, 200, 20, 250, "Autorized Signature", "F3", 8, hCenter
        If tRs1.Fields("CONFIRMADA") = "E" Then
            oDoc.WTextBox 750, 200, 25, 250, "COPIA CANCELADA", "F3", 25, hCenter, , vbBlue
        Else
            oDoc.WTextBox 750, 200, 25, 250, "COPY", "F3", 25, hCenter, , vbBlue
        End If
        'cierre del reporte
        oDoc.PDFClose
        oDoc.Show
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    VarTipo = ""
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwOCInternacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With lvwOCNacionales
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With lvwOCIndirectas
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "ID_PROVEEDOR", 0
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 1440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "COMENTARIO", 1440
    End With
    With Me.lvwCotizaciones
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Producto", 2200
        .ColumnHeaders.Add , , "Descripcion", 3500
        .ColumnHeaders.Add , , "CANTIDAD", 1000, 2
        .ColumnHeaders.Add , , "PRECIO", 1440, 2
        .ColumnHeaders.Add , , "SUBTOTAL", 1440, 2
    End With
    If NvoMen.Text1(11).Text = "N" Then
        Frame5.Visible = False
    Else
        Frame5.Visible = True
    End If
    If Hay_Ordenes_Compra Then
        Llenar_Lista_Compras "Internacionales"
        Llenar_Lista_Compras "Nacionales"
        Llenar_Lista_Compras "Indirectas"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Public Function Hay_Ordenes_Compra() As Boolean
On Error GoTo ManejaError
    Dim sBuscar  As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT  COUNT(ID_ORDEN_COMPRA) AS CONTA FROM ORDEN_COMPRA WHERE CONFIRMADA = 'P' OR Confirmada = 'S'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Hay_Ordenes_Compra = True
    Else
        Hay_Ordenes_Compra = False
    End If
Exit Function
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Function
Sub Llenar_Lista_Compras(Tipo As String)
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Select Case Tipo
        Case "Internacionales":
            Me.lvwOCInternacionales.ListItems.Clear
            sBuscar = "SELECT  OC.Id_Orden_Compra, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, oc.COMENTARIO FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Tipo = 'I' AND OC.Confirmada = 'S' ORDER BY OC.Id_Orden_Compra"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    Set ItMx = Me.lvwOCInternacionales.ListItems.Add(, , tRs.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(tRs.Fields("ID_PROVEEDOR"))
                    If Not IsNull(tRs.Fields("Nombre")) Then ItMx.SubItems(2) = Trim(tRs.Fields("Nombre"))
                    If Not IsNull(tRs.Fields("Total_Pagar")) Then ItMx.SubItems(3) = Trim(tRs.Fields("Total_Pagar"))
                    If Not IsNull(tRs.Fields("COMENTARIO")) Then ItMx.SubItems(4) = Trim(tRs.Fields("COMENTARIO"))
                    tRs.MoveNext
                Loop
            End If
        Case "Nacionales":
            Me.lvwOCNacionales.ListItems.Clear
            sBuscar = "SELECT  OC.Id_Orden_Compra, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, oc.COMENTARIO FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Tipo = 'N' AND OC.Confirmada = 'S' ORDER BY OC.Id_Orden_Compra"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    Set ItMx = Me.lvwOCNacionales.ListItems.Add(, , tRs.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(tRs.Fields("ID_PROVEEDOR"))
                    If Not IsNull(tRs.Fields("Nombre")) Then ItMx.SubItems(2) = Trim(tRs.Fields("Nombre"))
                    If Not IsNull(tRs.Fields("Total_Pagar")) Then ItMx.SubItems(3) = Trim(tRs.Fields("Total_Pagar"))
                    If Not IsNull(tRs.Fields("COMENTARIO")) Then ItMx.SubItems(4) = Trim(tRs.Fields("COMENTARIO"))
                    tRs.MoveNext
                Loop
            End If
        Case "Indirectas":
            Me.lvwOCIndirectas.ListItems.Clear
            sBuscar = "SELECT  OC.Id_Orden_Compra, OC.Id_Proveedor, P.Nombre, ((OC.Total - OC.Discount) + OC.TAX + OC.Freight + OC.Otros_Cargos) AS Total_Pagar, oc.COMENTARIO FROM ORDEN_COMPRA AS OC JOIN PROVEEDOR AS P ON P.Id_Proveedor = OC.Id_Proveedor WHERE OC.Tipo = 'X' AND OC.Confirmada = 'S' ORDER BY OC.Id_Orden_Compra"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    Set ItMx = Me.lvwOCIndirectas.ListItems.Add(, , tRs.Fields("ID_ORDEN_COMPRA"))
                    If Not IsNull(tRs.Fields("ID_PROVEEDOR")) Then ItMx.SubItems(1) = Trim(tRs.Fields("ID_PROVEEDOR"))
                    If Not IsNull(tRs.Fields("Nombre")) Then ItMx.SubItems(2) = Trim(tRs.Fields("Nombre"))
                    If Not IsNull(tRs.Fields("Total_Pagar")) Then ItMx.SubItems(3) = Trim(tRs.Fields("Total_Pagar"))
                    If Not IsNull(tRs.Fields("COMENTARIO")) Then ItMx.SubItems(4) = Trim(tRs.Fields("COMENTARIO"))
                    tRs.MoveNext
                Loop
            End If
        Case Else:
            MsgBox "ERROR GRAVE. LA APLICACIÓN TERMINARA", vbCritical, "SACC"
            End
    End Select
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CD.FileName = ""
    CD.DialogTitle = "Guardar como"
    CD.Filter = "Excel (*.xls) |*.xls|"
    Me.CD.ShowSave
    Ruta = Me.CD.FileName
    If lvwCotizaciones.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = lvwCotizaciones.ColumnHeaders.Count
            For Con = 1 To lvwCotizaciones.ColumnHeaders.Count
                StrCopi = StrCopi & lvwCotizaciones.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To lvwCotizaciones.ListItems.Count
                StrCopi = StrCopi & lvwCotizaciones.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & lvwCotizaciones.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwOCIndirectas_Click()
On Error GoTo ManejaError
    If lvwOCIndirectas.ListItems.Count > 0 Then
        nLvw = 1
        nOrdenCompra = Me.lvwOCIndirectas.SelectedItem
        TraeDatos nOrdenCompra
        VarTipo = "X"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwOCIndirectas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoOrden = Item
End Sub
Private Sub lvwOCInternacionales_Click()
On Error GoTo ManejaError
    If lvwOCInternacionales.ListItems.Count > 0 Then
        nLvw = 1
        nOrdenCompra = Me.lvwOCInternacionales.SelectedItem
        TraeDatos nOrdenCompra
        VarTipo = "I"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub lvwOCInternacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoOrden = Item
End Sub
Private Sub lvwOCNacionales_Click()
On Error GoTo ManejaError
    If lvwOCNacionales.ListItems.Count > 0 Then
        nLvw = 2
        nOrdenCompra = Me.lvwOCNacionales.SelectedItem
        TraeDatos nOrdenCompra
        VarTipo = "N"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub TraeDatos(NO As Integer)
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tLi As ListItem
    Dim Subtotal As Double
    sBuscar = "SELECT * FROM ORDEN_COMPRA WHERE ID_ORDEN_COMPRA = " & NO
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If .EOF And .BOF Then
            MsgBox "FALLA GRAVE DE INFORMACION, LLAME A SOPORTE", vbCritical, "SACC"
        Else
            lblID.Caption = .Fields("ID_ORDEN_COMPRA")
            lblFolio.Caption = .Fields("NUM_ORDEN")
            If .Fields("TIPO") = "I" Then
                opnInternacional.Value = True
            ElseIf .Fields("TIPO") = "N" Then
                opnNacional.Value = True
            Else
                opnIndirecta.Value = True
            End If
            lblID.Caption = NO
            If Not IsNull(.Fields("DISCOUNT")) Then
                txtDescuento.Text = .Fields("DISCOUNT")
            Else
                txtDescuento.Text = "0"
            End If
            If Not IsNull(.Fields("TAX")) Then
                txtImpuesto.Text = .Fields("TAX")
            Else
                txtImpuesto.Text = "0"
            End If
            If Not IsNull(.Fields("FREIGHT")) Then
                txtFlete.Text = .Fields("FREIGHT")
            Else
                txtFlete.Text = "0"
            End If
            If Not IsNull(.Fields("OTROS_CARGOS")) Then
                txtCargos.Text = .Fields("OTROS_CARGOS")
            Else
                txtCargos.Text = "0"
            End If
            If Not IsNull(.Fields("TOTAL")) Then
                txtSubtotal.Text = .Fields("TOTAL")
            Else
                txtSubtotal.Text = "0"
            End If
            txtTotal.Text = Format(CDbl(txtSubtotal.Text - txtDescuento.Text + txtImpuesto.Text + txtFlete.Text + txtCargos.Text), "###,###,##0.00")
            If Not IsNull(.Fields("ENVIARA")) Then txtEnviara.Text = .Fields("ENVIARA")
            If Not IsNull(.Fields("COMENTARIO")) Then txtComentarios.Text = .Fields("COMENTARIO")
            Subtotal = 0
            lvwCotizaciones.ListItems.Clear
            sBuscar = "SELECT * FROM ORDEN_COMPRA_DETALLE WHERE ID_ORDEN_COMPRA = " & NO
            Set tRs2 = cnn.Execute(sBuscar)
            With tRs2
                If Not (.EOF And .BOF) Then
                    Do While Not .EOF
                        Set tLi = lvwCotizaciones.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                        If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = Trim(.Fields("Descripcion"))
                        If Not IsNull(.Fields("CANTIDAD")) Then tLi.SubItems(2) = Trim(.Fields("CANTIDAD"))
                        If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(3) = Trim(.Fields("PRECIO"))
                        tLi.SubItems(4) = CDbl(.Fields("PRECIO")) * CDbl(.Fields("CANTIDAD"))
                         Subtotal = Subtotal + (CDbl(.Fields("PRECIO")) * CDbl(.Fields("CANTIDAD")))
                        .MoveNext
                    Loop
                End If
            End With
           txtSubtotal.Text = Subtotal
        End If
    End With
End Sub
Private Sub lvwOCNacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoOrden = Item
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtCargos_Change()
    'If opnInternacional.Value = False Then
    '    If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" Then
    '        txtImpuesto.Text = Format(((Val(txtSubtotal.Text) - Val(txtDescuento.Text)) + (CDbl(txtFlete.Text))) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    '    End If
    'End If
    'If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" And txtImpuesto.Text <> "" And txtCargos.Text <> "" Then
    '    txtTotal = (Val(txtSubtotal.Text) - Val(txtDescuento.Text)) + (Val(txtImpuesto.Text) + Val(txtFlete.Text) + Val(txtCargos.Text))
    'End If
End Sub
Private Sub txtDescuento_Change()
    Dim IVA As String
    If opnInternacional.Value = False Then
        If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" And txtImpuesto.Text <> "" And txtCargos.Text <> "" Then
            IVA = (Val(txtSubtotal.Text) - Val(txtDescuento.Text) + (Val(txtFlete.Text))) * CDbl(CDbl(VarMen.Text4(7).Text) / 100)
            txtImpuesto.Text = Format(IVA, "###,###,##0.00")
        End If
    End If
    If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" And txtImpuesto.Text <> "" And txtCargos.Text <> "" Then
        txtTotal = (Val(txtSubtotal.Text) - Val(txtDescuento.Text)) + (Val(txtImpuesto.Text) + Val(txtFlete.Text) + Val(txtCargos.Text))
    End If
End Sub
Private Sub txtFlete_Change()
    If opnInternacional.Value = False Then
        If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" And txtImpuesto.Text <> "" And txtCargos.Text <> "" Then
            txtImpuesto.Text = Format((Val(txtSubtotal.Text) - Val(txtDescuento.Text) + (Val(txtFlete.Text))) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
        End If
    End If
    If txtSubtotal.Text <> "" And txtDescuento.Text <> "" And txtFlete.Text <> "" And txtImpuesto.Text <> "" And txtCargos.Text <> "" Then
        txtTotal = (Val(txtSubtotal.Text) - Val(txtDescuento.Text)) + (Val(txtImpuesto.Text) + Val(txtFlete.Text) + Val(txtCargos.Text))
    End If
End Sub
Private Sub txtSubtotal_Change()
    'If opnInternacional.Value = False Then
    '    txtImpuesto.Text = Format((Val(txtSubtotal.Text) - Val(txtDescuento.Text) + Val(txtFlete.Text) + Val(txtCargos.Text)) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
    'End If
    'txtTotal = (Val(txtSubtotal.Text) - Val(txtDescuento.Text)) + (Val(txtImpuesto.Text) + Val(txtFlete.Text) + Val(txtCargos.Text))
End Sub
