VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AltaClien 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTA DE CLIENTE"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   57
      Top             =   3000
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "AltaClien.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "AltaClien.frx":030A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label26 
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
         TabIndex        =   58
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8400
      TabIndex        =   55
      Top             =   1800
      Width           =   975
      Begin VB.Label Label23 
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
         TabIndex        =   56
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "AltaClien.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "AltaClien.frx":26F6
         Top             =   240
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "AltaClien.frx":40B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label28"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label29"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label30"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(12)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(15)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(18)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Combo7"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo8"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "AltaClien.frx":40D4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(7)"
      Tab(1).Control(1)=   "Text1(9)"
      Tab(1).Control(2)=   "Text1(10)"
      Tab(1).Control(3)=   "Text1(11)"
      Tab(1).Control(4)=   "Text1(13)"
      Tab(1).Control(5)=   "Text1(14)"
      Tab(1).Control(6)=   "Text1(17)"
      Tab(1).Control(7)=   "Text1(19)"
      Tab(1).Control(8)=   "Text1(20)"
      Tab(1).Control(9)=   "Combo3"
      Tab(1).Control(10)=   "COLONIA"
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(15)=   "Label16"
      Tab(1).Control(16)=   "Label17"
      Tab(1).Control(17)=   "Label19"
      Tab(1).Control(18)=   "Label20"
      Tab(1).Control(19)=   "Label22"
      Tab(1).Control(20)=   "Label25"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Crédito"
      TabPicture(2)   =   "AltaClien.frx":40F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text2"
      Tab(2).Control(1)=   "Combo4"
      Tab(2).Control(2)=   "Check1"
      Tab(2).Control(3)=   "Text1(23)"
      Tab(2).Control(4)=   "Text1(16)"
      Tab(2).Control(5)=   "Combo1"
      Tab(2).Control(6)=   "Combo2"
      Tab(2).Control(7)=   "Label27"
      Tab(2).Control(8)=   "Label40"
      Tab(2).Control(9)=   "Comentarios"
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(11)=   "Label14"
      Tab(2).Control(12)=   "Label9"
      Tab(2).ControlCount=   13
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   4440
         TabIndex        =   65
         Top             =   1680
         Width           =   3615
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   6720
         TabIndex        =   63
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74640
         TabIndex        =   25
         Top             =   3480
         Width           =   2775
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3840
         TabIndex        =   10
         Top             =   3480
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   -74640
         TabIndex        =   24
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar leyendas en Facturas"
         Height          =   255
         Left            =   -71400
         TabIndex        =   27
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   2085
         Index           =   23
         Left            =   -71400
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -74640
         MaxLength       =   8
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74640
         TabIndex        =   22
         Text            =   "0"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74640
         TabIndex        =   23
         Text            =   "0"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -71880
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -74640
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -69240
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   -74640
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -70320
         MaxLength       =   18
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -68880
         MaxLength       =   5
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -71160
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   -74640
         MaxLength       =   100
         TabIndex        =   19
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   -69240
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -74640
         TabIndex        =   11
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CommandButton COLONIA 
         Caption         =   "Colonia Nueva"
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
         Left            =   -70800
         Picture         =   "AltaClien.frx":410C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   5880
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   120
         MaxLength       =   150
         TabIndex        =   0
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   150
         TabIndex        =   1
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Régimen Fiscal"
         Height          =   255
         Left            =   4440
         TabIndex        =   66
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Régimen Capital"
         Height          =   255
         Left            =   6720
         TabIndex        =   64
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "* Uso CFDi"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Cuenta Bancaria"
         Height          =   255
         Left            =   -74640
         TabIndex        =   61
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "Asignar Agente"
         Height          =   255
         Left            =   3840
         TabIndex        =   60
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label40 
         Caption         =   "Descuento por Tipo "
         Height          =   255
         Left            =   -74640
         TabIndex        =   59
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Comentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         Height          =   195
         Left            =   -71400
         TabIndex        =   54
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Limite de Crédito"
         Height          =   195
         Left            =   -74640
         TabIndex        =   53
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   -74640
         TabIndex        =   52
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dias Crédito"
         Height          =   195
         Left            =   -74640
         TabIndex        =   51
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Pais"
         Height          =   195
         Left            =   -69240
         TabIndex        =   50
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "* Dirección"
         Height          =   195
         Left            =   -74640
         TabIndex        =   49
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Ciudad"
         Height          =   195
         Left            =   -74640
         TabIndex        =   48
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "* Colonia"
         Height          =   195
         Left            =   -74640
         TabIndex        =   47
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Numero Exterior"
         Height          =   195
         Left            =   -70320
         TabIndex        =   46
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Numero Interior"
         Height          =   195
         Left            =   -68760
         TabIndex        =   45
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "* Codigo Postal"
         Height          =   195
         Left            =   -69240
         TabIndex        =   44
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dirección de Correo Electronico"
         Height          =   195
         Left            =   -71160
         TabIndex        =   43
         Top             =   2640
         Width           =   2250
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "* Estado"
         Height          =   195
         Left            =   -71880
         TabIndex        =   42
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Direccion Envio"
         Height          =   195
         Left            =   -74640
         TabIndex        =   41
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña de Web"
         Height          =   195
         Left            =   5880
         TabIndex        =   40
         Top             =   2640
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CURP"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   3840
         TabIndex        =   37
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Telefono Trabajo"
         Height          =   195
         Left            =   2040
         TabIndex        =   36
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Telefono Casa"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* R.F.C"
         Height          =   195
         Left            =   5400
         TabIndex        =   34
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre Comercial"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Nombre "
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave Cliente"
         Height          =   195
         Left            =   5880
         TabIndex        =   31
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "AltaClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ValDes As String
Private Sub Image8_Click()
On Error GoTo ManejaError
    If Text1(0).Text = "" Then
        If Text1(1).Text <> "" And Text1(15).Text <> "" And Text1(2).Text <> "" And Text1(3).Text <> "" And Combo3.Text <> "" And Text1(10).Text <> "" And Text1(11).Text <> "" And Text1(9).Text <> "" And Text1(7).Text <> "" And Combo6.Text <> "" And Combo8.Text <> "" Then
            Dim MostLey As String
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            Dim UsoCFDi As String
            Dim RegimenFiscal As String
            If Combo6.Text <> "" Then
                sBuscar = "SELECT Clave FROM SATUsoCFDi WHERE Descripcion = '" & Combo6.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    UsoCFDi = tRs.Fields("Clave")
                End If
            Else
                MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTREO", vbExclamation, "SACC"
                Exit Sub
            End If
            If Combo8.Text <> "" Then
                sBuscar = "SELECT Clave FROM SATRegimenFiscal WHERE Descripcion = '" & Combo8.Text & "'"
                Set tRs = cnn.Execute(sBuscar)
                If Not (tRs.EOF And tRs.BOF) Then
                    RegimenFiscal = tRs.Fields("Clave")
                Else
                    MsgBox "NO SE ENCONTRÒ EL REGIMEN FISCAL SELECIONADO", vbExclamation, "SACC"
                    Exit Sub
                End If
            Else
                MsgBox "FALTA INFORMACION NECESARIA PARA EL REGISTREO", vbExclamation, "SACC"
                Exit Sub
            End If
            If Combo1.Text = "" Then
                Combo1.Text = "0"
            End If
            If ValDes = "" Then
                ValDes = "0"
            End If
            If Check1.Value = 1 Then
                MostLey = "S"
            Else
                MostLey = "N"
            End If
            Text1(16).Text = Replace(Text1(16).Text, ",", "")
            ValDes = Replace(ValDes, ",", "")
            If VarMen.Text1(57).Text = "S" Then
                sBuscar = "INSERT INTO CLIENTE (CP, DIRECCION, NOMBRE, CONTACTO, CURP, NOMBRE_COMERCIAL, RFC, TELEFONO_CASA, TELEFONO_TRABAJO, FAX, WEB_PASSWORD, COLONIA, NUMERO_INTERIOR, NUMERO_EXTERIOR, CIUDAD, ESTADO, PAIS, DIRECCION_ENVIO, EMAIL, ID_AGENTE, LIMITE_CREDITO, DIAS_CREDITO, DESCUENTO, FECHA_ALTA, COMENTARIOS, LEYENDAS, ID_DESCUENTO, VALORACION, AGENTE,ASIG, NUM_CUENTA_PAGO_CLIENTE, UsoCFDi, RegimenCapital, RegimenFiscal) VALUES ('"
                sBuscar = sBuscar & Text1(10).Text & "', '" & Text1(11).Text & "', '" & Text1(15).Text & "', '" & Text1(8).Text & "', '" & Text1(12).Text & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(4).Text & "', '" & Text1(5).Text & "', '" & Text1(18).Text & "', '" & Combo3.Text & "', '" & Text1(14).Text & "', '" & Text1(13).Text & "', '" & Text1(9).Text & "', '" & Text1(7).Text & "', '" & Text1(20).Text & "', '" & Text1(19).Text & "', '" & Text1(17).Text & "', '" & VarMen.Text1(0).Text & "', " & Text1(16).Text & ", '" & Combo1.Text & "', " & ValDes & ", '" & Format(Date, "dd/mm/yyyy") & "', '" & Text1(23).Text & "', '" & MostLey & "', '" & Combo4.Text & "', 'R','" & Combo5.Text & "', 'S', '" & Text2.Text & "', '" & UsoCFDi & "', '" & Combo7.Text & "', '" & RegimenFiscal & "');"
            Else
                sBuscar = "INSERT INTO CLIENTE (CP, DIRECCION, NOMBRE, CONTACTO, CURP, NOMBRE_COMERCIAL, RFC, TELEFONO_CASA, TELEFONO_TRABAJO, FAX, WEB_PASSWORD, COLONIA, NUMERO_INTERIOR, NUMERO_EXTERIOR, CIUDAD, ESTADO, PAIS, DIRECCION_ENVIO, EMAIL, ID_AGENTE, FECHA_ALTA, COMENTARIOS, LEYENDAS, VALORACION, AGENTE, ASIG, UsoCFDi, RegimenCapital, RegimenFiscal) VALUES ('"
                sBuscar = sBuscar & Text1(10).Text & "', '" & Text1(11).Text & "', '" & Text1(15).Text & "', '" & Text1(8).Text & "', '" & Text1(12).Text & "', '" & Text1(1).Text & "', '" & Text1(2).Text & "', '" & Text1(3).Text & "', '" & Text1(4).Text & "', '" & Text1(5).Text & "', '" & Text1(18).Text & "', '" & Combo3.Text & "', '" & Text1(14).Text & "', '" & Text1(13).Text & "', '" & Text1(9).Text & "', '" & Text1(7).Text & "', '" & Text1(20).Text & "', '" & Text1(19).Text & "', '" & Text1(17).Text & "', '" & VarMen.Text1(0).Text & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & Text1(23).Text & "', 'S', 'R','" & Combo5.Text & "', 'S', '" & UsoCFDi & "', '" & Combo7.Text & "', '" & RegimenFiscal & "');"
            End If
            cnn.Execute (sBuscar)
            Text1(15).Text = ""
            Text1(12).Text = ""
            Text1(1).Text = ""
            Text1(8).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(5).Text = ""
            Text1(18).Text = ""
            Combo3.Text = ""
            Text1(10).Text = ""
            Text1(11).Text = ""
            Text1(14).Text = ""
            Text1(13).Text = ""
            Text1(9).Text = ""
            Text1(7).Text = ""
            Text1(20).Text = ""
            Text1(19).Text = ""
            Text1(17).Text = ""
            Text1(23).Text = ""
            Text1(16).Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo4.Text = ""
            Combo7.Text = ""
            Combo8.Text = ""
        Else
            MsgBox "Debe proporcionar toda la información marcada con asteriscos (*)", vbExclamation, "SACC"
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub COLONIA_Click()
On Error GoTo ManejaError
    FrmAgrColonia.Show vbModal
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Combo1.Clear
    Combo1.AddItem "0"
    Combo1.AddItem "15"
    Combo1.AddItem "30"
    Combo1.AddItem "45"
    Combo1.AddItem "60"
    Combo1.AddItem "90"
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Combo2_DropDown()
On Error GoTo ManejaError
    Combo2.Clear
    Combo2.AddItem "50 %"
    Combo2.AddItem "40 %"
    Combo2.AddItem "30 %"
    Combo2.AddItem "15 %"
    Combo2.AddItem "14 %"
    Combo2.AddItem "11 %"
    Combo2.AddItem "5 %"
    Combo2.AddItem "0 %"
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo2_Click()
On Error GoTo ManejaError
    Select Case Combo2.ListIndex
        Case Is = 0: ValDes = "50"
        Case Is = 1: ValDes = "40"
        Case Is = 2: ValDes = "30"
        Case Is = 3: ValDes = "25"
        Case Is = 3: ValDes = "20"
        Case Is = 3: ValDes = "15"
        Case Is = 4: ValDes = "14"
        Case Is = 5: ValDes = "11"
        Case Is = 6: ValDes = "5"
        Case Is = 7: ValDes = "0"
    End Select
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HFFE1E1
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &H80000005
End Sub
Private Sub Combo3_DropDown()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo3.Clear
    sBuscar = "SELECT NOMBRE FROM COLONIAS ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("NOMBRE")) Then
                Combo3.AddItem tRs.Fields("NOMBRE")
            End If
            tRs.MoveNext
        Loop
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo3_GotFocus()
    Combo3.BackColor = &HFFE1E1
End Sub
Private Sub Combo3_LostFocus()
    Combo3.BackColor = &H80000005
End Sub

Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo ManejaError
    If Index = 0 And Text1(0).Text = "" Then
        Text1(15).Locked = True
        Text1(12).Locked = True
        Text1(1).Locked = True
        Text1(8).Locked = True
        Text1(2).Locked = True
        Text1(3).Locked = True
        Text1(4).Locked = True
        Text1(5).Locked = True
        Text1(18).Locked = True
        Combo3.Locked = True
        Text1(10).Locked = True
        Text1(11).Locked = True
        Text1(14).Locked = True
        Text1(13).Locked = True
        Text1(9).Locked = True
        Text1(7).Locked = True
        Text1(20).Locked = True
        Text1(19).Locked = True
        Text1(17).Locked = True
        Text1(23).Locked = True
        Text1(16).Locked = True
        Combo1.Locked = True
        Combo2.Locked = True
    Else
        Text1(15).Locked = False
        Text1(12).Locked = False
        Text1(1).Locked = False
        Text1(8).Locked = False
        Text1(2).Locked = False
        Text1(3).Locked = False
        Text1(4).Locked = False
        Text1(5).Locked = False
        Text1(18).Locked = False
        Combo3.Locked = False
        Text1(10).Locked = False
        Text1(11).Locked = False
        Text1(14).Locked = False
        Text1(13).Locked = False
        Text1(9).Locked = False
        Text1(7).Locked = False
        Text1(20).Locked = False
        Text1(19).Locked = False
        Text1(17).Locked = False
        Text1(23).Locked = False
        Text1(16).Locked = False
        Combo1.Locked = False
        Combo2.Locked = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo ManejaError
    Text1(Index).BackColor = &HFFE1E1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    If Index = 3 Or Index = 4 Or Index = 5 Then
        Valido = "1234567890-()"
    Else
        If Index = 16 Or Index = 3 Or Index = 4 Or Index = 5 Then
            Valido = "1234567890."
        Else
            If Index = 17 Then
                Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz _-@.,;&"
            Else
                Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%@!?*+"
            End If
        End If
    End If
    If Index = 17 Then
        KeyAscii = Asc(Chr(KeyAscii))
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
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
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
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
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    sBuscar = "SELECT ID_DESCUENTO FROM DESCUENTOS GROUP BY ID_DESCUENTO ORDER BY ID_DESCUENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo4.AddItem tRs.Fields("ID_DESCUENTO")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT Descripcion FROM SATUsoCFDi ORDER BY Descripcion"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo6.AddItem tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT RegimenCapital FROM CLIENTE WHERE  RegimenCapital <> '' GROUP BY RegimenCapital ORDER BY RegimenCapital"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("RegimenCapital")) Then Combo7.AddItem tRs.Fields("RegimenCapital")
            tRs.MoveNext
        Loop
    End If
    sBuscar = "SELECT Descripcion FROM SATRegimenFiscal ORDER BY Descripcion"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("Descripcion")) Then Combo8.AddItem tRs.Fields("Descripcion")
            tRs.MoveNext
        Loop
    End If
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Combo5_DropDown()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Combo5.Clear
    sBuscar = "SELECT NOMBRE FROM USUARIOS WHERE PE7 = 'S' AND ESTADO = 'A' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Combo5.AddItem "<TODAS>"
     If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Combo5.AddItem tRs.Fields("NOMBRE")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890-"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
