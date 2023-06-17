VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTickets 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tickets"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   40
      Top             =   4200
      Width           =   975
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FrmTickets.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmTickets.frx":030A
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label28 
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
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   39
      Top             =   5400
      Width           =   975
      Begin VB.Label Label12 
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
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmTickets.frx":0899
         MousePointer    =   99  'Custom
         Picture         =   "FrmTickets.frx":0BA3
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   21
      Top             =   6600
      Width           =   975
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmTickets.frx":26E5
         MousePointer    =   99  'Custom
         Picture         =   "FrmTickets.frx":29EF
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Tickets"
      TabPicture(0)   =   "FrmTickets.frx":4AD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Nuevo"
      TabPicture(1)   =   "FrmTickets.frx":4AED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Combo1"
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(5)=   "Text3"
      Tab(1).Control(6)=   "Command3"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Historial"
      TabPicture(2)   =   "FrmTickets.frx":4B09
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "ListView2"
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(4)=   "DTPicker1"
      Tab(2).Control(5)=   "Command4"
      Tab(2).Control(6)=   "Picture2"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Reporte"
      TabPicture(3)   =   "FrmTickets.frx":4B25
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(1)=   "Label14"
      Tab(3).Control(2)=   "Label15"
      Tab(3).Control(3)=   "Label16"
      Tab(3).Control(4)=   "DTPicker4"
      Tab(3).Control(5)=   "DTPicker3"
      Tab(3).Control(6)=   "Combo2"
      Tab(3).Control(7)=   "ListView3"
      Tab(3).Control(8)=   "Command5"
      Tab(3).Control(9)=   "Combo3"
      Tab(3).Control(10)=   "Command6"
      Tab(3).ControlCount=   11
      Begin VB.CommandButton Command6 
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
         Left            =   -66480
         Picture         =   "FrmTickets.frx":4B41
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   7080
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -69480
         TabIndex        =   43
         Top             =   7080
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
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
         Left            =   -66480
         Picture         =   "FrmTickets.frx":7513
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   900
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   20
         Top             =   1380
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9763
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -73080
         TabIndex        =   18
         Top             =   900
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   "72 Horas"
         Height          =   255
         Left            =   7440
         TabIndex        =   5
         Top             =   7140
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "48 Horas"
         Height          =   255
         Left            =   7440
         TabIndex        =   4
         Top             =   6900
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "24 Horas"
         Height          =   255
         Left            =   7440
         TabIndex        =   3
         Top             =   6660
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         Height          =   3135
         Left            =   -74880
         ScaleHeight     =   3075
         ScaleWidth      =   9555
         TabIndex        =   26
         Top             =   4380
         Width           =   9615
         Begin VB.Label Label11 
            Caption         =   "Label11"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   7440
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   7215
         End
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
         Left            =   -71040
         Picture         =   "FrmTickets.frx":9EE5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   660
         Width           =   1095
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
         Left            =   -66600
         Picture         =   "FrmTickets.frx":C8B7
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   1965
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1860
         Width           =   8055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73560
         TabIndex        =   9
         Top             =   1380
         Width           =   8055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -72840
         TabIndex        =   8
         Top             =   780
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
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
         Left            =   8640
         Picture         =   "FrmTickets.frx":F289
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7140
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Responder"
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
         Left            =   8640
         Picture         =   "FrmTickets.frx":11C5B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6660
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   6660
         Width           =   7215
      End
      Begin VB.PictureBox Picture1 
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   9555
         TabIndex        =   1
         Top             =   3180
         Width           =   9615
         Begin VB.Label Label3 
            Caption         =   "Label3"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   7440
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   7215
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   12
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   44711
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         Top             =   780
         Width           =   9615
         _ExtentX        =   16960
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   -72480
         TabIndex        =   13
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   44711
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   15
         Top             =   1140
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   16
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   44711
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   -72480
         TabIndex        =   17
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   44711
      End
      Begin VB.Label Label16 
         Caption         =   "Nuevo Departamento"
         Height          =   255
         Left            =   -71160
         TabIndex        =   44
         Top             =   7080
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Del :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Al :"
         Height          =   255
         Left            =   -72840
         TabIndex        =   37
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Departamento destino"
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Al :"
         Height          =   255
         Left            =   -72840
         TabIndex        =   33
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Del :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Mensaje :"
         Height          =   255
         Left            =   -74520
         TabIndex        =   31
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Título :"
         Height          =   255
         Left            =   -74520
         TabIndex        =   30
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Departamento destino"
         Height          =   255
         Left            =   -74520
         TabIndex        =   29
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Respuesta"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   6420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Abiertos"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   540
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   10200
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim NoTicket As String
Dim IdUsuario As String
Dim DepaDest As String
Dim iScrollMax As Integer
Dim IdTicket As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
    Dim sBuscar As String
    If NoTicket <> "" Then
        sBuscar = "UPDATE TICKETS SET ESTATUS = 'F', FECHA_CIERRE = GETDATE(), ID_USUARIO_CIERRE = '" & VarMen.Text1(0).Text & "' WHERE ID_TICKET = '" & NoTicket & "'"
        cnn.Execute (sBuscar)
        Actualiza
    End If
End Sub
Private Sub Command2_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If NoTicket <> "" Then
        sBuscar = "INSERT INTO TICKETS_DETALLE (ID_TICKET, MENSAJE, ID_USUARIO, ID_USUARIO_RESPONDE) VALUES ('" & NoTicket & "', '" & Text1.Text & "', '" & VarMen.Text1(0).Text & "', '" & IdUsuario & "')"
        cnn.Execute (sBuscar)
        sBuscar = "UPDATE TICKETS SET FECHA_ULTIMA_RESPUESTA = GETDATE(), CONTADOR_MENSAJES = CONTADOR_MENSAJES + 1, ID_USUARIO_RESPONDIO = '" & VarMen.Text1(0).Text & "' WHERE ID_TICKET = '" & NoTicket & "'"
        cnn.Execute (sBuscar)
        sBuscar = "SELECT ID_USUARIO_ATIENDE FROM TICKETS WHERE ID_TICKET = '" & NoTicket & "' AND ID_USUARIO_ATIENDE IS NULL"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            sBuscar = "UPDATE TICKETS SET ID_USUARIO_ATIENDE = '" & VarMen.Text1(0).Text & "', ESTATUS = 'A' WHERE ID_TICKET = '" & NoTicket & "'"
            cnn.Execute (sBuscar)
        End If
        sBuscar = "SELECT TIEMPO_RESPUESTA FROM TICKETS WHERE ID_TICKET = '" & NoTicket & "' AND TIEMPO_RESPUESTA IS NULL"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            If Option1.Value Then
                sBuscar = "UPDATE TICKETS SET TIEMPO_RESPUESTA = '24' WHERE ID_TICKET = '" & NoTicket & "'"
            End If
            If Option2.Value Then
                sBuscar = "UPDATE TICKETS SET TIEMPO_RESPUESTA = '48' WHERE ID_TICKET = '" & NoTicket & "'"
            End If
            If Option3.Value Then
                sBuscar = "UPDATE TICKETS SET TIEMPO_RESPUESTA = '72' WHERE ID_TICKET = '" & NoTicket & "'"
            End If
            cnn.Execute (sBuscar)
        End If
        Text1.Text = ""
    End If
    RecargaMensajes
End Sub
Private Sub Command3_Click()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If Combo1.Text <> "" Then
        If Text2.Text <> "" And Text3.Text <> "" Then
            sBuscar = "INSERT INTO TICKETS (ID_USUARIO, TITULO, DEPARTAMENTO_DESTINO) VALUES ('" & VarMen.Text1(0).Text & "', '" & Text2.Text & "', '" & Combo1.Text & "')"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT TOP 1 ID_TICKET FROM TICKETS ORDER BY ID_TICKET DESC"
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                sBuscar = "INSERT INTO TICKETS_DETALLE (ID_TICKET, MENSAJE, ID_USUARIO) VALUES ('" & tRs.Fields("ID_TICKET") & "', '" & Text3.Text & "', '" & VarMen.Text1(0).Text & "')"
                cnn.Execute (sBuscar)
            End If
            Text2.Text = ""
            Text3.Text = ""
            Combo1.Text = ""
        Else
            MsgBox "DEBE DAR UN ASUNTO Y MENSAJE AL TICKET", vbExclamation, "SACC"
        End If
    Else
        MsgBox "DEBE DAR UN DEPARTAMENTO DESTINO DEL TICKET", vbExclamation, "SACC"
    End If
    'End If
End Sub
Private Sub Command4_Click()
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView2.ListItems.Clear
    'sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.ID_USUARIO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, TICKETS.TITULO, TICKETS.ESTATUS, TICKETS.DEPARTAMENTO_DESTINO FROM TICKETS INNER JOIN USUARIOS ON TICKETS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (TICKETS.DEPARTAMENTO_DESTINO = '" & VarMen.Text1(75).Text & "' AND TICKETS.ESTATUS NOT IN ('F')) OR TICKETS.ID_USUARIO_ATIENDE = '" & VarMen.Text1(0).Text & "' ORDER BY TICKETS.ID_TICKET DESC"
    sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.ID_USUARIO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, TICKETS.TITULO, TICKETS.ESTATUS, TICKETS.DEPARTAMENTO_DESTINO FROM TICKETS INNER JOIN USUARIOS ON TICKETS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (TICKETS.DEPARTAMENTO_DESTINO = '" & VarMen.Text1(75).Text & "') AND (TICKETS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') OR (TICKETS.ID_USUARIO = '" & VarMen.Text1(0).Text & "') AND (TICKETS.FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & "') ORDER BY TICKETS.ID_TICKET DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_TICKET"))
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(1) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_USUARIO")) Then tLi.SubItems(2) = tRs.Fields("ID_USUARIO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
            If Not IsNull(tRs.Fields("TITULO")) Then tLi.SubItems(4) = tRs.Fields("TITULO")
            If Not IsNull(tRs.Fields("ESTATUS")) Then
                If tRs.Fields("ESTATUS") = "I" Then
                    tLi.SubItems(5) = "NUEVO"
                Else
                    If tRs.Fields("ESTATUS") = "F" Then
                        tLi.SubItems(5) = "FINALIZADO"
                    Else
                        tLi.SubItems(5) = "EN PROCESO"
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("DEPARTAMENTO_DESTINO")) Then tLi.SubItems(6) = tRs.Fields("DEPARTAMENTO_DESTINO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command5_Click()
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView3.ListItems.Clear
    sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.TITULO, COUNT(TICKETS_DETALLE.ID_TICKET) AS RESPUESTAS, ISNULL(USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS, 'SIN RESPUESTA') AS ATIENDE, USUARIOS_1.NOMBRE + ' ' + USUARIOS_1.APELLIDOS AS SOLICITO, TICKETS.DEPARTAMENTO_DESTINO, TICKETS.ESTATUS, TICKETS.FECHA_CIERRE, TICKETS.TIEMPO_RESPUESTA FROM TICKETS INNER JOIN TICKETS_DETALLE ON TICKETS.ID_TICKET = TICKETS_DETALLE.ID_TICKET LEFT OUTER JOIN USUARIOS ON TICKETS.ID_USUARIO_ATIENDE = USUARIOS.ID_USUARIO INNER JOIN USUARIOS AS USUARIOS_1 ON TICKETS.ID_USUARIO = USUARIOS_1.ID_USUARIO WHERE (TICKETS.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " 23:59:59.999') "
    If Combo2.Text <> "" And Combo2.Text <> "<TODOS>" Then
        sBuscar = sBuscar & "AND TICKETS.DEPARTAMENTO_DESTINO = '" & Combo2.Text & "' "
    End If
    sBuscar = sBuscar & "GROUP BY TICKETS.FECHA, TICKETS.TITULO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, USUARIOS_1.NOMBRE, USUARIOS_1.APELLIDOS, TICKETS.DEPARTAMENTO_DESTINO, TICKETS.ID_TICKET, TICKETS.ESTATUS , TICKETS.FECHA_CIERRE, TICKETS.TIEMPO_RESPUESTA"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_TICKET"))
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(1) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("TITULO")) Then tLi.SubItems(2) = tRs.Fields("TITULO")
            If Not IsNull(tRs.Fields("RESPUESTAS")) Then tLi.SubItems(3) = tRs.Fields("RESPUESTAS")
            If Not IsNull(tRs.Fields("ATIENDE")) Then tLi.SubItems(4) = tRs.Fields("ATIENDE")
            If Not IsNull(tRs.Fields("SOLICITO")) Then tLi.SubItems(5) = tRs.Fields("SOLICITO")
            If Not IsNull(tRs.Fields("DEPARTAMENTO_DESTINO")) Then tLi.SubItems(6) = tRs.Fields("DEPARTAMENTO_DESTINO")
            If Not IsNull(tRs.Fields("ESTATUS")) Then
                If tRs.Fields("ESTATUS") = "I" Then
                    tLi.SubItems(7) = "NUEVO"
                Else
                    If tRs.Fields("ESTATUS") = "F" Then
                        tLi.SubItems(7) = "FINALIZADO"
                    Else
                        tLi.SubItems(7) = "EN PROCESO"
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("FECHA_CIERRE")) Then tLi.SubItems(8) = tRs.Fields("FECHA_CIERRE")
            If Not IsNull(tRs.Fields("TIEMPO_RESPUESTA")) Then tLi.SubItems(9) = tRs.Fields("TIEMPO_RESPUESTA")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Command6_Click()
    If IdTicket <> "" And Combo3.Text <> "" Then
        Dim sBuscar As String
        sBuscar = "UPDATE TICKETS SET  DEPARTAMENTO_DESTINO = '" & Combo3.Text & "' WHERE ID_TICKET = " & IdTicket
        cnn.Execute (sBuscar)
        IdTicket = ""
        Combo3.Text = ""
        Command5.Value = True
    Else
        MsgBox "FALTA INFORMACIÒN NECESARIA PARA EL REGISTRO", vbExclamation, "SACC"
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FunMoveLabel (KeyCode)
End Sub
Private Sub Image10_Click()
On Error GoTo ManejaError
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
    Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image26_Click()
On Error GoTo ManejaError
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim ConPag As Integer
    ConPag = 1
    totalgen = 0
    totalgenpro = 0
    Image1.Picture = LoadPicture(App.Path & "\REPORTES\" & NvoMen.TxtBaseDatos.Text & ".JPG")
    If Not oDoc.PDFCreate(App.Path & "\RepCuentasPagadas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Horizontal
    oDoc.WImage 70, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 1, 20, 840, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 1, 20, 840, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 8, hCenter
    oDoc.WTextBox 60, 420, 20, 420, "Fecha del " & DTPicker3.Value & " al " & DTPicker4.Value, "F3", 8, hCenter
    oDoc.WTextBox 70, 1, 20, 840, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 8, hCenter
    oDoc.WTextBox 80, 1, 20, 840, "Tel " & VarMen.TxtEmp(2).Text, "F2", 8, hCenter
    oDoc.WTextBox 90, 1, 20, 840, "Reporte de Tickets", "F2", 10, hCenter
    oDoc.WTextBox 70, 420, 20, 420, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 90, 420, 20, 420, Date, "F3", 8, hCenter
    ' Encabezado de pagina
    oDoc.WTextBox 110, 10, 20, 50, "Folio", "F2", 10, hLeft
    oDoc.WTextBox 110, 40, 20, 50, "Fecha", "F2", 10, hCenter
    oDoc.WTextBox 110, 90, 20, 270, "Titulo", "F2", 10, hCenter
    oDoc.WTextBox 110, 360, 20, 70, "# Respuestas", "F2", 10, hCenter
    oDoc.WTextBox 110, 430, 20, 100, "Atiende", "F2", 10, hCenter
    oDoc.WTextBox 110, 530, 20, 100, "Solicito", "F2", 10, hCenter
    oDoc.WTextBox 110, 620, 20, 70, "Dpto. Destino", "F2", 10, hCenter
    oDoc.WTextBox 110, 690, 20, 50, "Estatus", "F2", 10, hCenter
    oDoc.WTextBox 110, 740, 20, 70, "Fecha Cierre", "F2", 10, hCenter
    'oDoc.WTextBox 110, 720, 20, 100, "Tiempo Respuesta", "F2", 10, hCenter
    ' Cuerpo del reporte
    sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.TITULO, COUNT(TICKETS_DETALLE.ID_TICKET) AS RESPUESTAS, ISNULL(USUARIOS.NOMBRE + ' ' + USUARIOS.APELLIDOS, 'SIN RESPUESTA') AS ATIENDE, USUARIOS_1.NOMBRE + ' ' + USUARIOS_1.APELLIDOS AS SOLICITO, TICKETS.DEPARTAMENTO_DESTINO, TICKETS.ESTATUS, TICKETS.FECHA_CIERRE, TICKETS.TIEMPO_RESPUESTA FROM TICKETS INNER JOIN TICKETS_DETALLE ON TICKETS.ID_TICKET = TICKETS_DETALLE.ID_TICKET LEFT OUTER JOIN USUARIOS ON TICKETS.ID_USUARIO_ATIENDE = USUARIOS.ID_USUARIO INNER JOIN USUARIOS AS USUARIOS_1 ON TICKETS.ID_USUARIO = USUARIOS_1.ID_USUARIO WHERE (TICKETS.FECHA BETWEEN '" & DTPicker3.Value & "' AND '" & DTPicker4.Value & " 23:59:59.999') "
    If Combo2.Text <> "" And Combo2.Text <> "<TODOS>" Then
        sBuscar = sBuscar & "AND TICKETS.DEPARTAMENTO_DESTINO = '" & Combo2.Text & "' "
    End If
    sBuscar = sBuscar & "GROUP BY TICKETS.FECHA, TICKETS.TITULO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, USUARIOS_1.NOMBRE, USUARIOS_1.APELLIDOS, TICKETS.DEPARTAMENTO_DESTINO, TICKETS.ID_TICKET, TICKETS.ESTATUS , TICKETS.FECHA_CIERRE, TICKETS.TIEMPO_RESPUESTA"
    Set tRs = cnn.Execute(sBuscar)
    Posi = 120
    Total = 0
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 100
    oDoc.WLineTo 840, 100
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 840, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Posi = Posi + 10
            If Not IsNull(tRs.Fields("ID_TICKET")) Then oDoc.WTextBox Posi, 10, 20, 50, tRs.Fields("ID_TICKET"), "F2", 8, hLeft
            If Not IsNull(tRs.Fields("FECHA")) Then oDoc.WTextBox Posi, 30, 20, 90, Format(tRs.Fields("FECHA"), "dd/mm/yyyy hh:MM AM/PM"), "F2", 8, hLeft
            If Not IsNull(tRs.Fields("TITULO")) Then oDoc.WTextBox Posi, 130, 20, 270, Mid(tRs.Fields("TITULO"), 1, 55), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("RESPUESTAS")) Then oDoc.WTextBox Posi, 400, 20, 70, tRs.Fields("RESPUESTAS"), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("ATIENDE")) Then oDoc.WTextBox Posi, 420, 20, 100, Mid(tRs.Fields("ATIENDE"), 1, 17), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("SOLICITO")) Then oDoc.WTextBox Posi, 520, 20, 100, Mid(tRs.Fields("SOLICITO"), 1, 17), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("DEPARTAMENTO_DESTINO")) Then oDoc.WTextBox Posi, 620, 20, 70, Mid(tRs.Fields("DEPARTAMENTO_DESTINO"), 1, 10), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("ESTATUS")) Then oDoc.WTextBox Posi, 690, 20, 50, tRs.Fields("ESTATUS"), "F3", 8, hLeft
            If Not IsNull(tRs.Fields("FECHA_CIERRE")) Then oDoc.WTextBox Posi, 740, 20, 70, Format(tRs.Fields("FECHA_CIERRE"), "dd/mm/yyyy hh:MM AM/PM"), "F3", 8, hLeft
            'If Not IsNull(tRs.Fields("TIEMPO_RESPUESTA")) Then oDoc.WTextBox Posi, 740, 20, 100, tRs.Fields("TIEMPO_RESPUESTA"), "F3", 8, hLeft
            If Posi >= 760 Then
                oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 120
                oDoc.WTextBox 40, 1, 20, 840, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 1, 20, 840, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 8, hCenter
                oDoc.WTextBox 60, 1, 20, 840, "Fecha del " & DTPicker3.Value & " al " & DTPicker4.Value, "F3", 8, hCenter
                oDoc.WTextBox 70, 1, 20, 840, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 8, hCenter
                oDoc.WTextBox 80, 1, 20, 840, "Tel " & VarMen.TxtEmp(2).Text, "F2", 8, hCenter
                oDoc.WTextBox 90, 1, 20, 840, "Reporte de Tickets", "F2", 10, hCenter
                oDoc.WTextBox 70, 420, 20, 420, "Fecha de Impresion", "F3", 8, hCenter
                oDoc.WTextBox 90, 420, 20, 420, Date, "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 110, 10, 20, 50, "Folio", "F2", 10, hLeft
                oDoc.WTextBox 110, 40, 20, 50, "Fecha", "F2", 10, hCenter
                oDoc.WTextBox 110, 90, 20, 270, "Titulo", "F2", 10, hCenter
                oDoc.WTextBox 110, 360, 20, 70, "# Respuestas", "F2", 10, hCenter
                oDoc.WTextBox 110, 400, 20, 100, "Atiende", "F2", 10, hCenter
                oDoc.WTextBox 110, 500, 20, 100, "Solicito", "F2", 10, hCenter
                oDoc.WTextBox 110, 620, 20, 70, "Dpto. Destino", "F2", 10, hCenter
                oDoc.WTextBox 110, 690, 20, 50, "Estatus", "F2", 10, hCenter
                oDoc.WTextBox 110, 740, 20, 70, "Fecha Cierre", "F2", 10, hCenter
                'oDoc.WTextBox 110, 720, 20, 100, "Tiempo Respuesta", "F2", 10, hCenter
                oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
                oDoc.MoveTo 10, 100
                oDoc.WLineTo 580, 100
                oDoc.LineStroke
                oDoc.MoveTo 10, 125
                oDoc.WLineTo 580, 125
                oDoc.LineStroke
            End If
            tRs.MoveNext
        Loop
    End If
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ManejaError:
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    FunMoveLabel (KeyCode)
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    IdTicket = Item
End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    FunMoveLabel (KeyCode)
End Sub
Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    FunMoveLabel2 (KeyCode)
End Sub
Private Sub FunMoveLabel(KeyCode As Integer)
    Select Case KeyCode
    Case 38
    If Label2(1).Top <> 120 Then
        For e = 1 To Label2.Count - 1
            Label2(e).Top = Label2(e).Top + 100
            Label3(e).Top = Label3(e).Top + 100
        Next e
    End If
    Case 40
    For e = 1 To Label2.Count - 1
        Label2(e).Top = Label2(e).Top - 100
        Label3(e).Top = Label3(e).Top - 100
    Next e
    End Select
End Sub
Private Sub FunMoveLabel2(KeyCode As Integer)
    Select Case KeyCode
    Case 38
    If Label10(1).Top <> 120 Then
        For e = 1 To Label10.Count - 1
            Label10(e).Top = Label10(e).Top + 100
            Label11(e).Top = Label11(e).Top + 100
        Next e
    End If
    Case 40
    For e = 1 To Label10.Count - 1
        Label10(e).Top = Label10(e).Top - 100
        Label11(e).Top = Label11(e).Top - 100
    Next e
    End Select
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dim tRs As ADODB.Recordset
    Dim sBus As String
    Dim tLi As ListItem
    Set cnn = New ADODB.Connection
    DTPicker1.Value = Date - 30
    DTPicker2.Value = Date
    DTPicker3.Value = Date - 30
    DTPicker4.Value = Date
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
        .ColumnHeaders.Add , , "# Ticket", 1000
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Id Usuario", 0
        .ColumnHeaders.Add , , "Usuario", 1200
        .ColumnHeaders.Add , , "Titulo", 4200
        .ColumnHeaders.Add , , "Estado", 1200
        .ColumnHeaders.Add , , "Departamento", 1200
        .ColumnHeaders.Add , , "Tiempo de Respuesta", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "# Ticket", 1000
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Id Usuario", 0
        .ColumnHeaders.Add , , "Usuario", 1200
        .ColumnHeaders.Add , , "Titulo", 4200
        .ColumnHeaders.Add , , "Estado", 1200
        .ColumnHeaders.Add , , "Departamento", 1200
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "# Ticket", 1000
        .ColumnHeaders.Add , , "Fecha Inicio", 1200
        .ColumnHeaders.Add , , "Tìtulo", 4000
        .ColumnHeaders.Add , , "# Mensajes", 1200
        .ColumnHeaders.Add , , "Atiende", 1200
        .ColumnHeaders.Add , , "Solicitò", 1200
        .ColumnHeaders.Add , , "Departamento", 1200
        .ColumnHeaders.Add , , "Estatus", 1200
        .ColumnHeaders.Add , , "Fecha Cierre", 1200
        .ColumnHeaders.Add , , "Tiempo Respuesta", 1200
    End With
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    Combo2.AddItem "<TODOS>"
    sBuscar = "SELECT DEPARTAMENTO FROM DEPARTAMENTOS WHERE ESTATUS = 'A' AND TIPO = 'T' ORDER BY DEPARTAMENTO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Combo1.AddItem tRs.Fields("DEPARTAMENTO")
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Combo2.AddItem tRs.Fields("DEPARTAMENTO")
            If Not IsNull(tRs.Fields("DEPARTAMENTO")) Then Combo3.AddItem tRs.Fields("DEPARTAMENTO")
            tRs.MoveNext
        Loop
    End If
    Actualiza
End Sub
Private Sub Actualiza()
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    ListView1.ListItems.Clear
    'sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.ID_USUARIO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, TICKETS.TITULO, TICKETS.ESTATUS, TICKETS.DEPARTAMENTO_DESTINO FROM TICKETS INNER JOIN USUARIOS ON TICKETS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (TICKETS.DEPARTAMENTO_DESTINO = '" & VarMen.Text1(75).Text & "' AND TICKETS.ESTATUS NOT IN ('F')) OR TICKETS.ID_USUARIO_ATIENDE = '" & VarMen.Text1(0).Text & "' ORDER BY TICKETS.ID_TICKET DESC"
    sBuscar = "SELECT TICKETS.ID_TICKET, TICKETS.FECHA, TICKETS.ID_USUARIO, USUARIOS.NOMBRE, USUARIOS.APELLIDOS, TICKETS.TITULO, TICKETS.ESTATUS, TICKETS.DEPARTAMENTO_DESTINO, TICKETS.TIEMPO_RESPUESTA FROM TICKETS INNER JOIN USUARIOS ON TICKETS.ID_USUARIO = USUARIOS.ID_USUARIO WHERE (TICKETS.DEPARTAMENTO_DESTINO = '" & VarMen.Text1(75).Text & "') AND (TICKETS.ESTATUS NOT IN ('F')) OR (TICKETS.ID_USUARIO_ATIENDE = '" & VarMen.Text1(0).Text & "') AND (TICKETS.ESTATUS NOT IN ('F')) OR (TICKETS.ESTATUS NOT IN ('F')) AND (TICKETS.ID_USUARIO = '" & VarMen.Text1(0).Text & "') ORDER BY TICKETS.ID_TICKET DESC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_TICKET"))
            If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(1) = tRs.Fields("FECHA")
            If Not IsNull(tRs.Fields("ID_USUARIO")) Then tLi.SubItems(2) = tRs.Fields("ID_USUARIO")
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(3) = tRs.Fields("NOMBRE") & " " & tRs.Fields("APELLIDOS")
            If Not IsNull(tRs.Fields("TITULO")) Then tLi.SubItems(4) = tRs.Fields("TITULO")
            If Not IsNull(tRs.Fields("ESTATUS")) Then
                If tRs.Fields("ESTATUS") = "I" Then
                    tLi.SubItems(5) = "NUEVO"
                Else
                    If tRs.Fields("ESTATUS") = "F" Then
                        tLi.SubItems(5) = "FINALIZADO"
                    Else
                        tLi.SubItems(5) = "EN PROCESO"
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("DEPARTAMENTO_DESTINO")) Then tLi.SubItems(6) = tRs.Fields("DEPARTAMENTO_DESTINO")
            If Not IsNull(tRs.Fields("TIEMPO_RESPUESTA")) Then
                If tRs.Fields("TIEMPO_RESPUESTA") = "24" Then
                    Option1.Value = True
                Else
                    If tRs.Fields("TIEMPO_RESPUESTA") = "48" Then
                        Option2.Value = True
                    Else
                        Option3.Value = True
                    End If
                End If
            End If
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    NoTicket = Item
    IdUsuario = Item.SubItems(2)
    If VarMen.Text1(0).Text = IdUsuario Then
        Command1.Visible = True
    Else
        Command1.Visible = False
    End If
    RecargaMensajes
End Sub
Private Sub RecargaMensajes()
    Dim NoReng As Integer
    Dim Altura As Integer
    Dim i As Long
    i = 1
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    'FrmTickets.Controls.Remove (Label2)
    Altura = 120
    sBuscar = "SELECT MENSAJE, FECHA, ID_USUARIO, ID_USUARIO_RESPONDE FROM TICKETS_DETALLE WHERE ID_TICKET = '" & NoTicket & "' ORDER BY ID_TICKET_DETALLE"
    Set tRs = cnn.Execute(sBuscar)
    
    If Label2.Count > 1 Then
        For e = 1 To Label2.Count - 1
            Unload Label2(e)
            Unload Label3(e)
        Next e
    End If
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Load Label2(i)
            With Label2(i)
                NoReng = Len(tRs.Fields("MENSAJE")) \ 80
                NoReng = NoReng + 1
                .Top = Altura 'Label2(i - 1).Top + 390
                .Left = 120
                .Width = 7215
                .Height = 255 * NoReng
                'MsgBox VarMen.Text1(0).Text & " = " & tRs.Fields("ID_USUARIO")
                If CDbl(VarMen.Text1(0).Text) = CDbl(tRs.Fields("ID_USUARIO")) Then
                    .ForeColor = &H808080
                Else
                    .Font.Bold = True
                End If
                .Caption = tRs.Fields("MENSAJE")
                .Visible = True
            End With
            Load Label3(i)
            With Label3(i)
                .Top = Altura 'Label2(i - 1).Top + 390
                .Left = 7440
                .ForeColor = &H808080
                .Width = 2055
                .Height = 255 * NoReng
                .Caption = tRs.Fields("FECHA")
                .Visible = True
                
            End With
            Altura = Altura + 255 * NoReng
            i = i + 1
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim NoReng As Integer
    Dim Altura As Integer
    Dim i As Long
    i = 1
    Dim sBuscar As String
    Dim tLi As ListItem
    Dim tRs As ADODB.Recordset
    'FrmTickets.Controls.Remove (Label2)
    Altura = 120
    sBuscar = "SELECT MENSAJE, FECHA, ID_USUARIO, ID_USUARIO_RESPONDE FROM TICKETS_DETALLE WHERE ID_TICKET = '" & Item & "'"
    Set tRs = cnn.Execute(sBuscar)
    
    If Label10.Count > 1 Then
        For e = 1 To Label10.Count - 1
            Unload Label10(e)
            Unload Label11(e)
        Next e
    End If
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Load Label10(i)
            With Label10(i)
                NoReng = Len(tRs.Fields("MENSAJE")) \ 80
                NoReng = NoReng + 1
                .Top = Altura 'Label2(i - 1).Top + 390
                .Left = 120
                .Width = 7215
                .Height = 255 * NoReng
                'MsgBox VarMen.Text1(0).Text & " = " & tRs.Fields("ID_USUARIO")
                If CDbl(VarMen.Text1(0).Text) = CDbl(tRs.Fields("ID_USUARIO")) Then
                    .ForeColor = &H808080
                End If
                .Caption = tRs.Fields("MENSAJE")
                .Visible = True
            End With
            Load Label11(i)
            With Label11(i)
                .Top = Altura 'Label2(i - 1).Top + 390
                .Left = 7440
                .ForeColor = &H808080
                .Width = 2055
                .Height = 255 * NoReng
                .Caption = tRs.Fields("FECHA")
                .Visible = True
            End With
            Altura = Altura + 255 * NoReng
            i = i + 1
            tRs.MoveNext
        Loop
    End If
End Sub



