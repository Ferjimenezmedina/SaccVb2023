VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmConfigCorreo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de correo"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   3
      Top             =   720
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FrmConfigCorreo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmConfigCorreo.frx":030A
         Top             =   240
         Width           =   675
      End
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   1
      Top             =   1920
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmConfigCorreo.frx":1CCC
         MousePointer    =   99  'Custom
         Picture         =   "FrmConfigCorreo.frx":1FD6
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "FrmConfigCorreo.frx":40B8
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
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Puerto :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "SMTP : "
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Correo :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmConfigCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Text1.Text = GetSetting("APTONER", "ConfigSACC", "Correo", "")
    Text2.Text = GetSetting("APTONER", "ConfigSACC", "CorreoPass", "")
    Text3.Text = GetSetting("APTONER", "ConfigSACC", "SMTP", "")
    Text4.Text = GetSetting("APTONER", "ConfigSACC", "PuertoCorreo", "")
End Sub
Private Sub Image8_Click()
    SaveSetting "APTONER", "ConfigSACC", "Correo", Text1.Text
    SaveSetting "APTONER", "ConfigSACC", "CorreoPass", Text2.Text
    SaveSetting "APTONER", "ConfigSACC", "SMTP", Text3.Text
    SaveSetting "APTONER", "ConfigSACC", "PuertoCorreo", Text4.Text
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
