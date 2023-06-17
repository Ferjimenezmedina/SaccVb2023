VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FormResplado 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame29 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2280
      TabIndex        =   57
      Top             =   2640
      Width           =   975
      Begin VB.Image Image27 
         Height          =   795
         Left            =   120
         MouseIcon       =   "FormResplado.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":030A
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "hhh"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      Height          =   285
      Left            =   4440
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   3360
      Width           =   735
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1200
      TabIndex        =   54
      Top             =   2640
      Width           =   975
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
         TabIndex        =   55
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image26 
         Height          =   855
         Left            =   120
         MouseIcon       =   "FormResplado.frx":082F
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":0B39
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame27 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   1440
      Width           =   975
      Begin VB.Image Image25 
         Height          =   675
         Left            =   120
         MouseIcon       =   "FormResplado.frx":10C8
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":13D2
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
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
         TabIndex        =   53
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame26 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   50
      Top             =   1440
      Width           =   975
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regresar"
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
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image24 
         Height          =   765
         Left            =   240
         MouseIcon       =   "FormResplado.frx":2F80
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":328A
         Top             =   120
         Width           =   645
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame25 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   48
      Top             =   1440
      Width           =   975
      Begin VB.Image Image23 
         Height          =   630
         Left            =   120
         MouseIcon       =   "FormResplado.frx":4D18
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":5022
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pegar"
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
         TabIndex        =   49
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame24 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   46
      Top             =   1440
      Width           =   975
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copiar"
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
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image22 
         Height          =   825
         Left            =   120
         MouseIcon       =   "FormResplado.frx":6AA4
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":6DAE
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame Frame23 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   44
      Top             =   1440
      Width           =   975
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Respaldar"
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
      Begin VB.Image Image21 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FormResplado.frx":8F74
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":927E
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   42
      Top             =   2760
      Width           =   975
      Begin VB.Image Image7 
         Height          =   810
         Left            =   120
         MouseIcon       =   "FormResplado.frx":AD8C
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":B096
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label2 
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
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame22 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1200
      TabIndex        =   40
      Top             =   1440
      Width           =   975
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leer"
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
      Begin VB.Image Image20 
         Height          =   630
         Left            =   120
         MouseIcon       =   "FormResplado.frx":D1C0
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":D4CA
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   975
      Begin VB.Image imgLeer 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FormResplado.frx":EEA4
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":F1AE
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptar"
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
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2280
      TabIndex        =   36
      Top             =   1440
      Width           =   975
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes"
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
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image19 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FormResplado.frx":10C60
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":10F6A
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   28
      Top             =   2760
      Width           =   975
      Begin VB.Frame Frame19 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   33
         Top             =   0
         Width           =   975
         Begin VB.Image Image17 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FormResplado.frx":12A1C
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":12D26
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
      Begin VB.Frame Frame18 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   31
         Top             =   1320
         Width           =   975
         Begin VB.Image Image16 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FormResplado.frx":147D8
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":14AE2
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Aceptar"
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
            TabIndex        =   32
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   29
         Top             =   1320
         Width           =   975
         Begin VB.Label Label16 
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
            TabIndex        =   30
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image15 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FormResplado.frx":1630C
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":16616
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Image Image18 
         Height          =   750
         Left            =   120
         MouseIcon       =   "FormResplado.frx":17FD8
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":182E2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   26
      Top             =   120
      Width           =   975
      Begin VB.Image Image11 
         Height          =   675
         Left            =   120
         MouseIcon       =   "FormResplado.frx":1A00C
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":1A316
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label11 
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
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2280
      TabIndex        =   24
      Top             =   120
      Width           =   975
      Begin VB.Label Label9 
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
         MouseIcon       =   "FormResplado.frx":1BA8C
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":1BD96
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   22
      Top             =   2760
      Width           =   975
      Begin VB.Image Image14 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FormResplado.frx":1DE78
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":1E182
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8520
      TabIndex        =   20
      Top             =   120
      Width           =   975
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
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image13 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FormResplado.frx":1FD84
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":2008E
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   975
      Begin VB.Image Image12 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FormResplado.frx":21AD8
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":21DE2
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
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
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7440
      TabIndex        =   16
      Top             =   120
      Width           =   975
      Begin VB.Image Image10 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FormResplado.frx":23DD4
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":240DE
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
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
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   14
      Top             =   2760
      Width           =   975
      Begin VB.Image Image8 
         Height          =   720
         Left            =   120
         MouseIcon       =   "FormResplado.frx":25C20
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":25F2A
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
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   975
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
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
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FormResplado.frx":278EC
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":27BF6
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   975
      Begin VB.Image Image3 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FormResplado.frx":296A8
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":299B2
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label5 
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   975
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   10
         Top             =   1320
         Width           =   975
         Begin VB.Image Image5 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FormResplado.frx":2B734
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":2BA3E
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label7 
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
            TabIndex        =   11
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   975
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Aceptar"
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
            TabIndex        =   9
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image4 
            Height          =   675
            Left            =   120
            MouseIcon       =   "FormResplado.frx":2D400
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":2D70A
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   975
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
            TabIndex        =   7
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "FormResplado.frx":2EF34
            MousePointer    =   99  'Custom
            Picture         =   "FormResplado.frx":2F23E
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   780
         Left            =   120
         MouseIcon       =   "FormResplado.frx":30CF0
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":30FFA
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   975
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FormResplado.frx":32FEC
         MousePointer    =   99  'Custom
         Picture         =   "FormResplado.frx":332F6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
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
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormResplado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Image10_Click()
' Necesario para el correcto funcionmiento agregar al form lo siguiente :
' - Funcion ShellExecute (para abrir el archivo al terminar de ejecutar)
' - CommonDialog
' - ProgressBar
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
End Sub
