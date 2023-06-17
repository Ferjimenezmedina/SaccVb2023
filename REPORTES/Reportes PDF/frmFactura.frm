VERSION 5.00
Begin VB.Form frmFactura 
   Caption         =   "Facturar"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "COBRO_MCIA.rpt"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MASVENDIDO.rpt"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CARTVACCOMPRA.rpt"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARTVACVENTA.rpt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   120
      Picture         =   "frmFactura.frx":0000
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "FrmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim cont As Integer
    Dim Posi As Integer
    
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate("c:\Prueba.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
    oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
    oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
    oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "materia_prima@aptoner.com.mx", "F2", 10, hCenter
    oDoc.WTextBox 70, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 80, 380, 20, 250, Date, "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 140, 20, 20, 250, "EPSON", "F2", 10, hLeft
    oDoc.WTextBox 160, 20, 20, 200, "Clave del Producto", "F2", 10, hCenter
    oDoc.WTextBox 160, 20, 20, 700, "Cantidad", "F2", 10, hCenter
    oDoc.WTextBox 160, 20, 20, 950, "Precio al Publico", "F2", 10, hCenter
' Cuerpo del reporte
    Posi = 180
    Do While Not cont = 100
        oDoc.WTextBox Posi, 50, 20, 200, cont & "XXXXX", "F3", 10, hLeft
        oDoc.WTextBox Posi, 330, 20, 60, "2", "F3", 10, hRight
        oDoc.WTextBox Posi, 400, 20, 140, "$1,138,000,000.00", "F3", 10, hRight
        Posi = Posi + 15
        If Posi >= 760 Then
            oDoc.NewPage A4_Vertical
        ' Encabezado del reporte
            oDoc.WImage 80, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
            oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
            oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
            oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
            oDoc.WTextBox 80, 200, 20, 250, "materia_prima@aptoner.com.mx", "F2", 10, hCenter
            oDoc.WTextBox 70, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
            oDoc.WTextBox 80, 380, 20, 250, Date, "F3", 8, hCenter
        ' Encabezado de pagina
            oDoc.WTextBox 140, 20, 20, 250, "EPSON", "F2", 10, hLeft
            oDoc.WTextBox 160, 20, 20, 200, "Clave del Producto", "F2", 10, hCenter
            oDoc.WTextBox 160, 20, 20, 700, "Cantidad", "F2", 10, hCenter
            oDoc.WTextBox 160, 20, 20, 950, "Precio al Publico", "F2", 10, hCenter
            Posi = 180
        End If
        cont = cont + 1
    Loop
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Command2_Click()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim cont As Integer
    Dim Posi As Integer
    
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate("c:\Prueba.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
    oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
    oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
    oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "materia_prima@aptoner.com.mx", "F2", 10, hCenter
    oDoc.WTextBox 70, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 80, 380, 20, 250, Date, "F3", 8, hCenter
' Encabezado de pagina
    oDoc.WTextBox 140, 20, 20, 250, "EPSON", "F2", 10, hLeft
    oDoc.WTextBox 160, 20, 20, 200, "Clave del Producto", "F2", 10, hCenter
    oDoc.WTextBox 160, 20, 20, 500, "Cantidad", "F2", 10, hCenter
    oDoc.WTextBox 160, 20, 20, 720, "Precio de Compra", "F2", 10, hCenter
    oDoc.WTextBox 160, 20, 20, 950, "Precio al Publico", "F2", 10, hCenter
' Cuerpo del reporte
    Posi = 180
    Do While Not cont = 100
        oDoc.WTextBox Posi, 50, 20, 200, cont & "XXXXX", "F3", 10, hLeft
        oDoc.WTextBox Posi, 230, 20, 60, "2", "F3", 10, hRight
        oDoc.WTextBox Posi, 290, 20, 140, "$1,138,000,000.00", "F3", 10, hRight
        oDoc.WTextBox Posi, 400, 20, 140, "$1,138,000,000.00", "F3", 10, hRight
        Posi = Posi + 15
        If Posi >= 760 Then
            oDoc.NewPage A4_Vertical
        ' Encabezado del reporte
            oDoc.WImage 80, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
            oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
            oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
            oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
            oDoc.WTextBox 80, 200, 20, 250, "materia_prima@aptoner.com.mx", "F2", 10, hCenter
            oDoc.WTextBox 70, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
            oDoc.WTextBox 80, 380, 20, 250, Date, "F3", 8, hCenter
        ' Encabezado de pagina
            oDoc.WTextBox 140, 20, 20, 250, "EPSON", "F2", 10, hLeft
            oDoc.WTextBox 160, 20, 20, 200, "Clave del Producto", "F2", 10, hCenter
            oDoc.WTextBox 160, 20, 20, 500, "Cantidad", "F2", 10, hCenter
            oDoc.WTextBox 160, 20, 20, 720, "Precio de Compra", "F2", 10, hCenter
            oDoc.WTextBox 160, 20, 20, 950, "Precio al Publico", "F2", 10, hCenter
            Posi = 180
        End If
        cont = cont + 1
    Loop
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Command3_Click()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim cont As Integer
    Dim Posi As Integer
    
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate("c:\Prueba.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
' Encabezado del reporte
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.NewPage A4_Vertical
    oDoc.WImage 80, 40, 43, 161, "Logo"
    oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
    oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
    oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
    oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
    oDoc.WTextBox 60, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
    oDoc.WTextBox 70, 380, 20, 250, Date, "F3", 8, hCenter
    oDoc.WTextBox 90, 200, 20, 230, "PRODUCTOS MAS VENDIDOS", "F2", 10, hCenter
' Linea
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 110
    oDoc.WLineTo 580, 110
    oDoc.LineStroke
' Cuerpo del reporte
    Posi = 120
    Do While Not cont = 100
        oDoc.WTextBox Posi, 50, 20, 250, cont & "XXXXX", "F3", 10, hLeft
        Posi = Posi + 15
    ' Linea
        oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
        oDoc.MoveTo 50, Posi
        oDoc.WLineTo 250, Posi
        oDoc.LineStroke
        oDoc.WTextBox Posi, 50, 20, 350, cout & " DESCRIPCION", "F3", 10, hLeft
        oDoc.WTextBox Posi, 400, 20, 140, "1,138,000,000.00", "F3", 10, hRight
        Posi = Posi + 20
        If Posi >= 760 Then
            oDoc.NewPage A4_Vertical
        ' Encabezado del reporte
            oDoc.WImage 80, 40, 43, 161, "Logo"
            oDoc.WTextBox 40, 200, 20, 250, "ACTITUD POSITIVA EN TONER S. DE R.L. M.I.", "F2", 10, hCenter
            oDoc.WTextBox 50, 200, 20, 250, "Ortiz de Campos 1308, Col. San Felipe", "F2", 10, hCenter
            oDoc.WTextBox 60, 200, 20, 250, "Chihuahua Chih.", "F2", 10, hCenter
            oDoc.WTextBox 70, 200, 20, 250, "Tel (614) 414-82-41", "F2", 10, hCenter
            oDoc.WTextBox 60, 380, 20, 250, "Fecha de Impresion", "F3", 8, hCenter
            oDoc.WTextBox 70, 380, 20, 250, Date, "F3", 8, hCenter
            oDoc.WTextBox 90, 200, 20, 230, "PRODUCTOS MAS VENDIDOS", "F2", 10, hCenter
        ' Linea
            oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
            oDoc.MoveTo 10, 110
            oDoc.WLineTo 580, 110
            oDoc.LineStroke
            Posi = 120
        End If
        cont = cont + 1
    Loop
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub Command4_Click()
    
End Sub
