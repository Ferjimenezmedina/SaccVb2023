Attribute VB_Name = "mdlUtilerias"
Public BanLista As Byte
Public bInvCre As Boolean
Public INV As Integer
Public crReport As New CRAXDRT.Report
Public crApplication As New CRAXDRT.Application
Public PedDir As Integer
Public PedInd As Integer

Public Sub Limpiar_Campos(FORMA As Form)

Dim LimpCamp As Integer
For LimpCamp = 0 To FORMA.Controls.Count - 1
    If TypeOf FORMA.Controls(LimpCamp) Is TextBox Then
        FORMA.Controls(LimpCamp).Text = ""
    End If
Next

End Sub

Public Sub Abrir_Campos(FORMA As Form)
    
Dim AbriCamp As Integer
For AbriCamp = 0 To FORMA.Controls.Count - 1
    If TypeOf FORMA.Controls(AbriCamp) Is TextBox Then
        FORMA.Controls(AbriCamp).Locked = True
    End If
Next

End Sub

Public Function Mayusculas(C As Integer) As Integer

    Mayusculas = Asc(UCase(Chr(C)))
    
End Function

Sub Main()

    frmProd.Show
    
End Sub

Public Function MESES(MES As String)

    Select Case MES
        Case 1: MESES = "ENERO"
        Case 2: MESES = "FEBRERO"
        Case 3: MESES = "MARZO"
        Case 4: MESES = "ABRIL"
        Case 5: MESES = "MAYO"
        Case 6: MESES = "JUNIO"
        Case 7: MESES = "JULIO"
        Case 8: MESES = "AGOSTO"
        Case 9: MESES = "SEPTIEMBRE"
        Case 10: MESES = "OCTUBRE"
        Case 11: MESES = "NOVIEMBRE"
        Case 12: MESES = "DICIEMBRE"
    End Select

End Function

Public Function Hay_Pedidos_Directos() As Boolean

    deAPTONER.Hay_Pedidos_Directos
    With deAPTONER.rsHAY_PEDIDOS_DIRECTOS
        If Not (.BOF And .EOF) Then
            Hay_Pedidos_Directos = True
            PedDir = !pedidosdirectos
        Else
            Hay_Pedidos_Directos = False
        End If
        .Close
    End With
    
End Function

Public Function Hay_Pedidos_Indirectos() As Boolean

    deAPTONER.Hay_Pedidos_Indirectos
    With deAPTONER.rsHAY_PEDIDOS_INDIRECTOS
        If Not (.BOF And .EOF) Then
            Hay_Pedidos_Indirectos = True
            PedInd = !pedidosindirectos
        Else
            Hay_Pedidos_Indirectos = False
        End If
        .Close
    End With
    
End Function

Public Function Hay_Existencias(cProd As String, cSuc As String) As Boolean

    deAPTONER.Hay_Existencias cProd, cSuc
    With deAPTONER.rsHAY_EXISTENCIAS
        If !ID_EXISTENCIA <> 0 Then
            Hay_Existencias = True
        Else
            Hay_Existencias = False
        End If
        .Close
    End With
    
End Function

Public Function Hay_Usuarios() As Boolean
    
    deAPTONER.Hay_Usuarios
    With deAPTONER.rsHAY_USUARIOS
        If !Id_Usuario <> 0 Then
            Hay_Usuarios = True
        Else
            Hay_Usuarios = False
        End If
        .Close
    End With
    
End Function

Public Function Hay_Mensajes(Usuario As String) As Boolean
    
    deAPTONER.Hay_Mensajes Usuario, Format(Date, "dd/mm/yyyy")
    With deAPTONER.rsHAY_MENSAJES
        If !Mensajes <> 0 Then
            Hay_Mensajes = True
        Else
            Hay_Mensajes = False
        End If
        .Close
    End With
    
End Function
