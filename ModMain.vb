Option Explicit On 

Public Module ModMain
    Private msUnidad As String
    Public Function mfsRefina(ByVal lsDesc As String) As String
        'Dim lnI As Integer
        'Dim lsTemp As String

        'lsTemp = ""
        'For lnI = 1 To Len(lsDesc)
        '    If Mid(lsDesc, lnI, 1) = "'" Then
        '        lsTemp = lsTemp & "´"
        '    Else
        '        lsTemp = lsTemp & Mid(lsDesc, lnI, 1)
        '    End If
        'Next
        'mfsRefina = lsTemp

        mfsRefina = Replace(lsDesc, "'", "´")

    End Function

    Public Sub mrGrabaLineaLog(ByVal lsFicherolog As String, ByVal lsError As String)
        Dim lsFichero As String

        ' ************ visualiza el fichero por ventana **********
        lsFichero = "C:\Ejecutables .Net\logs\" & lsFicherolog
        Try
            FileOpen(1, lsFichero, OpenMode.Append)
            lsError = Format(Now, "dd/MM/yyyy hh:mm:ss") & " - " & lsError
            PrintLine(1, lsError)
            FileClose(1)
        Catch ex As Exception
            MsgBox(ex.Message & " - Abriendo fichero de logs - " & lsFicherolog, _
                   MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Visanfer .Net")
        End Try

    End Sub

    Public Function mfsLeeDsnRemoto() As String
        Dim lsCadena As String
        Dim lsRuta As String
        Dim lsDSN As String

        Try
            lsRuta = "C:\Ejecutables .Net\Configuraciones\ConexionRemota.ini"
            FileOpen(1, lsRuta, OpenMode.Input)
            ' *********** parametro de corte
            Input(1, lsCadena)
            lsDSN = Mid(lsCadena, 5)
            FileClose(1)
        Catch ex As Exception
        End Try

        Return lsDSN

    End Function

    Public Function mfsUnidad() As String
        ' DEVUELVE LA LETRA DE LA UNIDAD DONDE SE EJECUTAN LAS APLICACIONES
        If msUnidad = "" Then
            Dim loAplicacion As Windows.Forms.Application
            Return Mid(loAplicacion.StartupPath, 1, 2)
        Else
            Return msUnidad
        End If
    End Function

    Public Sub mrLimpiaControl(ByRef loControl As Windows.Forms.Control)
        Dim loControlAux As Windows.Forms.Control
        Dim loCaja As control.txtVisanfer
        Dim loChekbox As Windows.Forms.CheckBox

        For Each loControlAux In loControl.Controls
            If TypeOf (loControlAux) Is control.txtVisanfer Then
                loCaja = loControlAux
                If loCaja.tipo = control.txtVisanfer.TiposCaja.fecha Then
                    loControlAux.Text = Format(Now, "dd/MM/yyyy")
                Else
                    loControlAux.Text = ""
                End If
            ElseIf TypeOf (loControlAux) Is Windows.Forms.TextBox Then
                loControlAux.Text = ""
            ElseIf TypeOf (loControlAux) Is Windows.Forms.CheckBox Then
                loChekbox = loControlAux
                loChekbox.Checked = False
            Else
                mrLimpiaControl(loControlAux)
            End If
        Next
    End Sub

    Public Function mfsArgumentos() As String()
        ' lee los argumentos que se pasan en la linea de comandos
        ' y devuelve un array con los argumentos
        Dim lsSeparadores As String
        Dim lsComandos As String
        Dim lsArgumentos() As String
        'Dim lsI As Object

        lsSeparadores = " "
        lsComandos = Microsoft.VisualBasic.Command()
        lsArgumentos = lsComandos.Split(lsSeparadores.ToCharArray)
        Return lsArgumentos

    End Function

    Public Function mfnInt16(ByVal lsTexto As String) As Int16

        lsTexto = Trim(lsTexto)
        If lsTexto = "" Then
            mfnInt16 = 0
        Else
            If IsNumeric(lsTexto) Then
                mfnInt16 = CInt(lsTexto)
            Else
                mfnInt16 = 0
            End If
        End If

    End Function

    Public Function mfnInt32(ByVal lsTexto As String) As Int32

        lsTexto = Trim(lsTexto)
        If lsTexto = "" Then
            mfnInt32 = 0
        Else
            If IsNumeric(lsTexto) Then
                mfnInt32 = CInt(lsTexto)
            Else
                mfnInt32 = 0
            End If
        End If

    End Function

    Public Function mfnInteger(ByVal lsTexto As String) As Integer

        lsTexto = Trim(lsTexto)
        If lsTexto = "" Then
            mfnInteger = 0
        Else
            If IsNumeric(lsTexto) Then
                mfnInteger = CInt(lsTexto)
            Else
                mfnInteger = 0
            End If
        End If

    End Function

    Public Function mfsCeros(ByVal lsTexto As String, ByVal lnLongitud As Integer) As String
        Dim lnI As Integer
        Dim lnLen As Integer
        Dim lsCadena As String
        Dim lsResultado As String

        lnLen = Len(lsTexto)
        For lnI = 0 To lnLongitud - 1
            If lnI < lnLen Then
                lsCadena = Mid(lsTexto, lnLen - lnI, 1)
                lsResultado = lsCadena & lsResultado
            Else
                lsResultado = "0" & lsResultado
            End If
        Next
        mfsCeros = lsResultado

    End Function

    Public Function mfnLong(ByVal lsTexto As String) As Long

        lsTexto = Trim(lsTexto)
        If lsTexto = "" Then
            mfnLong = 0
        Else
            If IsNumeric(lsTexto) Then
                mfnLong = CLng(lsTexto)
            Else
                mfnLong = 0
            End If
        End If

    End Function

    Public Function mfnDouble(ByVal lsTexto As String) As Double

        lsTexto = Trim(lsTexto)
        If lsTexto = "" Then
            mfnDouble = 0
        Else
            If IsNumeric(lsTexto) Then
                mfnDouble = CDbl(lsTexto)
            Else
                mfnDouble = 0
            End If
        End If

    End Function

    Public Function mfnBoolean(ByVal lbValor As Boolean) As Integer

        mfnBoolean = IIf(lbValor, 1, 0)

    End Function

    Public Function mfbBoolean(ByVal lsTexto As String) As Boolean
        Dim lnValor As Integer

        lsTexto = Trim(lsTexto)
        Select Case lsTexto
            Case "", "F"
                lnValor = 0
            Case "V"
                lnValor = 1
            Case Else
                lnValor = CInt(lsTexto)
        End Select

        mfbBoolean = IIf(lnValor = 1, True, False)

    End Function

    Function mfdFecha(ByVal lsTexto As String) As Date

        If Not InStr(lsTexto, "/", CompareMethod.Text) > 0 Then
            On Error Resume Next
            lsTexto = Mid(lsTexto, 1, 2) & "/" & Mid(lsTexto, 3, 2) & "/" & Mid(lsTexto, 5, 4)
            On Error GoTo 0
        End If

        lsTexto = Trim(lsTexto)
        If lsTexto <> "" And lsTexto <> "//" Then
            If IsDate(lsTexto) Then
                mfdFecha = CDate(lsTexto)
            Else
                mfdFecha = "01/01/1900"
            End If
        Else
            mfdFecha = "01/01/1900"
        End If

    End Function

    Function mfsMoneda(ByVal lnValor As Double) As String

        mfsMoneda = Format(lnValor, "#,##0.00")

    End Function

    Public Function mfbVerificarCIF(ByVal lsValor As String, ByRef lsMensaje As String) As Boolean
        Dim lsLetra As String, lsNumero As String, lsDigito As String
        Dim lsDigitoAux As String
        Dim lnAuxNum As Integer
        Dim lnI As Integer
        Dim lnSuma As Integer
        Dim lsLetras As String
        Dim lsLetraAux As String

        lsLetras = "ABCDEFGHKLMPQSX"
        lsValor = UCase(lsValor)

        If Len(lsValor) < 9 Or Not IsNumeric(Mid(lsValor, 2, 7)) Then
            lsMensaje = "El dato introducido no corresponde a un CIF"
            mfbVerificarCIF = False
            Exit Function
        End If
        lsLetra = Mid(lsValor, 1, 1)       'letra del CIF
        lsNumero = Mid(lsValor, 2, 7)    'Codigo de Control
        lsDigito = Mid(lsValor, 9)                'CIF menos primera y ultima posiciones

        If InStr(lsLetras, lsLetra) = 0 Then 'comprobamos la letra del CIF (1ª posicion)
            lsMensaje = "la letra introducida no corresponde a un CIF"
            mfbVerificarCIF = False
            Exit Function
        End If

        lnI = 0
        For lnI = 1 To 7
            ' reviso que solo contenga numero el chorizo
            lsLetraAux = Mid(lsNumero, lnI, 1)
            If Not IsNumeric(lsLetraAux) Then
                lsMensaje = "Caracter '" & lsLetraAux & "' no permitido en el CIF"
                mfbVerificarCIF = False
                Exit Function
            End If
            If lnI Mod 2 = 0 Then
                lnSuma = lnSuma + CInt(Mid(lsNumero, lnI, 1))
            Else
                lnAuxNum = CInt(Mid(lsNumero, lnI, 1)) * 2
                lnSuma = lnSuma + (lnAuxNum \ 10) + (lnAuxNum Mod 10)
            End If
        Next
        lnSuma = (10 - (lnSuma Mod 10)) Mod 10

        Select Case lsLetra
            Case "K", "P", "Q", "S"
                lnSuma = lnSuma + 64
                lsDigitoAux = Chr(lnSuma)
            Case "X"
                lsDigitoAux = Mid(mfsCalculaNIF(lsNumero), 8, 1)
            Case Else
                lsDigitoAux = CStr(lnSuma)
        End Select

        If lsDigito = lsDigitoAux Then
            lsMensaje = "CIF correcto"
            mfbVerificarCIF = True
        Else
            lsMensaje = "El cif no es correcto, el dígito de control no coincide"
            mfbVerificarCIF = False
        End If

    End Function

    Public Function mfbVerificarNIF(ByVal lsValor As String, ByRef lsMensaje As String) As Boolean
        Dim lsAux As String

        lsMensaje = ""
        lsValor = UCase(lsValor) 'ponemos la letra en mayúscula
        lsAux = Mid(lsValor, Len(lsValor), 1)
        If Not IsNumeric(lsAux) Then
            lsAux = Mid(lsValor, 1, Len(lsValor) - 1) 'quitamos la letra del NIF
        Else
            lsAux = lsValor
        End If

        If Len(lsAux) = 8 And IsNumeric(lsAux) Then
            lsAux = mfsCalculaNIF(lsAux) 'calculamos la letra del NIF para comparar con la que tenemos
        Else
            MsgBox("El dato introducido no corresponde a un NIF")
            mfbVerificarNIF = False
            Exit Function
        End If

        If lsValor <> lsAux Then 'comparamos las letras
            lsMensaje = "El NIF " & lsValor & " es INCORRECTO" & vbCrLf & "D.N.I. Correcto: " & lsAux
            mfbVerificarNIF = False
        Else
            lsMensaje = "El NIF " & lsValor & " es CORRECTO"
            mfbVerificarNIF = True
        End If

    End Function

    Public Function mfsCalculaNIF(ByVal lsValor As String) As String
        Dim lsResto As Integer
        Dim lsLetraNIF As String

        lsLetraNIF = ""

        If lsValor = "" Then
            mfsCalculaNIF = ""
            Exit Function
        ElseIf Len(lsValor) < 8 Then
            mfsCalculaNIF = ""
            Exit Function
        ElseIf Not IsNumeric(lsValor) Then
            mfsCalculaNIF = ""
            Exit Function
        Else
            lsResto = Val(lsValor) Mod 23
            Select Case lsResto
                Case 0
                    lsLetraNIF = "T"
                Case 1
                    lsLetraNIF = "R"
                Case 2
                    lsLetraNIF = "W"
                Case 3
                    lsLetraNIF = "A"
                Case 4
                    lsLetraNIF = "G"
                Case 5
                    lsLetraNIF = "M"
                Case 6
                    lsLetraNIF = "Y"
                Case 7
                    lsLetraNIF = "F"
                Case 8
                    lsLetraNIF = "P"
                Case 9
                    lsLetraNIF = "D"
                Case 10
                    lsLetraNIF = "X"
                Case 11
                    lsLetraNIF = "B"
                Case 12
                    lsLetraNIF = "N"
                Case 13
                    lsLetraNIF = "J"
                Case 14
                    lsLetraNIF = "Z"
                Case 15
                    lsLetraNIF = "S"
                Case 16
                    lsLetraNIF = "Q"
                Case 17
                    lsLetraNIF = "V"
                Case 18
                    lsLetraNIF = "H"
                Case 19
                    lsLetraNIF = "L"
                Case 20
                    lsLetraNIF = "C"
                Case 21
                    lsLetraNIF = "K"
                Case 22
                    lsLetraNIF = "E"
            End Select
            mfsCalculaNIF = lsValor & lsLetraNIF
            Exit Function
        End If

    End Function

    Public Function mfbVerificaBanco(ByVal lsBanco As String, ByVal lsOficina As String, _
                                     ByVal lsControl As String, ByVal lsCuenta As String) As Boolean
        Dim lnD1 As Integer
        Dim lnD2 As Integer
        Dim lsCuenta2 As String

        lsCuenta2 = Format(mfnInteger(lsBanco), "0000")
        lsCuenta2 = lsCuenta2 & Format(mfnInteger(lsOficina), "0000")
        lsCuenta2 = lsCuenta2 & Format(mfnInteger(lsControl), "00")
        lsCuenta2 = lsCuenta2 & Format(mfnLong(lsCuenta), "0000000000")

        lnD1 = Val(Mid(lsCuenta2, 9, 1))
        lnD2 = Val(Mid(lsCuenta2, 10, 1))
        If mfnCalculaDigito(Left(lsCuenta2, 8), 2) = lnD1 And mfnCalculaDigito(Right(lsCuenta, 10), 0) = lnD2 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mfnCalculaDigito(ByVal lsCuenta As String, ByVal lnOffset As Integer) As Integer
        Dim lnSuma As Integer
        Dim lnI As Integer
        Dim lnDig As Integer
        Dim lnApesos(10) As Integer

        lnApesos(1) = 1
        lnApesos(2) = 2
        lnApesos(3) = 4
        lnApesos(4) = 8
        lnApesos(5) = 5
        lnApesos(6) = 10
        lnApesos(7) = 9
        lnApesos(8) = 7
        lnApesos(9) = 3
        lnApesos(10) = 6
        For lnI = 1 To (10 - lnOffset)
            lnSuma = lnSuma + Val(Mid(lsCuenta, lnI, 1)) * lnApesos(lnI + lnOffset)
        Next
        lnDig = 11 - (lnSuma - (11 * Int(lnSuma / 11)))
        If lnDig = 10 Then lnDig = 1
        If lnDig = 11 Then lnDig = 0
        mfnCalculaDigito = lnDig

    End Function



End Module



