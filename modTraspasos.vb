Option Explicit On 
Imports prjControl
Imports CrystalDecisions.CrystalReports.Engine

Public Module modTraspasos
    Public gnLlave As Integer          ' llave actual de seguridad
    Public goUsuario As New clsUsuario
    Public goProfile As New clsprofilelocal()
    Public goListado As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Public Sub Main()

        goUsuario = New clsUsuario
        goUsuario.mnCodigo = 206
        goUsuario.mrRecuperaDatos()

        'Dim loPrincipal As New frmTraspasos
        'loPrincipal.mrCargar(1, Nothing)

        mrEjecutaMenu("2", Nothing, 1)

    End Sub

    Public Function mfoMenu() As Collection

        Dim loColeccion As New Collection
        loColeccion.Add(New clsOpcionMenu("1", "Consulta", "Consulta/Alta de Traspasos."))
        loColeccion.Add(New clsOpcionMenu("2", "Entradas", "Consulta/Confirmación de traspasos pendientes de confirmar."))
        loColeccion.Add(New clsOpcionMenu("3", "Solicitud", "Solicitud de traspaso a otro almacén"))
        loColeccion.Add(New clsOpcionMenu("4", "Solicitudes Pendientes", "Solicitudes de otros almacenes pendientes de servir"))

        Return loColeccion

    End Function

    Public Sub mrEjecutaMenu(ByVal lsOpcion As String, ByVal loPadre As Form, ByVal lnEmpresa As Integer)

        Select Case lsOpcion
            Case "1"
                Dim loConsulta As New frmTraspasos
                loConsulta.MdiParent = loPadre
                loConsulta.mbCargaDirecta = True
                loConsulta.mrCargar(lnEmpresa, goUsuario)

            Case "2"

                Dim loConfEntradas As New frmConfEntradas
                loConfEntradas.MdiParent = loPadre
                loConfEntradas.mrCargar(lnEmpresa)

            Case "3"

                Dim loSolTraspasos As New frmSolTraspasos
                loSolTraspasos.MdiParent = loPadre
                loSolTraspasos.mrCargar(lnEmpresa)

            Case "4"

                mrMiraSolicitudes(loPadre, lnEmpresa)

        End Select

    End Sub

    Private Sub mrMiraSolicitudes(ByVal loPadre As Form, ByVal lnEmpresa As Integer)

        Dim loBusSolTraspasos As New clsBusSolTraspasos
        loBusSolTraspasos.mnEmpresa = lnEmpresa
        loBusSolTraspasos.mnDesde = goProfile.mnAlmacen
        loBusSolTraspasos.msEstado = "P"
        loBusSolTraspasos.mdDesdeFecha = DateAdd(DateInterval.Day, -60, Now)
        loBusSolTraspasos.mdHastaFecha = Now
        loBusSolTraspasos.mrBuscaSolTraspasos()
        If loBusSolTraspasos.mcolSolTraspasos.Count > 0 Then
            Dim loPendientes As New frmSolTraspasosPendientes
            loPendientes.MdiParent = loPadre
            loPendientes.mrCargar(loBusSolTraspasos)
        Else
            MsgBox("NO HAY SOLICITUDES DE TRASPASO PENDIENTES.", MsgBoxStyle.Information, "Visanfer .Net")
        End If

    End Sub

    Public Function mfsEan13(ByVal lsCodigoArticulo As String) As String

        Dim lnI As Integer
        Dim lnChecksum As Integer
        Dim lnPrimero As Integer
        Dim lsCodeBarre As String = ""
        Dim lbTablaA As Boolean
        Dim lnUltmimaPagina As Integer = 0

        If Len(lsCodigoArticulo) = 12 Then
            For lnI = 1 To 12
                If Asc(Mid(lsCodigoArticulo, lnI, 1)) < 48 Or Asc(Mid(lsCodigoArticulo, lnI, 1)) > 57 Then
                    lnI = 0
                    Exit For
                End If
            Next
            If lnI = 13 Then
                For lnI = 12 To 1 Step -2
                    lnChecksum = lnChecksum + Val(Mid(lsCodigoArticulo, lnI, 1))
                Next
                lnChecksum = lnChecksum * 3
                For lnI = 11 To 1 Step -2
                    lnChecksum = lnChecksum + Val(Mid(lsCodigoArticulo, lnI, 1))
                Next
                lsCodigoArticulo = lsCodigoArticulo & (10 - lnChecksum Mod 10) Mod 10
                lsCodeBarre = Microsoft.VisualBasic.Left(lsCodigoArticulo, 1) & Chr(65 + Val(Mid(lsCodigoArticulo, 2, 1)))
                lnPrimero = Val(Microsoft.VisualBasic.Left(lsCodigoArticulo, 1))
                For lnI = 3 To 7
                    lbTablaA = False
                    Select Case lnI
                        Case 3
                            Select Case lnPrimero
                                Case 0 To 3
                                    lbTablaA = True
                            End Select
                        Case 4
                            Select Case lnPrimero
                                Case 0, 4, 7, 8
                                    lbTablaA = True
                            End Select
                        Case 5
                            Select Case lnPrimero
                                Case 0, 1, 4, 5, 9
                                    lbTablaA = True
                            End Select
                        Case 6
                            Select Case lnPrimero
                                Case 0, 2, 5, 6, 7
                                    lbTablaA = True
                            End Select
                        Case 7
                            Select Case lnPrimero
                                Case 0, 3, 6, 8, 9
                                    lbTablaA = True
                            End Select
                    End Select
                    If lbTablaA Then
                        lsCodeBarre = lsCodeBarre & Chr(65 + Val(Mid(lsCodigoArticulo, lnI, 1)))
                    Else
                        lsCodeBarre = lsCodeBarre & Chr(75 + Val(Mid(lsCodigoArticulo, lnI, 1)))
                    End If
                Next
                lsCodeBarre = lsCodeBarre & "*"
                For lnI = 8 To 13
                    lsCodeBarre = lsCodeBarre & Chr(97 + Val(Mid(lsCodigoArticulo, lnI, 1)))
                Next
                lsCodeBarre = lsCodeBarre & "+"
            End If
        End If

        Return lsCodeBarre

    End Function

End Module
