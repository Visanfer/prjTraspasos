Option Explicit On 
Imports prjControl
Imports CrystalDecisions.CrystalReports.Engine

Public Module modTraspasos
    Public gnLlave As Integer          ' llave actual de seguridad
    Public goUsuario As New clsUsuario
    Public goProfile As New clsProfile()
    Public goListado As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    Public Sub Main()
        Dim loPrincipal As New frmTraspasos
        loPrincipal.mrCargar(1, Nothing)
    End Sub

End Module
