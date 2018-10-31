Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsSolTraspaso
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public mnDesde As Integer
    Public mnHasta As Integer
    Public mdFecha As Date
    Public msObservaciones As String
    Public msEstado As String
    Public mnVendedor As Integer
    Public mnOperario As Integer

    Public mbEsNuevo As Boolean
    Public mcolLineas As Collection
    Public Event evtBusSolTraspaso()       ' evento desencadenado despues una busqueda

#Region " Funciones y Rutinas varias "

    Public Function mpsCodigo() As String
        mpsCodigo = "clsSolTraspaso-" & mnEmpresa & "-" & mnCodigo
    End Function

    Public Sub New()
        mbEsNuevo = True
    End Sub

    Public Sub mrMandaEvento()
        RaiseEvent evtBusSolTraspaso()
    End Sub

    Public Sub mrBuscaSolTraspaso()
        Dim loBusSolTraspaso As New frmBusSolTraspasos
        loBusSolTraspaso.mrCargar(Me, mnEmpresa)
    End Sub

    Public Sub mrClonar(ByRef loSolTraspaso As clsSolTraspaso)
        loSolTraspaso.mnEmpresa = mnEmpresa
        loSolTraspaso.mnCodigo = mnCodigo
        loSolTraspaso.mnDesde = mnDesde
        loSolTraspaso.mnHasta = mnHasta
        loSolTraspaso.mdFecha = mdFecha
        loSolTraspaso.msObservaciones = msObservaciones
        loSolTraspaso.msEstado = msEstado
        loSolTraspaso.mnVendedor = mnVendedor
        loSolTraspaso.mnOperario = mnOperario
        loSolTraspaso.mbEsNuevo = mbEsNuevo
        loSolTraspaso.mcolLineas = mcolLineas
    End Sub

#End Region

#Region " Acceso a la base de Datos "

    Public Sub mrNuevoCodigo()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As mySqlDataReader
        Dim loComando As New mySqlCommand

        mnCodigo = 1
        ' ******** primero selecciono el registro ***************************
        lsSql = "select max(cod_tra) as ultimo from soltracabe where emp_tra = " & mnEmpresa
        loComando = New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            mnCodigo = mfnLong(loRecord("ultimo") & "") + 1
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrRecuperaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As mySqlDataReader

        lsSql = "select * from soltracabe where emp_tra = " & mnEmpresa & _
                " and cod_tra = " & mnCodigo
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        mbEsNuevo = True
        While loRecord.Read
            mnEmpresa = mfnInteger(loRecord("emp_tra") & "")
            mnCodigo = mfnLong(loRecord("cod_tra") & "")
            mnDesde = mfnInteger(loRecord("des_tra") & "")
            mnHasta = mfnInteger(loRecord("has_tra") & "")
            mdFecha = mfdFecha(loRecord("fec_tra") & "")
            msObservaciones = Trim(loRecord("obs_tra") & "")
            msEstado = Trim(loRecord("est_tra") & "")
            mnVendedor = mfnInteger(loRecord("ven_tra") & "")
            mnOperario = mfnInteger(loRecord("ope_tra") & "")
            mbEsNuevo = False
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrGrabaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into soltracabe values ('" & mnEmpresa & "','" & _
                    mnCodigo & "','" & _
                    mnDesde & "','" & _
                    mnHasta & "','" & _
                    Format(mdFecha, "yyyy/MM/dd") & "','" & _
                    msObservaciones & "','" & _
                    msEstado & "','" & _
                    mnVendedor & "','" & _
                    mnOperario & "')"
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        Else
            lsSql = "update soltracabe set des_tra = '" & mnDesde & _
                "', has_tra = '" & mnHasta & _
                "', fec_tra = '" & Format(mdFecha, "yyyy/MM/dd") & _
                "', obs_tra = '" & msObservaciones & _
                "', est_tra = '" & msEstado & _
                "', ven_tra = '" & mnVendedor & _
                "', ope_tra = '" & mnOperario & _
                "' where emp_tra = " & mnEmpresa & _
                " and cod_tra = " & mnCodigo
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        End If
        mbEsNuevo = False

    End Sub

    Public Sub mrBorrarDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        ' ahora borro de del fichero de cabeceras ****************
        lsSql = "delete from soltracabe where emp_tra = " & mnEmpresa & _
                " and cod_tra = " & mnCodigo
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()
        mbEsNuevo = True

    End Sub

    Public Sub mrRecuperaLineas()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loFuente As mySqlDataReader
        Dim loLinea As clsSolTraspasoLin

        lsSql = "select * from soltraline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " order by lin_lin asc"

        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader()
        mcolLineas = New Collection
        While loFuente.Read()
            loLinea = New clsSolTraspasoLin
            loLinea.mnEmpresa = mfnInteger(loFuente("emp_lin") & "")
            loLinea.mnCodigo = mfnLong(loFuente("cod_lin") & "")
            loLinea.mnLinea = mfnInteger(loFuente("lin_lin") & "")
            loLinea.mnArticulo = mfnLong(loFuente("art_lin") & "")
            loLinea.mnDetalle = mfnInteger(loFuente("det_lin") & "")
            loLinea.msDescripcion = Trim(loFuente("des_lin") & "")
            loLinea.mnCantidad = mfnDouble(loFuente("ctd_lin") & "")
            loLinea.msEstado = Trim(loFuente("est_lin") & "")
            loLinea.mbEsNuevo = False

            mcolLineas.Add(loLinea, loLinea.mpsCodigo)
        End While
        loFuente.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrBorraLineas()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        lsSql = "delete from soltraline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()

    End Sub

#End Region

End Class
