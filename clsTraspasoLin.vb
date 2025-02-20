Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsTraspasoLin
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public mnLinea As Integer
    Public mnArticulo As Long
    Public mnDetalle As Integer
    Public msDescripcion As String
    Public mnCantidad As Double
    Public msEstado As String

    Public mbEsNuevo As Boolean
    Private msLogConfirmado As String = ""

    Public Function mpsCodigo() As String
        mpsCodigo = "clsTraspasoLin-" & mnEmpresa & "-" & mnCodigo & "-" & mnLinea
    End Function

    Public Sub mrRecuperaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loFuente As mySqlDataReader

        lsSql = "select * from traline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " and lin_lin = " & mnLinea
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader()
        mbEsNuevo = True
        While loFuente.Read()
            mnEmpresa = mfnInteger(loFuente("emp_lin") & "")
            mnCodigo = mfnLong(loFuente("cod_lin") & "")
            mnLinea = mfnInteger(loFuente("lin_lin") & "")
            mnArticulo = mfnLong(loFuente("art_lin") & "")
            mnDetalle = mfnInteger(loFuente("det_lin") & "")
            msDescripcion = Trim(loFuente("des_lin") & "")
            mnCantidad = mfnDouble(loFuente("ctd_lin") & "")
            msEstado = Trim(loFuente("est_lin") & "")
            mbEsNuevo = False
        End While
        loFuente.Close()
        lconConexion.Close()
    End Sub

    Public Function mfnPendienteTraspasoArticulo() As Double
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Function

        Dim lsSql As String
        Dim loFuente As mySqlDataReader
        Dim lnPendiente As Double

        lsSql = "select sum(ctd_lin) as total from traline" & _
                " where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " and art_lin = " & mnArticulo & _
                " and det_lin = " & mnDetalle & _
                " and est_lin = 'N'"
        If mnCantidad > 0 Then lsSql = lsSql & " and ctd_lin = " & mnCantidad
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader()
        While loFuente.Read()
            lnPendiente = mfnDouble(loFuente("total") & "")
        End While
        loFuente.Close()
        lconConexion.Close()

        Return lnPendiente

    End Function

    Public Function mfnPendienteTraspaso(ByVal lnAlmacen As Integer) As Double
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Function

        Dim lsSql As String
        Dim loFuente As mySqlDataReader
        Dim lnPendiente As Double
        Dim lcolTraspasos As Collection
        Dim loTraspasoLin As clsTraspasoLin
        Dim loTraspaso As clsTraspaso

        lsSql = "select * from traline" & _
                " where emp_lin = " & mnEmpresa & _
                " and art_lin = " & mnArticulo & _
                " and det_lin = " & mnDetalle & _
                " and est_lin = 'N'"
        lcolTraspasos = New Collection
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader()
        While loFuente.Read()
            loTraspasoLin = New clsTraspasoLin
            loTraspasoLin.mnCodigo = mfnInteger(loFuente("cod_lin") & "")
            loTraspasoLin.mnLinea = mfnInteger(loFuente("lin_lin") & "")
            loTraspasoLin.mnCantidad = mfnDouble(loFuente("ctd_lin") & "")
            lcolTraspasos.Add(loTraspasoLin, loTraspasoLin.mpsCodigo)
        End While
        loFuente.Close()
        lconConexion.Close()

        For Each loTraspasoLin In lcolTraspasos
            loTraspaso = New clsTraspaso
            loTraspaso.mnEmpresa = mnEmpresa
            loTraspaso.mnCodigo = loTraspasoLin.mnCodigo
            loTraspaso.mrRecuperaDatos()
            If loTraspaso.mnHasta = lnAlmacen Then
                lnPendiente = lnPendiente + loTraspasoLin.mnCantidad
            End If
        Next

        Return lnPendiente

    End Function

    Public Sub mrBorraDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        lsSql = "delete from traline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " and lin_lin = " & mnLinea
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()
        mbEsNuevo = True

    End Sub

    Public Sub mrActualizaLineaTraspaso(ByVal lnAlmacen As Integer, ByVal lnUsuario As Integer)

        'Dim lsLog As String
        ' gabo un log con el usuario que lo hace *********
        'lsLog = " -T " & mnCodigo & " -L " & mnLinea
        'lsLog = lsLog & " -A " & mnArticulo & "." & mnDetalle & " -U " & lnUsuario
        'prjControl.mrGrabaLineaLog("ActuTraspasos.log", lsLog)

        If goUsuario Is Nothing Then goUsuario = New clsUsuario
        If goProfile Is Nothing Then goProfile = New clsprofilelocal
        If goUsuario.mbEsNuevo Then
            goUsuario.mnCodigo = lnUsuario
            goUsuario.mrRecuperaDatos()
        End If

        mrGrabaLog(lnUsuario)

        ' actualizo las existencias en el almacen destino, despues de que
        ' se haya confirmado su llegada en el almacen *******************

        ' cuando la linea tiene detalle, debo mirar si el detalle tiene existencias propias y el cun
        If mnDetalle > 0 Then
            Dim loDetalle As New prjArticulos.clsArticulo
            loDetalle.mnEmpresa = mnEmpresa
            loDetalle.mnCodigo = mnArticulo
            loDetalle.mnDetalle = mnDetalle
            loDetalle.mrRecuperaDatos()

            If Not loDetalle.msControlExistencias = "N" Then mnDetalle = 0
            ' da error en las vigas, lo quito pues en el resto de programas no se hace ************ 13-08-21
            'mnCantidad = mnCantidad * loDetalle.mnCantidadUnidad

        End If

        prjArticulos.goUsuario = modTraspasos.goUsuario
        prjArticulos.goProfile = modTraspasos.goProfile

        Dim loExistencias As New prjArticulos.clsExistencias
        loExistencias.mnEmpresa = mnEmpresa
        loExistencias.mnAlmacen = lnAlmacen
        loExistencias.mnArticulo = mnArticulo
        loExistencias.mnDetalle = mnDetalle
        loExistencias.mnExistencias = (-1) * mnCantidad
        loExistencias.mdFechaUltEntrada = Now
        loExistencias.mrActualizaExistencias()


    End Sub

    Public Function mfsLogConfirmado() As String

        If msLogConfirmado.Length = 0 Then

            Dim lsSql As String = "select * from traline_log where empresa = " & mnEmpresa &
                                " and codigo = " & mnCodigo &
                                " and linea = " & mnLinea &
                                 " and articulo = " & mnArticulo &
                                " and detalle = " & mnDetalle
            Dim loDatos As DataTable = New prjControl.clsControlBD().mfoRecuperaDatos(False, lsSql, "confirmacion")

            If loDatos.Rows.Count > 0 Then
                Dim loRow As DataRow = loDatos.Rows(0)
                msLogConfirmado = "LINEA CONFIRMADA EN DESTINO POR EL USUARIO " & loRow("usuario") & " - " & loRow("fecha") & " - " & loRow("hora")
            Else
                msLogConfirmado = "-"
            End If

        End If

        Return msLogConfirmado

    End Function

    Public Sub mrGrabaLog(ByVal lnOperario As Integer)
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loComando As New MySqlCommand
        Dim lsSql As String = "insert into traline_log " &
            "(empresa,codigo,linea,articulo,detalle,fecha,hora,usuario) " &
            "values (" & mnEmpresa & "," &
            mnCodigo & "," &
            mnLinea & "," &
            mnArticulo & "," &
            mnDetalle & ",'" &
            Format(Now, formatoFecha) & "','" &
            Format(Now, "HH:mm:ss") & "'," &
            goUsuario.mnCodigo & ")"

        loComando = New MySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()

    End Sub

    Public Sub mrActualizaLineaTraspasoOld(ByVal lnAlmacen As Integer, ByVal lnUsuario As Integer)
        Dim loExistencias As prjArticulos.clsExistencias
        Dim lsLog As String

        ' gabo un log con el usuario que lo hace *********
        lsLog = " -T " & mnCodigo & " -L " & mnLinea
        lsLog = lsLog & " -A " & mnArticulo & " -U " & lnUsuario
        prjControl.mrGrabaLineaLog("ActuTraspasos.log", lsLog)

        ' actualizo las existencias en el almacen destino, despues de que
        ' se haya confirmado su llegada en el almacen *******************
        loExistencias = New prjArticulos.clsExistencias
        loExistencias.mnEmpresa = mnEmpresa
        loExistencias.mnAlmacen = lnAlmacen
        loExistencias.mnArticulo = mnArticulo
        loExistencias.mnDetalle = mnDetalle
        loExistencias.mrRecuperaDatos()
        loExistencias.mnExistencias = loExistencias.mnExistencias + mnCantidad
        loExistencias.mdFechaUltEntrada = Now
        loExistencias.mrGrabaDatos()

    End Sub

    Public Sub mrGrabaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into traline values (" & mnEmpresa & ",'" & _
             mnCodigo & "','" & _
             mnLinea & "','" & _
             mnArticulo & "','" & _
             mnDetalle & "','" & _
             mfsRefina(msDescripcion) & "','" & _
             mnCantidad & "','" & _
             msEstado & "')"

            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        Else
            lsSql = "update traline set art_lin = '" & mnArticulo & _
                    "', det_lin = '" & mnDetalle & _
                    "', des_lin = '" & mfsRefina(msDescripcion) & _
                    "', ctd_lin = '" & mnCantidad & _
                    "', est_lin = '" & msEstado & _
                    "' where emp_lin = " & mnEmpresa & _
                    " and cod_lin = " & mnCodigo & _
                    " and lin_lin = " & mnLinea
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        End If
        mbEsNuevo = False
    End Sub

    Public Sub New()
        mbEsNuevo = True
    End Sub

End Class
