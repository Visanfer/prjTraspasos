Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsTraspaso
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public mnDesde As Integer
    Public mnHasta As Integer
    Public mdFecha As Date
    Public msObservaciones As String
    Public msEstado As String
    Public mnVendedor As Integer
    Public mnOperario As Integer
    Public msEstadoEnvio As String = ""
    Public mdFechaGrabacion As Date = Now
    Public mnOperarioEnvio As Integer = 0
    Public mdFechaEnvio As Date = Now
    Public mnOperarioRecepcion As Integer = 0
    Public mdFechaRecepcion As Date = Now

    Public mbEsNuevo As Boolean
    Public mcolLineas As Collection
    Public Event evtBusTraspaso()       ' evento desencadenado despues una busqueda

#Region " Funciones y Rutinas varias "

    Public Function mpsCodigo() As String
        mpsCodigo = "clsTraspaso-" & mnEmpresa & "-" & mnCodigo
    End Function

    Public Sub New()
        mbEsNuevo = True
    End Sub

    Public Sub mrMandaEvento()
        RaiseEvent evtBusTraspaso()
    End Sub

    Public Sub mrCrearTraspasoNuevo(ByVal lnEmpresa As Integer, ByRef loUsuario As clsUsuario)
        Dim loFormulario As New frmTraspasos
        loFormulario.mrCrearTraspasoNuevo(lnEmpresa, loUsuario, Me)
    End Sub

    Public Sub mrBuscaTraspaso()
        Dim loBusTraspaso As New frmBusTraspasos
        loBusTraspaso.mrCargar(Me, mnEmpresa)
    End Sub

    Public Sub mrClonar(ByRef loTraspaso As clsTraspaso)
        loTraspaso.mnEmpresa = mnEmpresa
        loTraspaso.mnCodigo = mnCodigo
        loTraspaso.mnDesde = mnDesde
        loTraspaso.mnHasta = mnHasta
        loTraspaso.mdFecha = mdFecha
        loTraspaso.msObservaciones = msObservaciones
        loTraspaso.msEstado = msEstado
        loTraspaso.mnVendedor = mnVendedor
        loTraspaso.mnOperario = mnOperario
        loTraspaso.msEstadoEnvio = msEstadoEnvio
        loTraspaso.mdFechaGrabacion = mdFechaGrabacion
        loTraspaso.mnOperarioEnvio = mnOperarioEnvio
        loTraspaso.mdFechaEnvio = mdFechaEnvio
        loTraspaso.mnOperarioRecepcion = mnOperarioRecepcion
        loTraspaso.mdFechaRecepcion = mdFechaRecepcion
        loTraspaso.mbEsNuevo = mbEsNuevo
        loTraspaso.mcolLineas = mcolLineas
    End Sub

    Public Sub mrActualizaTraspaso()

        ' guardo en un log el usuario que ha actualizado el traspaso *****************
        ' actualiza los almacenes y genera apuntes en el historico de movimientos ****
        For Each loLinea As clsTraspasoLin In mcolLineas

            mrActualizaLinea(loLinea)

            ' **** ahora actualizo el estado del traspaso ********
            msEstado = "A"
            mrGrabaDatos()
            ' ****************************************************
        Next

    End Sub

    Public Sub mrActualizaLinea(ByVal loLinea As clsTraspasoLin)

        Dim loExistencias As prjArticulos.clsExistencias
        Dim loMovimiento As prjArticulos.clsHisMovimientos

        Dim lnCantidad As Double = loLinea.mnCantidad
        Dim lnDetalle As Integer = loLinea.mnDetalle

        ' primero actualizo la salida *********************
        prjArticulos.goUsuario = modTraspasos.goUsuario
        prjArticulos.goProfile = modTraspasos.goProfile
        loExistencias = New prjArticulos.clsExistencias
        loExistencias.mnEmpresa = mnEmpresa
        loExistencias.mnAlmacen = mnDesde
        loExistencias.mnArticulo = loLinea.mnArticulo
        loExistencias.mnDetalle = lnDetalle
        loExistencias.mnExistencias = lnCantidad
        loExistencias.mdFechaUltSalida = Now
        loExistencias.mrActualizaExistencias()

        ' despues actualizo la entrada *********************
        ' ******** esto ya no se hace, ahora es necesario de una confirmacion
        ' del almacen destino ***********************************************
        ' domingo me dice que los traspasos a la fabrica se actualicen automaticamente 02/02/2016
        ' el motivo es por que generalmente son traspasos administrativos, es decir, se hace solo por que gines hace la
        ' salida desde la fabrica cuando manda un camion a la casa del cliente y el material no esta en la fabrica
        ' lo mismo pasa con los traspasos hacia el almacen 11 cuando el origen es el 10 - 21/12/2016

        Dim lbActualizaDestino As Boolean = False

        If loLinea.msEstado = "A" Or mnHasta = 8 Then lbActualizaDestino = True
        'If loLinea.msEstado = "A" Or mnHasta = 9 Then lbActualizaDestino = True  ' me dice gines que no lo ponga directo, que se encarga el
        If loLinea.msEstado = "A" Or (mnDesde = 10 And mnHasta = 11) Then lbActualizaDestino = True
        If loLinea.msEstado = "A" Or (mnDesde = 11 And mnHasta = 10) Then lbActualizaDestino = True
        If loLinea.msEstado = "A" Or (mnDesde = 1 And mnHasta = 13) Then lbActualizaDestino = True   ' 04-06-2020
        If loLinea.msEstado = "A" Or mnHasta = 14 Then lbActualizaDestino = True    ' 05/05/2022     TRASPASOS A ONDA

        If lbActualizaDestino Then
            ' si el estado de la linea es actualizado, entonces grabo las existencias en destino
            loExistencias = New prjArticulos.clsExistencias
            loExistencias.mnEmpresa = mnEmpresa
            loExistencias.mnAlmacen = mnHasta
            loExistencias.mnArticulo = loLinea.mnArticulo
            loExistencias.mnDetalle = loLinea.mnDetalle
            loExistencias.mnExistencias = (-1) * loLinea.mnCantidad
            loExistencias.mdFechaUltEntrada = Now
            loExistencias.mrActualizaExistencias()

            ' actualizo el estado de la linea, para QUE NO SALGA COMO PENDIENTE
            loLinea.msEstado = "A"
            loLinea.mrGrabaDatos()
        End If

        ' compruebo si el articulo es virtual y descuenta de otros ******************
        Dim loBusAgrArt As New prjArticulos.clsBusAgrArt
        Dim loAgrArt As New prjArticulos.clsAgrArt
        loAgrArt.mnEmpresa = mnEmpresa
        loAgrArt.mnCodigo = loLinea.mnArticulo
        loAgrArt.mnDetalle = loLinea.mnDetalle
        loBusAgrArt.mrBusca(loAgrArt)

        If loBusAgrArt.mcolAgrart.Count <> 0 Then
            For Each loAgrArt In loBusAgrArt.mcolAgrart
                ' *********** ahora apunto una linea de traspaso al historico *************
                loMovimiento = New prjArticulos.clsHisMovimientos
                loMovimiento.mnEmpresa = mnEmpresa
                loMovimiento.mdFecha = Now
                loMovimiento.mnDocumento = mnCodigo
                loMovimiento.mnAlbaran = 0
                loMovimiento.mnTipoMovimiento = "T"
                loMovimiento.mnArticulo = loAgrArt.mnArticulo
                loMovimiento.mnDetalle = loAgrArt.mnDetallePadre
                loMovimiento.msDescripcion = loLinea.msDescripcion
                loMovimiento.mnAlmacen1 = mnDesde
                loMovimiento.mnAlmacen2 = mnHasta
                loMovimiento.mnExistencias = loLinea.mnCantidad * loAgrArt.mnExistencias
                loMovimiento.mnOrden = loMovimiento.mfnContador()
                loMovimiento.mrGrabaDatos()
            Next
        Else
            ' *********** ahora apunto una linea de traspaso al historico *************
            loMovimiento = New prjArticulos.clsHisMovimientos
            loMovimiento.mnEmpresa = mnEmpresa
            loMovimiento.mdFecha = Now
            loMovimiento.mnDocumento = mnCodigo
            loMovimiento.mnAlbaran = 0
            loMovimiento.mnTipoMovimiento = "T"
            loMovimiento.mnArticulo = loLinea.mnArticulo
            loMovimiento.mnDetalle = loLinea.mnDetalle
            loMovimiento.msDescripcion = loLinea.msDescripcion
            loMovimiento.mnAlmacen1 = mnDesde
            loMovimiento.mnAlmacen2 = mnHasta
            loMovimiento.mnExistencias = lnCantidad
            loMovimiento.mnOrden = loMovimiento.mfnContador()
            loMovimiento.mrGrabaDatos()
        End If

    End Sub

    Public Sub mrActualizaTraspasoOld()
        ' guardo en un log el usuario que ha actualizado el traspaso *****************
        ' actualiza los almacenes y genera apuntes en el historico de movimientos ****
        Dim loExistencias As prjArticulos.clsExistencias
        Dim loMovimiento As prjArticulos.clsHisMovimientos
        Dim loLinea As clsTraspasoLin

        For Each loLinea In mcolLineas

            ' primero actualizo la salida *********************
            loExistencias = New prjArticulos.clsExistencias
            loExistencias.mnEmpresa = mnEmpresa
            loExistencias.mnAlmacen = mnDesde
            loExistencias.mnArticulo = loLinea.mnArticulo
            loExistencias.mnDetalle = loLinea.mnDetalle
            loExistencias.mrRecuperaDatos()
            loExistencias.mnExistencias = loExistencias.mnExistencias - loLinea.mnCantidad
            loExistencias.mdFechaUltSalida = Now
            loExistencias.mrGrabaDatos()

            ' despues actualizo la entrada *********************
            ' ******** esto ya no se hace, ahora es necesario de una confirmacion
            ' del almacen destino ***********************************************
            If loLinea.msEstado = "A" Then
                ' si el estado de la linea es actulizado, entonces grabo las existencias en destino
                loExistencias = New prjArticulos.clsExistencias
                loExistencias.mnEmpresa = mnEmpresa
                loExistencias.mnAlmacen = mnHasta
                loExistencias.mnArticulo = loLinea.mnArticulo
                loExistencias.mnDetalle = loLinea.mnDetalle
                loExistencias.mrRecuperaDatos()
                loExistencias.mnExistencias = loExistencias.mnExistencias + loLinea.mnCantidad
                loExistencias.mdFechaUltEntrada = Now
                loExistencias.mrGrabaDatos()
            End If

            ' *********** ahora apunto una linea de traspaso al historico *************
            loMovimiento = New prjArticulos.clsHisMovimientos
            loMovimiento.mnEmpresa = mnEmpresa
            loMovimiento.mdFecha = Now
            loMovimiento.mnDocumento = mnCodigo
            loMovimiento.mnAlbaran = 0
            loMovimiento.mnTipoMovimiento = "T"
            loMovimiento.mnArticulo = loLinea.mnArticulo
            loMovimiento.mnDetalle = loLinea.mnDetalle
            loMovimiento.msDescripcion = loLinea.msDescripcion
            loMovimiento.mnAlmacen1 = mnDesde
            loMovimiento.mnAlmacen2 = mnHasta
            loMovimiento.mnExistencias = loLinea.mnCantidad
            loMovimiento.mnOrden = loMovimiento.mfnContador()
            loMovimiento.mrGrabaDatos()
            ' **** ahora actualizo el estado del traspaso ********
            msEstado = "A"
            mrGrabaDatos()
            ' ****************************************************
        Next

    End Sub

#End Region

#Region " Acceso a la base de Datos "

    Public Sub mrNuevoCodigo()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As mySqlDataReader
        Dim loComando As New mySqlCommand

        ' ******** primero selecciono el registro ***************************
        lsSql = "select ult_con from tracontr where alm_con = 0"
        loComando = New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            mnCodigo = mfnInteger(loRecord("ult_con") & "") + 1
        End While
        loRecord.Close()
        ' ******* despues incremento el contador ****************************
        lsSql = "update tracontr set ult_con = " & mnCodigo & " where alm_con = 0"
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()  ' actualiza sin control de concurrencia
        lconConexion.Close()

    End Sub

    Public Sub mrRecuperaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As mySqlDataReader

        lsSql = "select * from tracabe where emp_tra = " & mnEmpresa & _
                " and cod_tra = " & mnCodigo
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        mbEsNuevo = True
        While loRecord.Read
            mrCargaDatos(loRecord)
            mbEsNuevo = False
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrCargaDatos(ByVal loRecord As mySqlDataReader)
        mnEmpresa = mfnInteger(loRecord("emp_tra") & "")
        mnCodigo = mfnLong(loRecord("cod_tra") & "")
        mnDesde = mfnInteger(loRecord("des_tra") & "")
        mnHasta = mfnInteger(loRecord("has_tra") & "")
        mdFecha = mfdFecha(loRecord("fec_tra") & "")
        msObservaciones = Trim(loRecord("obs_tra") & "")
        msEstado = Trim(loRecord("est_tra") & "")
        mnVendedor = mfnInteger(loRecord("ven_tra") & "")
        mnOperario = mfnInteger(loRecord("ope_tra") & "")
        msEstadoEnvio = Trim(loRecord("estenvio") & "")
        mdFechaGrabacion = mfdFecha(loRecord("fecgraba") & "")
        mnOperarioEnvio = mfnInteger(loRecord("operenvio") & "")
        mdFechaEnvio = mfdFecha(loRecord("fecenvio") & "")
        mnOperarioRecepcion = mfnInteger(loRecord("operrecibe") & "")
        mdFechaRecepcion = mfdFecha(loRecord("fecrecibe") & "")
    End Sub

    Public Sub mrGrabaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into tracabe values ('" & mnEmpresa & "','" & _
                    mnCodigo & "','" & _
                    mnDesde & "','" & _
                    mnHasta & "','" & _
                    Format(mdFecha, "yyyy/MM/dd") & "','" & _
                    msObservaciones & "','" & _
                    msEstado & "','" & _
                    mnVendedor & "','" & _
                    mnOperario & "','" & _
                    msEstadoEnvio & "','" & _
                    Format(mdFechaGrabacion, "yyyy/MM/dd HH:mm:ss") & "','" & _
                    mnOperarioEnvio & "','" & _
                    Format(mdFechaEnvio, "yyyy/MM/dd HH:mm:ss") & "','" & _
                    mnOperarioRecepcion & "','" & _
                    Format(mdFechaRecepcion, "yyyy/MM/dd HH:mm:ss") & "')"
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        Else
            lsSql = "update tracabe set des_tra = '" & mnDesde & _
                "', has_tra = '" & mnHasta & _
                "', fec_tra = '" & Format(mdFecha, "yyyy/MM/dd") & _
                "', obs_tra = '" & msObservaciones & _
                "', est_tra = '" & msEstado & _
                "', ven_tra = '" & mnVendedor & _
                "', ope_tra = '" & mnOperario & _
                "', estenvio = '" & msEstadoEnvio & _
                "', fecgraba = '" & Format(mdFechaGrabacion, "yyyy/MM/dd HH:mm:ss") & _
                "', operenvio = '" & mnOperarioEnvio & _
                "', fecenvio = '" & Format(mdFechaEnvio, "yyyy/MM/dd HH:mm:ss") & _
                "', operrecibe = '" & mnOperarioRecepcion & _
                "', fecrecibe = '" & Format(mdFechaRecepcion, "yyyy/MM/dd HH:mm:ss") & _
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
        lsSql = "delete from tracabe where emp_tra = " & mnEmpresa & _
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
        Dim loLinea As clsTraspasoLin

        lsSql = "select * from traline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " order by lin_lin asc"

        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader()
        mcolLineas = New Collection
        While loFuente.Read()
            loLinea = New clsTraspasoLin
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

        lsSql = "delete from traline where emp_lin = " & mnEmpresa &
                " and cod_lin = " & mnCodigo
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()

    End Sub

    Public Function mfoStatusPendientes() As DataTable

        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Return Nothing

        Dim ldDesde As Date = DateAdd(DateInterval.Day, -7, Now)
        Dim loResultado As New DataSet
        Dim lsSql As String = "select has_tra as almacen, alm.nombre, count(*) as numero from tracabe cab, traline lin, almacenes alm" &
            " where cab.emp_tra=lin.emp_lin And cab.cod_tra=lin.cod_lin and cab.emp_tra=alm.empresa" &
            " and cab.has_tra=alm.codigo and cab.emp_tra = " & mnEmpresa &
            " and cab.fec_tra>='" & Format(ldDesde, formatoFecha) &
            "' and lin.est_lin = 'N' group by has_tra"

        Dim loDataAdapter As New MySqlDataAdapter(lsSql, lconConexion)
        loDataAdapter.Fill(loResultado, "almacenes")
        lconConexion.Close()

        Return loResultado.Tables("almacenes")

    End Function

#End Region

End Class
