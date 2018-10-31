Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsBusTraspasos
    Public mnEmpresa As Integer
    Public mnCodigo As Integer
    Public mnDesdeCodigo As Integer
    Public mnHastaCodigo As Integer
    Public mdFecha As Date
    Public mdDesdeFecha As Date
    Public mdHastaFecha As Date
    Public mnDesde As Integer
    Public mnHasta As Integer
    Public mnOperario As Integer
    Public mnVendedor As Integer
    Public msObservaciones As String

    Public mcolTraspasos As Collection
    Public mcolLineas As Collection

    Public Sub mrBuscaTraspasos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loTraspaso As clsTraspaso
        Dim loRecord As mySqlDataReader
        Dim lsSql As String

        mcolTraspasos = New Collection
        lsSql = "select * from tracabe " & mfsWhereOrden() & " order by cod_tra desc"
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            loTraspaso = New clsTraspaso
            loTraspaso.mrCargaDatos(loRecord)
            loTraspaso.mbEsNuevo = False
            mcolTraspasos.Add(loTraspaso, loTraspaso.mpsCodigo)
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrBuscaTraspasosPendientes(ByVal lnOrden As Integer)
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loTraspaso As clsTraspaso
        Dim loLinea As clsTraspasoLin
        Dim loFuente As mySqlDataReader
        Dim lsSql As String

        mcolLineas = New Collection
        mcolTraspasos = New Collection
        lsSql = "select * from traline where emp_lin = " & mnEmpresa & _
                " and est_lin = 'N' and cod_lin in (" & _
                "select cod_tra from tracabe where emp_tra = " & mnEmpresa & _
                " and cod_tra in (select cod_lin from traline" & _
                " where emp_lin = " & mnEmpresa & _
                " and est_lin = 'N') and has_tra = " & mnHasta & ")"
        lsSql = "select b.* from tracabe a, traline b where" & _
                " a.emp_tra = b.emp_lin and a.cod_tra = b.cod_lin" & _
                " and b.emp_lin = " & mnEmpresa & _
                " and b.est_lin = 'N'" & _
                " and a.has_tra = " & mnHasta
        If lnOrden = 1 Then
            lsSql = lsSql & " order by b.cod_lin,b.lin_lin asc"
        Else
            lsSql = lsSql & " order by b.des_lin,b.cod_lin asc"
        End If
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loFuente = loComando.ExecuteReader
        While loFuente.Read
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

        For Each loLinea In mcolLineas
            ' ahora gestiono el traspaso y veo si es suyo **********
            loTraspaso = New clsTraspaso
            loTraspaso.mnEmpresa = loLinea.mnEmpresa
            loTraspaso.mnCodigo = loLinea.mnCodigo
            On Error Resume Next
            loTraspaso = mcolTraspasos(loTraspaso.mpsCodigo)
            On Error GoTo 0
            If loTraspaso.mbEsNuevo Then
                loTraspaso.mrRecuperaDatos()
                loTraspaso.mcolLineas = New Collection
                mcolTraspasos.Add(loTraspaso, loTraspaso.mpsCodigo)
            End If
            'If loTraspaso.mnHasta = mnHasta Then
            '    loTraspaso.mcolLineas.Add(loLinea, loLinea.mpsCodigo)
            'End If
        Next

        ' ahora borro los traspasos que no quiero ***************
        'For Each loTraspaso In mcolTraspasos
        '    If loTraspaso.mnHasta <> mnHasta Then
        '        mcolTraspasos.Remove(loTraspaso.mpsCodigo)
        '    End If
        'Next

    End Sub

    Private Function mfsWhereOrden() As String
        Dim lsWhere As String = ""

        If Not mfnInteger(mnEmpresa) = 0 Then
            lsWhere = lsWhere & " emp_tra = " & mnEmpresa & " and"
        End If

        If Not mfnInteger(mnHastaCodigo) = 0 Then
            If Not mfnInteger(mnDesdeCodigo) = 0 Then
                lsWhere = lsWhere & " cod_tra >= " & mnDesdeCodigo & " and"
                lsWhere = lsWhere & " cod_tra <= " & mnHastaCodigo & " and"
            End If
        Else
            If Not mfnInteger(mnCodigo) = 0 Then
                lsWhere = lsWhere & " cod_tra = " & mnCodigo & " and"
            End If
        End If

        If Format(mdHastaFecha, "dd/MM/yyyy") <> "01/01/1900" Then
            If Format(mdDesdeFecha, "dd/MM/yyyy") <> "01/01/1900" Then
                lsWhere = lsWhere & " fec_tra >= '" & Format(mdDesdeFecha, "yyyy/MM/dd") & "' and"
                lsWhere = lsWhere & " fec_tra <= '" & Format(mdHastaFecha, "yyyy/MM/dd") & "' and"
            End If
        Else
            If Format(mdFecha, "dd/MM/yyyy") <> "01/01/1900" Then
                lsWhere = lsWhere & " fec_tra = '" & Format(mdFecha, "yyyy/MM/dd") & "' and"
            End If
        End If

        If Not mfnInteger(mnDesde) = 0 Then
            lsWhere = lsWhere & " des_tra = " & mnDesde & " and"
        End If

        If Not mfnInteger(mnHasta) = 0 Then
            lsWhere = lsWhere & " has_tra = " & mnHasta & " and"
        End If

        If Not mfnInteger(mnOperario) = 0 Then
            lsWhere = lsWhere & " ope_tra = " & mnOperario & " and"
        End If

        If Not mfnInteger(mnVendedor) = 0 Then
            lsWhere = lsWhere & " ven_tra = " & mnVendedor & " and"
        End If

        If Not Trim(msObservaciones) = "" Then
            lsWhere = lsWhere & " obs_tra like '" & msObservaciones & "' and"
        End If

        lsWhere = Replace(lsWhere, "*", "%")
        If lsWhere <> "" Then
            lsWhere = "where " & Mid(lsWhere, 1, Len(lsWhere) - 4)
        End If
        mfsWhereOrden = lsWhere

    End Function

End Class
