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

        Dim loTraspaso As clsTraspaso
        Dim lsSql As String

        mcolTraspasos = New Collection
        lsSql = "select * from tracabe " & mfsWhereOrden() & " order by cod_tra desc"
        Dim loDatos As DataTable = New clsControlBD().mfoRecuperaDatos(False, lsSql, "tracabe")
        For Each loRegistro As DataRow In loDatos.Rows
            loTraspaso = New clsTraspaso
            loTraspaso.mrCargaDatos(loRegistro)
            loTraspaso.mbEsNuevo = False
            mcolTraspasos.Add(loTraspaso, loTraspaso.mpsCodigo)
        Next

    End Sub

    Public Sub mrBuscaTraspasosPendientes(ByVal lnOrden As Integer)

        Dim loTraspaso As clsTraspaso
        Dim loLinea As clsTraspasoLin


        mcolLineas = New Collection
        mcolTraspasos = New Collection
        Dim lsSql As String = "select * from tracabe left join traline" &
            " on tracabe.emp_tra=traline.emp_lin And tracabe.cod_tra=traline.cod_lin" &
            " where emp_tra = " & mnEmpresa &
            " and has_tra = " & mnHasta &
            " and est_lin='N'"
        If lnOrden = 1 Then
            lsSql = lsSql & " order by cod_lin,lin_lin asc"
        Else
            lsSql = lsSql & " order by des_lin,cod_lin asc"
        End If
        Dim loDatos As DataTable = New clsControlBD().mfoRecuperaDatos(False, lsSql, "tracabe")
        For Each loRegistro As DataRow In loDatos.Rows
            loLinea = New clsTraspasoLin
            loLinea.mnEmpresa = mfnInteger(loRegistro("emp_lin") & "")
            loLinea.mnCodigo = mfnLong(loRegistro("cod_lin") & "")
            loLinea.mnLinea = mfnInteger(loRegistro("lin_lin") & "")
            loLinea.mnArticulo = mfnLong(loRegistro("art_lin") & "")
            loLinea.mnDetalle = mfnInteger(loRegistro("det_lin") & "")
            loLinea.msDescripcion = Trim(loRegistro("des_lin") & "")
            loLinea.mnCantidad = mfnDouble(loRegistro("ctd_lin") & "")
            loLinea.msEstado = Trim(loRegistro("est_lin") & "")
            loLinea.mbEsNuevo = False
            mcolLineas.Add(loLinea, loLinea.mpsCodigo)

            loTraspaso = New clsTraspaso
            loTraspaso.mrCargaDatos(loRegistro)
            loTraspaso.mbEsNuevo = False
            On Error Resume Next
            mcolTraspasos.Add(loTraspaso, loTraspaso.mpsCodigo)
            On Error GoTo 0
        Next

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
                lsWhere = lsWhere & " fec_tra >= '" & Format(mdDesdeFecha, formatoFecha) & "' and"
                lsWhere = lsWhere & " fec_tra <= '" & Format(mdHastaFecha, formatoFecha) & "' and"
            End If
        Else
            If Format(mdFecha, "dd/MM/yyyy") <> "01/01/1900" Then
                lsWhere = lsWhere & " fec_tra = '" & Format(mdFecha, formatoFecha) & "' and"
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
