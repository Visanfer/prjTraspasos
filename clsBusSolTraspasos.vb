Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsBusSolTraspasos
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
    Public msEstado As String
    Public msObservaciones As String
    Public mbNuevaConexion As Boolean = False

    Public mcolSolTraspasos As Collection
    Public mcolLineas As Collection

    Public Sub mrBuscaSolTraspasos()
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loSolTraspaso As clsSolTraspaso
        Dim loRecord As MySqlDataReader
        Dim lsSql As String


        mcolSolTraspasos = New Collection
        lsSql = "select * from soltracabe " & mfsWhereOrden() & _
                " order by cod_tra"
        Dim loComando As New MySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            loSolTraspaso = New clsSolTraspaso
            loSolTraspaso.mnEmpresa = mfnInteger(loRecord("emp_tra") & "")
            loSolTraspaso.mnCodigo = mfnLong(loRecord("cod_tra") & "")
            loSolTraspaso.mnDesde = mfnInteger(loRecord("des_tra") & "")
            loSolTraspaso.mnHasta = mfnInteger(loRecord("has_tra") & "")
            loSolTraspaso.mdFecha = mfdFecha(loRecord("fec_tra") & "")
            loSolTraspaso.msObservaciones = Trim(loRecord("obs_tra") & "")
            loSolTraspaso.msEstado = Trim(loRecord("est_tra") & "")
            loSolTraspaso.mnVendedor = mfnInteger(loRecord("ven_tra") & "")
            loSolTraspaso.mnOperario = mfnInteger(loRecord("ope_tra") & "")
            loSolTraspaso.mbEsNuevo = False
            mcolSolTraspasos.Add(loSolTraspaso, loSolTraspaso.mpsCodigo)
        End While
        loRecord.Close()
        lconConexion.Close()

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

        If Not Trim(msEstado) = "" Then
            lsWhere = lsWhere & " est_tra = '" & msEstado & "' and"
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
