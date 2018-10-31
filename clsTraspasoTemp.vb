Option Explicit On
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsTraspasoTemp
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public mnOrigen As Integer
    Public mnDestino As Integer
    Public mdFecha As Date = Now
    Public msHora As String = Format(Now, "HH:mm:ss")
    Public mnAbierto As Integer = 1
    Public mnTraspaso As Long
    Public mnUsuario As Integer

    Public mbEsNuevo As Boolean = True

    Public Function mfnCodigoNuevo() As Long
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Function

        Dim lsSql As String
        Dim loRecord As MySqlDataReader
        Dim lnCodigo As Long = 0

        lsSql = "select max(codigo) as ultimo from traspasotemp where empresa = " & mnEmpresa

        Dim loComando As New MySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader

        While loRecord.Read
            lnCodigo = mfnLong(loRecord("ultimo") & "")
        End While
        loRecord.Close()
        lconConexion.Close()

        Return lnCodigo + 1

    End Function

    Public Sub mrRecuperaDatos()
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As MySqlDataReader


        If mnTraspaso > 0 Then
            lsSql = "select * from traspasotemp where traspaso = " & mnTraspaso
        Else
            lsSql = "select * from traspasotemp where codigo = " & mnCodigo
        End If

        Dim loComando As New MySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        mbEsNuevo = True
        While loRecord.Read
            mrCargaDatos(loRecord)
            mbEsNuevo = False
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrCargaDatos(ByVal loRecord As MySqlDataReader)
        mnEmpresa = mfnInteger(loRecord("empresa") & "")
        mnCodigo = mfnLong(loRecord("codigo") & "")
        mnOrigen = mfnInteger(loRecord("origen") & "")
        mnDestino = mfnInteger(loRecord("destino") & "")
        mdFecha = mfdFecha(loRecord("fecha") & "")
        msHora = Trim(loRecord("hora") & "")
        mnAbierto = mfnInteger(loRecord("abierto") & "")
        mnTraspaso = mfnLong(loRecord("traspaso") & "")
        mnUsuario = mfnInteger(loRecord("usuario") & "")
    End Sub

    Public Sub mrGrabaDatos()

        If mnCodigo = 0 Then mnCodigo = mfnCodigoNuevo()
        If mnTraspaso = 0 Then
            ' recupero un codigo de albaran de proveedor para luego grabar definitivamente
            Dim loTraspaso As New clsTraspaso
            loTraspaso.mnEmpresa = mnEmpresa
            loTraspaso.mrNuevoCodigo()
            mnTraspaso = loTraspaso.mnCodigo
        End If

        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New MySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into traspasotemp values (" & mnEmpresa & "," &
                    mnCodigo & "," &
                    mnOrigen & "," &
                    mnDestino & ",'" &
                    Format(mdFecha, formatoFecha) & "','" &
                    msHora & "',1," &
                    mnTraspaso & "," &
                    mnUsuario & ")"
        Else
            lsSql = "update traspasotemp set origen = " & mnOrigen &
                ", destino = " & mnDestino &
                ", fecha = '" & Format(mdFecha, formatoFecha) &
                "', hora = '" & msHora &
                "', abierto = " & mnAbierto &
                ", traspaso = " & mnTraspaso &
                ", usuario = " & mnUsuario &
                " where empresa = " & mnEmpresa &
                " and codigo = " & mnCodigo
        End If
        mbEsNuevo = False

        loComando = New MySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()

    End Sub

End Class
