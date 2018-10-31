Option Explicit On
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsBusTraspasoTemp

    Public mnEmpresa As Integer
    Public mnOrigen As Integer
    Public mnDestino As Integer
    Public mnUsuario As Integer

    Public mcolTraspasosTemp As Collection

    Public Sub mrBuscaTraspasosAbiertos()
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loTraspaso As clsTraspasoTemp
        Dim loRecord As MySqlDataReader
        Dim lsSql As String


        mcolTraspasosTemp = New Collection
        lsSql = "select * from traspasotemp where empresa = " & mnEmpresa &
                " and origen = " & mnOrigen
        If mnDestino > 0 Then lsSql = lsSql & " and destino = " & mnDestino
        If mnUsuario > 0 Then lsSql = lsSql & " and usuario = " & mnUsuario
        lsSql = lsSql & " and fecha = '" & Format(Now, formatoFecha) & "' and abierto=1 order by codigo desc"
        Dim loComando As New MySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            loTraspaso = New clsTraspasoTemp
            loTraspaso.mrCargaDatos(loRecord)
            loTraspaso.mbEsNuevo = False
            mcolTraspasosTemp.Add(loTraspaso)
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

End Class
