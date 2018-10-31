Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsSolTraspasoLin
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public mnLinea As Integer
    Public mnArticulo As Long
    Public mnDetalle As Integer
    Public msDescripcion As String
    Public mnCantidad As Double
    Public msEstado As String

    Public mbEsNuevo As Boolean

    Public Function mpsCodigo() As String
        mpsCodigo = "clsSolTraspasoLin-" & mnEmpresa & "-" & mnCodigo & "-" & mnLinea
    End Function

    Public Sub mrRecuperaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loFuente As mySqlDataReader

        lsSql = "select * from soltraline where emp_lin = " & mnEmpresa & _
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

    Public Sub mrBorraDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        lsSql = "delete from soltraline where emp_lin = " & mnEmpresa & _
                " and cod_lin = " & mnCodigo & _
                " and lin_lin = " & mnLinea
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()
        mbEsNuevo = True

    End Sub

    Public Sub mrGrabaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into soltraline values (" & mnEmpresa & ",'" & _
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
            lsSql = "update soltraline set art_lin = '" & mnArticulo & _
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
