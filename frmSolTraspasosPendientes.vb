Option Explicit On 
Imports System.Windows.Forms.SendKeys
Imports prjControl

Public Class frmSolTraspasosPendientes
    Inherits System.Windows.Forms.Form
    Private mnEmpresa As Int32      ' empresa de gestion
    Private moBusSolTraspasos As New clsBusSolTraspasos
    Private moSolTraspasoAux As New clsSolTraspaso

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents lblTeclas As System.Windows.Forms.Label
    Friend WithEvents ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents grpLineas As System.Windows.Forms.GroupBox
    Friend WithEvents lstLineas As System.Windows.Forms.ListView
    Friend WithEvents lstBusqueda As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CODIGO As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHA As System.Windows.Forms.ColumnHeader
    Friend WithEvents OBS As System.Windows.Forms.ColumnHeader
    Friend WithEvents DESDE As System.Windows.Forms.ColumnHeader
    Friend WithEvents HASTA As System.Windows.Forms.ColumnHeader
    Friend WithEvents OPER As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblPrograma As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTeclas = New System.Windows.Forms.Label
        Me.ESTADO = New System.Windows.Forms.ColumnHeader
        Me.panCampos = New System.Windows.Forms.Panel
        Me.grpLineas = New System.Windows.Forms.GroupBox
        Me.lstLineas = New System.Windows.Forms.ListView
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.lstBusqueda = New System.Windows.Forms.ListView
        Me.CODIGO = New System.Windows.Forms.ColumnHeader
        Me.FECHA = New System.Windows.Forms.ColumnHeader
        Me.OBS = New System.Windows.Forms.ColumnHeader
        Me.DESDE = New System.Windows.Forms.ColumnHeader
        Me.HASTA = New System.Windows.Forms.ColumnHeader
        Me.OPER = New System.Windows.Forms.ColumnHeader
        Me.lblPrograma = New System.Windows.Forms.Label
        Me.panCampos.SuspendLayout()
        Me.grpLineas.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.Color.Silver
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(5, 710)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(1007, 24)
        Me.lblTeclas.TabIndex = 28
        Me.lblTeclas.Text = "CONFIRMAR - F8                                       CTRL+L - VER LINEAS         " & _
        "                            ESC - SALIDA"
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ESTADO
        '
        Me.ESTADO.Text = "ESTADO"
        Me.ESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ESTADO.Width = 100
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.grpLineas)
        Me.panCampos.Controls.Add(Me.lstBusqueda)
        Me.panCampos.Location = New System.Drawing.Point(5, 40)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(1007, 665)
        Me.panCampos.TabIndex = 30
        '
        'grpLineas
        '
        Me.grpLineas.Controls.Add(Me.lstLineas)
        Me.grpLineas.Location = New System.Drawing.Point(8, 336)
        Me.grpLineas.Name = "grpLineas"
        Me.grpLineas.Size = New System.Drawing.Size(989, 321)
        Me.grpLineas.TabIndex = 10
        Me.grpLineas.TabStop = False
        Me.grpLineas.Text = "LINEAS DEL TRASPASO"
        '
        'lstLineas
        '
        Me.lstLineas.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lstLineas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstLineas.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader1})
        Me.lstLineas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstLineas.GridLines = True
        Me.lstLineas.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstLineas.HideSelection = False
        Me.lstLineas.Location = New System.Drawing.Point(8, 24)
        Me.lstLineas.MultiSelect = False
        Me.lstLineas.Name = "lstLineas"
        Me.lstLineas.Size = New System.Drawing.Size(973, 288)
        Me.lstLineas.TabIndex = 8
        Me.lstLineas.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ARTICULO"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "DESCRIPCION"
        Me.ColumnHeader3.Width = 500
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "CTD."
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader4.Width = 100
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "OBSERVACIONES"
        Me.ColumnHeader1.Width = 250
        '
        'lstBusqueda
        '
        Me.lstBusqueda.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lstBusqueda.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstBusqueda.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CODIGO, Me.FECHA, Me.OBS, Me.DESDE, Me.HASTA, Me.OPER, Me.ESTADO})
        Me.lstBusqueda.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstBusqueda.GridLines = True
        Me.lstBusqueda.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstBusqueda.HideSelection = False
        Me.lstBusqueda.Location = New System.Drawing.Point(8, 5)
        Me.lstBusqueda.MultiSelect = False
        Me.lstBusqueda.Name = "lstBusqueda"
        Me.lstBusqueda.Size = New System.Drawing.Size(989, 323)
        Me.lstBusqueda.TabIndex = 7
        Me.lstBusqueda.View = System.Windows.Forms.View.Details
        '
        'CODIGO
        '
        Me.CODIGO.Text = "CODIGO"
        '
        'FECHA
        '
        Me.FECHA.Text = "FECHA"
        Me.FECHA.Width = 80
        '
        'OBS
        '
        Me.OBS.Text = "OBSERVACIONES"
        Me.OBS.Width = 500
        '
        'DESDE
        '
        Me.DESDE.Text = "DESDE"
        Me.DESDE.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.DESDE.Width = 80
        '
        'HASTA
        '
        Me.HASTA.Text = "HASTA"
        Me.HASTA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.HASTA.Width = 80
        '
        'OPER
        '
        Me.OPER.Text = "OPER"
        Me.OPER.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(5, 6)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(1007, 32)
        Me.lblPrograma.TabIndex = 29
        Me.lblPrograma.Text = "SOLICITUDES DE TRASPASO PENDIENTES"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmSolTraspasosPendientes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 741)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTeclas)
        Me.Controls.Add(Me.panCampos)
        Me.Controls.Add(Me.lblPrograma)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSolTraspasosPendientes"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SOLICITUD DE TRASPASOS PENDIENTES"
        Me.panCampos.ResumeLayout(False)
        Me.grpLineas.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCargar(ByVal loBusSolTraspaso As clsBusSolTraspasos)
        mrAsignaEventos()
        mnEmpresa = loBusSolTraspaso.mnEmpresa
        moBusSolTraspasos = loBusSolTraspaso
        mrRellenaList()
        Send("{DOWN}")
        Me.ShowDialog()
    End Sub

    Private Sub mrAsignaEventos()
        Dim loControl As Object
        Dim loControlAux As Object
        Dim loCaja As control.txtVisanfer

        For Each loControl In panCampos.Controls
            If TypeOf loControl Is GroupBox Then
                For Each loControlAux In loControl.Controls
                    If TypeOf loControlAux Is control.txtVisanfer Then
                        loCaja = loControlAux
                        AddHandler loCaja.KeyDown, AddressOf mrLeeTecla
                    End If
                Next
            End If
            If TypeOf loControl Is control.txtVisanfer Then
                loCaja = loControl
                AddHandler loCaja.KeyDown, AddressOf mrLeeTecla
            End If
        Next
        AddHandler lstBusqueda.KeyDown, AddressOf mrLeeTecla
        AddHandler lstLineas.KeyDown, AddressOf mrLeeTecla

    End Sub

    Private Sub mrMoverCampos()

        ' recupero los datos de las lineas **************************
        If moSolTraspasoAux.mcolLineas Is Nothing Then moSolTraspasoAux.mrRecuperaLineas()
        mrRellenaLineas()

    End Sub

    Private Sub mrPasaCodigo()
        Dim loItem As ListViewItem

        If lstBusqueda.SelectedItems.Count > 0 Then
            loItem = lstBusqueda.SelectedItems(0)
            moSolTraspasoAux = moBusSolTraspasos.mcolSolTraspasos(loItem.Tag)
        End If

    End Sub

    Private Sub mrLeeTecla(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loCaja As control.txtVisanfer
        Dim lolstView As ListView
        Dim lsControl As String = ""

        If TypeOf sender Is control.txtVisanfer Then
            loCaja = sender
            lsControl = loCaja.Name
        End If
        If TypeOf sender Is ListView Then
            lolstView = sender
            lsControl = lolstView.Name
        End If

        Select Case e.KeyValue
            Case Keys.L And e.Control = True     'CONTROL + L
                lstLineas.Focus()
                Send("{DOWN}")
            Case Keys.Enter
                Select Case lsControl
                    Case "lstBusqueda"
                        mrPasaCodigo()
                        Dim loSolTraspaso As New frmSolTraspasos
                        loSolTraspaso.mrCargarSolicitud(moSolTraspasoAux)
                End Select
            Case Keys.Escape      ' Escape Salgo de la busqueda
                If lsControl = "lstLineas" Then
                    lstBusqueda.Focus()
                Else
                    Me.Close()
                End If
            Case Keys.F8     ' confirmar linea
                mrConfirmaLinea()
        End Select

    End Sub

    Private Sub mrConfirmaLinea()

        Dim loItem As ListViewItem

        If lstBusqueda.SelectedItems.Count > 0 Then
            loItem = lstBusqueda.SelectedItems(0)
            moSolTraspasoAux = moBusSolTraspasos.mcolSolTraspasos(loItem.Tag)

            If moSolTraspasoAux.msEstado = "P" Then
                moSolTraspasoAux.msEstado = "A"
                moSolTraspasoAux.mrGrabaDatos()

                mrRellenaList()
            End If
        End If

    End Sub

    Private Sub mrRellenaLineas()
        Dim loLinea As clsSolTraspasoLin
        Dim loItem As ListViewItem
        Dim loSubItem As ListViewItem.ListViewSubItem

        lstLineas.BeginUpdate()
        ' Eliminar el contenido previo
        lstLineas.Items.Clear()
        lstLineas.View = View.Details
        lstLineas.GridLines = True
        lstLineas.FullRowSelect = True
        For Each loLinea In moSolTraspasoAux.mcolLineas
            loItem = New ListViewItem
            loItem.Text = loLinea.mnArticulo                ' articulo
            If loLinea.mnDetalle > 0 Then loItem.Text = loLinea.mnArticulo & "." & loLinea.mnDetalle
            loItem.Tag = loLinea.mpsCodigo
            loSubItem = New ListViewItem.ListViewSubItem    ' descripcion
            loSubItem.Text = loLinea.msDescripcion
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' cantidad
            loSubItem.Text = loLinea.mnCantidad
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' OBSERVACIONES
            loSubItem.Text = loLinea.msEstado
            loItem.SubItems.Add(loSubItem)
            lstLineas.Items.Add(loItem)
        Next
        lstLineas.EndUpdate()

    End Sub

    Private Sub mrRellenaList()
        Dim loTraspaso As clsSolTraspaso
        Dim loItem As ListViewItem
        Dim loSubItem As ListViewItem.ListViewSubItem

        lstBusqueda.BeginUpdate()
        ' Eliminar el contenido previo
        lstBusqueda.Items.Clear()
        lstBusqueda.View = View.Details
        lstBusqueda.GridLines = True
        lstBusqueda.FullRowSelect = True
        For Each loTraspaso In moBusSolTraspasos.mcolSolTraspasos
            loItem = New ListViewItem
            loItem.Text = loTraspaso.mnCodigo              ' codigo
            loItem.Tag = loTraspaso.mpsCodigo
            loSubItem = New ListViewItem.ListViewSubItem    ' fecha
            loSubItem.Text = Format(loTraspaso.mdFecha, "dd/MM/yyyy")
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' observaciones
            loSubItem.Text = loTraspaso.msObservaciones
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' origen
            loSubItem.Text = loTraspaso.mnDesde
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' destino
            loSubItem.Text = loTraspaso.mnHasta
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' operador
            loSubItem.Text = loTraspaso.mnOperario
            loItem.SubItems.Add(loSubItem)
            loSubItem = New ListViewItem.ListViewSubItem    ' ESTADO
            loSubItem.Text = IIf(loTraspaso.msEstado = "A", "ATENDIDO", "PENDIENTE")
            loItem.SubItems.Add(loSubItem)
            loItem.BackColor = IIf(loTraspaso.msEstado = "A", Color.White, Color.Orange)
            lstBusqueda.Items.Add(loItem)
        Next
        lstBusqueda.EndUpdate()

    End Sub

#End Region

#Region " Eventos del formulario "

    Private Sub frmBusSolTraspasos_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.Text = goUsuario.msNombre
    End Sub

#End Region

#Region " Eventos de Teclado para los controles "

    Private Sub lstBusqueda_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstBusqueda.SelectedIndexChanged
        mrPasaCodigo()
        mrMoverCampos()
    End Sub

#End Region


End Class
