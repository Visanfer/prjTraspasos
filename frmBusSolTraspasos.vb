Option Explicit On

Imports System.Windows.Forms.SendKeys
Imports prjControl
Imports prjAlmacen

Public Class frmBusSolTraspasos
    Inherits System.Windows.Forms.Form
    Private mnEmpresa As Int32      ' empresa de gestion
    Private moSolTraspaso As New clsSolTraspaso
    Private moSolTraspasoAux As New clsSolTraspaso
    Private moBusSolTraspasos As New clsBusSolTraspasos
    Private mbEncontrado As Boolean
    Dim mcolAlmacenes As Collection

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
    Friend WithEvents lblPrograma As System.Windows.Forms.Label
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDesA2 As control.txtVisanfer
    Friend WithEvents txtDesA1 As control.txtVisanfer
    Friend WithEvents txtHasta As control.txtVisanfer
    Friend WithEvents txtDesde As control.txtVisanfer
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtOperador As control.txtVisanfer
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtComentario As control.txtVisanfer
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtFecha As control.txtVisanfer
    Friend WithEvents txtCodigo As control.txtVisanfer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grpLineas As System.Windows.Forms.GroupBox
    Friend WithEvents lstBusqueda As System.Windows.Forms.ListView
    Friend WithEvents CODIGO As System.Windows.Forms.ColumnHeader
    Friend WithEvents FECHA As System.Windows.Forms.ColumnHeader
    Friend WithEvents OBS As System.Windows.Forms.ColumnHeader
    Friend WithEvents DESDE As System.Windows.Forms.ColumnHeader
    Friend WithEvents HASTA As System.Windows.Forms.ColumnHeader
    Friend WithEvents OPER As System.Windows.Forms.ColumnHeader
    Friend WithEvents ESTADO As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtEstado As control.txtVisanfer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lstLineas As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTeclas = New System.Windows.Forms.Label
        Me.lblPrograma = New System.Windows.Forms.Label
        Me.panCampos = New System.Windows.Forms.Panel
        Me.txtEstado = New control.txtVisanfer
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDesA2 = New control.txtVisanfer
        Me.txtDesA1 = New control.txtVisanfer
        Me.txtHasta = New control.txtVisanfer
        Me.txtDesde = New control.txtVisanfer
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtOperador = New control.txtVisanfer
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtComentario = New control.txtVisanfer
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtFecha = New control.txtVisanfer
        Me.txtCodigo = New control.txtVisanfer
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
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
        Me.ESTADO = New System.Windows.Forms.ColumnHeader
        Me.panCampos.SuspendLayout()
        Me.grpLineas.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.Color.Silver
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(6, 711)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(1007, 24)
        Me.lblTeclas.TabIndex = 25
        Me.lblTeclas.Text = "BUSCAR - F5            CTRL+L - VER LINEAS           LIMPIAR DATOS - F7          " & _
        "   ESC - SALIDA"
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(6, 7)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(1007, 32)
        Me.lblPrograma.TabIndex = 26
        Me.lblPrograma.Text = "BUSQUEDA DE SOLICITUDES DE TRASPASOS"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.txtEstado)
        Me.panCampos.Controls.Add(Me.Label4)
        Me.panCampos.Controls.Add(Me.Label2)
        Me.panCampos.Controls.Add(Me.txtDesA2)
        Me.panCampos.Controls.Add(Me.txtDesA1)
        Me.panCampos.Controls.Add(Me.txtHasta)
        Me.panCampos.Controls.Add(Me.txtDesde)
        Me.panCampos.Controls.Add(Me.Label24)
        Me.panCampos.Controls.Add(Me.txtOperador)
        Me.panCampos.Controls.Add(Me.Label7)
        Me.panCampos.Controls.Add(Me.txtComentario)
        Me.panCampos.Controls.Add(Me.Label14)
        Me.panCampos.Controls.Add(Me.txtFecha)
        Me.panCampos.Controls.Add(Me.txtCodigo)
        Me.panCampos.Controls.Add(Me.Label3)
        Me.panCampos.Controls.Add(Me.Label1)
        Me.panCampos.Controls.Add(Me.grpLineas)
        Me.panCampos.Controls.Add(Me.lstBusqueda)
        Me.panCampos.Location = New System.Drawing.Point(6, 41)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(1007, 665)
        Me.panCampos.TabIndex = 27
        '
        'txtEstado
        '
        Me.txtEstado.AutoSelec = False
        Me.txtEstado.BackColor = System.Drawing.Color.White
        Me.txtEstado.Blink = False
        Me.txtEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtEstado.DesdeCodigo = CType(0, Long)
        Me.txtEstado.DesdeFecha = New Date(CType(0, Long))
        Me.txtEstado.HastaCodigo = CType(0, Long)
        Me.txtEstado.HastaFecha = New Date(CType(0, Long))
        Me.txtEstado.Location = New System.Drawing.Point(368, 47)
        Me.txtEstado.MaxLength = 1
        Me.txtEstado.Name = "txtEstado"
        Me.txtEstado.Size = New System.Drawing.Size(34, 20)
        Me.txtEstado.TabIndex = 4
        Me.txtEstado.Text = ""
        Me.txtEstado.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtEstado.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtEstado.ValorMax = 999999999
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(136, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(224, 16)
        Me.Label4.TabIndex = 124
        Me.Label4.Text = "ESTADO (P-PENDIENTE, A-ATENDIDO):"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(544, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 16)
        Me.Label2.TabIndex = 122
        Me.Label2.Text = "HASTA:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDesA2
        '
        Me.txtDesA2.AutoSelec = False
        Me.txtDesA2.BackColor = System.Drawing.Color.White
        Me.txtDesA2.Blink = False
        Me.txtDesA2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDesA2.DesdeCodigo = CType(0, Long)
        Me.txtDesA2.DesdeFecha = New Date(CType(0, Long))
        Me.txtDesA2.HastaCodigo = CType(0, Long)
        Me.txtDesA2.HastaFecha = New Date(CType(0, Long))
        Me.txtDesA2.Location = New System.Drawing.Point(648, 47)
        Me.txtDesA2.MaxLength = 30
        Me.txtDesA2.Name = "txtDesA2"
        Me.txtDesA2.ReadOnly = True
        Me.txtDesA2.Size = New System.Drawing.Size(329, 20)
        Me.txtDesA2.TabIndex = 121
        Me.txtDesA2.TabStop = False
        Me.txtDesA2.Text = ""
        Me.txtDesA2.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA2.ValorMax = 999999999
        '
        'txtDesA1
        '
        Me.txtDesA1.AutoSelec = False
        Me.txtDesA1.BackColor = System.Drawing.Color.White
        Me.txtDesA1.Blink = False
        Me.txtDesA1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDesA1.DesdeCodigo = CType(0, Long)
        Me.txtDesA1.DesdeFecha = New Date(CType(0, Long))
        Me.txtDesA1.HastaCodigo = CType(0, Long)
        Me.txtDesA1.HastaFecha = New Date(CType(0, Long))
        Me.txtDesA1.Location = New System.Drawing.Point(648, 12)
        Me.txtDesA1.MaxLength = 30
        Me.txtDesA1.Name = "txtDesA1"
        Me.txtDesA1.ReadOnly = True
        Me.txtDesA1.Size = New System.Drawing.Size(329, 20)
        Me.txtDesA1.TabIndex = 120
        Me.txtDesA1.TabStop = False
        Me.txtDesA1.Text = ""
        Me.txtDesA1.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA1.ValorMax = 999999999
        '
        'txtHasta
        '
        Me.txtHasta.AutoSelec = False
        Me.txtHasta.BackColor = System.Drawing.Color.White
        Me.txtHasta.Blink = False
        Me.txtHasta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtHasta.DesdeCodigo = CType(0, Long)
        Me.txtHasta.DesdeFecha = New Date(CType(0, Long))
        Me.txtHasta.HastaCodigo = CType(0, Long)
        Me.txtHasta.HastaFecha = New Date(CType(0, Long))
        Me.txtHasta.Location = New System.Drawing.Point(608, 47)
        Me.txtHasta.MaxLength = 2
        Me.txtHasta.Name = "txtHasta"
        Me.txtHasta.Size = New System.Drawing.Size(40, 20)
        Me.txtHasta.TabIndex = 5
        Me.txtHasta.Text = ""
        Me.txtHasta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtHasta.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtHasta.ValorMax = 999999999
        '
        'txtDesde
        '
        Me.txtDesde.AutoSelec = False
        Me.txtDesde.BackColor = System.Drawing.Color.White
        Me.txtDesde.Blink = False
        Me.txtDesde.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDesde.DesdeCodigo = CType(0, Long)
        Me.txtDesde.DesdeFecha = New Date(CType(0, Long))
        Me.txtDesde.HastaCodigo = CType(0, Long)
        Me.txtDesde.HastaFecha = New Date(CType(0, Long))
        Me.txtDesde.Location = New System.Drawing.Point(608, 12)
        Me.txtDesde.MaxLength = 2
        Me.txtDesde.Name = "txtDesde"
        Me.txtDesde.Size = New System.Drawing.Size(40, 20)
        Me.txtDesde.TabIndex = 2
        Me.txtDesde.Text = ""
        Me.txtDesde.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDesde.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtDesde.ValorMax = 999999999
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(544, 16)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(59, 16)
        Me.Label24.TabIndex = 119
        Me.Label24.Text = "DESDE:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOperador
        '
        Me.txtOperador.AutoSelec = False
        Me.txtOperador.BackColor = System.Drawing.Color.White
        Me.txtOperador.Blink = False
        Me.txtOperador.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtOperador.DesdeCodigo = CType(0, Long)
        Me.txtOperador.DesdeFecha = New Date(CType(0, Long))
        Me.txtOperador.HastaCodigo = CType(0, Long)
        Me.txtOperador.HastaFecha = New Date(CType(0, Long))
        Me.txtOperador.Location = New System.Drawing.Point(96, 47)
        Me.txtOperador.MaxLength = 3
        Me.txtOperador.Name = "txtOperador"
        Me.txtOperador.Size = New System.Drawing.Size(34, 20)
        Me.txtOperador.TabIndex = 3
        Me.txtOperador.Text = ""
        Me.txtOperador.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOperador.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtOperador.ValorMax = 999999999
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(20, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 118
        Me.Label7.Text = "OPERARIO:"
        '
        'txtComentario
        '
        Me.txtComentario.AutoSelec = False
        Me.txtComentario.BackColor = System.Drawing.Color.White
        Me.txtComentario.Blink = False
        Me.txtComentario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtComentario.DesdeCodigo = CType(0, Long)
        Me.txtComentario.DesdeFecha = New Date(CType(0, Long))
        Me.txtComentario.HastaCodigo = CType(0, Long)
        Me.txtComentario.HastaFecha = New Date(CType(0, Long))
        Me.txtComentario.Location = New System.Drawing.Point(140, 82)
        Me.txtComentario.MaxLength = 30
        Me.txtComentario.Name = "txtComentario"
        Me.txtComentario.Size = New System.Drawing.Size(837, 20)
        Me.txtComentario.TabIndex = 6
        Me.txtComentario.Text = ""
        Me.txtComentario.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtComentario.ValorMax = 999999999
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(20, 86)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(112, 16)
        Me.Label14.TabIndex = 116
        Me.Label14.Text = "OBSERVACIONES:"
        '
        'txtFecha
        '
        Me.txtFecha.AutoSelec = False
        Me.txtFecha.BackColor = System.Drawing.Color.White
        Me.txtFecha.Blink = False
        Me.txtFecha.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFecha.DesdeCodigo = CType(0, Long)
        Me.txtFecha.DesdeFecha = New Date(CType(0, Long))
        Me.txtFecha.HastaCodigo = CType(0, Long)
        Me.txtFecha.HastaFecha = New Date(CType(0, Long))
        Me.txtFecha.Location = New System.Drawing.Point(368, 12)
        Me.txtFecha.MaxLength = 21
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(72, 20)
        Me.txtFecha.TabIndex = 1
        Me.txtFecha.Text = ""
        Me.txtFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFecha.tipo = control.txtVisanfer.TiposCaja.RangoFecha
        Me.txtFecha.ValorMax = 999999999
        '
        'txtCodigo
        '
        Me.txtCodigo.AutoSelec = False
        Me.txtCodigo.BackColor = System.Drawing.Color.White
        Me.txtCodigo.Blink = False
        Me.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodigo.DesdeCodigo = CType(0, Long)
        Me.txtCodigo.DesdeFecha = New Date(CType(0, Long))
        Me.txtCodigo.HastaCodigo = CType(0, Long)
        Me.txtCodigo.HastaFecha = New Date(CType(0, Long))
        Me.txtCodigo.Location = New System.Drawing.Point(96, 12)
        Me.txtCodigo.MaxLength = 13
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(64, 20)
        Me.txtCodigo.TabIndex = 0
        Me.txtCodigo.Text = ""
        Me.txtCodigo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCodigo.tipo = control.txtVisanfer.TiposCaja.RangoCodigo
        Me.txtCodigo.ValorMax = 999999999
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(288, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 16)
        Me.Label3.TabIndex = 114
        Me.Label3.Text = "FECHA:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(20, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 110
        Me.Label1.Text = "CODIGO:"
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
        Me.lstLineas.TabIndex = 9
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
        Me.lstBusqueda.Location = New System.Drawing.Point(8, 112)
        Me.lstBusqueda.MultiSelect = False
        Me.lstBusqueda.Name = "lstBusqueda"
        Me.lstBusqueda.Size = New System.Drawing.Size(989, 216)
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
        'ESTADO
        '
        Me.ESTADO.Text = "ESTADO"
        Me.ESTADO.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ESTADO.Width = 100
        '
        'frmBusSolTraspasos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1018, 743)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTeclas)
        Me.Controls.Add(Me.lblPrograma)
        Me.Controls.Add(Me.panCampos)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBusSolTraspasos"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BUSQUEDA DE SOLICITUDES DE TRASPASO"
        Me.panCampos.ResumeLayout(False)
        Me.grpLineas.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCargar(ByRef loSolTraspaso As clsSolTraspaso, ByVal lnEmpresa As Int32)
        txtFecha.Text = Format(Now, "dd/MM/yyyy")
        mrAsignaEventos()
        mnEmpresa = lnEmpresa
        moSolTraspaso = loSolTraspaso
        mrLimpiaFormulario()
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
        Dim loAlmacen As clsAlmacen
        Dim loBusAlmacen As clsBusAlmacenes

        ' primero miro si tengo la coleccion de almacenes completa *********
        If mcolAlmacenes Is Nothing Then
            loBusAlmacen = New clsBusAlmacenes
            loBusAlmacen.mnEmpresa = mnEmpresa
            loBusAlmacen.mrRecuperaAlmacenes()
            mcolAlmacenes = loBusAlmacen.mcolAlmacenes
        End If
        ' muevo los datos a las cajas de texto
        txtCodigo.Text = moSolTraspasoAux.mnCodigo
        txtFecha.Text = Format(moSolTraspasoAux.mdFecha, "dd/MM/yyyy")
        txtDesde.Text = moSolTraspasoAux.mnDesde
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moSolTraspasoAux.mnDesde
        loAlmacen = mcolAlmacenes(loAlmacen.mpsCodigo)
        txtDesA1.Text = loAlmacen.msNombre
        txtHasta.Text = moSolTraspasoAux.mnHasta
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moSolTraspasoAux.mnHasta
        loAlmacen = mcolAlmacenes(loAlmacen.mpsCodigo)
        txtDesA2.Text = loAlmacen.msNombre
        txtOperador.Text = moSolTraspasoAux.mnOperario
        txtComentario.Text = Replace(moSolTraspasoAux.msObservaciones, vbCrLf, " ")
        txtEstado.Text = moSolTraspasoAux.msEstado

        ' recupero los datos de las lineas **************************
        If moSolTraspasoAux.mcolLineas Is Nothing Then moSolTraspasoAux.mrRecuperaLineas()
        mrRellenaLineas()

    End Sub

    Private Sub mrLimpiaFormulario()
        Dim loControl As Object
        Dim loControlAux As Object
        Dim loCaja As control.txtVisanfer

        lstBusqueda.Items.Clear()
        lstLineas.Items.Clear()
        For Each loControl In panCampos.Controls
            If TypeOf loControl Is GroupBox Then
                For Each loControlAux In loControl.Controls
                    If TypeOf loControlAux Is control.txtVisanfer Then
                        loCaja = loControlAux
                        loCaja.Text = ""
                    End If
                Next
            End If
            If TypeOf loControl Is control.txtVisanfer Then
                loCaja = loControl
                loCaja.Text = ""
            End If
        Next
        txtFecha.Text = Format(Now, "dd/MM/yyyy")
        txtDesde.Text = ""  'goProfile.mnAlmacen

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
            Case Keys.Enter      ' Intro
                Select Case lsControl
                    Case "lstBusqueda"
                        mrPasaCodigo()
                        ' Aqui le paso los valores de la busqueda y mando el evento
                        moSolTraspasoAux.mrClonar(moSolTraspaso)
                        moSolTraspaso.mrMandaEvento()
                        Me.Close()
                End Select
                Send("{TAB}")
            Case Keys.Escape      ' Escape Salgo de la busqueda
                If lsControl = "lstLineas" Then
                    lstBusqueda.Focus()
                Else
                    Me.Close()
                End If
            Case Keys.F7     ' Limpia formulario
                mrLimpiaFormulario()
                txtCodigo.Focus()
            Case Keys.F5     ' Busqueda
                mrBuscar()
                If mbEncontrado Then
                    lstBusqueda.Focus()
                    Send("{DOWN}")
                Else
                    txtCodigo.Focus()
                End If
        End Select

    End Sub

    Private Sub mrBuscar()

        mbEncontrado = False
        moBusSolTraspasos = New clsBusSolTraspasos
        moBusSolTraspasos.mnEmpresa = mnEmpresa
        moBusSolTraspasos.mnCodigo = mfnInteger(txtCodigo.Text)
        moBusSolTraspasos.mnDesdeCodigo = mfnInteger(txtCodigo.DesdeCodigo)
        moBusSolTraspasos.mnHastaCodigo = mfnInteger(txtCodigo.HastaCodigo)
        moBusSolTraspasos.mdFecha = mfdFecha(txtFecha.Text)
        moBusSolTraspasos.mdDesdeFecha = mfdFecha(txtFecha.DesdeFecha)
        moBusSolTraspasos.mdHastaFecha = mfdFecha(txtFecha.HastaFecha)
        moBusSolTraspasos.mnDesde = mfnInteger(txtDesde.Text)
        moBusSolTraspasos.mnHasta = mfnInteger(txtHasta.Text)
        moBusSolTraspasos.mnOperario = mfnInteger(txtOperador.Text)
        moBusSolTraspasos.mnVendedor = 0
        moBusSolTraspasos.msObservaciones = Trim(txtComentario.Text)
        moBusSolTraspasos.msEstado = Trim(txtEstado.Text)
        If Not mfbConsultaVacia() Then
            Cursor = Cursors.WaitCursor
            moBusSolTraspasos.mrBuscaSolTraspasos()
            Cursor = Cursors.Default
            If moBusSolTraspasos.mcolSolTraspasos.Count > 0 Then
                mbEncontrado = True
                mrRellenaList()
            Else
                MsgBox("No existen registros que cumplan esos criterios.", vbCritical + vbOKOnly, _
                        "Gestion de Traspasos")
                mrLimpiaFormulario()
            End If
        End If

    End Sub

    Private Function mfbConsultaVacia() As Boolean
        Dim loControl As Windows.Forms.Control
        Dim lmRespuesta As MsgBoxResult
        Dim lbVacio As Boolean

        lbVacio = True
        For Each loControl In panCampos.Controls
            If Mid(loControl.Name, 1, 3) = "txt" Then
                If loControl.Text <> "" Then lbVacio = False
            End If
        Next

        If lbVacio Then
            lmRespuesta = MsgBox("No ha especificado ningun parametro de busqueda. ¿Desea Continuar?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, txtCodigo.Text)
            If lmRespuesta = MsgBoxResult.Yes Then lbVacio = False
        End If

        ' de momento no filtto los campos basicos
        'lbVacio = False
        mfbConsultaVacia = lbVacio

    End Function

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
            loSubItem.Text = Replace(loTraspaso.msObservaciones, vbCrLf, " ")
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
        txtCodigo.Focus()
    End Sub

#End Region

#Region " Eventos de Teclado para los controles "

    Private Sub lstBusqueda_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstBusqueda.SelectedIndexChanged
        mrPasaCodigo()
        mrMoverCampos()
    End Sub

#End Region

End Class
