Option Explicit On
Imports System.Windows.Forms.SendKeys
Imports prjControl
Imports prjPrinterNet
Imports prjEmpresas
Imports prjArticulos
Imports prjAlmacen
Imports Microsoft.Reporting.WinForms

Public Class frmSolTraspasos
    Inherits System.Windows.Forms.Form
    Private mnEmpresa As Int32      ' empresa de gestion
    Dim WithEvents moSolTraspaso As clsSolTraspaso
    Dim WithEvents moSolTraspasoAux As clsSolTraspaso
    Dim moAlmacen1 As clsAlmacen
    Dim moAlmacen2 As clsAlmacen
    Dim WithEvents moArticulo As clsArticulo
    Dim mbGrabando As Boolean = False
    Public Enum EstadoVentana       ' estados posibles de la ventana
        Consulta = 1
        Mantenimiento = 2
        NuevoRegistro = 3
        Lineas = 4
        Salida = 5
    End Enum
    Dim mtEstado As EstadoVentana
    Dim mbAviso1 As Boolean = False
    Dim mbAviso2 As Boolean = False
    Dim mcolLineas As Collection
    ' variables de impresion ************************
    Private WithEvents moSelImpresora As prjControl.frmSelImpresora
    Dim moPrinter As New clsPrinter       ' objeto para imprimir
    Dim moImpresora As New prjPrinterNet.clsImpresora   ' Objeto Impresora
    Dim moDataTraspaso As dtsTraspaso

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
    Friend WithEvents lblPrograma As System.Windows.Forms.Label
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdDesbloqueo As System.Windows.Forms.Button
    Friend WithEvents cmdBloqueo As System.Windows.Forms.Button
    Friend WithEvents lblTraspaso As System.Windows.Forms.Label
    Friend WithEvents lblAviso2 As System.Windows.Forms.Label
    Friend WithEvents lblAviso1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grdLineas As prjGrid.ctlGrid
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtOperador As control.txtVisanfer
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtEstado As control.txtVisanfer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtComentario As control.txtVisanfer
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtFecha As control.txtVisanfer
    Friend WithEvents txtCodigo As control.txtVisanfer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tmrAviso As System.Windows.Forms.Timer
    Friend WithEvents lblTitulo As System.Windows.Forms.Label
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents lblTeclas As System.Windows.Forms.Label
    Friend WithEvents txtDesde As control.txtVisanfer
    Friend WithEvents txtHasta As control.txtVisanfer
    Friend WithEvents txtExpo1 As control.txtVisanfer
    Friend WithEvents txtExpo2 As control.txtVisanfer
    Friend WithEvents txtExis1 As control.txtVisanfer
    Friend WithEvents txtExis2 As control.txtVisanfer
    Friend WithEvents txtDesA1 As control.txtVisanfer
    Friend WithEvents txtDesA2 As control.txtVisanfer
    Friend WithEvents txtDesOperario As control.txtVisanfer
    Friend WithEvents lblConfirmado As System.Windows.Forms.Label
    Friend WithEvents lblConfirma As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSolTraspasos))
        Me.lblPrograma = New System.Windows.Forms.Label()
        Me.panCampos = New System.Windows.Forms.Panel()
        Me.lblConfirmado = New System.Windows.Forms.Label()
        Me.lblConfirma = New System.Windows.Forms.Label()
        Me.txtDesOperario = New control.txtVisanfer()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtExpo1 = New control.txtVisanfer()
        Me.txtExpo2 = New control.txtVisanfer()
        Me.cmdDesbloqueo = New System.Windows.Forms.Button()
        Me.cmdBloqueo = New System.Windows.Forms.Button()
        Me.lblTraspaso = New System.Windows.Forms.Label()
        Me.lblAviso2 = New System.Windows.Forms.Label()
        Me.lblAviso1 = New System.Windows.Forms.Label()
        Me.txtExis1 = New control.txtVisanfer()
        Me.txtExis2 = New control.txtVisanfer()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grdLineas = New prjGrid.ctlGrid()
        Me.txtDesA1 = New control.txtVisanfer()
        Me.txtDesA2 = New control.txtVisanfer()
        Me.txtDesde = New control.txtVisanfer()
        Me.txtHasta = New control.txtVisanfer()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtOperador = New control.txtVisanfer()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtEstado = New control.txtVisanfer()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtComentario = New control.txtVisanfer()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtFecha = New control.txtVisanfer()
        Me.txtCodigo = New control.txtVisanfer()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tmrAviso = New System.Windows.Forms.Timer(Me.components)
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.lblTeclas = New System.Windows.Forms.Label()
        Me.panCampos.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(2, 1)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(661, 56)
        Me.lblPrograma.TabIndex = 46
        Me.lblPrograma.Text = "GESTION"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.lblConfirmado)
        Me.panCampos.Controls.Add(Me.lblConfirma)
        Me.panCampos.Controls.Add(Me.txtDesOperario)
        Me.panCampos.Controls.Add(Me.Label9)
        Me.panCampos.Controls.Add(Me.Label10)
        Me.panCampos.Controls.Add(Me.Label8)
        Me.panCampos.Controls.Add(Me.Label6)
        Me.panCampos.Controls.Add(Me.txtExpo1)
        Me.panCampos.Controls.Add(Me.txtExpo2)
        Me.panCampos.Controls.Add(Me.cmdDesbloqueo)
        Me.panCampos.Controls.Add(Me.cmdBloqueo)
        Me.panCampos.Controls.Add(Me.lblTraspaso)
        Me.panCampos.Controls.Add(Me.lblAviso2)
        Me.panCampos.Controls.Add(Me.lblAviso1)
        Me.panCampos.Controls.Add(Me.txtExis1)
        Me.panCampos.Controls.Add(Me.txtExis2)
        Me.panCampos.Controls.Add(Me.Label2)
        Me.panCampos.Controls.Add(Me.grdLineas)
        Me.panCampos.Controls.Add(Me.txtDesA1)
        Me.panCampos.Controls.Add(Me.txtDesA2)
        Me.panCampos.Controls.Add(Me.txtDesde)
        Me.panCampos.Controls.Add(Me.txtHasta)
        Me.panCampos.Controls.Add(Me.Label24)
        Me.panCampos.Controls.Add(Me.txtOperador)
        Me.panCampos.Controls.Add(Me.Label7)
        Me.panCampos.Controls.Add(Me.txtEstado)
        Me.panCampos.Controls.Add(Me.Label5)
        Me.panCampos.Controls.Add(Me.txtComentario)
        Me.panCampos.Controls.Add(Me.Label14)
        Me.panCampos.Controls.Add(Me.txtFecha)
        Me.panCampos.Controls.Add(Me.txtCodigo)
        Me.panCampos.Controls.Add(Me.Label3)
        Me.panCampos.Controls.Add(Me.Label1)
        Me.panCampos.Location = New System.Drawing.Point(2, 57)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(1004, 632)
        Me.panCampos.TabIndex = 44
        '
        'lblConfirmado
        '
        Me.lblConfirmado.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblConfirmado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblConfirmado.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirmado.Location = New System.Drawing.Point(852, 264)
        Me.lblConfirmado.Name = "lblConfirmado"
        Me.lblConfirmado.Size = New System.Drawing.Size(141, 35)
        Me.lblConfirmado.TabIndex = 122
        Me.lblConfirmado.Text = "ATENDIDO"
        Me.lblConfirmado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblConfirmado.Visible = False
        '
        'lblConfirma
        '
        Me.lblConfirma.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblConfirma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblConfirma.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirma.Location = New System.Drawing.Point(852, 400)
        Me.lblConfirma.Name = "lblConfirma"
        Me.lblConfirma.Size = New System.Drawing.Size(141, 90)
        Me.lblConfirma.TabIndex = 121
        Me.lblConfirma.Text = "(F8)         CONFIRMAR LECTURA DE LA SOLICITUD"
        Me.lblConfirma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblConfirma.Visible = False
        '
        'txtDesOperario
        '
        Me.txtDesOperario.AutoSelec = False
        Me.txtDesOperario.BackColor = System.Drawing.Color.White
        Me.txtDesOperario.Blink = False
        Me.txtDesOperario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDesOperario.DesdeCodigo = CType(0, Long)
        Me.txtDesOperario.DesdeFecha = New Date(CType(0, Long))
        Me.txtDesOperario.HastaCodigo = CType(0, Long)
        Me.txtDesOperario.HastaFecha = New Date(CType(0, Long))
        Me.txtDesOperario.Location = New System.Drawing.Point(124, 60)
        Me.txtDesOperario.MaxLength = 30
        Me.txtDesOperario.Name = "txtDesOperario"
        Me.txtDesOperario.Size = New System.Drawing.Size(236, 20)
        Me.txtDesOperario.TabIndex = 119
        Me.txtDesOperario.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesOperario.ValorMax = 999999999.0R
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Info
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(888, 44)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 16)
        Me.Label9.TabIndex = 118
        Me.Label9.Text = "EXPOSI."
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Info
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(824, 44)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 16)
        Me.Label10.TabIndex = 117
        Me.Label10.Text = "EXISTENCIAS"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Info
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(888, 4)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(43, 16)
        Me.Label8.TabIndex = 116
        Me.Label8.Text = "EXPOSI."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Info
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(824, 4)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 16)
        Me.Label6.TabIndex = 115
        Me.Label6.Text = "EXISTENCIAS"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtExpo1
        '
        Me.txtExpo1.AutoSelec = False
        Me.txtExpo1.BackColor = System.Drawing.Color.White
        Me.txtExpo1.Blink = False
        Me.txtExpo1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExpo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExpo1.DesdeCodigo = CType(0, Long)
        Me.txtExpo1.DesdeFecha = New Date(CType(0, Long))
        Me.txtExpo1.HastaCodigo = CType(0, Long)
        Me.txtExpo1.HastaFecha = New Date(CType(0, Long))
        Me.txtExpo1.Location = New System.Drawing.Point(888, 20)
        Me.txtExpo1.MaxLength = 6
        Me.txtExpo1.Name = "txtExpo1"
        Me.txtExpo1.ReadOnly = True
        Me.txtExpo1.Size = New System.Drawing.Size(43, 20)
        Me.txtExpo1.TabIndex = 114
        Me.txtExpo1.TabStop = False
        Me.txtExpo1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExpo1.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExpo1.ValorMax = 999999999.0R
        '
        'txtExpo2
        '
        Me.txtExpo2.AutoSelec = False
        Me.txtExpo2.BackColor = System.Drawing.Color.White
        Me.txtExpo2.Blink = False
        Me.txtExpo2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExpo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExpo2.DesdeCodigo = CType(0, Long)
        Me.txtExpo2.DesdeFecha = New Date(CType(0, Long))
        Me.txtExpo2.HastaCodigo = CType(0, Long)
        Me.txtExpo2.HastaFecha = New Date(CType(0, Long))
        Me.txtExpo2.Location = New System.Drawing.Point(888, 60)
        Me.txtExpo2.MaxLength = 6
        Me.txtExpo2.Name = "txtExpo2"
        Me.txtExpo2.ReadOnly = True
        Me.txtExpo2.Size = New System.Drawing.Size(43, 20)
        Me.txtExpo2.TabIndex = 113
        Me.txtExpo2.TabStop = False
        Me.txtExpo2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExpo2.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExpo2.ValorMax = 999999999.0R
        '
        'cmdDesbloqueo
        '
        Me.cmdDesbloqueo.Image = CType(resources.GetObject("cmdDesbloqueo.Image"), System.Drawing.Image)
        Me.cmdDesbloqueo.Location = New System.Drawing.Point(936, 52)
        Me.cmdDesbloqueo.Name = "cmdDesbloqueo"
        Me.cmdDesbloqueo.Size = New System.Drawing.Size(28, 28)
        Me.cmdDesbloqueo.TabIndex = 112
        Me.cmdDesbloqueo.TabStop = False
        '
        'cmdBloqueo
        '
        Me.cmdBloqueo.Image = CType(resources.GetObject("cmdBloqueo.Image"), System.Drawing.Image)
        Me.cmdBloqueo.Location = New System.Drawing.Point(960, 52)
        Me.cmdBloqueo.Name = "cmdBloqueo"
        Me.cmdBloqueo.Size = New System.Drawing.Size(28, 28)
        Me.cmdBloqueo.TabIndex = 111
        Me.cmdBloqueo.TabStop = False
        Me.cmdBloqueo.Visible = False
        '
        'lblTraspaso
        '
        Me.lblTraspaso.BackColor = System.Drawing.SystemColors.Info
        Me.lblTraspaso.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTraspaso.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTraspaso.Location = New System.Drawing.Point(8, 216)
        Me.lblTraspaso.Name = "lblTraspaso"
        Me.lblTraspaso.Size = New System.Drawing.Size(987, 32)
        Me.lblTraspaso.TabIndex = 110
        Me.lblTraspaso.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAviso2
        '
        Me.lblAviso2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAviso2.ForeColor = System.Drawing.Color.Red
        Me.lblAviso2.Location = New System.Drawing.Point(936, 24)
        Me.lblAviso2.Name = "lblAviso2"
        Me.lblAviso2.Size = New System.Drawing.Size(30, 16)
        Me.lblAviso2.TabIndex = 109
        Me.lblAviso2.Text = "!!!!!"
        Me.lblAviso2.Visible = False
        '
        'lblAviso1
        '
        Me.lblAviso1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAviso1.ForeColor = System.Drawing.Color.Red
        Me.lblAviso1.Location = New System.Drawing.Point(936, 64)
        Me.lblAviso1.Name = "lblAviso1"
        Me.lblAviso1.Size = New System.Drawing.Size(30, 16)
        Me.lblAviso1.TabIndex = 108
        Me.lblAviso1.Text = "!!!!!"
        Me.lblAviso1.Visible = False
        '
        'txtExis1
        '
        Me.txtExis1.AutoSelec = False
        Me.txtExis1.BackColor = System.Drawing.Color.White
        Me.txtExis1.Blink = False
        Me.txtExis1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExis1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExis1.DesdeCodigo = CType(0, Long)
        Me.txtExis1.DesdeFecha = New Date(CType(0, Long))
        Me.txtExis1.HastaCodigo = CType(0, Long)
        Me.txtExis1.HastaFecha = New Date(CType(0, Long))
        Me.txtExis1.Location = New System.Drawing.Point(824, 20)
        Me.txtExis1.MaxLength = 6
        Me.txtExis1.Name = "txtExis1"
        Me.txtExis1.ReadOnly = True
        Me.txtExis1.Size = New System.Drawing.Size(65, 20)
        Me.txtExis1.TabIndex = 107
        Me.txtExis1.TabStop = False
        Me.txtExis1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExis1.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExis1.ValorMax = 999999999.0R
        '
        'txtExis2
        '
        Me.txtExis2.AutoSelec = False
        Me.txtExis2.BackColor = System.Drawing.Color.White
        Me.txtExis2.Blink = False
        Me.txtExis2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtExis2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExis2.DesdeCodigo = CType(0, Long)
        Me.txtExis2.DesdeFecha = New Date(CType(0, Long))
        Me.txtExis2.HastaCodigo = CType(0, Long)
        Me.txtExis2.HastaFecha = New Date(CType(0, Long))
        Me.txtExis2.Location = New System.Drawing.Point(824, 60)
        Me.txtExis2.MaxLength = 6
        Me.txtExis2.Name = "txtExis2"
        Me.txtExis2.ReadOnly = True
        Me.txtExis2.Size = New System.Drawing.Size(65, 20)
        Me.txtExis2.TabIndex = 106
        Me.txtExis2.TabStop = False
        Me.txtExis2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExis2.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExis2.ValorMax = 999999999.0R
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(371, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 16)
        Me.Label2.TabIndex = 103
        Me.Label2.Text = "HASTA EL ALMACEN:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grdLineas
        '
        Me.grdLineas.BackColor = System.Drawing.SystemColors.HighlightText
        Me.grdLineas.Columnas = 0
        Me.grdLineas.Editable = False
        Me.grdLineas.Location = New System.Drawing.Point(8, 264)
        Me.grdLineas.Name = "grdLineas"
        Me.grdLineas.Size = New System.Drawing.Size(830, 362)
        Me.grdLineas.TabIndex = 6
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
        Me.txtDesA1.Location = New System.Drawing.Point(545, 20)
        Me.txtDesA1.MaxLength = 30
        Me.txtDesA1.Name = "txtDesA1"
        Me.txtDesA1.ReadOnly = True
        Me.txtDesA1.Size = New System.Drawing.Size(277, 20)
        Me.txtDesA1.TabIndex = 101
        Me.txtDesA1.TabStop = False
        Me.txtDesA1.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA1.ValorMax = 999999999.0R
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
        Me.txtDesA2.Location = New System.Drawing.Point(545, 60)
        Me.txtDesA2.MaxLength = 30
        Me.txtDesA2.Name = "txtDesA2"
        Me.txtDesA2.ReadOnly = True
        Me.txtDesA2.Size = New System.Drawing.Size(277, 20)
        Me.txtDesA2.TabIndex = 100
        Me.txtDesA2.TabStop = False
        Me.txtDesA2.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA2.ValorMax = 999999999.0R
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
        Me.txtDesde.Location = New System.Drawing.Point(504, 20)
        Me.txtDesde.MaxLength = 2
        Me.txtDesde.Name = "txtDesde"
        Me.txtDesde.Size = New System.Drawing.Size(40, 20)
        Me.txtDesde.TabIndex = 3
        Me.txtDesde.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDesde.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtDesde.ValorMax = 999999999.0R
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
        Me.txtHasta.Location = New System.Drawing.Point(504, 60)
        Me.txtHasta.MaxLength = 2
        Me.txtHasta.Name = "txtHasta"
        Me.txtHasta.ReadOnly = True
        Me.txtHasta.Size = New System.Drawing.Size(40, 20)
        Me.txtHasta.TabIndex = 2
        Me.txtHasta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtHasta.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtHasta.ValorMax = 999999999.0R
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(371, 24)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(128, 16)
        Me.Label24.TabIndex = 99
        Me.Label24.Text = "MANDAR DESDE:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        Me.txtOperador.Location = New System.Drawing.Point(88, 60)
        Me.txtOperador.MaxLength = 3
        Me.txtOperador.Name = "txtOperador"
        Me.txtOperador.ReadOnly = True
        Me.txtOperador.Size = New System.Drawing.Size(34, 20)
        Me.txtOperador.TabIndex = 4
        Me.txtOperador.TabStop = False
        Me.txtOperador.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOperador.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtOperador.ValorMax = 999999999.0R
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(44, 16)
        Me.Label7.TabIndex = 44
        Me.Label7.Text = "OPER.:"
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
        Me.txtEstado.Location = New System.Drawing.Point(956, 108)
        Me.txtEstado.MaxLength = 1
        Me.txtEstado.Name = "txtEstado"
        Me.txtEstado.ReadOnly = True
        Me.txtEstado.Size = New System.Drawing.Size(28, 20)
        Me.txtEstado.TabIndex = 7
        Me.txtEstado.TabStop = False
        Me.txtEstado.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtEstado.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtEstado.ValorMax = 999999999.0R
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(896, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 16)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "ESTADO:"
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
        Me.txtComentario.Location = New System.Drawing.Point(156, 108)
        Me.txtComentario.MaxLength = 500
        Me.txtComentario.Multiline = True
        Me.txtComentario.Name = "txtComentario"
        Me.txtComentario.Size = New System.Drawing.Size(724, 99)
        Me.txtComentario.TabIndex = 5
        Me.txtComentario.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtComentario.ValorMax = 999999999.0R
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(40, 112)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(112, 16)
        Me.Label14.TabIndex = 40
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
        Me.txtFecha.Location = New System.Drawing.Point(288, 20)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(72, 20)
        Me.txtFecha.TabIndex = 1
        Me.txtFecha.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFecha.tipo = control.txtVisanfer.TiposCaja.fecha
        Me.txtFecha.ValorMax = 999999999.0R
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
        Me.txtCodigo.Location = New System.Drawing.Point(88, 20)
        Me.txtCodigo.MaxLength = 6
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(64, 20)
        Me.txtCodigo.TabIndex = 0
        Me.txtCodigo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCodigo.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtCodigo.ValorMax = 999999999.0R
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(232, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "FECHA:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "CODIGO:"
        '
        'tmrAviso
        '
        Me.tmrAviso.Interval = 500
        '
        'lblTitulo
        '
        Me.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTitulo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTitulo.Location = New System.Drawing.Point(663, 1)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(343, 56)
        Me.lblTitulo.TabIndex = 45
        Me.lblTitulo.Text = "VISANFER, S.A. - 2003"
        Me.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblInfo
        '
        Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInfo.Location = New System.Drawing.Point(2, 689)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(1004, 16)
        Me.lblInfo.TabIndex = 43
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(2, 705)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(1004, 24)
        Me.lblTeclas.TabIndex = 42
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmSolTraspasos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.ControlBox = False
        Me.Controls.Add(Me.panCampos)
        Me.Controls.Add(Me.lblTitulo)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.lblTeclas)
        Me.Controls.Add(Me.lblPrograma)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1024, 768)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1024, 768)
        Me.Name = "frmSolTraspasos"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SOLICITUD DE TRASPASO DE MATERIAL"
        Me.panCampos.ResumeLayout(False)
        Me.panCampos.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCargarSolicitud(ByVal loSolicitud As clsSolTraspaso)

        mnEmpresa = loSolicitud.mnEmpresa
        moSolTraspaso = loSolicitud
        mrAsignaEventos()
        mrPintaFormulario()
        If moSolTraspaso.mbEsNuevo Then
            mrNuevoRegistro()
        Else
            mrConsulta()
        End If
        mrMoverCampos(1)
        txtCodigo.Focus()

        Me.ShowDialog()

    End Sub

    Public Sub mrCargar(ByVal lnEmpresa As Integer)

        mnEmpresa = lnEmpresa

        goProfile = New clsProfileLocal
        goProfile.mrRecuperaDatos()

        mrAsignaEventos()
        mrPintaFormulario()
        mrConsulta()

        If Me.MdiParent Is Nothing Then
            Me.ShowDialog()
        Else
            Me.Show()
        End If

    End Sub

    Private Sub mrPintaFormulario()
        Dim loEmpresa As New prjEmpresas.clsEmpresa

        ' pongo los datos de la empresa ********
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()
        lblPrograma.Text = "SOLICITUD DE TRASPASOS"
        lblTitulo.Text = loEmpresa.msNombre

        mrPreparaGrid()

    End Sub

    Private Sub mrLeeTecla(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loCaja As control.txtVisanfer
        Dim loGrid As prjGrid.ctlGrid
        Dim lsControl As String = ""
        Dim lbSalta As Boolean = True

        If TypeOf sender Is control.txtVisanfer Then
            loCaja = sender
            lsControl = loCaja.Name
        End If
        If TypeOf sender Is prjGrid.ctlGrid Then
            loGrid = sender
            lsControl = loGrid.Name
        End If

        Select Case e.KeyValue
            Case Keys.L And e.Control = True     'CONTROL + L
                If mtEstado = EstadoVentana.Consulta And Not moSolTraspaso.mbEsNuevo Then
                    mtEstado = EstadoVentana.Lineas
                    grdLineas.mrPonFoco(0, 0)
                End If
            Case Keys.P And e.Control = True     'CONTROL + P
                mrImprimirCopia()
            Case Keys.M And e.Control = True     'CONTROL + M
                mrMantenimiento()
            Case Keys.Escape
                If Not mbGrabando Then
                    If lsControl = "grdLineas" Then
                        If mtEstado = EstadoVentana.Lineas Then
                            mtEstado = EstadoVentana.Consulta
                        End If
                        txtCodigo.Focus()
                    Else
                        mrConsulta()
                        mtEstado = EstadoVentana.Salida
                        mrLimpiaFormulario()
                        Me.Close()
                    End If
                End If
            Case Keys.Enter
                Select Case lsControl
                    Case "grdLineas"
                        If grdLineas.mnCol = 0 And mtEstado <> EstadoVentana.Consulta Then
                            If Not mfbCargaArticulo() Then
                                ' no hago nada
                            End If
                        End If
                        If grdLineas.mnCol = 2 And mtEstado <> EstadoVentana.Consulta Then
                            Dim lnCantidad As Double = mfnDouble(grdLineas.marMemoria(2, grdLineas.mnRow))
                            Dim lnArticulo As Long = mfnCodigoArticulo(grdLineas.marMemoria(0, grdLineas.mnRow))
                            Dim lnDetalle As Integer = mfnCodigoDetalle(grdLineas.marMemoria(0, grdLineas.mnRow))

                            mrMiraExistencias(lnArticulo, lnDetalle, lnCantidad)

                            If lnCantidad = 0 Then
                                grdLineas.mrPonFoco(2, grdLineas.mnRow)
                            Else
                                mrPendiente(grdLineas.mnRow, "N")
                                grdLineas.marMemoria(2, grdLineas.mnRow) = Format(lnCantidad, "#,##0.00")
                            End If
                        End If
                        If grdLineas.mnCol = 3 And mtEstado <> EstadoVentana.Consulta Then
                            If grdLineas.mnRow = grdLineas.mnFilasDatos - 1 Then
                                grdLineas.mrAñadirFila()
                            End If
                            grdLineas.mrRefrescaGrid()
                        End If
                    Case "txtCodigo"
                        mrCargaSolTraspaso()
                    Case "txtDesde"
                        mrCargaAlmacen(mfnLong(txtDesde.Text), 1, e)
                    Case "txtHasta"
                        mrCargaAlmacen(mfnLong(txtHasta.Text), 2, e)
                    Case "txtComentario"
                        lbSalta = False

                        Dim lsObservaciones As String
                        lsObservaciones = txtComentario.Text
                        If InStr(lsObservaciones, vbCrLf & vbCrLf) > 0 Then
                            If mtEstado <> EstadoVentana.Consulta Then lbSalta = True
                        End If
                End Select
                If (lsControl <> "grdLineas") And lbSalta Then Send("{TAB}")
            Case Keys.F1
                If lsControl = "grdLineas" Then
                    mrInsercion()
                Else
                    mrNuevoRegistro()
                End If
            Case Keys.F2
                If lsControl = "grdLineas" And mtEstado <> EstadoVentana.Consulta Then mrBorrado()
            Case Keys.F5
                If Not mbGrabando Then mrGrabar(e)
            Case Keys.F8
                If mtEstado = EstadoVentana.Consulta Then mrConfirmaSolicitud(True)
            Case Keys.F9
                Select Case lsControl
                    Case "txtCodigo"
                        If mtEstado = EstadoVentana.Consulta Then mrBuscaSolTraspaso()
                    Case "grdLineas"
                        If (mtEstado <> EstadoVentana.Consulta) And (grdLineas.mnCol = 0) Then mrBuscaArticulos()
                End Select
        End Select

    End Sub

    Private Sub mrConfirmaSolicitud(ByVal lbAviso As Boolean)

        If moSolTraspaso.msEstado = "P" Then
            If (goProfile.mnAlmacen <> moSolTraspaso.mnDesde) And lbAviso Then
                MsgBox("LOS TRASPASOS SOLO SE PUEDEN CONFIRMAR DESDE EL ALMACEN DE SALIDA.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            Else
                moSolTraspaso.msEstado = "A"
                moSolTraspaso.mrGrabaDatos()

                mrMoverCampos(1)
            End If
        End If

    End Sub

    Private Sub mrMiraExistencias(ByVal lnArticulo As Long, ByVal lnDetalle As Integer, ByVal lnCantidad As Double)
        Dim lnAlmacenSalida As Integer = mfnInteger(txtDesde.Text)
        Dim loExistencias As clsExistencias

        Dim loAgrArt As clsAgrArt
        loAgrArt = New clsAgrArt
        loAgrArt.mnEmpresa = mnEmpresa
        loAgrArt.mnCodigo = lnArticulo
        loAgrArt.mnDetalle = lnDetalle
        loAgrArt.mnLinea = 1
        loAgrArt.mrRecuperaDatos()
        If Not loAgrArt.mbEsNuevo Then
            loAgrArt.mnLinea = 0
            Dim loBusAgrArt As clsBusAgrArt
            loBusAgrArt = New clsBusAgrArt
            loBusAgrArt.mrBusca(loAgrArt)
            For Each loAgrArt In loBusAgrArt.mcolAgrart
                ' miramos si el articulo tiene existencias fijas ***********
                Dim loArticulo As New clsArticulo
                loArticulo.mnEmpresa = mnEmpresa
                loArticulo.mnCodigo = loAgrArt.mnArticulo
                loArticulo.mnDetalle = loAgrArt.mnDetallePadre
                loArticulo.mrRecuperaDatos()
                If loArticulo.msControlExistencias = "N" Then
                    loExistencias = New clsExistencias
                    loExistencias.mnEmpresa = mnEmpresa
                    loExistencias.mnAlmacen = lnAlmacenSalida 'goProfile.mnAlmacen
                    loExistencias.mnArticulo = loAgrArt.mnArticulo
                    loExistencias.mnDetalle = loAgrArt.mnDetallePadre
                    loExistencias.mrRecuperaDatos()
                    If loExistencias.mnExistencias < (lnCantidad * loAgrArt.mnExistencias) Then
                        Dim loAviso As New frmAvisoControl
                        loAviso.msAviso = "¡¡¡ ATENCION !!!"
                        loAviso.msMensaje = "ATENCION: EN ESTE MOMENTO NO HAY EXISTENCIAS DE ESTE " & vbCrLf & _
                                            "ARTICULO EN EL ALMACEN AL QUE ESTAS PIDIENDO." & vbCrLf
                        loAviso.msMensaje = loAviso.msMensaje & " (EXISTENCIAS: " & Format(loExistencias.mnExistencias, "#,##0.00") & ")"
                        loAviso.msTeclas = "PULSE (F5) PARA SALIR"
                        loAviso.mbParpadeo = True
                        loAviso.moTecla = Keys.F5
                        loAviso.mrAvisar()
                        Exit For
                    End If
                End If
            Next
        Else

            ' miramos si el articulo tiene existencias fijas ***********
            Dim loArticulo As New clsArticulo
            loArticulo.mnEmpresa = mnEmpresa
            loArticulo.mnCodigo = lnArticulo
            loArticulo.mnDetalle = lnDetalle
            loArticulo.mrRecuperaDatos()
            If loArticulo.msControlExistencias = "S" Then Exit Sub

            'Si no es un articulo Virtual leemos existencias normalmente
            loExistencias = New clsExistencias
            loExistencias.mnEmpresa = mnEmpresa
            loExistencias.mnAlmacen = lnAlmacenSalida 'goProfile.mnAlmacen
            loExistencias.mnArticulo = lnArticulo
            loExistencias.mnDetalle = lnDetalle
            loExistencias.mrRecuperaDatos()

            If loExistencias.mnExistencias < lnCantidad Then
                Dim loAviso As New frmAvisoControl
                loAviso.msAviso = "¡¡¡ ATENCION !!!"
                loAviso.msMensaje = "ATENCION: EN ESTE MOMENTO NO HAY EXISTENCIAS DE ESTE " & vbCrLf & _
                                    "ARTICULO EN EL ALMACEN AL QUE ESTAS PIDIENDO." & vbCrLf
                loAviso.msMensaje = loAviso.msMensaje & " (EXISTENCIAS: " & Format(loExistencias.mnExistencias, "#,##0.00") & ")"
                loAviso.msTeclas = "PULSE (F5) PARA SALIR"
                loAviso.mbParpadeo = True
                loAviso.moTecla = Keys.F5
                loAviso.mrAvisar()
            Else
                If loExistencias.mnExposicion > 0 And loExistencias.mnExistencias <= loExistencias.mnExposicion Then
                    Dim lmRespuesta As MsgBoxResult

                    lmRespuesta = MsgBox("¿ESTE ARTICULO ES DE EXPOSICION?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Gestion Visanfer")
                    If lmRespuesta = MsgBoxResult.Yes Then
                        'En principio no controlamos las existencias en exposicion negativas
                        'porque en las modificaciones siempre nos descuadrarian.
                        'Las existencias de exposicion de momento solo son orientativas.
                        'Dicho Domingo 26/01/2009
                    Else
                        Dim loAviso As New frmAvisoControl
                        loAviso.msAviso = "¡¡¡ ATENCION !!!"
                        loAviso.msMensaje = "ATENCION: CASI SEGURO QUE SE ESTA EQUIVOCANDO DE ALMACEN. " & vbCrLf & _
                                            "REVISE EL ALMACEN DE SALIDA DE ESTE ARTICULO. EN CASO   " & vbCrLf & _
                                            "CONTRARIO REVISE LAS EXISTENCIAS DE ESTE ARTICULO EN EL" & vbCrLf & _
                                            "ALMACEN SUMANDO ADEMAS LAS QUE TIENE EN EXPOSICION. " & _
                                            "(EXISTENCIAS TOTALES:" & Format(loExistencias.mnExistencias, "#,##0.00") & " EN EXPOSICION:" & Format(loExistencias.mnExistencias, "#,##0.00") & ")"
                        loAviso.msTeclas = "PULSE (F5) PARA SALIR"
                        loAviso.mbParpadeo = True
                        loAviso.moTecla = Keys.F5
                        loAviso.mrAvisar()
                    End If
                End If
            End If

        End If

    End Sub

    Private Sub mrBuscaSolTraspaso()
        moSolTraspaso = New clsSolTraspaso
        moSolTraspaso.mnEmpresa = mnEmpresa
        moSolTraspaso.mrBuscaSolTraspaso()
    End Sub

    Private Sub moSolTraspaso_evtBusTraspaso() Handles moSolTraspaso.evtBusSolTraspaso
        lblTeclas.Text = " CTRL-M Modificacion de Datos      CTRL-L Ver Lineas    F1-Alta Nuevo        " & _
                         " CTRL-P Imprimir        ESC-Salida"
        mrMoverCampos(1)
    End Sub

    Private Sub mrNuevoRegistro()

        ' Relleno de los comandos de las teclas *************
        'If goUsuario.mfbAccesoPermitido(1, True) Then
        mtEstado = EstadoVentana.NuevoRegistro
        gnLlave = 0
        mrLimpiaFormulario()
        ' añado por lo menos una linea para que tenga el grid ******
        grdLineas.mrAñadirFila()
        grdLineas.mrRefrescaGrid()
        ' **********************************************************
        moAlmacen1 = New clsAlmacen
        moAlmacen2 = New clsAlmacen
        moSolTraspaso = New clsSolTraspaso
        txtCodigo.Text = "0"
        txtFecha.Text = Format(Now, "dd/MM/yyyy")
        txtOperador.Text = goUsuario.mnCodigo
        mrCargaOperario()
        txtEstado.Text = "P"
        lblConfirmado.Visible = False
        txtHasta.Text = goProfile.mnAlmacen
        mrCargaAlmacen(mfnLong(txtHasta.Text), 2, Nothing)
        lblTeclas.BackColor = Color.GreenYellow
        lblTeclas.Text = "F5-GRABACION                                ESC-SALIDA"
        lblPrograma.Text = "SOL. TRASPASOS - NUEVO REGISTRO"
        lblPrograma.BackColor = Color.GreenYellow
        panCampos.Enabled = True
        grdLineas.Editable = True
        txtDesde.Focus()
        'End If

    End Sub

    Private Sub mrCargaOperario()

        Dim loOperario As New clsUsuario
        loOperario.mnCodigo = mfnInteger(txtOperador.Text)
        loOperario.mrRecuperaDatos()
        txtDesOperario.Text = loOperario.msNombre

    End Sub

    Private Sub mrCargaSolTraspaso()
        Dim lnCodigo As Long

        lnCodigo = mfnInteger(txtCodigo.Text)
        If lnCodigo > 0 Then
            If mtEstado <> EstadoVentana.Mantenimiento Then
                moSolTraspaso = New clsSolTraspaso
                moSolTraspaso.mnEmpresa = mnEmpresa
                moSolTraspaso.mnCodigo = lnCodigo
                moSolTraspaso.mrRecuperaDatos()
                mrLimpiaFormulario()
                If moSolTraspaso.mbEsNuevo Then
                    MsgBox("Traspaso no Encontrado", MsgBoxStyle.Exclamation, "Gestion Visanfer")
                    txtCodigo.Focus()
                Else
                    moSolTraspaso.mrRecuperaLineas()
                    mrMoverCampos(1)
                    lblTeclas.Text = " CTRL-M Modificacion de Datos      CTRL-L Ver Lineas    F1-Alta Nuevo        " & _
                                     " CTRL-P Imprimir        ESC-Salida"
                End If
            End If
            txtCodigo.SelectAll()
        End If

    End Sub

    Private Sub mrLimpiaFormulario()
        Dim loControl As Windows.Forms.Control
        Dim loCaja As control.txtVisanfer

        For Each loControl In panCampos.Controls
            If TypeOf loControl Is control.txtVisanfer Then
                loCaja = loControl
                loCaja.Text = ""
            End If
        Next
        grdLineas.mrClear(prjGrid.ctlGrid.TipoBorrado.Contenido)
        lblAviso1.Visible = False
        lblAviso2.Visible = False
        txtExis1.BackColor = Color.White
        txtExis2.BackColor = Color.White
        lblTraspaso.Text = ""
        lblConfirma.Visible = False

    End Sub

    Private Sub mrCargaAlmacen(ByVal lnCodigo As Integer, ByVal lnNumero As Integer, _
                               ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loAlmacen As clsAlmacen

        If moAlmacen1 Is Nothing Then moAlmacen1 = New clsAlmacen
        If moAlmacen2 Is Nothing Then moAlmacen2 = New clsAlmacen

        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = lnCodigo
        loAlmacen.mrRecuperaDatos(False)
        If lnNumero = 1 Then
            moAlmacen1 = loAlmacen
            txtDesde.Text = loAlmacen.mnCodigo
            txtDesA1.Text = loAlmacen.msNombre
        Else
            moAlmacen2 = loAlmacen
            txtHasta.Text = loAlmacen.mnCodigo
            txtDesA2.Text = loAlmacen.msNombre
        End If
        If loAlmacen.mbEsNuevo Then
            e.Handled = True
            MsgBox("Registro no Encontrado", MsgBoxStyle.Exclamation, "Gestion Visanfer")
            If lnNumero = 1 Then
                txtFecha.Focus()
            Else
                txtDesde.Focus()
            End If
        End If
        'If lnNumero = 1 Then
        '    txtDesA1.SelectAll()
        'Else
        '    txtDesA2.SelectAll()
        'End If

        If moAlmacen1.mnCodigo = moAlmacen2.mnCodigo Then
            e.Handled = True
            MsgBox("Almacenes iguales, cambie alguno.", MsgBoxStyle.Exclamation, "Gestion Visanfer")
        End If
        lblTraspaso.Text = txtDesA1.Text & "   ····>   " & txtDesA2.Text

    End Sub

    Private Sub moArticulo_evtBusArticulo() Handles moArticulo.evtBusArticulo
        Dim lnFila As Integer
        Dim lnInicio As Integer
        Dim loArticulo As clsArticulo
        Dim lnContador As Integer = 0

        lnFila = grdLineas.mnRow
        lnInicio = lnFila
        For Each loArticulo In moArticulo.moBusMultiple.mcolArticulos
            If loArticulo.mbSeleccionado Then
                If lnContador > 0 Then grdLineas.mrAñadirFila()
                lnContador = lnContador + 1
                grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo
                If loArticulo.mnDetalle > 0 Then grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo & "." & loArticulo.mnDetalle
                grdLineas.marMemoria(1, lnFila) = loArticulo.msDescripcion
                grdLineas.marMemoria(2, lnFila) = "0"
                mrPendiente(lnFila, "S")
                'grdLineas.mrAñadirFila()
                lnFila = lnFila + 1
            End If
        Next
        If lnContador = 0 Then
            grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo
            If moArticulo.mnDetalle > 0 Then grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo & "." & moArticulo.mnDetalle
            grdLineas.marMemoria(1, lnFila) = moArticulo.msDescripcion
            grdLineas.marMemoria(2, lnFila) = "0"
            mfbCargaArticulo()
        End If
        grdLineas.mrRefrescaGrid()
        grdLineas.mrPonFoco(0, lnInicio + 1)

    End Sub

    Private Sub mrBuscaArticulos()
        ' le digo el fichero del que quiero que lea ***********
        moArticulo = New clsArticulo
        moArticulo.mnEmpresa = mnEmpresa
        moArticulo.mrBuscaArticulo(True, goUsuario)
    End Sub

    Private Sub mrGrabar(ByVal e As System.Windows.Forms.KeyEventArgs)
        ' graba toda la carga en la base de datos *****************

        mbGrabando = True
        If mtEstado = EstadoVentana.NuevoRegistro Or _
           mtEstado = EstadoVentana.Mantenimiento Then
            ' compruebo que los campos obligatorios estan cumplimentados
            If mfbObligatorios(e) Then
                Dim loTraspasoLin As clsSolTraspasoLin

                mrMoverCampos(2)    ' paso los valores al objeto
                If mcolLineas.Count > 0 Then
                    Cursor = Cursors.WaitCursor
                    Dim lbError As Boolean
                    If Not mtEstado = EstadoVentana.NuevoRegistro Then
                        moSolTraspaso.mrGrabaDatos() ' grabo el contenido
                        ' ahora borro las lineas y las grabo otra vez
                        moSolTraspaso.mrBorraLineas()
                        ' ahora grabo las lineas ******************
                        For Each loTraspasoLin In mcolLineas
                            loTraspasoLin.mnCodigo = moSolTraspaso.mnCodigo
                            loTraspasoLin.mrGrabaDatos()
                        Next
                        moSolTraspaso.mrRecuperaLineas()
                    End If
                    If Not lbError Then ' si la grabacion es correcta
                        If mtEstado = EstadoVentana.NuevoRegistro Then
                            ' si es un registro nuevo recupero su codigo
                            moSolTraspaso.mnEmpresa = mnEmpresa
                            moSolTraspaso.mrNuevoCodigo()
                            moSolTraspaso.mbEsNuevo = True
                            ' ahora grabo las lineas ******************
                            For Each loTraspasoLin In mcolLineas
                                loTraspasoLin.mnCodigo = moSolTraspaso.mnCodigo
                                loTraspasoLin.mrGrabaDatos()
                            Next
                            txtCodigo.Text = moSolTraspaso.mnCodigo
                            moSolTraspaso.mrGrabaDatos() ' grabo el contenido
                            moSolTraspaso.mrRecuperaLineas()
                            ' le pongo la captura de eventos porque algunas veces lo hace mal hasta aqui ******
                            System.Windows.Forms.Application.DoEvents()
                            ' ahora lo que hago es lanzar la impresion del pedido ********
                            'mrImprimirRemoto()
                        End If
                        ' despues de grabar resituo el foco en le inicio
                        mrMoverCampos(1)
                        mrConsulta()
                    End If
                    Cursor = Cursors.Default
                    txtCodigo.Focus()
                Else
                    MsgBox("Debe tener por lo menos alguna linea correcta.", MsgBoxStyle.Critical, "Visanfer .Net")
                End If
            End If
        End If
        mbGrabando = False

    End Sub

    Private Sub mrImprimirCopia()
        ' imrprimo los listados en ambas impresoras ******************
        mrSelImpresora()

    End Sub

    Private Sub mrImprimirRemoto()
        ' imrprimo los listados en ambas impresoras ******************
        mrSelImpresora()

    End Sub

    Private Sub mrSelImpresora()
        ' selecciona la impresora ************************************

        moSelImpresora = New prjControl.frmSelImpresora

        moSelImpresora.mnImpresora = goProfile.mnImpresora
        If goProfile.mnImpresora = 80 Then
            moSelImpresora.msPapel = "A5"
        Else
            moSelImpresora.msPapel = "A4"
        End If
        moSelImpresora.msDestino = "I"
        moSelImpresora.mnCopias = 1
        moSelImpresora.mrSeleccionar("IMP. SOL. TRASPASOS")

        mrConfirmaSolicitud(False)

    End Sub

    Private Sub mrImprimirRpt(ByVal lnImpresora As Integer, ByVal lnCopias As Integer, _
                              ByVal lsSalida As String, ByVal lsPapel As String)

        Dim loLinea As clsSolTraspasoLin
        Dim loTablaCabecera As DataTable
        Dim loTablaLineas As DataTable

        ' confirmo el estado de la solicitud de traspaso
        If (moSolTraspaso.msEstado = "P") And (moSolTraspaso.mnDesde = goProfile.mnAlmacen) Then
            moSolTraspaso.msEstado = "A"
            moSolTraspaso.mrGrabaDatos()

            mrMoverCampos(1)
        End If

        Cursor = Cursors.WaitCursor
        panCampos.Enabled = False
        ' RELLENO LOS DATASET CON LOS DATOS DE LAS CAJAS

        moDataTraspaso = New dtsTraspaso
        ' primero vacio las tablas ***************
        loTablaCabecera = moDataTraspaso.Tables("Cabecera")
        loTablaCabecera.Rows.Clear()
        loTablaLineas = moDataTraspaso.Tables("Lineas")
        loTablaLineas.Rows.Clear()
        ' relleno las tablas del dataset con los valores del listado ********
        Dim loRow As DataRow

        ' recupero los almacenes ******************

        ' ahora paso los valores a las tablas
        loRow = loTablaCabecera.NewRow
        loRow("Codigo") = moSolTraspaso.mnCodigo
        loRow("Fecha") = moSolTraspaso.mdFecha
        loRow("Tipo") = "SOLICITUD"
        loRow("Titulo") = "SOLICITUD TRASPASO"
        ' **************************************
        Dim loAlmacen As clsAlmacen
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moSolTraspaso.mnDesde
        loAlmacen.mrRecuperaDatos(False)
        loRow("Desde") = loAlmacen.msNombre
        ' **************************************
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moSolTraspaso.mnHasta
        loAlmacen.mrRecuperaDatos(False)
        loRow("Hasta") = loAlmacen.msNombre
        ' **************************************
        loRow("Vendedor") = moSolTraspaso.mnVendedor
        loRow("Operario") = moSolTraspaso.mnOperario
        loRow("Observaciones") = moSolTraspaso.msObservaciones
        loTablaCabecera.Rows.Add(loRow)

        Dim lnI As Integer = 1
        For Each loLinea In moSolTraspaso.mcolLineas
            ' añado un registro a la tabla de lineas
            loRow = loTablaLineas.NewRow()
            loRow("Codigo") = loLinea.mnCodigo
            loRow("Linea") = lnI
            If loLinea.mnDetalle > 0 Then
                loRow("CodArticulo") = loLinea.mnArticulo & "." & loLinea.mnDetalle
            Else
                loRow("CodArticulo") = loLinea.mnArticulo
            End If
            loRow("Articulo") = loLinea.mnArticulo
            loRow("Detalle") = loLinea.mnDetalle
            loRow("Cantidad") = loLinea.mnCantidad
            loRow("Descripcion") = loLinea.msDescripcion
            ' leo la referencia del articulo en el proveedor actual
            loRow("Referencia") = mfsReferencia(loLinea)
            ' leo las existencias en el almacen destino
            loRow("Existencias") = mfnExistencias(loLinea)
            loTablaLineas.Rows.Add(loRow)
            lnI = lnI + 1
        Next

        'Dim loListado As New rptSolTraspaso
        'loListado.SetDataSource(moDataTraspaso)

        ' impresion especial para la fabrica ********
        Dim lbApaisado As Boolean = True
        If moImpresora.mnCodigo = 80 Then
            lsPapel = "A5"
            moImpresora.msOrientacion = "H"
            moImpresora.msPapel = "A5"
        Else
            lsPapel = "A4"
            moImpresora.msPapel = "A4"
        End If

        'If moSelImpresora.msDestino = "I" Then
        '    prjPrinterNet.mrImprimeReport(goUsuario, CType(loListado, CrystalDecisions.CrystalReports.Engine.ReportDocument), moImpresora, moSelImpresora.mnCopias)
        'Else
        '    prjPrinterNet.mrVisualizaReport(goUsuario, CType(loListado, CrystalDecisions.CrystalReports.Engine.ReportDocument), moImpresora, moSelImpresora.mnCopias)
        'End If


        Dim loVisor As New prjPrinterNet.frmVisorReport
        loVisor.moReport.LocalReport.ReportEmbeddedResource = "prjTraspasos.rdlcSolTraspaso.rdlc"
        loVisor.moReport.LocalReport.DataSources.Clear()
        loVisor.moReport.LocalReport.DataSources.Add(New ReportDataSource("Cabecera", loTablaCabecera))
        loVisor.moReport.LocalReport.DataSources.Add(New ReportDataSource("Lineas", loTablaLineas))
        loVisor.moReport.LocalReport.EnableExternalImages = True
        If lsSalida.Equals("I") Then
            loVisor.mrImprimir(lsPapel, lbApaisado, moImpresora.msCola)
        Else
            loVisor.mrVisualizar(lsPapel, lbApaisado, moImpresora.msCola)
        End If

        panCampos.Enabled = True
        Cursor = Cursors.Default

    End Sub

    Private Function mfnExistencias(ByVal loLinea As clsSolTraspasoLin) As Double

        Dim loExistencias As New clsExistencias
        loExistencias.mnEmpresa = mnEmpresa
        loExistencias.mnAlmacen = moAlmacen2.mnCodigo
        loExistencias.mnArticulo = loLinea.mnArticulo
        loExistencias.mnDetalle = loLinea.mnDetalle
        loExistencias.mrRecuperaDatos()

        Return loExistencias.mnExistencias

    End Function

    Private Function mfsReferencia(ByVal loLinea As clsSolTraspasoLin) As String

        Dim loArticulo As New clsArticulo
        loArticulo.mnEmpresa = mnEmpresa
        loArticulo.mnCodigo = loLinea.mnArticulo
        loArticulo.mnDetalle = loLinea.mnDetalle
        loArticulo.mrRecuperaDatos()
        Dim loTarifa As New prjLibroProv.clsTarifa
        loTarifa.mnEmpresa = mnEmpresa
        loTarifa.mnProveedor = loArticulo.mnProveedor
        loTarifa.mnArticulo = loArticulo.mnCodigo
        loTarifa.mrRecuperaDatos()

        Return loTarifa.msReferencia

    End Function

    Private Sub moSelImpresora_evtBusImpresora() Handles moSelImpresora.evtBusImpresora
        moImpresora = New prjPrinterNet.clsImpresora
        moImpresora.mnEmpresa = mnEmpresa
        moImpresora.mnCodigo = moSelImpresora.mnImpresora
        moImpresora.mrRecuperaDatos()
        If moImpresora.mbEsNuevo Then
            MsgBox("Impresora no instalada", MsgBoxStyle.Critical, "Visanfer .Net")
            Exit Sub
        End If
        If moSelImpresora.msPapel <> "" Then moImpresora.msPapel = moSelImpresora.msPapel
        If moImpresora.msTipo = "L" Then
            If moSelImpresora.msPapel = "A5" Then moImpresora.msOrientacion = "V"
            mrImprimirRpt(moImpresora.mnCodigo, moSelImpresora.mnCopias, moSelImpresora.msDestino, moSelImpresora.msPapel)
        Else
            mrImprimirSolTraspaso()
        End If
    End Sub

    Private Sub mrImprimirSolTraspaso()
        Dim loLinea As clsSolTraspasoLin
        Dim lnTopLineas As Integer
        Dim lnPagina As Integer
        Dim loMensaje As New frmMensajes
        Dim loUnidad As prjArticulos.clsUnimed
        Dim loArticulo As prjArticulos.clsArticulo
        Dim loBusUnimed As prjArticulos.clsBusUnimed
        Dim lnLineas As Integer

        ' ******* esta dll necesita que se le pase la conexion ************
        loMensaje.mrAbrir("Imprimiendo, por favor espere ...")

        ' ******** recupero la cola de la impresora seleccionada **********
        moImpresora.mnEmpresa = mnEmpresa
        moImpresora.mnCodigo = moSelImpresora.mnImpresora
        moImpresora.mrRecuperaDatos()
        ' *********** inicio del proceso de impresion *************************
        moPrinter = New clsPrinter
        moPrinter.msTipoImpresora = moImpresora.msTipo
        moPrinter.mrInicio(goProfile.mfsLogin)
        moPrinter.msCola = moImpresora.msCola
        moPrinter.mpbComprimida = True
        moPrinter.mpbProporcional = True

        If moImpresora.mbEsNuevo Then
            MsgBox("Impresora no instalada", MsgBoxStyle.Critical, "Visanfer .Net")
        Else
            ' ********* parametros de impresion ********************
            lnPagina = 1
            If moImpresora.msTipo = "L" Then
                lnTopLineas = 70                ' numero de lineas por pagina
            Else
                lnTopLineas = 48                ' numero de lineas por pagina
            End If
            'lnTopLineas = 66                ' numero de lineas por pagina

            moPrinter.mpnLineas = lnTopLineas    ' albaranes 48, A4 66
            moPrinter.mnCopias = moSelImpresora.mnCopias     ' Copias a imprimir

            ' ************  cabecera ************************************
            mrCabecera(lnPagina)
            ' ************  lineas del listado **************************
            loBusUnimed = New prjArticulos.clsBusUnimed
            loBusUnimed.mnEmpresa = moSolTraspaso.mnEmpresa
            lnLineas = 1
            For Each loLinea In moSolTraspaso.mcolLineas
                ' recupero la unidad de medida del articulo ***************
                loArticulo = New prjArticulos.clsArticulo
                loArticulo.mnEmpresa = moSolTraspaso.mnEmpresa
                loArticulo.mnCodigo = loLinea.mnArticulo
                loArticulo.mrRecuperaDatos()
                loUnidad = loBusUnimed.mfoUnimed(loArticulo.mnTipoUnidad)
                ' ******** impresion **************************************
                moPrinter.mrPrint(0, loLinea.mnArticulo, 0, 7, modGeneral.Alineacion.Derecha)
                moPrinter.mrPrint(-1, Format(loLinea.mnCantidad, "#,##0.00"), 12, 10, modGeneral.Alineacion.Derecha)
                moPrinter.mrPrint(-1, loUnidad.msAbreviatura, 25, 3)
                moPrinter.mrPrint(-1, loLinea.msDescripcion, 30, 45)
                ' ************ Control de paginas *************************
                'If moPrinter.mpnLineaActual > lnTopLineas - 3 Then
                If lnLineas = 22 Then
                    lnLineas = 0
                    moPrinter.mrFormFeed()
                    lnPagina = lnPagina + 1
                    mrCabecera(lnPagina)
                End If
                lnLineas = lnLineas + 1
            Next

            ' ************ final del listado *******************************
            moPrinter.mrFinal()

            If moSelImpresora.msDestino = "I" Then
                moPrinter.mrImprimir()
            Else
                moPrinter.mrVisualizar()
            End If
            'Me.Close()

        End If

        loMensaje.mrCerrar()

    End Sub

    Private Sub mrCabecera(ByVal lnPagina As Integer)

        moPrinter.mrPrint(0, "")
        moPrinter.mrPrint(0, "TRASPASO", 58, 10, modGeneral.Alineacion.Centro)

        moPrinter.mrPrint(-1, "TRASPASO", 83, 15)
        moPrinter.mrPrint(-1, moSolTraspaso.mnOperario, 128, 3)
        moPrinter.mrPrint(-1, moSolTraspaso.mnVendedor, 133, 3)

        moPrinter.mrPrint(0, moSolTraspaso.mnCodigo, 58, 10, modGeneral.Alineacion.Centro)
        moPrinter.mrPrint(0, "DESDE ALMACEN: " & txtDesA1.Text, 75, 45)
        moPrinter.mpnTamaño = modGeneral.Tamaño.cpi_20
        moPrinter.mrPrint(0, "HASTA ALMACEN: " & txtDesA2.Text, 75, 45)
        moPrinter.mrPrint(0, "")
        moPrinter.mrPrint(0, Format(moSolTraspaso.mdFecha, "dd/MM/yyyy"), 58, 10)
        moPrinter.mrPrint(-1, "PEDIDO POR: " & moSolTraspaso.msObservaciones, 75, 42)
        moPrinter.mrPrint(0, "")
        moPrinter.mrPrint(0, "")

    End Sub

    Private Sub mrCargaLineas()
        Dim lnI As Integer
        Dim lnJ As Integer
        Dim lsTipo As String
        Dim loTraspasoLin As clsSolTraspasoLin

        mcolLineas = New Collection
        lnJ = 1
        For lnI = 0 To grdLineas.mnFilasDatos - 1
            ' primero veo que tipo de linea es ********************
            lsTipo = grdLineas.marMemoria(4, lnI)
            If lsTipo <> "B" Then
                loTraspasoLin = New clsSolTraspasoLin
                loTraspasoLin.mnEmpresa = mnEmpresa
                'loTraspasoLin.mnCodigo = moSolTraspaso.mnCodigo
                loTraspasoLin.mnLinea = lnJ
                loTraspasoLin.mnArticulo = mfnCodigoArticulo(grdLineas.marMemoria(0, lnI))
                loTraspasoLin.mnDetalle = mfnCodigoDetalle(grdLineas.marMemoria(0, lnI))
                loTraspasoLin.msDescripcion = Trim(grdLineas.marMemoria(1, lnI))
                loTraspasoLin.mnCantidad = mfnDouble(grdLineas.marMemoria(2, lnI))
                loTraspasoLin.msEstado = Trim(grdLineas.marMemoria(3, lnI))
                If loTraspasoLin.mnArticulo <> 0 Then
                    mcolLineas.Add(loTraspasoLin)
                    lnJ = lnJ + 1
                End If
            End If
        Next
    End Sub

    Private Function mfbObligatorios(ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        ' Compruebo que los campos obligatorios estan terminados ****************
        ' los campos los controlo por el orden que he estimado oportun0 *********

        If Trim(txtFecha.Text) = "" Then
            MsgBox("DEBE PONER ALGUNA FECHA.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtFecha.SelectAll()
            txtFecha.Focus()
            e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
            Return False
        End If

        If mfnInteger(txtDesde.Text) = 0 Then
            MsgBox("DEBE PONER ALGUN ALMACEN DE ORIGEN.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtDesde.SelectAll()
            txtDesde.Focus()
            e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
            Return False
        End If

        If mfnInteger(txtHasta.Text) = 0 Then
            MsgBox("DEBE PONER ALGUN ALMACEN DE DESTINO.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtHasta.SelectAll()
            txtHasta.Focus()
            e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
            Return False
        End If

        If Trim(txtComentario.Text) = "" Then
            MsgBox("DEBE PONER QUIEN HA PEDIDO EL MATERIAL.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtComentario.SelectAll()
            txtComentario.Focus()
            e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
            Return False
        End If

        Return True

    End Function

    Private Sub mrInsercion()
        ' inserto una nueva linea
        Dim lnFila As Integer

        lnFila = grdLineas.mnRow

        grdLineas.mrInsertarFila(lnFila)
        grdLineas.marMemoria(4, lnFila) = "I"

    End Sub

    Private Sub mrBorrado()
        Dim loColor As Color
        Dim lnLinea As Integer
        Dim lnCol As Integer
        Dim lnI As Integer
        Dim lsTipo As String

        ' borro la linea actual **********************
        If mtEstado <> EstadoVentana.Consulta Then
            lnLinea = grdLineas.mnRow
            lnCol = grdLineas.mnCol
            ' ******* coloreo la linea de color rojo *
            lsTipo = grdLineas.marMemoria(4, lnLinea)
            Select Case lsTipo
                Case "B"        ' si la linea esta borrada se pone normal
                    grdLineas.marMemoria(4, lnLinea) = ""
                    loColor = Color.FromName("Window")
                Case "I"        ' si la linea es nueva se borra *********
                    grdLineas.mrBorrarFila(grdLineas.mnRow)
                Case Else       ' si la linea en normal se marca ********
                    grdLineas.marMemoria(4, lnLinea) = "B"
                    loColor = Color.DarkOrange
            End Select
            If lsTipo <> "I" Then
                For lnI = 0 To 5
                    grdLineas.ColorCelda(lnI, grdLineas.mnRow).mnBackColor = loColor
                Next
            End If
            grdLineas.mrRefrescaGrid()
        End If

    End Sub

    Private Sub mrConsulta()

        ' Relleno de los comandos de las teclas *************
        tmrAviso.Enabled = False
        mtEstado = EstadoVentana.Consulta
        grdLineas.Editable = False
        lblTeclas.BackColor = Color.FromKnownColor(KnownColor.ActiveBorder)
        lblTeclas.ForeColor = Color.Black
        lblTeclas.Text = " CTRL-M Modificacion de Datos      CTRL-L Ver Lineas    F1-Alta Nuevo        " & _
                         " CTRL-P Imprimir        ESC-Salida"
        lblPrograma.Text = "SOLICITUD DE TRASPASOS"
        lblPrograma.BackColor = Color.FromKnownColor(KnownColor.ActiveBorder)

    End Sub

    Private Sub mrMantenimiento()

        'mrAsignaconexion(gconADO)
        'If goUsuario.mfbAccesoPermitido(9, True) Then
        If mtEstado = EstadoVentana.Mantenimiento Then
            mrConsulta()
        Else
            If moSolTraspaso.msEstado = "A" Then
                MsgBox("Esta Solicitud ya ha sido atendida en el almacen de destino, no se puede modificar.", MsgBoxStyle.Exclamation, "Gestion Visanfer")
            Else
                mtEstado = EstadoVentana.Mantenimiento ' entra en modo mantenimiento
                grdLineas.Editable = True
                lblTeclas.BackColor = Color.Tomato
                lblTeclas.ForeColor = Color.White
                lblTeclas.Text = " F1.-INSERTA LINEA          F2.-BORRA LINEA      " & _
                                 "      F5-.GRABA             ESC-.SALIDA"
                lblPrograma.Text = "SOL. TRASPASOS - MANTENIMIENTO"
                lblPrograma.BackColor = Color.Tomato
            End If
        End If
        'End If

    End Sub

    Private Sub mrPreparaGrid()

        grdLineas.Columnas = 7

        grdLineas.marTitulos(0).Texto = "ARTIC."
        grdLineas.marTitulos(0).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(0).Ancho = 100
        grdLineas.marTitulos(0).Longitud = 13
        grdLineas.marTitulos(0).Editable = True

        grdLineas.marTitulos(1).Texto = "DESCRIPCION"
        grdLineas.marTitulos(1).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(1).Ancho = 335
        grdLineas.marTitulos(1).Longitud = 100
        grdLineas.marTitulos(1).Editable = True

        grdLineas.marTitulos(2).Texto = "CTD."
        grdLineas.marTitulos(2).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(2).Ancho = 100
        grdLineas.marTitulos(2).Longitud = 10
        grdLineas.marTitulos(2).Editable = True

        grdLineas.marTitulos(3).Texto = "OBSERVACIONES"
        grdLineas.marTitulos(3).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(3).Ancho = 280
        grdLineas.marTitulos(3).Longitud = 50
        grdLineas.marTitulos(3).Editable = True

        grdLineas.marTitulos(4).Texto = "mpsCodigo"
        grdLineas.marTitulos(4).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(4).Ancho = 0
        grdLineas.marTitulos(4).Longitud = 15
        grdLineas.marTitulos(4).Editable = False

        grdLineas.marTitulos(5).Texto = "marcas"
        grdLineas.marTitulos(5).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(5).Ancho = 0
        grdLineas.marTitulos(5).Longitud = 1
        grdLineas.marTitulos(5).Editable = False

        grdLineas.marTitulos(6).Texto = "A"
        grdLineas.marTitulos(6).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(6).Ancho = 0
        grdLineas.marTitulos(6).Longitud = 6
        grdLineas.marTitulos(6).Editable = False

        'grdLineas.mnAjustarColumna = 2
        grdLineas.mrPintaGrid()

    End Sub

    Private Sub mrMoverCampos(ByVal lnTipo As Integer)
        Dim loLinea As clsSolTraspasoLin
        Dim lnI As Integer

        If lnTipo = 1 Then
            ' primero los campos de la cabecera **********
            txtCodigo.Text = moSolTraspaso.mnCodigo
            txtFecha.Text = Format(moSolTraspaso.mdFecha, "dd/MM/yyyy")
            txtEstado.Text = moSolTraspaso.msEstado
            txtComentario.Text = moSolTraspaso.msObservaciones
            ' ****************
            lblConfirmado.Visible = (moSolTraspaso.msEstado = "A")
            txtDesde.Text = moSolTraspaso.mnDesde
            moAlmacen1 = New clsAlmacen
            moAlmacen1.mnEmpresa = mnEmpresa
            moAlmacen1.mnCodigo = moSolTraspaso.mnDesde
            moAlmacen1.mrRecuperaDatos(False)
            txtDesA1.Text = moAlmacen1.msNombre
            ' ****************
            txtHasta.Text = moSolTraspaso.mnHasta
            moAlmacen2 = New clsAlmacen
            moAlmacen2.mnEmpresa = mnEmpresa
            moAlmacen2.mnCodigo = moSolTraspaso.mnHasta
            moAlmacen2.mrRecuperaDatos(False)
            txtDesA2.Text = moAlmacen2.msNombre
            lblTraspaso.Text = txtDesA1.Text & "   ····>   " & txtDesA2.Text
            lblConfirma.Visible = (moSolTraspaso.msEstado = "P")
            ' ****************
            txtOperador.Text = moSolTraspaso.mnOperario
            mrCargaOperario()
            ' despues los campos de las lineas ***********
            lnI = 0
            For Each loLinea In moSolTraspaso.mcolLineas
                grdLineas.mrAñadirFila()
                grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo
                If loLinea.mnDetalle > 0 Then grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo & "." & loLinea.mnDetalle
                grdLineas.marMemoria(1, lnI) = loLinea.msDescripcion
                grdLineas.marMemoria(2, lnI) = Format(loLinea.mnCantidad, "#,##0.00")
                grdLineas.marMemoria(3, lnI) = loLinea.msEstado
                grdLineas.marMemoria(4, lnI) = loLinea.mpsCodigo
                grdLineas.marMemoria(5, lnI) = ""
                grdLineas.marMemoria(6, lnI) = loLinea.mnArticulo
                lnI = lnI + 1
            Next
            ' despues añado una linea nueva que es la que pongo para meter nuevos datos
            grdLineas.mrAñadirFila()
            grdLineas.mrRefrescaGrid()
        Else
            ' primero los campos de la cabecera **************
            moSolTraspaso.mnCodigo = txtCodigo.Text
            moSolTraspaso.mdFecha = txtFecha.Text
            moSolTraspaso.msEstado = txtEstado.Text
            moSolTraspaso.msObservaciones = txtComentario.Text
            moSolTraspaso.mnDesde = txtDesde.Text
            moSolTraspaso.mnHasta = txtHasta.Text
            moSolTraspaso.mnOperario = txtOperador.Text
            moSolTraspaso.mnVendedor = 0
            ' ************************************************
            mrCargaLineas()
        End If

    End Sub

    Private Sub mrPendiente(ByVal lnFila As Integer, ByVal lsValor As String)
        Dim lnI As Integer
        Dim loColor As Color

        If lsValor = "S" Then
            loColor = Color.Red
        Else
            loColor = Color.Black
        End If

        For lnI = 0 To 5
            grdLineas.ColorCelda(lnI, lnFila).mnForeColor = loColor
        Next

    End Sub

    Private Function mfbCargaArticulo() As Boolean
        Dim loArticulo As New clsArticulo
        Dim lnArticuloTemp As Integer
        Dim lnDetalleTemp As Integer
        Dim lnArticulo As Long
        Dim lnDetalle As Integer
        Dim lnLinea As Integer
        Dim loExistencias As clsExistencias
        Dim lsCodigo As String

        If moAlmacen1 Is Nothing Then moAlmacen1 = New clsAlmacen
        If moAlmacen2 Is Nothing Then moAlmacen2 = New clsAlmacen

        If moAlmacen1.mnCodigo = 0 Then
            MsgBox("Debe indicar los almacenes que intervienen en el traspaso.", MsgBoxStyle.Critical, "Visanfer .Net")
            Return False
        End If

        lnLinea = grdLineas.mnRow
        lsCodigo = Trim(grdLineas.marMemoria(0, lnLinea))
        If lsCodigo = "" Then
            grdLineas.marMemoria(1, lnLinea) = ""
            grdLineas.mrRefrescaGrid()
            grdLineas.mbSaltoAutomatico = False
            grdLineas.mrSeleccionaTexto()
            Return False
        End If

        lnArticulo = mfnCodigoArticulo(lsCodigo)
        lnDetalle = mfnCodigoDetalle(lsCodigo)
        ' primero miro si se ha metido el codigo de barras ****************
        If (lnDetalle = 0) And (Len(lsCodigo) > 7) Then
            Dim loCodAlter As New prjAlbaranes.clsCodalter
            loCodAlter.msEan = lsCodigo
            loCodAlter.mrRecuperaDatos()
            If loCodAlter.mbEsNuevo Then
                grdLineas.marMemoria(0, lnLinea) = ""
                MsgBox("Articulo no encontrado.", MsgBoxStyle.Information, "Visanfer .Net")
            Else
                grdLineas.marMemoria(0, lnLinea) = loCodAlter.mnArticulo
                lnArticulo = loCodAlter.mnArticulo
                If loCodAlter.mnDetalle > 0 Then
                    grdLineas.marMemoria(0, lnLinea) = loCodAlter.mnArticulo & "." & loCodAlter.mnDetalle
                    lnDetalle = loCodAlter.mnDetalle
                End If
            End If
            grdLineas.mrRefrescaGrid()
        End If

        lnArticuloTemp = mfnCodigoArticulo(grdLineas.marMemoria(5, lnLinea))
        lnDetalleTemp = mfnCodigoDetalle(grdLineas.marMemoria(5, lnLinea))

        mbAviso1 = False
        mbAviso2 = False
        lblAviso1.Visible = False
        lblAviso2.Visible = False
        tmrAviso.Enabled = False

        'If lnArticulo <> lnArticuloTemp Then

        loArticulo.mnEmpresa = mnEmpresa
        loArticulo.mnCodigo = lnArticulo
        loArticulo.mnDetalle = lnDetalle
        If lnArticulo = 0 Then
            loArticulo.mbEsNuevo = True
        Else
            loArticulo.mrRecuperaDatos()
        End If
        grdLineas.marMemoria(5, lnLinea) = lnArticulo
        mrPendiente(lnLinea, "S")
        If loArticulo.mbEsNuevo Then
            grdLineas.marMemoria(5, lnLinea) = ""
            If lnArticulo > 0 Then
                grdLineas.marMemoria(1, lnLinea) = "ARTICULO NO ENCONTRADO"
                grdLineas.mrRefrescaGrid()
                grdLineas.mbSaltoAutomatico = False
                grdLineas.mrSeleccionaTexto()
            End If
            Return False
        Else

            If loArticulo.mnDetalle = 0 Then loArticulo.mrRecuperaDetalles(False)
            If loArticulo.mbTieneDetalle Then
                If lnDetalle <> 0 Then
                    loArticulo.mnDetalle = lnDetalle
                Else
                    Dim loSelDetalles As New frmSelDetalle
                    loSelDetalles.mnTop = grdLineas.Top + panCampos.Top + Me.Top
                    loSelDetalles.mnLeft = grdLineas.Left + Me.Left + panCampos.Left
                    loSelDetalles.mnEmpresa = mnEmpresa
                    loSelDetalles.mnArticulo = loArticulo.mnCodigo
                    loSelDetalles.mrCargar()

                    If loSelDetalles.mnDetalle > 0 Then
                        grdLineas.marMemoria(0, lnLinea) = loSelDetalles.mnArticulo & "." & loSelDetalles.mnDetalle
                        loArticulo.mnDetalle = loSelDetalles.mnDetalle
                        loArticulo.mrRecuperaDatos()
                    End If

                End If


                If loArticulo.mbEsNuevo Then

                    grdLineas.marMemoria(5, lnLinea) = ""
                    If lnArticulo > 0 Then
                        grdLineas.marMemoria(1, lnLinea) = "ARTICULO NO ENCONTRADO"
                        grdLineas.mrRefrescaGrid()
                        grdLineas.mbSaltoAutomatico = False
                        grdLineas.mrSeleccionaTexto()
                    End If
                    Return False

                End If

            End If

            grdLineas.marMemoria(1, lnLinea) = loArticulo.msDescripcion

            ' no acepto articulos virtuales
            Dim loAuxArticulo As New clsArticulo
            loAuxArticulo.mnEmpresa = mnEmpresa
            loAuxArticulo.mnCodigo = lnArticulo
            loAuxArticulo.mnDetalle = lnDetalle
            loAuxArticulo.mrRecuperaDatos()
            If loAuxArticulo.mbEsNuevo Then loAuxArticulo.msControlExistencias = "N"
            If loAuxArticulo.msControlExistencias = "S" Then
                Dim loAgrArt As New clsAgrArt
                loAgrArt.mnEmpresa = mnEmpresa
                loAgrArt.mnCodigo = lnArticulo
                loAgrArt.mnDetalle = lnDetalle
                Dim loBusAgrArt As New clsBusAgrArt
                loBusAgrArt.mrBusca(loAgrArt)

                If loBusAgrArt.mcolAgrart.Count > 0 Then
                    MsgBox("NO SE PUEDEN HACER TRASPASOS DE ARTICULOS VIRTUALES.", MsgBoxStyle.Critical, "Visanfer .Net")
                    grdLineas.marMemoria(1, lnLinea) = "ARTICULO NO ENCONTRADO"
                    grdLineas.mrRefrescaGrid()
                    grdLineas.mbSaltoAutomatico = False
                    grdLineas.mrSeleccionaTexto()
                    Return False
                End If
            End If

            ' ************************************************************
            loExistencias = New clsExistencias
            loExistencias.mnEmpresa = mnEmpresa
            loExistencias.mnAlmacen = moAlmacen1.mnCodigo
            loExistencias.mnArticulo = lnArticulo
            loExistencias.mnDetalle = lnDetalle
            loExistencias.mrRecuperaDatos()
            txtExis1.Text = Format(loExistencias.mnExistencias, "#,##0.00")
            txtExpo1.Text = Format(loExistencias.mnExposicion, "#,##0")
            If loExistencias.mnExistencias <= 0 Then
                txtExis1.BackColor = Color.Red
                mbAviso1 = True
            Else
                txtExis1.BackColor = Color.White
            End If
            If loExistencias.mnExposicion > 0 Then
                txtExpo1.BackColor = Color.GreenYellow
            Else
                txtExpo1.BackColor = Color.White
            End If

            loExistencias.mnAlmacen = moAlmacen2.mnCodigo
            loExistencias.mrRecuperaDatos()
            txtExis2.Text = Format(loExistencias.mnExistencias, "#,##0.00")
            txtExpo2.Text = Format(loExistencias.mnExposicion, "#,##0")
            If loExistencias.mnExistencias < 0 Then
                txtExis2.BackColor = Color.Red
                mbAviso2 = True
            Else
                txtExis2.BackColor = Color.White
            End If
            ' ************************************************************
        End If
        grdLineas.mrRefrescaGrid()

        If mbAviso1 Or mbAviso2 Then tmrAviso.Enabled = True
        Return True

    End Function

    Private Sub tmrAviso_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrAviso.Tick

        If mbAviso1 Then lblAviso1.Visible = Not lblAviso1.Visible
        If mbAviso2 Then lblAviso2.Visible = Not lblAviso2.Visible

    End Sub

    Private Sub mrAsignaEventos()
        Dim loControl As Object
        Dim loCaja As control.txtVisanfer

        For Each loControl In panCampos.Controls
            If TypeOf loControl Is control.txtVisanfer Then
                loCaja = loControl
                AddHandler loCaja.KeyDown, AddressOf mrLeeTecla
            End If
        Next

    End Sub

#End Region

#Region " Eventos de Formulario "

    Private Sub frmSolTraspasos_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyData = 131138 Then goUsuario.mrBloquear(gnLlave)
    End Sub

    Private Sub frmSolTraspasos_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.Text = goUsuario.msNombre
    End Sub

    Private Sub cmdDesbloqueo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDesbloqueo.Click
        cmdDesbloqueo.Visible = False
        cmdBloqueo.Visible = True
        txtHasta.ReadOnly = False
        txtHasta.Focus()
    End Sub

    Private Sub cmdBloqueo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBloqueo.Click
        cmdDesbloqueo.Visible = True
        cmdBloqueo.Visible = False
        txtHasta.ReadOnly = True
        txtHasta.Focus()
    End Sub

#End Region

#Region " Eventos de Teclado para las cajas de texto "

    Private Sub txtCodigo_evtSalida() Handles txtCodigo.evtSalida
        If mtEstado = EstadoVentana.Consulta Then
            txtCodigo.Focus()
            txtCodigo.SelectAll()
        End If
    End Sub

#End Region

#Region " Eventos del Grid "

    Private Sub grdLineas_evtKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdLineas.evtKeyDown
        mrLeeTecla(grdLineas, e)
    End Sub

    Private Sub grdLineas_evtLeaveCell() Handles grdLineas.evtLeaveCell
        'If mtEstado = EstadoVentana.Mantenimiento Then mrAnalizaDatos()
    End Sub

#End Region

End Class
