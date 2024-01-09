Option Explicit On

Imports System.Windows.Forms.SendKeys
Imports prjControl
Imports prjPrinterNet
Imports prjEmpresas
Imports prjArticulos
Imports prjAlmacen

Public Class frmTraspasos
    Inherits System.Windows.Forms.Form

    Private mnEmpresa As Int32      ' empresa de gestion
    Public mbCargaDirecta As Boolean = False

    Dim WithEvents moTraspaso As clsTraspaso
    Dim WithEvents moTraspasoAux As clsTraspaso
    Dim moAlmacen1 As clsAlmacen
    Dim moAlmacen2 As clsAlmacen
    Dim moEmpContable As clsEmpContable
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
    Dim mbPrimeraVez As Boolean
    Dim mbAviso1 As Boolean = False
    Dim mbAviso2 As Boolean = False
    Dim mcolLineas As Collection
    ' variables de impresion ************************
    Private WithEvents moSelImpresora As prjControl.frmSelImpresora
    Dim moImpresora As New prjPrinterNet.clsImpresora   ' Objeto Impresora
    Dim mbCompleto As Boolean
    Dim mbCargaAutomatica As Boolean = False
    Dim mbConsultaTraspaso As Boolean = False
    Dim moDataTraspaso As dtsTraspaso
    Dim moBusSolTraspasos As clsBusSolTraspasos
    Dim mcolOperarios As New Collection

    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents grpFechas As System.Windows.Forms.GroupBox
    Friend WithEvents lblOperarioEnvio As System.Windows.Forms.Label
    Friend WithEvents txtOperarioEnvia As control.txtVisanfer
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtOperarioGraba As control.txtVisanfer
    Friend WithEvents lblEstado As System.Windows.Forms.Label
    Friend WithEvents txtFechaEnvio As control.txtVisanfer
    Friend WithEvents lblFechaEnvio As System.Windows.Forms.Label
    Friend WithEvents txtFechaGrabacion As control.txtVisanfer
    Friend WithEvents cmdEnviado As System.Windows.Forms.Button
    Friend WithEvents grpEnvio As System.Windows.Forms.GroupBox
    Friend WithEvents grpRecibo As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRecibido As System.Windows.Forms.Button
    Friend WithEvents lblOperarioRecepcion As System.Windows.Forms.Label
    Friend WithEvents txtOperarioRecibe As control.txtVisanfer
    Friend WithEvents txtFechaRecepcion As control.txtVisanfer
    Friend WithEvents lblFechaRecepcion As System.Windows.Forms.Label
    Friend WithEvents lblConfirmado As Label
    Friend WithEvents panTitulo As Panel
    Friend WithEvents lblTitle As Label

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
    Friend WithEvents grdLineas As prjGrid.ctlGrid
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtComentario As control.txtVisanfer
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtCodigo As control.txtVisanfer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblDescripcion As System.Windows.Forms.Label
    Friend WithEvents lblTitulo As System.Windows.Forms.Label
    Friend WithEvents lblTeclas As System.Windows.Forms.Label
    Friend WithEvents txtDesA2 As control.txtVisanfer
    Friend WithEvents txtDesA1 As control.txtVisanfer
    Friend WithEvents txtHasta As control.txtVisanfer
    Friend WithEvents txtDesde As control.txtVisanfer
    Friend WithEvents txtEstado As control.txtVisanfer
    Friend WithEvents txtFecha As control.txtVisanfer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtOperador As control.txtVisanfer
    Friend WithEvents txtVendedor As control.txtVisanfer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tabMenu As prjToolBar.ctlToolBar
    Friend WithEvents txtExis1 As control.txtVisanfer
    Friend WithEvents txtExis2 As control.txtVisanfer
    Friend WithEvents tmrAviso As System.Windows.Forms.Timer
    Friend WithEvents lblAviso2 As System.Windows.Forms.Label
    Friend WithEvents lblAviso1 As System.Windows.Forms.Label
    Friend WithEvents lblTraspaso As System.Windows.Forms.Label
    Friend WithEvents cmdBloqueo As System.Windows.Forms.Button
    Friend WithEvents cmdDesbloqueo As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtExpo1 As control.txtVisanfer
    Friend WithEvents txtExpo2 As control.txtVisanfer
    Friend WithEvents lblSolicitudes As System.Windows.Forms.Label
    Friend WithEvents lblSolicitud As System.Windows.Forms.Label
    Friend WithEvents tmrSolicitud As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTraspasos))
        Me.lblPrograma = New System.Windows.Forms.Label()
        Me.panCampos = New System.Windows.Forms.Panel()
        Me.lblConfirmado = New System.Windows.Forms.Label()
        Me.grpRecibo = New System.Windows.Forms.GroupBox()
        Me.cmdRecibido = New System.Windows.Forms.Button()
        Me.lblOperarioRecepcion = New System.Windows.Forms.Label()
        Me.txtOperarioRecibe = New control.txtVisanfer()
        Me.txtFechaRecepcion = New control.txtVisanfer()
        Me.lblFechaRecepcion = New System.Windows.Forms.Label()
        Me.grpEnvio = New System.Windows.Forms.GroupBox()
        Me.cmdEnviado = New System.Windows.Forms.Button()
        Me.lblFechaEnvio = New System.Windows.Forms.Label()
        Me.txtFechaEnvio = New control.txtVisanfer()
        Me.txtOperarioEnvia = New control.txtVisanfer()
        Me.lblOperarioEnvio = New System.Windows.Forms.Label()
        Me.grpFechas = New System.Windows.Forms.GroupBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtOperarioGraba = New control.txtVisanfer()
        Me.txtFechaGrabacion = New control.txtVisanfer()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblSolicitud = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtExpo2 = New control.txtVisanfer()
        Me.txtExpo1 = New control.txtVisanfer()
        Me.lblEstado = New System.Windows.Forms.Label()
        Me.cmdDesbloqueo = New System.Windows.Forms.Button()
        Me.cmdBloqueo = New System.Windows.Forms.Button()
        Me.lblTraspaso = New System.Windows.Forms.Label()
        Me.lblAviso2 = New System.Windows.Forms.Label()
        Me.lblAviso1 = New System.Windows.Forms.Label()
        Me.txtExis2 = New control.txtVisanfer()
        Me.txtExis1 = New control.txtVisanfer()
        Me.txtVendedor = New control.txtVisanfer()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grdLineas = New prjGrid.ctlGrid()
        Me.txtDesA2 = New control.txtVisanfer()
        Me.txtDesA1 = New control.txtVisanfer()
        Me.txtHasta = New control.txtVisanfer()
        Me.txtDesde = New control.txtVisanfer()
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
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.lblTeclas = New System.Windows.Forms.Label()
        Me.tabMenu = New prjToolBar.ctlToolBar()
        Me.tmrAviso = New System.Windows.Forms.Timer(Me.components)
        Me.lblSolicitudes = New System.Windows.Forms.Label()
        Me.tmrSolicitud = New System.Windows.Forms.Timer(Me.components)
        Me.panTitulo = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.panCampos.SuspendLayout()
        Me.grpRecibo.SuspendLayout()
        Me.grpEnvio.SuspendLayout()
        Me.grpFechas.SuspendLayout()
        Me.panTitulo.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(12, 34)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(673, 32)
        Me.lblPrograma.TabIndex = 39
        Me.lblPrograma.Text = "GESTION"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.lblConfirmado)
        Me.panCampos.Controls.Add(Me.grpRecibo)
        Me.panCampos.Controls.Add(Me.grpEnvio)
        Me.panCampos.Controls.Add(Me.grpFechas)
        Me.panCampos.Controls.Add(Me.lblSolicitud)
        Me.panCampos.Controls.Add(Me.Label9)
        Me.panCampos.Controls.Add(Me.Label10)
        Me.panCampos.Controls.Add(Me.Label8)
        Me.panCampos.Controls.Add(Me.Label6)
        Me.panCampos.Controls.Add(Me.txtExpo2)
        Me.panCampos.Controls.Add(Me.txtExpo1)
        Me.panCampos.Controls.Add(Me.lblEstado)
        Me.panCampos.Controls.Add(Me.cmdDesbloqueo)
        Me.panCampos.Controls.Add(Me.cmdBloqueo)
        Me.panCampos.Controls.Add(Me.lblTraspaso)
        Me.panCampos.Controls.Add(Me.lblAviso2)
        Me.panCampos.Controls.Add(Me.lblAviso1)
        Me.panCampos.Controls.Add(Me.txtExis2)
        Me.panCampos.Controls.Add(Me.txtExis1)
        Me.panCampos.Controls.Add(Me.txtVendedor)
        Me.panCampos.Controls.Add(Me.Label4)
        Me.panCampos.Controls.Add(Me.Label2)
        Me.panCampos.Controls.Add(Me.grdLineas)
        Me.panCampos.Controls.Add(Me.txtDesA2)
        Me.panCampos.Controls.Add(Me.txtDesA1)
        Me.panCampos.Controls.Add(Me.txtHasta)
        Me.panCampos.Controls.Add(Me.txtDesde)
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
        Me.panCampos.Location = New System.Drawing.Point(12, 114)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(1001, 608)
        Me.panCampos.TabIndex = 37
        '
        'lblConfirmado
        '
        Me.lblConfirmado.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblConfirmado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblConfirmado.Location = New System.Drawing.Point(8, 573)
        Me.lblConfirmado.Name = "lblConfirmado"
        Me.lblConfirmado.Size = New System.Drawing.Size(588, 20)
        Me.lblConfirmado.TabIndex = 123
        Me.lblConfirmado.Text = "----"
        Me.lblConfirmado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpRecibo
        '
        Me.grpRecibo.Controls.Add(Me.cmdRecibido)
        Me.grpRecibo.Controls.Add(Me.lblOperarioRecepcion)
        Me.grpRecibo.Controls.Add(Me.txtOperarioRecibe)
        Me.grpRecibo.Controls.Add(Me.txtFechaRecepcion)
        Me.grpRecibo.Controls.Add(Me.lblFechaRecepcion)
        Me.grpRecibo.Location = New System.Drawing.Point(602, 468)
        Me.grpRecibo.Name = "grpRecibo"
        Me.grpRecibo.Size = New System.Drawing.Size(390, 121)
        Me.grpRecibo.TabIndex = 122
        Me.grpRecibo.TabStop = False
        Me.grpRecibo.Visible = False
        '
        'cmdRecibido
        '
        Me.cmdRecibido.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdRecibido.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRecibido.Image = CType(resources.GetObject("cmdRecibido.Image"), System.Drawing.Image)
        Me.cmdRecibido.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRecibido.Location = New System.Drawing.Point(16, 20)
        Me.cmdRecibido.Name = "cmdRecibido"
        Me.cmdRecibido.Size = New System.Drawing.Size(361, 35)
        Me.cmdRecibido.TabIndex = 117
        Me.cmdRecibido.Text = "F6 - MARCAR COMO RECIBIDO"
        Me.cmdRecibido.UseVisualStyleBackColor = False
        Me.cmdRecibido.Visible = False
        '
        'lblOperarioRecepcion
        '
        Me.lblOperarioRecepcion.Location = New System.Drawing.Point(32, 88)
        Me.lblOperarioRecepcion.Name = "lblOperarioRecepcion"
        Me.lblOperarioRecepcion.Size = New System.Drawing.Size(64, 16)
        Me.lblOperarioRecepcion.TabIndex = 116
        Me.lblOperarioRecepcion.Text = "OPERARIO"
        Me.lblOperarioRecepcion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOperarioRecibe
        '
        Me.txtOperarioRecibe.AutoSelec = False
        Me.txtOperarioRecibe.BackColor = System.Drawing.Color.White
        Me.txtOperarioRecibe.Blink = False
        Me.txtOperarioRecibe.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOperarioRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtOperarioRecibe.DesdeCodigo = CType(0, Long)
        Me.txtOperarioRecibe.DesdeFecha = New Date(CType(0, Long))
        Me.txtOperarioRecibe.HastaCodigo = CType(0, Long)
        Me.txtOperarioRecibe.HastaFecha = New Date(CType(0, Long))
        Me.txtOperarioRecibe.Location = New System.Drawing.Point(105, 84)
        Me.txtOperarioRecibe.MaxLength = 30
        Me.txtOperarioRecibe.moBlink = Nothing
        Me.txtOperarioRecibe.Name = "txtOperarioRecibe"
        Me.txtOperarioRecibe.ReadOnly = True
        Me.txtOperarioRecibe.Size = New System.Drawing.Size(257, 20)
        Me.txtOperarioRecibe.TabIndex = 115
        Me.txtOperarioRecibe.TabStop = False
        Me.txtOperarioRecibe.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtOperarioRecibe.ValorMax = 999999999.0R
        '
        'txtFechaRecepcion
        '
        Me.txtFechaRecepcion.AutoSelec = False
        Me.txtFechaRecepcion.BackColor = System.Drawing.Color.White
        Me.txtFechaRecepcion.Blink = False
        Me.txtFechaRecepcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFechaRecepcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFechaRecepcion.DesdeCodigo = CType(0, Long)
        Me.txtFechaRecepcion.DesdeFecha = New Date(CType(0, Long))
        Me.txtFechaRecepcion.HastaCodigo = CType(0, Long)
        Me.txtFechaRecepcion.HastaFecha = New Date(CType(0, Long))
        Me.txtFechaRecepcion.Location = New System.Drawing.Point(158, 61)
        Me.txtFechaRecepcion.MaxLength = 10
        Me.txtFechaRecepcion.moBlink = Nothing
        Me.txtFechaRecepcion.Name = "txtFechaRecepcion"
        Me.txtFechaRecepcion.Size = New System.Drawing.Size(204, 20)
        Me.txtFechaRecepcion.TabIndex = 113
        Me.txtFechaRecepcion.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFechaRecepcion.tipo = control.txtVisanfer.TiposCaja.fecha
        Me.txtFechaRecepcion.ValorMax = 999999999.0R
        '
        'lblFechaRecepcion
        '
        Me.lblFechaRecepcion.Location = New System.Drawing.Point(32, 65)
        Me.lblFechaRecepcion.Name = "lblFechaRecepcion"
        Me.lblFechaRecepcion.Size = New System.Drawing.Size(120, 16)
        Me.lblFechaRecepcion.TabIndex = 114
        Me.lblFechaRecepcion.Text = "FECHA RECEPCION:"
        '
        'grpEnvio
        '
        Me.grpEnvio.Controls.Add(Me.cmdEnviado)
        Me.grpEnvio.Controls.Add(Me.lblFechaEnvio)
        Me.grpEnvio.Controls.Add(Me.txtFechaEnvio)
        Me.grpEnvio.Controls.Add(Me.txtOperarioEnvia)
        Me.grpEnvio.Controls.Add(Me.lblOperarioEnvio)
        Me.grpEnvio.Location = New System.Drawing.Point(602, 356)
        Me.grpEnvio.Name = "grpEnvio"
        Me.grpEnvio.Size = New System.Drawing.Size(390, 122)
        Me.grpEnvio.TabIndex = 121
        Me.grpEnvio.TabStop = False
        Me.grpEnvio.Visible = False
        '
        'cmdEnviado
        '
        Me.cmdEnviado.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdEnviado.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEnviado.Image = CType(resources.GetObject("cmdEnviado.Image"), System.Drawing.Image)
        Me.cmdEnviado.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEnviado.Location = New System.Drawing.Point(17, 19)
        Me.cmdEnviado.Name = "cmdEnviado"
        Me.cmdEnviado.Size = New System.Drawing.Size(361, 35)
        Me.cmdEnviado.TabIndex = 111
        Me.cmdEnviado.Text = "F8 - MARCAR COMO ENVIADO"
        Me.cmdEnviado.UseVisualStyleBackColor = False
        Me.cmdEnviado.Visible = False
        '
        'lblFechaEnvio
        '
        Me.lblFechaEnvio.Location = New System.Drawing.Point(32, 64)
        Me.lblFechaEnvio.Name = "lblFechaEnvio"
        Me.lblFechaEnvio.Size = New System.Drawing.Size(120, 16)
        Me.lblFechaEnvio.TabIndex = 10
        Me.lblFechaEnvio.Text = "FECHA DE ENVIO:"
        '
        'txtFechaEnvio
        '
        Me.txtFechaEnvio.AutoSelec = False
        Me.txtFechaEnvio.BackColor = System.Drawing.Color.White
        Me.txtFechaEnvio.Blink = False
        Me.txtFechaEnvio.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFechaEnvio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFechaEnvio.DesdeCodigo = CType(0, Long)
        Me.txtFechaEnvio.DesdeFecha = New Date(CType(0, Long))
        Me.txtFechaEnvio.HastaCodigo = CType(0, Long)
        Me.txtFechaEnvio.HastaFecha = New Date(CType(0, Long))
        Me.txtFechaEnvio.Location = New System.Drawing.Point(158, 60)
        Me.txtFechaEnvio.MaxLength = 10
        Me.txtFechaEnvio.moBlink = Nothing
        Me.txtFechaEnvio.Name = "txtFechaEnvio"
        Me.txtFechaEnvio.Size = New System.Drawing.Size(204, 20)
        Me.txtFechaEnvio.TabIndex = 9
        Me.txtFechaEnvio.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFechaEnvio.tipo = control.txtVisanfer.TiposCaja.fecha
        Me.txtFechaEnvio.ValorMax = 999999999.0R
        '
        'txtOperarioEnvia
        '
        Me.txtOperarioEnvia.AutoSelec = False
        Me.txtOperarioEnvia.BackColor = System.Drawing.Color.White
        Me.txtOperarioEnvia.Blink = False
        Me.txtOperarioEnvia.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOperarioEnvia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtOperarioEnvia.DesdeCodigo = CType(0, Long)
        Me.txtOperarioEnvia.DesdeFecha = New Date(CType(0, Long))
        Me.txtOperarioEnvia.HastaCodigo = CType(0, Long)
        Me.txtOperarioEnvia.HastaFecha = New Date(CType(0, Long))
        Me.txtOperarioEnvia.Location = New System.Drawing.Point(105, 86)
        Me.txtOperarioEnvia.MaxLength = 30
        Me.txtOperarioEnvia.moBlink = Nothing
        Me.txtOperarioEnvia.Name = "txtOperarioEnvia"
        Me.txtOperarioEnvia.ReadOnly = True
        Me.txtOperarioEnvia.Size = New System.Drawing.Size(257, 20)
        Me.txtOperarioEnvia.TabIndex = 107
        Me.txtOperarioEnvia.TabStop = False
        Me.txtOperarioEnvia.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtOperarioEnvia.ValorMax = 999999999.0R
        '
        'lblOperarioEnvio
        '
        Me.lblOperarioEnvio.Location = New System.Drawing.Point(32, 90)
        Me.lblOperarioEnvio.Name = "lblOperarioEnvio"
        Me.lblOperarioEnvio.Size = New System.Drawing.Size(64, 16)
        Me.lblOperarioEnvio.TabIndex = 108
        Me.lblOperarioEnvio.Text = "OPERARIO"
        Me.lblOperarioEnvio.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grpFechas
        '
        Me.grpFechas.Controls.Add(Me.Label16)
        Me.grpFechas.Controls.Add(Me.txtOperarioGraba)
        Me.grpFechas.Controls.Add(Me.txtFechaGrabacion)
        Me.grpFechas.Controls.Add(Me.Label11)
        Me.grpFechas.Location = New System.Drawing.Point(602, 285)
        Me.grpFechas.Name = "grpFechas"
        Me.grpFechas.Size = New System.Drawing.Size(390, 80)
        Me.grpFechas.TabIndex = 120
        Me.grpFechas.TabStop = False
        Me.grpFechas.Text = " HISTORIAL DEL TRASPASO "
        Me.grpFechas.Visible = False
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(32, 49)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 16)
        Me.Label16.TabIndex = 106
        Me.Label16.Text = "OPERARIO"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOperarioGraba
        '
        Me.txtOperarioGraba.AutoSelec = False
        Me.txtOperarioGraba.BackColor = System.Drawing.Color.White
        Me.txtOperarioGraba.Blink = False
        Me.txtOperarioGraba.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOperarioGraba.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtOperarioGraba.DesdeCodigo = CType(0, Long)
        Me.txtOperarioGraba.DesdeFecha = New Date(CType(0, Long))
        Me.txtOperarioGraba.HastaCodigo = CType(0, Long)
        Me.txtOperarioGraba.HastaFecha = New Date(CType(0, Long))
        Me.txtOperarioGraba.Location = New System.Drawing.Point(105, 45)
        Me.txtOperarioGraba.MaxLength = 30
        Me.txtOperarioGraba.moBlink = Nothing
        Me.txtOperarioGraba.Name = "txtOperarioGraba"
        Me.txtOperarioGraba.ReadOnly = True
        Me.txtOperarioGraba.Size = New System.Drawing.Size(257, 20)
        Me.txtOperarioGraba.TabIndex = 105
        Me.txtOperarioGraba.TabStop = False
        Me.txtOperarioGraba.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtOperarioGraba.ValorMax = 999999999.0R
        '
        'txtFechaGrabacion
        '
        Me.txtFechaGrabacion.AutoSelec = False
        Me.txtFechaGrabacion.BackColor = System.Drawing.Color.White
        Me.txtFechaGrabacion.Blink = False
        Me.txtFechaGrabacion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFechaGrabacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFechaGrabacion.DesdeCodigo = CType(0, Long)
        Me.txtFechaGrabacion.DesdeFecha = New Date(CType(0, Long))
        Me.txtFechaGrabacion.HastaCodigo = CType(0, Long)
        Me.txtFechaGrabacion.HastaFecha = New Date(CType(0, Long))
        Me.txtFechaGrabacion.Location = New System.Drawing.Point(158, 19)
        Me.txtFechaGrabacion.MaxLength = 10
        Me.txtFechaGrabacion.moBlink = Nothing
        Me.txtFechaGrabacion.Name = "txtFechaGrabacion"
        Me.txtFechaGrabacion.Size = New System.Drawing.Size(204, 20)
        Me.txtFechaGrabacion.TabIndex = 7
        Me.txtFechaGrabacion.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtFechaGrabacion.tipo = control.txtVisanfer.TiposCaja.fecha
        Me.txtFechaGrabacion.ValorMax = 999999999.0R
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(32, 23)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 16)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "FECHA GRABACION:"
        '
        'lblSolicitud
        '
        Me.lblSolicitud.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblSolicitud.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSolicitud.Location = New System.Drawing.Point(13, 20)
        Me.lblSolicitud.Name = "lblSolicitud"
        Me.lblSolicitud.Size = New System.Drawing.Size(73, 20)
        Me.lblSolicitud.TabIndex = 119
        Me.lblSolicitud.Text = "SOLICITUD:"
        Me.lblSolicitud.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblSolicitud.Visible = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(824, 44)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 16)
        Me.Label9.TabIndex = 118
        Me.Label9.Text = "EXPOSI."
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(760, 44)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 16)
        Me.Label10.TabIndex = 117
        Me.Label10.Text = "EXISTENCIAS"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(824, 4)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(43, 16)
        Me.Label8.TabIndex = 116
        Me.Label8.Text = "EXPOSI."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(760, 4)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 16)
        Me.Label6.TabIndex = 115
        Me.Label6.Text = "EXISTENCIAS"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.txtExpo2.Location = New System.Drawing.Point(824, 60)
        Me.txtExpo2.MaxLength = 6
        Me.txtExpo2.moBlink = Nothing
        Me.txtExpo2.Name = "txtExpo2"
        Me.txtExpo2.ReadOnly = True
        Me.txtExpo2.Size = New System.Drawing.Size(43, 20)
        Me.txtExpo2.TabIndex = 114
        Me.txtExpo2.TabStop = False
        Me.txtExpo2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExpo2.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExpo2.ValorMax = 999999999.0R
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
        Me.txtExpo1.Location = New System.Drawing.Point(824, 20)
        Me.txtExpo1.MaxLength = 6
        Me.txtExpo1.moBlink = Nothing
        Me.txtExpo1.Name = "txtExpo1"
        Me.txtExpo1.ReadOnly = True
        Me.txtExpo1.Size = New System.Drawing.Size(43, 20)
        Me.txtExpo1.TabIndex = 113
        Me.txtExpo1.TabStop = False
        Me.txtExpo1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExpo1.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExpo1.ValorMax = 999999999.0R
        '
        'lblEstado
        '
        Me.lblEstado.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstado.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEstado.Location = New System.Drawing.Point(602, 247)
        Me.lblEstado.Name = "lblEstado"
        Me.lblEstado.Size = New System.Drawing.Size(391, 24)
        Me.lblEstado.TabIndex = 42
        Me.lblEstado.Text = "ESTADO DEL TRASPASO"
        Me.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblEstado.Visible = False
        '
        'cmdDesbloqueo
        '
        Me.cmdDesbloqueo.Image = CType(resources.GetObject("cmdDesbloqueo.Image"), System.Drawing.Image)
        Me.cmdDesbloqueo.Location = New System.Drawing.Point(872, 16)
        Me.cmdDesbloqueo.Name = "cmdDesbloqueo"
        Me.cmdDesbloqueo.Size = New System.Drawing.Size(28, 28)
        Me.cmdDesbloqueo.TabIndex = 112
        Me.cmdDesbloqueo.TabStop = False
        '
        'cmdBloqueo
        '
        Me.cmdBloqueo.Image = CType(resources.GetObject("cmdBloqueo.Image"), System.Drawing.Image)
        Me.cmdBloqueo.Location = New System.Drawing.Point(896, 16)
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
        Me.lblTraspaso.Location = New System.Drawing.Point(8, 200)
        Me.lblTraspaso.Name = "lblTraspaso"
        Me.lblTraspaso.Size = New System.Drawing.Size(985, 32)
        Me.lblTraspaso.TabIndex = 110
        Me.lblTraspaso.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAviso2
        '
        Me.lblAviso2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAviso2.ForeColor = System.Drawing.Color.Red
        Me.lblAviso2.Location = New System.Drawing.Point(871, 62)
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
        Me.lblAviso1.Location = New System.Drawing.Point(871, 22)
        Me.lblAviso1.Name = "lblAviso1"
        Me.lblAviso1.Size = New System.Drawing.Size(30, 16)
        Me.lblAviso1.TabIndex = 108
        Me.lblAviso1.Text = "!!!!!"
        Me.lblAviso1.Visible = False
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
        Me.txtExis2.Location = New System.Drawing.Point(760, 60)
        Me.txtExis2.MaxLength = 6
        Me.txtExis2.moBlink = Nothing
        Me.txtExis2.Name = "txtExis2"
        Me.txtExis2.ReadOnly = True
        Me.txtExis2.Size = New System.Drawing.Size(65, 20)
        Me.txtExis2.TabIndex = 107
        Me.txtExis2.TabStop = False
        Me.txtExis2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExis2.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExis2.ValorMax = 999999999.0R
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
        Me.txtExis1.Location = New System.Drawing.Point(760, 20)
        Me.txtExis1.MaxLength = 6
        Me.txtExis1.moBlink = Nothing
        Me.txtExis1.Name = "txtExis1"
        Me.txtExis1.ReadOnly = True
        Me.txtExis1.Size = New System.Drawing.Size(65, 20)
        Me.txtExis1.TabIndex = 106
        Me.txtExis1.TabStop = False
        Me.txtExis1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExis1.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtExis1.ValorMax = 999999999.0R
        '
        'txtVendedor
        '
        Me.txtVendedor.AutoSelec = False
        Me.txtVendedor.BackColor = System.Drawing.Color.White
        Me.txtVendedor.Blink = False
        Me.txtVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtVendedor.DesdeCodigo = CType(0, Long)
        Me.txtVendedor.DesdeFecha = New Date(CType(0, Long))
        Me.txtVendedor.HastaCodigo = CType(0, Long)
        Me.txtVendedor.HastaFecha = New Date(CType(0, Long))
        Me.txtVendedor.Location = New System.Drawing.Point(296, 60)
        Me.txtVendedor.MaxLength = 3
        Me.txtVendedor.moBlink = Nothing
        Me.txtVendedor.Name = "txtVendedor"
        Me.txtVendedor.Size = New System.Drawing.Size(34, 20)
        Me.txtVendedor.TabIndex = 4
        Me.txtVendedor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtVendedor.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtVendedor.ValorMax = 999999999.0R
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(240, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 16)
        Me.Label4.TabIndex = 105
        Me.Label4.Text = "VEND.:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(448, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 16)
        Me.Label2.TabIndex = 103
        Me.Label2.Text = "HASTA:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grdLineas
        '
        Me.grdLineas.BackColor = System.Drawing.SystemColors.HighlightText
        Me.grdLineas.Columnas = 0
        Me.grdLineas.Editable = False
        Me.grdLineas.Location = New System.Drawing.Point(8, 247)
        Me.grdLineas.Name = "grdLineas"
        Me.grdLineas.Size = New System.Drawing.Size(588, 326)
        Me.grdLineas.TabIndex = 6
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
        Me.txtDesA2.Location = New System.Drawing.Point(536, 60)
        Me.txtDesA2.MaxLength = 30
        Me.txtDesA2.moBlink = Nothing
        Me.txtDesA2.Name = "txtDesA2"
        Me.txtDesA2.ReadOnly = True
        Me.txtDesA2.Size = New System.Drawing.Size(226, 20)
        Me.txtDesA2.TabIndex = 101
        Me.txtDesA2.TabStop = False
        Me.txtDesA2.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA2.ValorMax = 999999999.0R
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
        Me.txtDesA1.Location = New System.Drawing.Point(536, 20)
        Me.txtDesA1.MaxLength = 30
        Me.txtDesA1.moBlink = Nothing
        Me.txtDesA1.Name = "txtDesA1"
        Me.txtDesA1.ReadOnly = True
        Me.txtDesA1.Size = New System.Drawing.Size(226, 20)
        Me.txtDesA1.TabIndex = 100
        Me.txtDesA1.TabStop = False
        Me.txtDesA1.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesA1.ValorMax = 999999999.0R
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
        Me.txtHasta.Location = New System.Drawing.Point(496, 60)
        Me.txtHasta.MaxLength = 2
        Me.txtHasta.moBlink = Nothing
        Me.txtHasta.Name = "txtHasta"
        Me.txtHasta.Size = New System.Drawing.Size(40, 20)
        Me.txtHasta.TabIndex = 3
        Me.txtHasta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtHasta.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtHasta.ValorMax = 999999999.0R
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
        Me.txtDesde.Location = New System.Drawing.Point(496, 20)
        Me.txtDesde.MaxLength = 2
        Me.txtDesde.moBlink = Nothing
        Me.txtDesde.Name = "txtDesde"
        Me.txtDesde.ReadOnly = True
        Me.txtDesde.Size = New System.Drawing.Size(40, 20)
        Me.txtDesde.TabIndex = 2
        Me.txtDesde.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDesde.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtDesde.ValorMax = 999999999.0R
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(448, 24)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(59, 16)
        Me.Label24.TabIndex = 99
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
        Me.txtOperador.Location = New System.Drawing.Point(88, 60)
        Me.txtOperador.MaxLength = 3
        Me.txtOperador.moBlink = Nothing
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
        Me.txtEstado.Location = New System.Drawing.Point(952, 92)
        Me.txtEstado.MaxLength = 1
        Me.txtEstado.moBlink = Nothing
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
        Me.Label5.Location = New System.Drawing.Point(888, 96)
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
        Me.txtComentario.Location = New System.Drawing.Point(156, 92)
        Me.txtComentario.MaxLength = 500
        Me.txtComentario.moBlink = Nothing
        Me.txtComentario.Multiline = True
        Me.txtComentario.Name = "txtComentario"
        Me.txtComentario.Size = New System.Drawing.Size(713, 97)
        Me.txtComentario.TabIndex = 5
        Me.txtComentario.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtComentario.ValorMax = 999999999.0R
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(40, 96)
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
        Me.txtFecha.Location = New System.Drawing.Point(296, 20)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.moBlink = Nothing
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.ReadOnly = True
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
        Me.txtCodigo.moBlink = Nothing
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(64, 20)
        Me.txtCodigo.TabIndex = 0
        Me.txtCodigo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCodigo.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtCodigo.ValorMax = 999999999.0R
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(240, 24)
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
        'lblDescripcion
        '
        Me.lblDescripcion.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDescripcion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescripcion.Location = New System.Drawing.Point(12, 66)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.Size = New System.Drawing.Size(673, 24)
        Me.lblDescripcion.TabIndex = 40
        Me.lblDescripcion.Text = "Descripcion del proceso."
        Me.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTitulo
        '
        Me.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTitulo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTitulo.Location = New System.Drawing.Point(685, 34)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(328, 56)
        Me.lblTitulo.TabIndex = 38
        Me.lblTitulo.Text = "VISANFER, S.A. - 2003"
        Me.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(12, 722)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(1001, 40)
        Me.lblTeclas.TabIndex = 35
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tabMenu
        '
        Me.tabMenu.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabMenu.Location = New System.Drawing.Point(12, 90)
        Me.tabMenu.Name = "tabMenu"
        Me.tabMenu.Size = New System.Drawing.Size(1001, 24)
        Me.tabMenu.TabIndex = 0
        '
        'tmrAviso
        '
        Me.tmrAviso.Interval = 500
        '
        'lblSolicitudes
        '
        Me.lblSolicitudes.BackColor = System.Drawing.Color.Red
        Me.lblSolicitudes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSolicitudes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSolicitudes.Location = New System.Drawing.Point(512, 66)
        Me.lblSolicitudes.Name = "lblSolicitudes"
        Me.lblSolicitudes.Size = New System.Drawing.Size(173, 24)
        Me.lblSolicitudes.TabIndex = 41
        Me.lblSolicitudes.Text = "SOLICITUDES TRASPASO"
        Me.lblSolicitudes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblSolicitudes.Visible = False
        '
        'tmrSolicitud
        '
        Me.tmrSolicitud.Interval = 500
        '
        'panTitulo
        '
        Me.panTitulo.BackColor = System.Drawing.Color.FromArgb(CType(CType(13, Byte), Integer), CType(CType(93, Byte), Integer), CType(CType(142, Byte), Integer))
        Me.panTitulo.Controls.Add(Me.lblTitle)
        Me.panTitulo.Dock = System.Windows.Forms.DockStyle.Top
        Me.panTitulo.Location = New System.Drawing.Point(0, 0)
        Me.panTitulo.Name = "panTitulo"
        Me.panTitulo.Size = New System.Drawing.Size(1024, 30)
        Me.panTitulo.TabIndex = 42
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(9, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(156, 13)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "NOMBRE DEL OPERARIO"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmTraspasos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1024, 768)
        Me.ControlBox = False
        Me.Controls.Add(Me.panTitulo)
        Me.Controls.Add(Me.lblSolicitudes)
        Me.Controls.Add(Me.tabMenu)
        Me.Controls.Add(Me.lblPrograma)
        Me.Controls.Add(Me.panCampos)
        Me.Controls.Add(Me.lblDescripcion)
        Me.Controls.Add(Me.lblTitulo)
        Me.Controls.Add(Me.lblTeclas)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.MaximumSize = New System.Drawing.Size(1024, 768)
        Me.MinimumSize = New System.Drawing.Size(1024, 768)
        Me.Name = "frmTraspasos"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gestion de Traspasos"
        Me.panCampos.ResumeLayout(False)
        Me.panCampos.PerformLayout()
        Me.grpRecibo.ResumeLayout(False)
        Me.grpRecibo.PerformLayout()
        Me.grpEnvio.ResumeLayout(False)
        Me.grpEnvio.PerformLayout()
        Me.grpFechas.ResumeLayout(False)
        Me.grpFechas.PerformLayout()
        Me.panTitulo.ResumeLayout(False)
        Me.panTitulo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Eventos del Menu "

    Private Sub tabMenu_evtEnter(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.EventArgs) Handles tabMenu.evtEnter
        lblDescripcion.Text = loTab.Tag
    End Sub

    Private Sub tabMenu_evtClick(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.EventArgs) Handles tabMenu.evtClick
        mrEjecutaAccion(loTab.Indice)
    End Sub

    Private Sub tabMenu_evtKeyDown(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tabMenu.evtKeyDown
        Dim loTabTemp As prjToolBar.clsTabPage

        Select Case e.KeyValue
            Case 13     ' Intro
                'mrEjecutaAccion(loTab.Indice)
            Case 27     ' Escape
                'End
                Me.Close()
            Case 67     ' Consulta
                loTabTemp = tabMenu.mcolTabs(1)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(1)
            Case 69     ' Entradas
                loTabTemp = tabMenu.mcolTabs(2)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(2)
            Case 79     ' SOLICITUDES
                loTabTemp = tabMenu.mcolTabs(3)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(3)
            Case 80     ' PENDIENTES
                loTabTemp = tabMenu.mcolTabs(4)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(4)
            Case 83     ' salida
                loTabTemp = tabMenu.mcolTabs(5)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(5)
        End Select

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCrearTraspasoNuevo(ByVal lnEmpresa As Integer,
                                  ByRef loUsuario As clsUsuario, ByRef loTraspaso As clsTraspaso)
        ' cargo los datos del nuevo pedido y lo paso a nuevo registro ******
        moTraspasoAux = loTraspaso
        mbCargaAutomatica = True
        mrCargar(lnEmpresa, loUsuario)

    End Sub

    Public Sub mrConsultaTraspaso(ByVal lnEmpresa As Integer,
                                  ByRef loUsuario As clsUsuario, ByRef loTraspaso As clsTraspaso)
        ' cargo los datos del traspaso y lo consulto
        moTraspasoAux = loTraspaso
        mbConsultaTraspaso = True
        mrCargar(lnEmpresa, loUsuario)

    End Sub

    Public Sub mrCargar(ByVal lnEmpresa As Integer, ByRef loUsuario As clsUsuario)

        ' la llave actual de seguridad *****************************
        gnLlave = 0
        goUsuario = loUsuario
        mnEmpresa = lnEmpresa
        ' **** recupero los datos del profile **************
        modTraspasos.goProfile.mrRecuperaDatos()
        ' **************************************************
        ' ******** seleccion de la empresa contable *****************************
        Dim loEmpresa As New clsEmpresa
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()
        moEmpContable = New clsEmpContable
        moEmpContable.mnCodigo = loEmpresa.mnEmpresaContable
        moEmpContable.mrRecuperaDatos()
        lblTitulo.Text = moEmpContable.msNombre
        ' **************************************************************************

        If Me.MdiParent Is Nothing Then
            Me.ShowDialog()
        Else
            Me.Show()
        End If


    End Sub

    Private Function mfoOperario(ByVal lnOperario As Integer) As clsUsuario

        If mcolLineas Is Nothing Then mcolLineas = New Collection

        Dim loOperario As New clsUsuario
        loOperario.mnCodigo = lnOperario

        On Error Resume Next
        loOperario = mcolOperarios(loOperario.mpsCodigo)
        If Err.Number > 0 Then
            loOperario.mrRecuperaDatos()
            mcolLineas.Add(loOperario, loOperario.mpsCodigo)
        End If
        On Error GoTo 0

        Return loOperario

    End Function

    Private Sub mrEjecutaAccion(ByVal lnNumero As Integer)

        Select Case lnNumero
            Case 1
                mrConsulta()
                txtCodigo.Focus()
            Case 2
                mrEntradas()
            Case 3
                mrSolicitud()
            Case 4
                mrSolicitudPendiente()
            Case 5
                Me.Close()
        End Select

    End Sub

    Private Sub mrSolicitudPendiente()

        If moBusSolTraspasos Is Nothing Then moBusSolTraspasos = New clsBusSolTraspasos
        If moBusSolTraspasos.mcolSolTraspasos Is Nothing Then moBusSolTraspasos.mcolLineas = New Collection
        If moBusSolTraspasos.mcolSolTraspasos.Count > 0 Then
            Dim loPendientes As New frmSolTraspasosPendientes
            loPendientes.mrCargar(moBusSolTraspasos)
        Else
            MsgBox("NO HAY SOLICITUDES DE TRASPASO PENDIENTES.", MsgBoxStyle.Information, "Visanfer .Net")
        End If

    End Sub

    Private Sub mrSolicitud()

        Dim loSolTraspasos As New frmSolTraspasos
        loSolTraspasos.mrCargar(mnEmpresa)

    End Sub

    Private Sub mrEntradas()
        Dim loConfEntradas As New frmConfEntradas

        loConfEntradas.mrCargar(mnEmpresa)

    End Sub

    Private Sub mrPintaFormulario()
        Dim loEmpresa As New prjEmpresas.clsEmpresa
        Dim loTab As prjToolBar.clsTabPage

        ' pongo los datos de la empresa ********
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()
        lblPrograma.Text = "TRASPASOS"
        lblTitulo.Text = loEmpresa.msNombre
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "CONSULTA DE TRASPASOS."
        loTab.Titulo = "(C)ONSULTA"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "CONFIRMACION DE ENTRADAS DE MATERIAL."
        loTab.Titulo = "(E)NTRADAS"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "SOLICITUD DE MATERIAL."
        loTab.Titulo = "S(O)LICITUD"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "SOLICITUD DE MATERIAL PENDIENTES."
        loTab.Titulo = "(P)ENDIENTES"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "SALIDA DEL PROGRAMA."
        loTab.Titulo = "(S)ALIDA"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ***** ahora pinto el tool en pantalla
        tabMenu.mrPintaTool()

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
            Case Keys.J And e.Control = True
                Dim loFormAviso As New prjControl.frmAvisoEvento
                loFormAviso.mrMostrar("TIENES SOLICITUDES DE TRASPASO PENDIENTES", 10, 0, Color.Green)
            Case Keys.S And e.Control = True     'CONTROL + s
                lblSolicitud.Visible = Not lblSolicitud.Visible
            Case Keys.L And e.Control = True     'CONTROL + L
                If mtEstado = EstadoVentana.Consulta And Not moTraspaso.mbEsNuevo Then
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

                        If mbCargaDirecta Then
                            Me.Close()
                        Else
                            mrConsulta()
                            mtEstado = EstadoVentana.Salida
                            mrLimpiaFormulario()
                            If mbCargaAutomatica Or mbConsultaTraspaso Then
                                Me.Close()
                            Else
                                tabMenu.evtFocus()
                            End If
                        End If

                    End If
                End If

            Case Keys.Enter
                Select Case lsControl
                    Case "grdLineas"
                        If grdLineas.mnCol = 0 And mtEstado <> EstadoVentana.Consulta Then
                            If Not mfbCargaArticulo("") Then
                                ' no hago nada
                            End If
                        End If
                        If grdLineas.mnCol = 2 And mtEstado <> EstadoVentana.Consulta Then
                            Dim lnFila As Integer = grdLineas.mnRow

                            Dim lnCantidad As Double = mfnDouble(grdLineas.marMemoria(2, lnFila))
                            Dim lnArticulo As Long = mfnCodigoArticulo(grdLineas.marMemoria(0, lnFila))
                            Dim lnDetalle As Integer = mfnCodigoDetalle(grdLineas.marMemoria(0, lnFila))

                            mrMiraExistencias(lnArticulo, lnDetalle, lnCantidad)

                            If lnCantidad = 0 Then
                                grdLineas.mrPonFoco(2, lnFila)
                            Else
                                grdLineas.marMemoria(2, lnFila) = Format(lnCantidad, "#,##0.00")
                                If lnFila = grdLineas.mnFilasDatos - 1 Then
                                    grdLineas.mrAñadirFila()
                                End If
                                mrPendiente(lnFila, "N")
                                grdLineas.mrRefrescaGrid()
                            End If
                        End If
                    Case "txtCodigo"
                        If lblSolicitud.Visible Then
                            mrCargaSolicitud()
                        Else
                            mrCargaTraspaso()
                        End If
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
            Case Keys.F6
                mrMarcarRecibido()
            Case Keys.F8
                mrMarcarEnviado()
            Case Keys.F9
                Select Case lsControl
                    Case "txtCodigo"
                        If mtEstado = EstadoVentana.Consulta Then mrBuscaTraspaso()
                    Case "grdLineas"
                        If (mtEstado <> EstadoVentana.Consulta) And (grdLineas.mnCol = 0) Then mrBuscaArticulosNuevo()
                End Select
        End Select

    End Sub

    Private Sub mrBuscaArticulosNuevo()

        Dim loBuscadorArticulos As New frmBuscadorArticulos
        loBuscadorArticulos.mnEmpresa = mnEmpresa
        loBuscadorArticulos.mbMultiple = True
        loBuscadorArticulos.mrCargar()

        If loBuscadorArticulos.mnSeleccionados > 0 Then

            Dim lnFila As Integer
            Dim lnInicio As Integer
            Dim loArticulo As clsArticulo
            Dim lnContador As Integer = 0

            lnFila = grdLineas.mnRow
            lnInicio = lnFila
            For Each loArticulo In loBuscadorArticulos.mcolSeleccionados
                If lnContador > 0 Then grdLineas.mrAñadirFila()
                lnContador = lnContador + 1

                If loArticulo.mnDetalle > 0 Then
                    grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo & "." & loArticulo.mnDetalle
                Else
                    grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo
                End If
                grdLineas.marMemoria(1, lnFila) = loArticulo.msDescripcion
                grdLineas.marMemoria(2, lnFila) = "0"
                mrPendiente(lnFila, "S")
                'grdLineas.mrAñadirFila()
                lnFila = lnFila + 1
            Next
            If lnContador = 0 Then
                If moArticulo.mnDetalle > 0 Then
                    grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo & "." & moArticulo.mnDetalle
                Else
                    grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo
                End If
                grdLineas.marMemoria(1, lnFila) = moArticulo.msDescripcion
                grdLineas.marMemoria(2, lnFila) = "0"
                mfbCargaArticulo(moArticulo.msDescripcion)
            End If
            grdLineas.mrRefrescaGrid()
            grdLineas.mrPonFoco(0, lnInicio + 1)

        End If

    End Sub

    Private Sub mrMarcarRecibido()

        If Not moTraspaso.mbEsNuevo Then
            If goProfile.mnAlmacen <> moTraspaso.mnHasta Then
                MsgBox("EL TRASPASO SOLO SE PUEDE MARCAR COMO RECIBIDO DESDE EL ALMACEN DE DESTINO", MsgBoxStyle.Information, "Visanfer.Net")
                Exit Sub
            End If
            If moTraspaso.msEstadoEnvio = "R" Then
                MsgBox("ESTE TRASPASO YA ESTA MARCADO COMO RECIBIDO", MsgBoxStyle.Information, "Visanfer.Net")
                Exit Sub
            End If

            moTraspaso.msEstadoEnvio = "R"
            moTraspaso.mnOperarioRecepcion = goUsuario.mnCodigo
            moTraspaso.mdFechaRecepcion = Now
            moTraspaso.mrGrabaDatos()

            mrConsulta()
            mrMoverCampos(1)

            MsgBox("TRASPASO MARCADO COMO RECIBIDO.", MsgBoxStyle.Information, "Visanfer.Net")
        End If

    End Sub

    Private Sub mrMarcarEnviado()

        If Not moTraspaso.mbEsNuevo Then
            If goProfile.mnAlmacen <> moTraspaso.mnDesde Then
                MsgBox("EL TRASPASO SOLO SE PUEDE MARCAR COMO ENVIADO DESDE EL ALMACEN DE ORIGEN", MsgBoxStyle.Information, "Visanfer.Net")
                Exit Sub
            End If
            If moTraspaso.msEstadoEnvio = "E" Then
                MsgBox("ESTE TRASPASO YA ESTA MARCADO COMO ENVIADO", MsgBoxStyle.Information, "Visanfer.Net")
                Exit Sub
            End If
            If moTraspaso.msEstadoEnvio = "R" Then
                MsgBox("ESTE TRASPASO YA ESTA MARCADO COMO RECIBIDO", MsgBoxStyle.Information, "Visanfer.Net")
                Exit Sub
            End If
            moTraspaso.msEstadoEnvio = "E"
            moTraspaso.mnOperarioEnvio = goUsuario.mnCodigo
            moTraspaso.mdFechaEnvio = Now
            moTraspaso.mrGrabaDatos()

            mrConsulta()
            mrMoverCampos(1)

            MsgBox("TRASPASO MARCADO COMO ENVIADO.", MsgBoxStyle.Information, "Visanfer.Net")
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
                loArticulo.mnDetalle = loAgrArt.mnDetalle
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
                        loAviso.msMensaje = "ATENCION: EN ESTE MOMENTO NO HAY EXISTENCIAS DE ESTE " & vbCrLf &
                                            "ARTICULO. REVISE EL ARTICULO Y LAS EXISTENCIAS DEL MISMO " & vbCrLf &
                                            "O REVISE EL ALMACEN DE SALIDA. SI TIENE ALGUNA INCIDENCIA " & vbCrLf &
                                            "AVISE A EMILIO O SOLANO. " & vbCrLf
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
                loAviso.msMensaje = "ATENCION: EN ESTE MOMENTO NO HAY EXISTENCIAS DE ESTE " & vbCrLf &
                                    "ARTICULO. REVISE EL ARTICULO Y LAS EXISTENCIAS DEL MISMO " & vbCrLf &
                                    "O REVISE EL ALMACEN DE SALIDA. SI TIENE ALGUNA INCIDENCIA " & vbCrLf &
                                    "AVISE A EMILIO O SOLANO. " & vbCrLf
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
                        loAviso.msMensaje = "ATENCION: CASI SEGURO QUE SE ESTA EQUIVOCANDO DE ALMACEN. " & vbCrLf &
                                            "REVISE EL ALMACEN DE VENTA DE ESTE ALBARAN. EN CASO   " & vbCrLf &
                                            "CONTRARIO REVISE LAS EXISTENCIAS DE ESTE ARTICULO EN SU" & vbCrLf &
                                            "ALMACEN SUMANDO ADEMAS LAS QUE TIENE EN EXPOSICION. " &
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

    Private Sub mrBuscaTraspaso()
        moTraspaso = New clsTraspaso
        moTraspaso.mnEmpresa = mnEmpresa
        moTraspaso.mrBuscaTraspaso()
    End Sub

    Private Sub moTraspaso_evtBusTraspaso() Handles moTraspaso.evtBusTraspaso
        mrLimpiaFormulario()
        lblTeclas.Text = " CTRL-M Modificacion de Datos              CTRL-L Ver Lineas               F1-Alta Nuevo    " &
                         "            CTRL-P Imprimir                 CTRL+S Carga Solicitud                ESC-Salida"
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
        moTraspaso = New clsTraspaso
        txtCodigo.Text = "0"
        txtFecha.Text = Format(Now, "dd/MM/yyyy")
        txtOperador.Text = goUsuario.mnCodigo
        txtEstado.Text = "N"
        txtDesde.Text = goProfile.mnAlmacen
        mrCargaAlmacen(mfnLong(txtDesde.Text), 1, Nothing)
        lblTeclas.BackColor = Color.GreenYellow
        lblTeclas.Text = "F5-GRABACION                                ESC-SALIDA"
        lblPrograma.Text = "TRASPASOS - NUEVO REGISTRO"
        lblPrograma.BackColor = Color.GreenYellow
        panCampos.Enabled = True
        grdLineas.Editable = True
        txtDesde.Focus()
        'End If

    End Sub

    Private Sub mrCargaSolicitud()

        Dim lnCodigo As Long

        lnCodigo = mfnInteger(txtCodigo.Text)
        If lnCodigo > 0 Then
            Dim loSolTraspaso As New clsSolTraspaso
            loSolTraspaso.mnEmpresa = mnEmpresa
            loSolTraspaso.mnCodigo = lnCodigo
            loSolTraspaso.mrRecuperaDatos()
            If loSolTraspaso.mbEsNuevo Then
                MsgBox("Solicitud no Encontrada", MsgBoxStyle.Exclamation, "Gestion Visanfer")
                txtCodigo.Focus()
            Else
                loSolTraspaso.mrRecuperaLineas()
                mrLimpiaFormulario()
                mrNuevoRegistro()
                mrMoverCamposSolicitud(loSolTraspaso)
                txtHasta.Focus()
            End If
        Else
            txtCodigo.SelectAll()
        End If

    End Sub

    Private Sub mrCargaTraspaso()
        Dim lnCodigo As Long

        lnCodigo = mfnInteger(txtCodigo.Text)
        If lnCodigo > 0 Then
            If mtEstado <> EstadoVentana.Mantenimiento Then
                moTraspaso = New clsTraspaso
                moTraspaso.mnEmpresa = mnEmpresa
                moTraspaso.mnCodigo = lnCodigo
                moTraspaso.mrRecuperaDatos()
                mrLimpiaFormulario()
                If moTraspaso.mbEsNuevo Then
                    MsgBox("Traspaso no Encontrado", MsgBoxStyle.Exclamation, "Gestion Visanfer")
                    txtCodigo.Focus()
                Else
                    moTraspaso.mrRecuperaLineas()
                    mrMoverCampos(1)
                    lblTeclas.Text = " CTRL-M Modificacion de Datos              CTRL-L Ver Lineas               F1-Alta Nuevo    " &
                                     "            CTRL-P Imprimir                 CTRL+S Carga Solicitud                ESC-Salida"
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
        lblEstado.Visible = False
        grpFechas.Visible = False
        grpEnvio.Visible = False
        grdLineas.mrClear(prjGrid.ctlGrid.TipoBorrado.Contenido)
        lblAviso1.Visible = False
        lblAviso2.Visible = False
        lblSolicitud.Visible = False
        txtExis1.BackColor = Color.White
        txtExis2.BackColor = Color.White
        lblTraspaso.Text = ""

    End Sub

    Private Sub mrCargaAlmacen(ByVal lnCodigo As Integer, ByVal lnNumero As Integer,
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

                If loArticulo.mnDetalle > 0 Then
                    grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo & "." & loArticulo.mnDetalle
                Else
                    grdLineas.marMemoria(0, lnFila) = loArticulo.mnCodigo
                End If
                grdLineas.marMemoria(1, lnFila) = loArticulo.msDescripcion
                grdLineas.marMemoria(2, lnFila) = "0"
                mrPendiente(lnFila, "S")
                'grdLineas.mrAñadirFila()
                lnFila = lnFila + 1
            End If
        Next
        If lnContador = 0 Then
            If moArticulo.mnDetalle > 0 Then
                grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo & "." & moArticulo.mnDetalle
            Else
                grdLineas.marMemoria(0, lnFila) = moArticulo.mnCodigo
            End If
            grdLineas.marMemoria(1, lnFila) = moArticulo.msDescripcion
            grdLineas.marMemoria(2, lnFila) = "0"
            mfbCargaArticulo(moArticulo.msDescripcion)
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

        If mtEstado = EstadoVentana.NuevoRegistro Or
           mtEstado = EstadoVentana.Mantenimiento Then
            ' compruebo que los campos obligatorios estan cumplimentados
            If mfbObligatorios(e) Then
                Dim loTraspasoLin As clsTraspasoLin

                mrMoverCampos(2)    ' paso los valores al objeto
                If mcolLineas.Count > 0 Then
                    Cursor = Cursors.WaitCursor
                    Dim lbError As Boolean
                    If Not mtEstado = EstadoVentana.NuevoRegistro Then
                        moTraspaso.mrGrabaDatos() ' grabo el contenido
                        ' ahora borro las lineas y las grabo otra vez
                        moTraspaso.mrBorraLineas()
                        ' ahora grabo las lineas ******************
                        For Each loTraspasoLin In mcolLineas
                            loTraspasoLin.mnCodigo = moTraspaso.mnCodigo
                            loTraspasoLin.mrGrabaDatos()
                        Next
                        moTraspaso.mrRecuperaLineas()
                    End If
                    If Not lbError Then ' si la grabacion es correcta
                        If mtEstado = EstadoVentana.NuevoRegistro Then

                            ' si es un registro nuevo recupero su codigo
                            ' le pongo los valores de la hora a la que se graba
                            moTraspaso.msEstadoEnvio = "G"
                            moTraspaso.mdFechaGrabacion = Now

                            moTraspaso.mnEmpresa = mnEmpresa
                            moTraspaso.mrNuevoCodigo()
                            moTraspaso.mbEsNuevo = True
                            ' ahora grabo las lineas ******************
                            For Each loTraspasoLin In mcolLineas
                                loTraspasoLin.mnCodigo = moTraspaso.mnCodigo
                                loTraspasoLin.mrGrabaDatos()
                            Next
                            ' *****************************************************
                            txtCodigo.Text = moTraspaso.mnCodigo
                            moTraspaso.mrGrabaDatos() ' grabo el contenido
                            moTraspaso.mrRecuperaLineas()

                            ' **** ahora actualizo las existencias del traspaso **********
                            moTraspaso.mrActualizaTraspaso()

                            '' le pongo la captura de eventos porque algunas veces lo hace mal hasta aqui ******
                            'System.Windows.Forms.Application.DoEvents()

                            '' ahora lo que hago es lanzar la impresion del pedido ********
                            'Dim lsResult As MsgBoxResult = MsgBox("¿DESEA IMPRIMIR EL TRASPASO?",
                            '            MsgBoxStyle.Information + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Visanfer .Net")
                            'If lsResult = vbYes Then mrImprimirRemoto()

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

        If moTraspaso Is Nothing Or moTraspaso.mnCodigo = 0 Then
            MsgBox("¡¡¡ATENCION!!!, DEBES GRABAR PRIMERO.", MsgBoxStyle.Exclamation, "Visanfer.Net")
            Exit Sub
        End If

        ' reviso si el traspaso esta abierto por el proceso automatico de gestion
        Dim loTraspasoTemp As New clsTraspasoTemp
        loTraspasoTemp.mnEmpresa = mnEmpresa
        loTraspasoTemp.mnTraspaso = moTraspaso.mnCodigo
        loTraspasoTemp.mrRecuperaDatos()
        ' si existe traspaso temporal y esta abierto, pregunto si lo cierro
        If (Not loTraspasoTemp.mbEsNuevo) AndAlso (loTraspasoTemp.mnAbierto = 1) Then
            Dim lsResultado As MsgBoxResult = MsgBox("ESTE TRASPASO ESTA ABIERTO ¿LO CIERRO?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Visanfer.Net")
            If lsResultado = MsgBoxResult.Yes Then
                loTraspasoTemp.mnAbierto = 0
                loTraspasoTemp.mrGrabaDatos()
            End If
        End If

        ' imrprimo los listados en ambas impresoras ******************
        mbCompleto = False
        mrSelImpresora()

    End Sub

    Private Sub mrImprimirRemoto()
        ' imrprimo los listados en ambas impresoras ******************
        mbCompleto = True
        mrSelImpresora()

    End Sub

    Private Sub mrSelImpresora()
        ' selecciona la impresora ************************************

        moSelImpresora = New prjControl.frmSelImpresora
        If mbCompleto Then
            moSelImpresora.mnImpresora = goProfile.mnImpresora    'moAlmacen1.mnImpresoraTraspasos
            moSelImpresora.msDestino = "I"
            moSelImpresora.mnCopias = 1
            moSelImpresora.msPapel = "A4"
            moSelImpresora.mrSeleccionar("IMP. TRASPASOS - ORIGEN")

            ' lo anulo por que no lo quieren
            'moSelImpresora.mnImpresora = moAlmacen2.mnImpresoraTraspasos
            'moSelImpresora.msDestino = "I"
            'moSelImpresora.mnCopias = 1
            'moSelImpresora.mrSeleccionar("IMP. TRASPASOS - DESTINO")
        Else
            moSelImpresora.mnImpresora = goProfile.mnImpresora
            moSelImpresora.msDestino = "I"
            moSelImpresora.mnCopias = 1
            moSelImpresora.msPapel = "A4"
            moSelImpresora.mrSeleccionar("IMP. TRASPASOS - COPIA")
        End If

    End Sub

    Private Sub mrImprimirRpt(ByVal lnImpresora As Integer, ByVal lnCopias As Integer,
                              ByVal lsSalida As String, ByVal lsPapel As String)

        Dim loLinea As clsTraspasoLin
        Dim loTablaCabecera As DataTable
        Dim loTablaLineas As DataTable
        Dim loTablaPie As DataTable

        Cursor = Cursors.WaitCursor
        panCampos.Enabled = False
        ' RELLENO LOS DATASET CON LOS DATOS DE LAS CAJAS

        moDataTraspaso = New dtsTraspaso
        ' primero vacio las tablas ***************
        loTablaCabecera = moDataTraspaso.Tables("Cabecera")
        loTablaCabecera.Rows.Clear()
        loTablaLineas = moDataTraspaso.Tables("Lineas")
        loTablaLineas.Rows.Clear()
        loTablaPie = moDataTraspaso.Tables("Pie")
        loTablaPie.Rows.Clear()

        ' relleno las tablas del dataset con los valores del listado ********
        Dim loRow As DataRow

        ' ahora paso los valores a las tablas
        loRow = loTablaCabecera.NewRow
        loRow("Codigo") = moTraspaso.mnCodigo
        loRow("Fecha") = moTraspaso.mdFecha
        loRow("Tipo") = "TRASPASO"
        If mbCompleto Then
            loRow("Titulo") = "TRASPASO"
        Else
            loRow("Titulo") = "T R A S P A S O - C O P I A"
        End If
        ' **************************************
        Dim loAlmacen As clsAlmacen
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moTraspaso.mnDesde
        loAlmacen.mrRecuperaDatos(False)
        loRow("Desde") = loAlmacen.msNombre
        ' **************************************
        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = moTraspaso.mnHasta
        loAlmacen.mrRecuperaDatos(False)
        loRow("Hasta") = loAlmacen.msNombre
        ' **************************************
        loRow("Vendedor") = moTraspaso.mnVendedor
        loRow("Operario") = moTraspaso.mnOperario
        loRow("Observaciones") = moTraspaso.msObservaciones
        loTablaCabecera.Rows.Add(loRow)

        ' ahora relleno el pie ***********************
        loRow = loTablaPie.NewRow
        loRow("Codigo") = moTraspaso.mnCodigo
        loRow("Empresa") = "Visanfer, S.A."

        Dim loDireccion As New prjPedProveedores.clsPedAlmacen
        loDireccion.mnEmpresa = mnEmpresa
        loDireccion.mnCodigo = moTraspaso.mnDesde
        loDireccion.mrRecuperaDatos()
        loRow("DesdeDireccion") = "ORIGEN DE LA CARGA: " & loDireccion.msDireccion & " - " & loDireccion.msCodigoPostal & " " & loDireccion.msPoblacion & " (" & loDireccion.msProvincia & ")"

        loDireccion.mnCodigo = moTraspaso.mnHasta
        loDireccion.mrRecuperaDatos()
        loRow("HastaDireccion") = "DESTINO DE LA CARGA: " & loDireccion.msDireccion & " - " & loDireccion.msCodigoPostal & " " & loDireccion.msPoblacion & " (" & loDireccion.msProvincia & ")"
        loTablaPie.Rows.Add(loRow)
        ' ********************************************

        Dim lnI As Integer = 1
        For Each loLinea In moTraspaso.mcolLineas
            ' añado un registro a la tabla de lineas
            loRow = loTablaLineas.NewRow()
            loRow("Codigo") = loLinea.mnCodigo
            loRow("Linea") = lnI
            loRow("Articulo") = loLinea.mnArticulo
            loRow("Detalle") = loLinea.mnDetalle
            loRow("Cantidad") = loLinea.mnCantidad
            loRow("Descripcion") = loLinea.msDescripcion

            Dim loEan13 As New clsEan13
            loEan13.mnArticulo = loLinea.mnArticulo
            loEan13.mnDetalle = loLinea.mnDetalle
            loEan13.mrRecuperaDatos(2)

            If loEan13.mbEsNuevo Then
                If loLinea.mnDetalle > 0 Then
                    Dim lnCodigo As Long = loLinea.mnArticulo * 1000 + loLinea.mnDetalle
                    Dim lsCodigo As String = "99" & Format(lnCodigo, "0000000000")
                    lsCodigo = lsCodigo ' & ChecksumEAN13(lsCodigo)

                    loRow("codigobarras") = mfsEan13(lsCodigo)
                Else
                    Dim lnCodigo As Long = loLinea.mnArticulo * 1000
                    Dim lsCodigo As String = "99" & Format(lnCodigo, "0000000000")
                    lsCodigo = lsCodigo '& ChecksumEAN13(lsCodigo)

                    loRow("codigobarras") = mfsEan13(lsCodigo)
                End If
            Else
                'loRow("codigobarras") = "00"
                loRow("codigobarras") = ""
            End If

            loTablaLineas.Rows.Add(loRow)
            lnI = lnI + 1
        Next

        Dim oReport1 As New Microsoft.Reporting.WinForms.ReportDataSource("Cabecera", loTablaCabecera)
        Dim oReport2 As New Microsoft.Reporting.WinForms.ReportDataSource("Lineas", loTablaLineas)
        Dim oReport3 As New Microsoft.Reporting.WinForms.ReportDataSource("Pie", loTablaPie)

        Dim loAsembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim lsReport As IO.Stream = loAsembly.GetManifestResourceStream("prjTraspasos.rdlcTraspaso.rdlc")

        If moSelImpresora.msDestino = "V" Then

            Dim loVisor As New frmVisorReport
            'loVisor.moReport.LocalReport.ReportEmbeddedResource = "prjTraspasos.rdlcTraspaso.rdlc"
            loVisor.moReport.LocalReport.LoadReportDefinition(lsReport)
            loVisor.moReport.LocalReport.DataSources.Clear()
            loVisor.moReport.LocalReport.DataSources.Add(oReport1)
            loVisor.moReport.LocalReport.DataSources.Add(oReport2)
            loVisor.moReport.LocalReport.DataSources.Add(oReport3)
            loVisor.moReport.RefreshReport()
            loVisor.mrVisualizar("A4", False, moImpresora.msCola)

        Else

            Dim loReport As New Microsoft.Reporting.WinForms.LocalReport
            'loReport.ReportEmbeddedResource = "prjTraspasos.rdlcTraspaso.rdlc"
            loReport.LoadReportDefinition(lsReport)
            loReport.DataSources.Clear()
            loReport.DataSources.Add(oReport1)
            loReport.DataSources.Add(oReport2)
            loReport.DataSources.Add(oReport3)
            loReport.Refresh()
            print_microsoft_report(loReport, "A4", False, moImpresora.msCola)

        End If

        panCampos.Enabled = True
        Cursor = Cursors.Default

    End Sub

    Private Sub moSelImpresora_evtBusImpresora() Handles moSelImpresora.evtBusImpresora
        moImpresora = New prjPrinterNet.clsImpresora
        moImpresora.mnEmpresa = mnEmpresa
        moImpresora.mnCodigo = moSelImpresora.mnImpresora
        moImpresora.mrRecuperaDatos()

        If moSelImpresora.msPapel <> "" Then moImpresora.msPapel = moSelImpresora.msPapel
        mrImprimirRpt(moImpresora.mnCodigo, moSelImpresora.mnCopias, moSelImpresora.msDestino, moSelImpresora.msPapel)

    End Sub

    Private Sub mrCargaLineas()
        Dim lnI As Integer
        Dim lnJ As Integer
        Dim lsTipo As String
        Dim loTraspasoLin As clsTraspasoLin

        mcolLineas = New Collection
        lnJ = 1
        For lnI = 0 To grdLineas.mnFilasDatos - 1
            ' primero veo que tipo de linea es ********************
            lsTipo = grdLineas.marMemoria(4, lnI)
            If lsTipo <> "B" Then
                loTraspasoLin = New clsTraspasoLin
                loTraspasoLin.mnEmpresa = mnEmpresa
                'loTraspasoLin.mnCodigo = moTraspaso.mnCodigo
                loTraspasoLin.mnLinea = lnJ
                loTraspasoLin.mnArticulo = mfnCodigoArticulo(grdLineas.marMemoria(0, lnI))
                loTraspasoLin.mnDetalle = mfnCodigoDetalle(grdLineas.marMemoria(0, lnI))
                loTraspasoLin.msDescripcion = Trim(grdLineas.marMemoria(1, lnI))
                loTraspasoLin.mnCantidad = mfnDouble(grdLineas.marMemoria(2, lnI))
                ' el estado lo pongo siempre pendiente de actualizar ************
                If moAlmacen2.mnCodigo = 8 Then
                    loTraspasoLin.msEstado = "A"
                Else
                    loTraspasoLin.msEstado = "N"
                End If
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

        'If mfnInteger(txtOperador.Text) = 0 Then
        '    MsgBox("DEBE PONER ALGUN OPERADOR.", MsgBoxStyle.Exclamation, "Visanfer .Net")
        '    txtOperador.SelectAll()
        '    txtOperador.Focus()
        '    e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
        '    Return False
        'End If

        If mfnInteger(txtVendedor.Text) = 0 Then
            MsgBox("DEBE PONER ALGUN VENDEDOR.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtVendedor.SelectAll()
            txtVendedor.Focus()
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

        ' ME LO PIDIO DOMINGO , PERO 15 MINUTOS MAS TARDE ME DIJO QUE NO, LE IBA A CREAR UN LLAVE NUEVA **********
        'Dim lnAlmacenDestino As Integer = mfnInteger(txtHasta.Text)
        'If lnAlmacenDestino = 9 Then

        '    If Not goUsuario.mfbAccesoPermitido(140, True) Then
        '        txtHasta.SelectAll()
        '        txtHasta.Focus()
        '        e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
        '        Return False
        '    End If

        'End If

        If Trim(txtComentario.Text) = "" Then
            MsgBox("DEBE PONER QUIEN HA PEDIDO EL MATERIAL.", MsgBoxStyle.Exclamation, "Visanfer .Net")
            txtComentario.SelectAll()
            txtComentario.Focus()
            e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
            Return False
        End If

        ' ahora recorro todo el grid buscando lineas incompletas
        For lnI As Integer = 0 To grdLineas.mnFilasDatos - 1
            For lnJ As Integer = 0 To 4
                If grdLineas.ColorCelda(lnJ, lnI).mnForeColor.Name = "Red" AndAlso grdLineas.marMemoria(lnJ, lnI) <> "" Then
                    MsgBox("HAY LINEAS CON ERRORES, REVISALAS.", MsgBoxStyle.Exclamation, "Visanfer .Net")
                    grdLineas.Focus()
                    e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
                    Return False
                End If
            Next

        Next

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
        lblTeclas.Text = " CTRL-M Modificacion de Datos              CTRL-L Ver Lineas               F1-Alta Nuevo    " &
                         "            CTRL-P Imprimir                 CTRL+S Carga Solicitud                ESC-Salida"
        lblPrograma.Text = "TRASPASOS"
        lblPrograma.BackColor = Color.FromKnownColor(KnownColor.ActiveBorder)

    End Sub

    Private Sub mrMantenimiento()

        'mrAsignaconexion(gconADO)
        'If goUsuario.mfbAccesoPermitido(9, True) Then
        If mtEstado = EstadoVentana.Mantenimiento Then
            mrConsulta()
        Else
            If moTraspaso.msEstado = "A" Then
                MsgBox("Traspaso Actualizado.", MsgBoxStyle.Exclamation, "Gestion Visanfer")
            Else
                mtEstado = EstadoVentana.Mantenimiento ' entra en modo mantenimiento
                grdLineas.Editable = True
                lblTeclas.BackColor = Color.Tomato
                lblTeclas.ForeColor = Color.White
                lblTeclas.Text = " F1.-INSERTA LINEA          F2.-BORRA LINEA      " &
                                 "      F5-.GRABA             ESC-.SALIDA"
                lblPrograma.Text = "TRASPASOS - MANTENIMIENTO"
                lblPrograma.BackColor = Color.Tomato
            End If
        End If
        'End If

    End Sub

    Private Sub mrSolicitudesPendientes()

        ' miro si hay solicitudes de traspaso pendientes ***********
        moBusSolTraspasos = New clsBusSolTraspasos
        moBusSolTraspasos.mnEmpresa = mnEmpresa
        moBusSolTraspasos.mnDesde = goProfile.mnAlmacen
        moBusSolTraspasos.msEstado = "P"
        moBusSolTraspasos.mdDesdeFecha = "01/01/1900"
        moBusSolTraspasos.mdHastaFecha = "31/12/9999"
        moBusSolTraspasos.mrBuscaSolTraspasos()
        If moBusSolTraspasos.mcolSolTraspasos.Count > 0 Then
            lblSolicitudes.Visible = True
            tmrSolicitud.Enabled = True
        Else
            tmrSolicitud.Enabled = False
            lblSolicitudes.Visible = False
        End If

    End Sub

    Private Sub mrPreparaGrid()

        grdLineas.Columnas = 7

        grdLineas.marTitulos(0).Texto = "ARTIC."
        grdLineas.marTitulos(0).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(0).Ancho = 70
        grdLineas.marTitulos(0).Longitud = 13
        grdLineas.marTitulos(0).Editable = True

        grdLineas.marTitulos(1).Texto = "DESCRIPCION"
        grdLineas.marTitulos(1).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(1).Ancho = 412
        grdLineas.marTitulos(1).Longitud = 100
        grdLineas.marTitulos(1).Editable = True

        grdLineas.marTitulos(2).Texto = "CTD."
        grdLineas.marTitulos(2).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(2).Ancho = 90
        grdLineas.marTitulos(2).Longitud = 10
        grdLineas.marTitulos(2).Editable = True

        grdLineas.marTitulos(3).Texto = "mpsCodigo"
        grdLineas.marTitulos(3).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(3).Ancho = 0
        grdLineas.marTitulos(3).Longitud = 15
        grdLineas.marTitulos(3).Editable = False

        grdLineas.marTitulos(4).Texto = "marcas"
        grdLineas.marTitulos(4).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(4).Ancho = 0
        grdLineas.marTitulos(4).Longitud = 1
        grdLineas.marTitulos(4).Editable = False

        grdLineas.marTitulos(5).Texto = "A"
        grdLineas.marTitulos(5).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(5).Ancho = 0
        grdLineas.marTitulos(5).Longitud = 6
        grdLineas.marTitulos(5).Editable = False

        grdLineas.marTitulos(6).Texto = "CODIGOVIEJO"
        grdLineas.marTitulos(6).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(6).Ancho = 0
        grdLineas.marTitulos(6).Longitud = 6
        grdLineas.marTitulos(6).Editable = False

        'grdLineas.mnAjustarColumna = 2
        grdLineas.mrPintaGrid()

    End Sub

    Private Sub mrMoverCamposSolicitud(ByVal loSolicitud As clsSolTraspaso)
        Dim loLinea As clsSolTraspasoLin
        Dim lnI As Integer

        ' primero los campos de la cabecera **********
        txtFecha.Text = Format(Now, "dd/MM/yyyy")
        txtEstado.Text = loSolicitud.msEstado
        txtComentario.Text = loSolicitud.msObservaciones
        ' ****************
        txtDesde.Text = loSolicitud.mnDesde
        moAlmacen1 = New clsAlmacen
        moAlmacen1.mnEmpresa = mnEmpresa
        moAlmacen1.mnCodigo = loSolicitud.mnDesde
        moAlmacen1.mrRecuperaDatos(False)
        txtDesA1.Text = moAlmacen1.msNombre
        ' ****************
        txtHasta.Text = loSolicitud.mnHasta
        moAlmacen2 = New clsAlmacen
        moAlmacen2.mnEmpresa = mnEmpresa
        moAlmacen2.mnCodigo = loSolicitud.mnHasta
        moAlmacen2.mrRecuperaDatos(False)
        txtDesA2.Text = moAlmacen2.msNombre
        lblTraspaso.Text = txtDesA1.Text & "   ····>   " & txtDesA2.Text
        ' ****************
        txtOperador.Text = loSolicitud.mnOperario
        txtVendedor.Text = ""
        ' despues los campos de las lineas ***********
        lnI = 0
        For Each loLinea In loSolicitud.mcolLineas
            grdLineas.mrAñadirFila()
            grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo
            If loLinea.mnDetalle > 0 Then grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo & "." & loLinea.mnDetalle
            grdLineas.marMemoria(1, lnI) = loLinea.msDescripcion
            grdLineas.marMemoria(2, lnI) = Format(loLinea.mnCantidad, "#,##0.00")
            grdLineas.marMemoria(3, lnI) = loLinea.mpsCodigo
            grdLineas.marMemoria(4, lnI) = ""
            grdLineas.marMemoria(5, lnI) = loLinea.mnArticulo
            lnI = lnI + 1
        Next
        ' despues añado una linea nueva que es la que pongo para meter nuevos datos
        grdLineas.mrAñadirFila()
        grdLineas.mrRefrescaGrid()

    End Sub


    Private Sub mrMoverCampos(ByVal lnTipo As Integer)
        Dim loLinea As clsTraspasoLin
        Dim lnI As Integer

        If lnTipo = 1 Then
            ' primero los campos de la cabecera **********
            txtCodigo.Text = moTraspaso.mnCodigo
            txtFecha.Text = Format(moTraspaso.mdFecha, "dd/MM/yyyy")
            txtEstado.Text = moTraspaso.msEstado
            txtComentario.Text = moTraspaso.msObservaciones
            ' ****************
            txtDesde.Text = moTraspaso.mnDesde
            moAlmacen1 = New clsAlmacen
            moAlmacen1.mnEmpresa = mnEmpresa
            moAlmacen1.mnCodigo = moTraspaso.mnDesde
            moAlmacen1.mrRecuperaDatos(False)
            txtDesA1.Text = moAlmacen1.msNombre
            ' ****************
            txtHasta.Text = moTraspaso.mnHasta
            moAlmacen2 = New clsAlmacen
            moAlmacen2.mnEmpresa = mnEmpresa
            moAlmacen2.mnCodigo = moTraspaso.mnHasta
            moAlmacen2.mrRecuperaDatos(False)
            txtDesA2.Text = moAlmacen2.msNombre
            lblTraspaso.Text = txtDesA1.Text & "   ····>   " & txtDesA2.Text
            ' ****************
            txtOperador.Text = moTraspaso.mnOperario
            txtVendedor.Text = moTraspaso.mnVendedor

            ' pongo los estados del envio ***************

            Dim loOperario As clsUsuario = mfoOperario(moTraspaso.mnOperario)
            txtFechaGrabacion.Text = Format(moTraspaso.mdFechaGrabacion, "dd/MM/yyyy HH:mm:ss")
            txtOperarioGraba.Text = loOperario.mnCodigo & " - " & loOperario.msNombre

            If moTraspaso.mnOperarioEnvio > 0 Then
                txtFechaEnvio.Text = Format(moTraspaso.mdFechaEnvio, "dd/MM/yyyy HH:mm:ss")
                loOperario = mfoOperario(moTraspaso.mnOperarioEnvio)
                txtOperarioEnvia.Text = loOperario.mnCodigo & " - " & loOperario.msNombre
            End If
            If moTraspaso.mnOperarioRecepcion > 0 Then
                txtFechaRecepcion.Text = Format(moTraspaso.mdFechaRecepcion, "dd/MM/yyyy HH:mm:ss")
                loOperario = mfoOperario(moTraspaso.mnOperarioRecepcion)
                txtOperarioRecibe.Text = loOperario.mnCodigo & " - " & loOperario.msNombre
            End If

            cmdEnviado.Visible = False
            cmdRecibido.Visible = False
            grpFechas.Visible = (moTraspaso.msEstadoEnvio <> "")
            grpEnvio.Visible = False

            txtFechaRecepcion.Visible = True
            txtOperarioRecibe.Visible = True
            txtFechaEnvio.Visible = True
            txtOperarioEnvia.Visible = True
            lblFechaRecepcion.Visible = True
            lblOperarioRecepcion.Visible = True
            lblFechaEnvio.Visible = True
            lblOperarioEnvio.Visible = True
            lblEstado.Visible = True

            Select Case moTraspaso.msEstadoEnvio
                Case "G"
                    cmdEnviado.Visible = True
                    cmdRecibido.Visible = True
                    lblEstado.Text = "TRASPASO GRABADO"

                    grpEnvio.Visible = True
                    txtFechaEnvio.Visible = False
                    txtOperarioEnvia.Visible = False
                    txtFechaRecepcion.Visible = False
                    txtOperarioRecibe.Visible = False
                    lblFechaEnvio.Visible = False
                    lblOperarioEnvio.Visible = False
                    lblFechaRecepcion.Visible = False
                    lblOperarioRecepcion.Visible = False
                Case "E"
                    cmdRecibido.Visible = True
                    lblEstado.Text = "TRASPASO ENVIADO"

                    grpEnvio.Visible = True
                    txtFechaRecepcion.Visible = False
                    txtOperarioRecibe.Visible = False
                    lblFechaRecepcion.Visible = False
                    lblOperarioRecepcion.Visible = False
                Case "R"
                    lblEstado.Text = "TRASPASO RECIBIDO"
                Case Else
                    lblEstado.Visible = False
            End Select
            mrMuestraConfirmacion()

            ' despues los campos de las lineas ***********
            lnI = 0
            For Each loLinea In moTraspaso.mcolLineas
                grdLineas.mrAñadirFila()
                grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo
                If loLinea.mnDetalle > 0 Then grdLineas.marMemoria(0, lnI) = loLinea.mnArticulo & "." & loLinea.mnDetalle
                grdLineas.marMemoria(1, lnI) = loLinea.msDescripcion
                grdLineas.marMemoria(2, lnI) = Format(loLinea.mnCantidad, "#,##0.00")
                grdLineas.marMemoria(3, lnI) = loLinea.mpsCodigo
                grdLineas.marMemoria(4, lnI) = ""
                grdLineas.marMemoria(5, lnI) = loLinea.mnArticulo
                lnI = lnI + 1
            Next
            ' despues añado una linea nueva que es la que pongo para meter nuevos datos
            grdLineas.mrAñadirFila()
            grdLineas.mrRefrescaGrid()
        Else
            ' primero los campos de la cabecera **************
            moTraspaso.mnCodigo = txtCodigo.Text
            moTraspaso.mdFecha = Now
            moTraspaso.msEstado = txtEstado.Text
            moTraspaso.msObservaciones = txtComentario.Text
            moTraspaso.mnDesde = txtDesde.Text
            moTraspaso.mnHasta = txtHasta.Text
            moTraspaso.mnOperario = txtOperador.Text
            moTraspaso.mnVendedor = txtVendedor.Text
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
            On Error Resume Next
            grdLineas.ColorCelda(lnI, lnFila).mnForeColor = loColor
            On Error GoTo 0
        Next

    End Sub

    Private Function mfbCargaArticulo(ByVal lsDescripcion As String) As Boolean
        Dim loArticulo As New clsArticulo
        Dim lnArticuloTemp As Integer
        Dim lnDetalleTemp As Integer
        Dim lnArticulo As Long
        Dim lnDetalle As Integer
        Dim lnLinea As Integer
        Dim loExistencias As clsExistencias
        Dim lsCodigo As String
        Dim lsCodigoAnterior As String

        If moAlmacen1 Is Nothing Then moAlmacen1 = New clsAlmacen
        If moAlmacen2 Is Nothing Then moAlmacen2 = New clsAlmacen

        If moAlmacen1.mnCodigo = 0 Then
            MsgBox("Debe indicar los almacenes que intervienen en el traspaso.", MsgBoxStyle.Critical, "Visanfer .Net")
            Return False
        End If

        lnLinea = grdLineas.mnRow
        lsCodigo = Trim(grdLineas.marMemoria(0, lnLinea))
        lsCodigoAnterior = Trim(grdLineas.marMemoria(6, lnLinea))

        If lsCodigo = lsCodigoAnterior Then Return True
        grdLineas.marMemoria(6, lnLinea) = lsCodigo

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

            If lsDescripcion = "" Then lsDescripcion = loArticulo.msDescripcion
            grdLineas.marMemoria(1, lnLinea) = lsDescripcion

            ' recomiendo no hacer traspasos con articulos virtuales
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
                    Dim lsUnidades As String = ""
                    For Each loAgrArt In loBusAgrArt.mcolAgrart
                        lsUnidades = lsUnidades & loAgrArt.mnArticulo & "." & loAgrArt.mnDetalle & ","
                    Next
                    lsUnidades = Mid(lsUnidades, 1, lsUnidades.Length - 1)
                    MsgBox("ATENCION, ESTE ARTICULO ES AGRUPACION DE OTROS, SE RECOMIENDA HACER EL TRASPASO CON EL CODIGO DE LAS UNIDADES." & vbCrLf & "(" & lsUnidades & ")", MsgBoxStyle.Critical, "Visanfer .Net")
                End If
            End If

            ' ************************************************************
            loExistencias = New clsExistencias
            loExistencias.mnEmpresa = mnEmpresa
            loExistencias.mnAlmacen = moAlmacen1.mnCodigo
            loExistencias.mnArticulo = loArticulo.mnCodigo
            loExistencias.mnDetalle = loArticulo.mnDetalle
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
            loExistencias.mnExistencias = 0  ' los pongo a cero por si acaso no hay registro creado en el almacen destino
            loExistencias.mnExposicion = 0

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

    Private Sub mrMuestraConfirmacion()

        Dim lnLinea As Integer = grdLineas.mnRow + 1
        If lnLinea <= moTraspaso.mcolLineas.Count Then
            Dim loLinea As New clsTraspasoLin
            loLinea = moTraspaso.mcolLineas(grdLineas.mnRow + 1)
            lblConfirmado.Text = loLinea.mfsLogConfirmado()
        End If

    End Sub

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

    Private Sub cmdEnviado_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEnviado.Click
        mrMarcarEnviado()
    End Sub

    Private Sub cmdRecibido_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mrMarcarRecibido()
    End Sub

#End Region

#Region " Eventos de Formulario "

    Private Sub frmTraspasos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        mbPrimeraVez = True
        moTraspaso = New clsTraspaso
        'If mnTipoCompilacion = 1 Then mrCargar(1, goUsuario) ' comentar al poner a dll
        mrAsignaEventos()
        mrPintaFormulario()
        If mbCargaAutomatica Then
            mrNuevoRegistro()
            moTraspaso = moTraspasoAux
            mrMoverCampos(1)
            tabMenu.Enabled = False
            txtDesde.Focus()
        ElseIf mbConsultaTraspaso Then
            mrConsulta()
            txtCodigo.Text = moTraspasoAux.mnCodigo
            mrCargaTraspaso()
            tabMenu.Enabled = False
            txtCodigo.Focus()
        Else

            If mbCargaDirecta Then
                tabMenu.Visible = False
                mrConsulta()
            Else
                tabMenu.evtFocus()
            End If

        End If

    End Sub

    Private Sub frmTraspasos_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyData = 131138 Then goUsuario.mrBloquear(gnLlave)
    End Sub

    Private Sub frmTraspasos_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If goUsuario Is Nothing Then goUsuario = New clsUsuario
        If goUsuario.mbEsNuevo Then goUsuario.mrBloquear(gnLlave)
        lblTitle.Text = goUsuario.msNombre
        mrSolicitudesPendientes()
    End Sub

    Private Sub cmdDesbloqueo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDesbloqueo.Click
        cmdDesbloqueo.Visible = False
        cmdBloqueo.Visible = True
        txtDesde.ReadOnly = False
        txtDesde.Focus()
    End Sub

    Private Sub cmdBloqueo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBloqueo.Click
        cmdDesbloqueo.Visible = True
        cmdBloqueo.Visible = False
        txtDesde.ReadOnly = True
        txtDesde.Focus()
    End Sub

#End Region

#Region " Eventos de Teclado para las cajas de texto "

    Private Sub txtCodigo_evtSalida() Handles txtCodigo.evtSalida
        If mtEstado = EstadoVentana.Consulta Then
            txtCodigo.Focus()
            txtCodigo.SelectAll()
        End If
    End Sub

    Private Sub tmrSolicitud_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrSolicitud.Tick
        lblSolicitudes.Visible = Not lblSolicitudes.Visible
    End Sub

#End Region

#Region " Eventos del Grid "

    Private Sub grdLineas_evtKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdLineas.evtKeyDown
        mrLeeTecla(grdLineas, e)
    End Sub

    Private Sub grdLineas_evtLeaveCell() Handles grdLineas.evtLeaveCell
        'If mtEstado = EstadoVentana.Mantenimiento Then mrAnalizaDatos()
    End Sub

    Private Sub grdLineas_evtTextChanged() Handles grdLineas.evtTextChanged
        'If (mtEstado = EstadoVentana.Mantenimiento) Or (mtEstado = EstadoVentana.NuevoRegistro) Then
        '    If Not mbCargaAutomatica Then
        '        mrPendiente(grdLineas.mnRow, "S")
        '        grdLineas.mrRefrescaGrid()
        '    End If
        'End If
    End Sub

    Private Sub grdLineas_evtRowColChange(lnColOld As Integer, lnRowOld As Integer) Handles grdLineas.evtRowColChange

        If (mtEstado = EstadoVentana.Lineas) AndAlso (moTraspaso.mnCodigo > 0) Then
            mrMuestraConfirmacion()
        End If

    End Sub

#End Region

End Class
