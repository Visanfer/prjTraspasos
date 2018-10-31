Option Explicit On

Imports System.Windows.Forms.SendKeys
Imports prjControl
Imports prjPrinterNet
Imports prjEmpresas
Imports prjArticulos
Imports prjAlmacen

Public Class frmConfEntradas
    Inherits System.Windows.Forms.Form
    Private mnEmpresa As Int32      ' empresa de gestion
    Dim moAlmacen As clsAlmacen
    Dim moEmpContable As clsEmpContable
    Dim moBusTraspasos As clsBusTraspasos
    Dim mbPrimeraVez As Boolean
    ' variables de impresion ************************
    Private WithEvents moSelImpresora As prjControl.frmSelImpresora
    Dim moPrinter As New clsPrinter       ' objeto para imprimir
    Dim moImpresora As New prjPrinterNet.clsImpresora   ' Objeto Impresora
    Dim msRaya As String
    Dim mnLineas As Integer
    Dim mnOrden As Integer = 1

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
    Friend WithEvents lblDescripcion As System.Windows.Forms.Label
    Friend WithEvents lblTitulo As System.Windows.Forms.Label
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents lblTraspaso As System.Windows.Forms.Label
    Friend WithEvents grdLineas As prjGrid.ctlGrid
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents lblTeclas As System.Windows.Forms.Label
    Friend WithEvents cmdDesbloqueo As System.Windows.Forms.Button
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtAlmacen As control.txtVisanfer
    Friend WithEvents txtDesAlmacen As control.txtVisanfer
    Friend WithEvents cmdBloqueo As System.Windows.Forms.Button
    Friend WithEvents lblDescripcionOrd As System.Windows.Forms.Label
    Friend WithEvents lblCodigoOrd As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfEntradas))
        Me.lblPrograma = New System.Windows.Forms.Label()
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.panCampos = New System.Windows.Forms.Panel()
        Me.lblDescripcionOrd = New System.Windows.Forms.Label()
        Me.lblCodigoOrd = New System.Windows.Forms.Label()
        Me.cmdBloqueo = New System.Windows.Forms.Button()
        Me.txtDesAlmacen = New control.txtVisanfer()
        Me.cmdDesbloqueo = New System.Windows.Forms.Button()
        Me.txtAlmacen = New control.txtVisanfer()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.lblTraspaso = New System.Windows.Forms.Label()
        Me.grdLineas = New prjGrid.ctlGrid()
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
        Me.lblPrograma.Location = New System.Drawing.Point(9, 9)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(727, 32)
        Me.lblPrograma.TabIndex = 46
        Me.lblPrograma.Text = "GESTION"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDescripcion
        '
        Me.lblDescripcion.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDescripcion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescripcion.Location = New System.Drawing.Point(9, 41)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.Size = New System.Drawing.Size(727, 24)
        Me.lblDescripcion.TabIndex = 47
        Me.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTitulo
        '
        Me.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTitulo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTitulo.Location = New System.Drawing.Point(733, 9)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(273, 56)
        Me.lblTitulo.TabIndex = 45
        Me.lblTitulo.Text = "VISANFER, S.A. - 2003"
        Me.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.lblDescripcionOrd)
        Me.panCampos.Controls.Add(Me.lblCodigoOrd)
        Me.panCampos.Controls.Add(Me.cmdBloqueo)
        Me.panCampos.Controls.Add(Me.txtDesAlmacen)
        Me.panCampos.Controls.Add(Me.cmdDesbloqueo)
        Me.panCampos.Controls.Add(Me.txtAlmacen)
        Me.panCampos.Controls.Add(Me.Label24)
        Me.panCampos.Controls.Add(Me.lblTraspaso)
        Me.panCampos.Controls.Add(Me.grdLineas)
        Me.panCampos.Location = New System.Drawing.Point(9, 64)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(997, 624)
        Me.panCampos.TabIndex = 44
        '
        'lblDescripcionOrd
        '
        Me.lblDescripcionOrd.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.lblDescripcionOrd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDescripcionOrd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescripcionOrd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblDescripcionOrd.Location = New System.Drawing.Point(402, 63)
        Me.lblDescripcionOrd.Name = "lblDescripcionOrd"
        Me.lblDescripcionOrd.Size = New System.Drawing.Size(493, 21)
        Me.lblDescripcionOrd.TabIndex = 119
        Me.lblDescripcionOrd.Text = "orden"
        Me.lblDescripcionOrd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblDescripcionOrd.Visible = False
        '
        'lblCodigoOrd
        '
        Me.lblCodigoOrd.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.lblCodigoOrd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCodigoOrd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCodigoOrd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblCodigoOrd.Location = New System.Drawing.Point(276, 63)
        Me.lblCodigoOrd.Name = "lblCodigoOrd"
        Me.lblCodigoOrd.Size = New System.Drawing.Size(61, 21)
        Me.lblCodigoOrd.TabIndex = 118
        Me.lblCodigoOrd.Text = "orden"
        Me.lblCodigoOrd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblCodigoOrd.Visible = False
        '
        'cmdBloqueo
        '
        Me.cmdBloqueo.Image = CType(resources.GetObject("cmdBloqueo.Image"), System.Drawing.Image)
        Me.cmdBloqueo.Location = New System.Drawing.Point(517, 4)
        Me.cmdBloqueo.Name = "cmdBloqueo"
        Me.cmdBloqueo.Size = New System.Drawing.Size(28, 28)
        Me.cmdBloqueo.TabIndex = 117
        Me.cmdBloqueo.TabStop = False
        Me.cmdBloqueo.Visible = False
        '
        'txtDesAlmacen
        '
        Me.txtDesAlmacen.AutoSelec = False
        Me.txtDesAlmacen.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtDesAlmacen.Blink = False
        Me.txtDesAlmacen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDesAlmacen.DesdeCodigo = CType(0, Long)
        Me.txtDesAlmacen.DesdeFecha = New Date(CType(0, Long))
        Me.txtDesAlmacen.HastaCodigo = CType(0, Long)
        Me.txtDesAlmacen.HastaFecha = New Date(CType(0, Long))
        Me.txtDesAlmacen.Location = New System.Drawing.Point(268, 8)
        Me.txtDesAlmacen.MaxLength = 30
        Me.txtDesAlmacen.Name = "txtDesAlmacen"
        Me.txtDesAlmacen.ReadOnly = True
        Me.txtDesAlmacen.Size = New System.Drawing.Size(235, 20)
        Me.txtDesAlmacen.TabIndex = 116
        Me.txtDesAlmacen.TabStop = False
        Me.txtDesAlmacen.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDesAlmacen.ValorMax = 999999999.0R
        '
        'cmdDesbloqueo
        '
        Me.cmdDesbloqueo.Image = CType(resources.GetObject("cmdDesbloqueo.Image"), System.Drawing.Image)
        Me.cmdDesbloqueo.Location = New System.Drawing.Point(517, 4)
        Me.cmdDesbloqueo.Name = "cmdDesbloqueo"
        Me.cmdDesbloqueo.Size = New System.Drawing.Size(28, 28)
        Me.cmdDesbloqueo.TabIndex = 115
        Me.cmdDesbloqueo.TabStop = False
        '
        'txtAlmacen
        '
        Me.txtAlmacen.AutoSelec = False
        Me.txtAlmacen.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtAlmacen.Blink = False
        Me.txtAlmacen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAlmacen.DesdeCodigo = CType(0, Long)
        Me.txtAlmacen.DesdeFecha = New Date(CType(0, Long))
        Me.txtAlmacen.HastaCodigo = CType(0, Long)
        Me.txtAlmacen.HastaFecha = New Date(CType(0, Long))
        Me.txtAlmacen.Location = New System.Drawing.Point(224, 8)
        Me.txtAlmacen.MaxLength = 2
        Me.txtAlmacen.Name = "txtAlmacen"
        Me.txtAlmacen.ReadOnly = True
        Me.txtAlmacen.Size = New System.Drawing.Size(40, 20)
        Me.txtAlmacen.TabIndex = 0
        Me.txtAlmacen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtAlmacen.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtAlmacen.ValorMax = 999999999.0R
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(16, 12)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(208, 16)
        Me.Label24.TabIndex = 114
        Me.Label24.Text = "ALMACEN RECEPTOR DE MATERIAL:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTraspaso
        '
        Me.lblTraspaso.BackColor = System.Drawing.SystemColors.Info
        Me.lblTraspaso.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTraspaso.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTraspaso.Location = New System.Drawing.Point(8, 32)
        Me.lblTraspaso.Name = "lblTraspaso"
        Me.lblTraspaso.Size = New System.Drawing.Size(984, 30)
        Me.lblTraspaso.TabIndex = 110
        Me.lblTraspaso.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grdLineas
        '
        Me.grdLineas.BackColor = System.Drawing.SystemColors.HighlightText
        Me.grdLineas.Columnas = 0
        Me.grdLineas.Editable = False
        Me.grdLineas.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdLineas.Location = New System.Drawing.Point(8, 84)
        Me.grdLineas.Name = "grdLineas"
        Me.grdLineas.Size = New System.Drawing.Size(984, 524)
        Me.grdLineas.TabIndex = 1
        '
        'lblInfo
        '
        Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInfo.Location = New System.Drawing.Point(9, 691)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(997, 16)
        Me.lblInfo.TabIndex = 43
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(9, 707)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(997, 24)
        Me.lblTeclas.TabIndex = 42
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmConfEntradas
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1018, 740)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblPrograma)
        Me.Controls.Add(Me.lblDescripcion)
        Me.Controls.Add(Me.lblTitulo)
        Me.Controls.Add(Me.panCampos)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.lblTeclas)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmConfEntradas"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Confirmacion de Entradas en Almacen"
        Me.panCampos.ResumeLayout(False)
        Me.panCampos.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCargar(ByVal lnEmpresa As Integer)

        ' la llave actual de seguridad *****************************
        gnLlave = 0
        ' pasa como parametro la conexion y la empresa ************************
        mnEmpresa = lnEmpresa
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

        Me.ShowDialog()

    End Sub

    Private Sub mrPintaFormulario()
        Dim loEmpresa As New prjEmpresas.clsEmpresa

        ' pongo los datos de la empresa ********
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()
        lblPrograma.Text = "TRASPASOS - PENDIENTES"
        lblDescripcion.Text = "LISTA TODOS LOS ARTICULOS TRASPASADOS PENDIENTES DE ACTUALIZAR"
        lblTitulo.Text = loEmpresa.msNombre
        txtAlmacen.Text = goProfile.mnAlmacenEntradas
        txtDesAlmacen.Text = mfsDesAlmacen(goProfile.mnAlmacenEntradas)
        lblTeclas.Text = "F1-ORDEN    F5-ACTUALIZAR DATOS    F8-MARCAR LINEA     CTRL+P-IMPRIMIR      ESC-SALIDA"
        ' ***************************************************************

        mrPreparaGrid()

    End Sub

    Private Sub mrLeeTecla(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loCaja As control.txtVisanfer
        Dim loGrid As prjGrid.ctlGrid
        Dim lsControl As String = ""
        Dim lbSalta As Boolean = False

        If TypeOf sender Is control.txtVisanfer Then
            loCaja = sender
            lsControl = loCaja.Name
        End If
        If TypeOf sender Is prjGrid.ctlGrid Then
            loGrid = sender
            lsControl = loGrid.Name
        End If

        Select Case e.KeyValue
            Case Keys.P And e.Control = True     'CONTROL + P
                mrSelImpresora()
            Case Keys.Enter
                Select Case lsControl
                    Case "txtAlmacen"
                        mrCargaAlmacen(mfnLong(txtAlmacen.Text), e)
                End Select
                If lsControl <> "grdLineas" Then Send("{TAB}")
            Case Keys.Escape
                Select Case lsControl
                    Case "txtAlmacen"
                        Me.Close()
                    Case "grdLineas"
                        txtAlmacen.Focus()
                End Select
            Case Keys.F1     ' cambia el orden
                mrCambiaOrden()
            Case Keys.F8
                If lsControl = "grdLineas" Then
                    mrMarcar()
                    Send("{DOWN}")
                End If
            Case Keys.F5
                mrGrabar()
        End Select

    End Sub

    Private Sub moSelImpresora_evtBusImpresora() Handles moSelImpresora.evtBusImpresora
        mrImprimir()
    End Sub

    Private Sub mrImprimir()
        Dim loLinea As clsTraspasoLin
        Dim loTraspaso As clsTraspaso
        Dim lnTopLineas As Integer
        Dim lsObservaciones As String
        Dim lsFecha As String
        Dim lsCodigo As String
        Dim lsArticulo As String
        Dim lsDesArticulo As String = ""
        Dim lsCantidad As String
        Dim lnPagina As Integer
        Dim lnI As Integer
        Dim loMensaje As New frmMensajes

        ' ******* esta dll necesita que se le pase la conexion ************
        loMensaje.mrAbrir("Imprimiendo, por favor espere ...")

        ' ******** recupero la cola de la impresora seleccionada **********
        moImpresora.mnEmpresa = mnEmpresa
        moImpresora.mnCodigo = moSelImpresora.mnImpresora
        moImpresora.mrRecuperaDatos()
        ' *********** inicio del proceso de impresion *************************
        moPrinter.mrInicio(goProfile.msLogin)
        moPrinter.msCola = moImpresora.msCola
        moPrinter.mpbComprimida = True
        moPrinter.mpbProporcional = True
        If moImpresora.mbEsNuevo Then
            MsgBox("Impresora no instalada", MsgBoxStyle.Critical, "Visanfer .Net")
        Else

            For lnI = 1 To 130
                msRaya = msRaya & "="
            Next
            ' ********* parametros de impresion ********************
            lnPagina = 1
            If moImpresora.msTipo = "L" Then
                lnTopLineas = 70                ' numero de lineas por pagina
            Else
                lnTopLineas = 66                ' numero de lineas por pagina
            End If
            'lnTopLineas = 66                ' numero de lineas por pagina

            moPrinter.mpnLineas = lnTopLineas    ' albaranes 48, A4 66
            moPrinter.mnCopias = moSelImpresora.mnCopias     ' Copias a imprimir

            ' ************  cabecera ************************************
            mrCabecera(lnPagina)
            ' ************  lineas del listado **************************

            For Each loLinea In moBusTraspasos.mcolLineas

                loTraspaso = New clsTraspaso
                loTraspaso.mnEmpresa = loLinea.mnEmpresa
                loTraspaso.mnCodigo = loLinea.mnCodigo
                loTraspaso.mrRecuperaDatos()

                lsObservaciones = mfsDesAlmacen(loTraspaso.mnDesde) & " (" & loTraspaso.msObservaciones & ")"
                lsFecha = Format(loTraspaso.mdFecha, "dd/MM/yyyy")
                lsCodigo = loLinea.mnCodigo
                lsArticulo = loLinea.mnArticulo
                'lsDesArticulo = mfsDesArticulo(loLinea.mnArticulo, loLinea.mnDetalle)
                lsCantidad = Format(loLinea.mnCantidad, "#,##0.00")
                ' ******** impresion **************************************
                moPrinter.mrPrint(0, lsObservaciones, 0, 40, modGeneral.Alineacion.Izquierda)
                moPrinter.mrPrint(-1, lsFecha, 41, 10, modGeneral.Alineacion.Centro)
                moPrinter.mrPrint(-1, lsCodigo, 52, 6, modGeneral.Alineacion.Centro)
                moPrinter.mrPrint(-1, lsArticulo, 59, 6, modGeneral.Alineacion.Centro)
                moPrinter.mrPrint(-1, lsDesArticulo, 66, 45)
                moPrinter.mrPrint(-1, lsCantidad, 112, 15, modGeneral.Alineacion.Derecha)
                ' ************ Control de paginas *************************
                If moPrinter.mpnLineaActual > lnTopLineas - 3 Then
                    moPrinter.mrFormFeed()
                    lnPagina = lnPagina + 1
                    mrCabecera(lnPagina)
                End If
            Next

            ' ************ pie del listado *******************************
            mrPie(lnPagina)
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

    Private Sub mrPie(ByVal lnPagina As Integer)

        moPrinter.mrPrint(0, msRaya, 0, 130)
        moPrinter.mrPrint(0, "- FINAL DEL LISTADO -", 60, 50)
        moPrinter.mrPrint(0, msRaya, 0, 130)

    End Sub

    Private Sub mrCabecera(ByVal lnPagina As Integer)

        moPrinter.mrPrint(0, msRaya, 0, 130)

        'moPrinter.mrPrint(0, "VISANFER.COM", 0, 20)
        moPrinter.mrPrint(0, Format(Now, "dd/MM/yyyy"), 5, 12)
        moPrinter.mrPrint(-1, "- TRASPASOS PENDIENTES - " & moAlmacen.msNombre, 25, 50)
        moPrinter.mrPrint(-1, "PAG. " & lnPagina, 90, 20)

        moPrinter.mrPrint(0, msRaya, 0, 130)

        moPrinter.mrPrint(0, "")
        moPrinter.mrPrint(0, "ORIGEN", 0, 40)
        moPrinter.mrPrint(-1, "FECHA", 41, 10)
        moPrinter.mrPrint(-1, "COD.", 52, 6)
        moPrinter.mrPrint(-1, "ART.", 59, 6)
        moPrinter.mrPrint(-1, "DESCRIPCION", 66, 45)
        moPrinter.mrPrint(-1, "CTD.", 112, 15)

        moPrinter.mrPrint(0, "")
        moPrinter.mrPrint(-1, "------", 0, 10)
        moPrinter.mrPrint(-1, "------", 41, 10)
        moPrinter.mrPrint(-1, "-------", 52, 6)
        moPrinter.mrPrint(-1, "---------", 59, 6)
        moPrinter.mrPrint(-1, "---------", 66, 45)
        moPrinter.mrPrint(-1, "---------", 112, 15)

    End Sub

    Private Sub mrSelImpresora()
        ' selecciona la impresora
        moSelImpresora = New prjControl.frmSelImpresora
        moSelImpresora.mrSeleccionar("IMP. TRASPASOS PENDIENTES")

    End Sub

    Private Sub mrCambiaOrden()
        Select Case mnOrden
            Case 1      ' codigo
                mnOrden = 2
            Case 2      ' descripcion
                mnOrden = 1
        End Select
        Cursor = Cursors.WaitCursor
        mrCargaAlmacen(mfnLong(txtAlmacen.Text), Nothing)
        Cursor = Cursors.Default
    End Sub

    Private Sub mrMarcar()
        Dim loColor As Color
        Dim lnLinea As Integer
        Dim lnCol As Integer
        Dim lnI As Integer
        Dim lsTipo As String
        Dim lsCodigo As String
        Dim lnLlave As Integer

        Select Case moAlmacen.mnCodigo
            Case 1
                lnLlave = 36
            Case 2
                lnLlave = 37
            Case 3
                lnLlave = 42
            Case 4
                lnLlave = 38
            Case 5
                lnLlave = 53
            Case 6
                lnLlave = 39
            Case 7, 10, 11
                lnLlave = 40
            Case 8
                lnLlave = 137
            Case 9
                lnLlave = 41
        End Select

        If goUsuario.mfbAccesoPermitido(lnLlave, True) Then
            ' marco la linea actual **********************
            lnLinea = grdLineas.mnRow
            lnCol = grdLineas.mnCol
            ' ******* coloreo la linea de color rojo *
            lsCodigo = grdLineas.marMemoria(2, lnLinea)
            If lsCodigo <> "" Then
                lsTipo = grdLineas.marMemoria(7, lnLinea)
                Select Case lsTipo
                    Case "B"        ' si la linea esta borrada se pone normal
                        grdLineas.marMemoria(7, lnLinea) = ""
                        loColor = Color.FromName("Window")
                    Case Else       ' si la linea en normal se marca ********
                        grdLineas.marMemoria(7, lnLinea) = "B"
                        loColor = Color.DarkOrange
                End Select
                If lsTipo <> "I" Then
                    For lnI = 0 To 7
                        grdLineas.ColorCelda(lnI, grdLineas.mnRow).mnBackColor = loColor
                    Next
                End If
                grdLineas.mrRefrescaGrid()
            End If
        End If

    End Sub


    Private Sub mrGrabar()
        Dim lnI As Integer
        Dim lsTipo As String
        Dim lsCodigo As String
        Dim loTraspasoLin As clsTraspasoLin
        Dim loTraspaso As clsTraspaso

        For lnI = 0 To grdLineas.mnFilasDatos - 1
            ' primero veo que tipo de linea es ********************
            lsTipo = grdLineas.marMemoria(7, lnI)
            If lsTipo = "B" Then
                loTraspaso = New clsTraspaso
                loTraspaso.mnEmpresa = mnEmpresa
                loTraspaso.mnCodigo = mfnInteger(grdLineas.marMemoria(2, lnI))
                loTraspaso = moBusTraspasos.mcolTraspasos(loTraspaso.mpsCodigo)
                ' ahora actualizo el fichero de almacenes ***************
                ' lo quito temporalmente pues ya se recupera por si mismo
                lsCodigo = grdLineas.marMemoria(6, lnI)
                If loTraspaso.mcolLineas Is Nothing Then loTraspaso.mrRecuperaLineas()
                If loTraspaso.mcolLineas.Count = 0 Then loTraspaso.mrRecuperaLineas()
                loTraspasoLin = loTraspaso.mcolLineas(lsCodigo)
                ' compruebo que no se haya actualizado ya *********************
                loTraspasoLin.mrRecuperaDatos()
                If loTraspasoLin.msEstado <> "A" Then
                    loTraspasoLin.mrActualizaLineaTraspaso(moAlmacen.mnCodigo, goUsuario.mnCodigo)
                    ' ahora actualizo el estado de la linea ****************
                    loTraspasoLin.msEstado = "A"
                    loTraspasoLin.mrGrabaDatos()
                End If
            End If
        Next
        ' refresco el estado de las nuevas lineas *******
        If moAlmacen Is Nothing Then Exit Sub
        mrCargaAlmacen(moAlmacen.mnCodigo, Nothing)
        grdLineas.Focus()

    End Sub

    Private Sub mrCargaAlmacen(ByVal lnCodigo As Integer, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loAlmacen As clsAlmacen
        Dim lnI As Integer
        Dim loLinea As clsTraspasoLin
        Dim loTraspaso As clsTraspaso

        If moAlmacen Is Nothing Then moAlmacen = New clsAlmacen

        loAlmacen = New clsAlmacen
        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = lnCodigo
        loAlmacen.mrRecuperaDatos(False)
        moAlmacen = loAlmacen
        txtAlmacen.Text = loAlmacen.mnCodigo
        txtDesAlmacen.Text = loAlmacen.msNombre
        If loAlmacen.mbEsNuevo Then
            e.Handled = True
            MsgBox("Registro no Encontrado", MsgBoxStyle.Exclamation, "Gestion Visanfer")
            txtAlmacen.Focus()
        Else
            ' recupero todos los traspasos pendientes de actualizar
            moBusTraspasos = New clsBusTraspasos
            moBusTraspasos.mnEmpresa = mnEmpresa
            moBusTraspasos.mnHasta = moAlmacen.mnCodigo
            moBusTraspasos.mrBuscaTraspasosPendientes(mnOrden)
            ' ***********************************************************
            If mnOrden = 1 Then
                lblCodigoOrd.Visible = True
                lblDescripcionOrd.Visible = False
            Else
                lblCodigoOrd.Visible = False
                lblDescripcionOrd.Visible = True
            End If
            ' pinto el distintivo de orden ******************************
            ' despues los campos de las lineas ***********
            lnI = 0
            grdLineas.mrClear(prjGrid.ctlGrid.TipoBorrado.Contenido)
            For Each loLinea In moBusTraspasos.mcolLineas
                loTraspaso = New clsTraspaso
                loTraspaso.mnEmpresa = loLinea.mnEmpresa
                loTraspaso.mnCodigo = loLinea.mnCodigo
                loTraspaso.mrRecuperaDatos()

                grdLineas.mrAñadirFila()
                grdLineas.marMemoria(0, lnI) = mfsDesAlmacen(loTraspaso.mnDesde) & " (" & Replace(loTraspaso.msObservaciones, vbCrLf, " ") & ")"
                grdLineas.marMemoria(1, lnI) = Format(loTraspaso.mdFecha, "dd/MM/yy")
                grdLineas.marMemoria(2, lnI) = loLinea.mnCodigo
                grdLineas.marMemoria(3, lnI) = loLinea.mnArticulo
                If loLinea.mnDetalle > 0 Then grdLineas.marMemoria(3, lnI) = loLinea.mnArticulo & "." & loLinea.mnDetalle
                'grdLineas.marMemoria(4, lnI) = mfsDesArticulo(loLinea.mnArticulo, loLinea.mnDetalle)
                grdLineas.marMemoria(4, lnI) = loLinea.msDescripcion
                grdLineas.marMemoria(5, lnI) = Format(loLinea.mnCantidad, "#,##0.00")
                grdLineas.marMemoria(6, lnI) = loLinea.mpsCodigo
                grdLineas.marMemoria(7, lnI) = ""
                lnI = lnI + 1
            Next


            'For Each loTraspaso In moBusTraspasos.mcolTraspasos
            '    For Each loLinea In loTraspaso.mcolLineas
            '        grdLineas.mrAñadirFila()
            '        grdLineas.marMemoria(0, lnI) = mfsDesAlmacen(loTraspaso.mnDesde) & " (" & loTraspaso.msObservaciones & ")"
            '        grdLineas.marMemoria(1, lnI) = Format(loTraspaso.mdFecha, "dd/MM/yyyy")
            '        grdLineas.marMemoria(2, lnI) = loLinea.mnCodigo
            '        grdLineas.marMemoria(3, lnI) = loLinea.mnArticulo
            '        grdLineas.marMemoria(4, lnI) = mfsDesArticulo(loLinea.mnArticulo)
            '        grdLineas.marMemoria(5, lnI) = Format(loLinea.mnCantidad, "#,##0.00")
            '        grdLineas.marMemoria(6, lnI) = loLinea.mpsCodigo
            '        grdLineas.marMemoria(7, lnI) = ""
            '        lnI = lnI + 1
            '    Next
            'Next
            ' despues añado una linea nueva que es la que pongo para meter nuevos datos
            grdLineas.mrAñadirFila()
            grdLineas.mrRefrescaGrid()
        End If

        lblTraspaso.Text = "****   " & txtDesAlmacen.Text & "   ****"

    End Sub

    Private Function mfsDesAlmacen(ByVal lnAlmacen As Integer) As String
        Dim loAlmacen As New clsAlmacen

        loAlmacen.mnEmpresa = mnEmpresa
        loAlmacen.mnCodigo = lnAlmacen
        loAlmacen.mrRecuperaDatos(False)
        Return loAlmacen.msNombre

    End Function

    Private Sub mrLimpiaFormulario()
        txtAlmacen.Text = ""
        txtDesAlmacen.Text = ""
        grdLineas.mrClear(prjGrid.ctlGrid.TipoBorrado.Contenido)
    End Sub

    Private Sub mrPreparaGrid()

        grdLineas.Columnas = 8

        grdLineas.marTitulos(0).Texto = "ORIGEN"
        grdLineas.marTitulos(0).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(0).Ancho = 190
        grdLineas.marTitulos(0).Longitud = 50
        grdLineas.marTitulos(0).Editable = False

        grdLineas.marTitulos(1).Texto = "FECHA"
        grdLineas.marTitulos(1).Alineacion = HorizontalAlignment.Center
        grdLineas.marTitulos(1).Ancho = 80
        grdLineas.marTitulos(1).Longitud = 10
        grdLineas.marTitulos(1).Editable = False

        grdLineas.marTitulos(2).Texto = "COD"
        grdLineas.marTitulos(2).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(2).Ancho = 60
        grdLineas.marTitulos(2).Longitud = 6
        grdLineas.marTitulos(2).Editable = False

        grdLineas.marTitulos(3).Texto = "ART"
        grdLineas.marTitulos(3).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(3).Ancho = 70
        grdLineas.marTitulos(3).Longitud = 11
        grdLineas.marTitulos(3).Editable = False

        grdLineas.marTitulos(4).Texto = "DESCRIPCION"
        grdLineas.marTitulos(4).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(4).Ancho = 260
        grdLineas.marTitulos(4).Longitud = 50
        grdLineas.marTitulos(4).Editable = False

        grdLineas.marTitulos(5).Texto = "CTD."
        grdLineas.marTitulos(5).Alineacion = HorizontalAlignment.Right
        grdLineas.marTitulos(5).Ancho = 80
        grdLineas.marTitulos(5).Longitud = 10
        grdLineas.marTitulos(5).Editable = False

        grdLineas.marTitulos(6).Texto = "mpsCodigo"
        grdLineas.marTitulos(6).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(6).Ancho = 0
        grdLineas.marTitulos(6).Longitud = 15
        grdLineas.marTitulos(6).Editable = False

        grdLineas.marTitulos(7).Texto = "marcas"
        grdLineas.marTitulos(7).Alineacion = HorizontalAlignment.Left
        grdLineas.marTitulos(7).Ancho = 0
        grdLineas.marTitulos(7).Longitud = 1
        grdLineas.marTitulos(7).Editable = False

        grdLineas.mnAjustarColumna = 4
        grdLineas.mrPintaGrid()

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

    Private Sub frmConfEntradas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        mbPrimeraVez = True
        mrAsignaEventos()
        mrPintaFormulario()
        txtAlmacen.Focus()

    End Sub

    Private Sub frmConfEntradas_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyData = 131138 Then goUsuario.mrBloquear(gnLlave)
    End Sub

    Private Sub frmConfEntradas_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.Text = goUsuario.msNombre
    End Sub

    Private Sub cmdDesbloqueo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDesbloqueo.Click
        cmdDesbloqueo.Visible = False
        cmdBloqueo.Visible = True
        txtAlmacen.ReadOnly = False
        txtAlmacen.Focus()
    End Sub

    Private Sub cmdBloqueo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBloqueo.Click
        cmdDesbloqueo.Visible = True
        cmdBloqueo.Visible = False
        txtAlmacen.ReadOnly = True
        txtAlmacen.Focus()
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
