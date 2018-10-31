Option Explicit On 
Imports Microsoft.Data.Odbc
Imports System.Windows.Forms.SendKeys
Imports System.Drawing.Printing
Imports prjPrinterNet

Public Class frmSelImpresora
    Inherits System.Windows.Forms.Form
    Public msDestino As String = ""
    Public mnImpresora As Integer = 0
    Public mnCopias As Integer = 0
    Public Event evtBusImpresora()   ' evento desencadenado despues una seleccion
    Dim WithEvents moImpresora As prjPrinterNet.clsImpresora

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
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents txtCopias As control.txtVisanfer
    Friend WithEvents txtImpresora As control.txtVisanfer
    Friend WithEvents txtDestino As control.txtVisanfer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblPrograma As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblTeclas = New System.Windows.Forms.Label()
        Me.panCampos = New System.Windows.Forms.Panel()
        Me.txtCopias = New control.txtVisanfer()
        Me.txtImpresora = New control.txtVisanfer()
        Me.txtDestino = New control.txtVisanfer()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblPrograma = New System.Windows.Forms.Label()
        Me.panCampos.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(4, 191)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(360, 25)
        Me.lblTeclas.TabIndex = 25
        Me.lblTeclas.Text = "F5 - ACEPTAR    F9 - IMPRESORAS     ESC - SALIDA"
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtCopias, Me.txtImpresora, Me.txtDestino, Me.Label5, Me.Label4, Me.Label3, Me.Label7})
        Me.panCampos.Location = New System.Drawing.Point(4, 36)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(360, 155)
        Me.panCampos.TabIndex = 24
        '
        'txtCopias
        '
        Me.txtCopias.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtCopias.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCopias.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCopias.Location = New System.Drawing.Point(307, 84)
        Me.txtCopias.MaxLength = 2
        Me.txtCopias.Name = "txtCopias"
        Me.txtCopias.Size = New System.Drawing.Size(32, 21)
        Me.txtCopias.TabIndex = 2
        Me.txtCopias.Text = ""
        Me.txtCopias.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtCopias.ValorMax = 99
        '
        'txtImpresora
        '
        Me.txtImpresora.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtImpresora.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtImpresora.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImpresora.Location = New System.Drawing.Point(152, 84)
        Me.txtImpresora.MaxLength = 2
        Me.txtImpresora.Name = "txtImpresora"
        Me.txtImpresora.Size = New System.Drawing.Size(32, 21)
        Me.txtImpresora.TabIndex = 1
        Me.txtImpresora.Text = ""
        Me.txtImpresora.tipo = control.txtVisanfer.TiposCaja.Numerico
        Me.txtImpresora.ValorMax = 99
        '
        'txtDestino
        '
        Me.txtDestino.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDestino.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDestino.Location = New System.Drawing.Point(112, 36)
        Me.txtDestino.MaxLength = 1
        Me.txtDestino.Name = "txtDestino"
        Me.txtDestino.Size = New System.Drawing.Size(32, 21)
        Me.txtDestino.TabIndex = 0
        Me.txtDestino.Text = ""
        Me.txtDestino.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDestino.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtDestino.ValorMax = 999999999
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(20, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(124, 16)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "NUM. IMPRESORA.:"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(207, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "NUM. COPIAS:"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(152, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(173, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "V. VENTANA  I.- IMPRESORA"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(22, 38)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(92, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "DESTINO:"
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(4, 4)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(360, 32)
        Me.lblPrograma.TabIndex = 23
        Me.lblPrograma.Text = "DESCRIPCION DEL PROCESO"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmSelImpresora
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(369, 220)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblTeclas, Me.panCampos, Me.lblPrograma})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmSelImpresora"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SELECCION DE IMPRESORA"
        Me.panCampos.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Rutinas Varias "

    Public Sub mrMandaEvento()
        RaiseEvent evtBusImpresora()
    End Sub

    Public Sub mrSeleccionar(ByVal lsPrograma As String)
        lblPrograma.Text = lsPrograma
        Me.ShowDialog()
    End Sub

    Private Sub moImpresora_evtBusImpresora() Handles moImpresora.evtBusImpresora
        txtImpresora.Text = moImpresora.mnCodigo
    End Sub

    Private Sub mrBuscar()
        moImpresora = New prjPrinterNet.clsImpresora()
        moImpresora.mrBuscaImpresora()
    End Sub

    Private Sub frmLisArticulos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mrPintaFormulario()
    End Sub

    Private Sub mrPintaFormulario()
        ' cargo los valores previos si los hay
        txtDestino.Text = msDestino
        txtImpresora.Text = mnImpresora
        txtCopias.Text = mnCopias
        If msDestino = "" Then txtDestino.Text = "I"
        If mnImpresora = 0 Then txtImpresora.Text = goProfile.mnImpresora
        If mnCopias = 0 Then txtCopias.Text = "1"
    End Sub

    Private Sub mrLeeTecla(ByVal e As System.Windows.Forms.KeyEventArgs, ByVal lnIndex As Int16)

        Select Case e.KeyValue
            Case 13     ' Intro
                Send("{TAB}")
            Case 27     ' Escape Salgo de la busqueda
                Me.Close()
            Case 38     ' Flecha para arriba
                Send("+{TAB}")
            Case 40     ' Flecha para abajo
                Send("{TAB}")
            Case 116        ' Aceptar
                msDestino = txtDestino.Text
                mnImpresora = mfnInteger(txtImpresora.Text)
                mnCopias = mfnInteger(txtCopias.Text)
                Me.Close()
                mrMandaEvento()
            Case 120    ' Busqueda  F9
                Select Case lnIndex
                    Case 1          ' txtimpresora
                        mrBuscar()
                End Select
        End Select

    End Sub

#End Region

#Region " Eventos de Teclado para los controles "

    Private Sub txtDestino_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDestino.KeyDown
        mrLeeTecla(e, 0)
    End Sub

    Private Sub txtImpresora_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtImpresora.KeyDown
        mrLeeTecla(e, 1)
    End Sub

    Private Sub txtCopias_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCopias.KeyDown
        mrLeeTecla(e, 2)
    End Sub

#End Region

End Class
