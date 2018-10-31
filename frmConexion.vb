Option Explicit On 
Imports Microsoft.Data.Odbc

Public Class frmConexion
    Inherits System.Windows.Forms.Form
    Private menvEntorno As Environment

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
    Friend WithEvents panConexion As System.Windows.Forms.Panel
    Friend WithEvents lbls As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.panConexion = New System.Windows.Forms.Panel()
        Me.lbls = New System.Windows.Forms.Label()
        Me.panConexion.SuspendLayout()
        Me.SuspendLayout()
        '
        'panConexion
        '
        Me.panConexion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panConexion.Controls.AddRange(New System.Windows.Forms.Control() {Me.lbls})
        Me.panConexion.Location = New System.Drawing.Point(10, 8)
        Me.panConexion.Name = "panConexion"
        Me.panConexion.Size = New System.Drawing.Size(365, 258)
        Me.panConexion.TabIndex = 1
        '
        'lbls
        '
        Me.lbls.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbls.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbls.Location = New System.Drawing.Point(20, 184)
        Me.lbls.Name = "lbls"
        Me.lbls.Size = New System.Drawing.Size(322, 58)
        Me.lbls.TabIndex = 1
        Me.lbls.Text = "Conexion"
        Me.lbls.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmConexion
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(383, 274)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.panConexion})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmConexion"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmConexion"
        Me.panConexion.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Soporte para la actualización "
    Private Shared m_vb6FormDefInstance As frmConexion
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmConexion
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmConexion()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmConexion)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Public Sub mrConectar()
        Me.Show()
        mrConectando()
    End Sub

    Private Sub mrConectando()
        Dim lsCadena As String
        Dim lsRuta As String
        Dim loAplicacion As Windows.Forms.Application

        lbls.Text = "Conectando a la base de datos"
        lbls.Refresh()

        ' *********** conexion a INFORMIX 7 ***************************
        gsDSN = goProfile.msDSN
        gsUID = "gestion"
        gsPWD = "buitre"
        ' *************************************************************

        gconADO = New OdbcConnection()
        'gconADO.ConnectionString = "Provider=ifxoledbc;password=buitre;User ID=gestion;" & _
        '                         "Data Source=basedatos@prueba;Persist Security Info=true"
        gconADO.ConnectionString = "DSN=" & gsDSN
        Try
            gconADO.Open()
        Catch ex As Exception
            MsgBox("Imposible Conectar - " & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Conexion")
            Me.Close()
        End Try

        ' *********** futura conexion a SQL - SERVER *********
        'gconsql.ConnectionString = "user id=gestion;password=buitre;initial catalog=basedatos;data source=10.0.0.105;Connect Timeout=30"
        'gconSql.Open()
        ' ****************************************************

        lbls.Text = "Conexion finalizada con exito"
        lbls.Refresh()
        Me.Close()

    End Sub

End Class
