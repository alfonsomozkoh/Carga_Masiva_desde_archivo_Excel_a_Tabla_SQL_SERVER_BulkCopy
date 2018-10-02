<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CargaMasivaDesdeExcel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TxtArchivo = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Log = New System.Windows.Forms.ListBox()
        Me.txtCadena = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.BtnConn = New System.Windows.Forms.Button()
        Me.txtTestConn = New System.Windows.Forms.TextBox()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CbxTablaDestino = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'TxtArchivo
        '
        Me.TxtArchivo.Location = New System.Drawing.Point(157, 105)
        Me.TxtArchivo.Name = "TxtArchivo"
        Me.TxtArchivo.Size = New System.Drawing.Size(570, 20)
        Me.TxtArchivo.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(12, 105)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(136, 20)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Selecciona archivo Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Log
        '
        Me.Log.FormattingEnabled = True
        Me.Log.Location = New System.Drawing.Point(12, 160)
        Me.Log.Name = "Log"
        Me.Log.Size = New System.Drawing.Size(715, 368)
        Me.Log.TabIndex = 2
        '
        'txtCadena
        '
        Me.txtCadena.Location = New System.Drawing.Point(157, 27)
        Me.txtCadena.Name = "txtCadena"
        Me.txtCadena.Size = New System.Drawing.Size(570, 20)
        Me.txtCadena.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Ingrese cadena de conexión"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(54, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Nombre tabla SQL"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(12, 131)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(136, 23)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "Realizar carga masiva"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BtnConn
        '
        Me.BtnConn.Location = New System.Drawing.Point(12, 52)
        Me.BtnConn.Name = "BtnConn"
        Me.BtnConn.Size = New System.Drawing.Size(136, 20)
        Me.BtnConn.TabIndex = 10
        Me.BtnConn.Text = "Probar conexión"
        Me.BtnConn.UseVisualStyleBackColor = True
        '
        'txtTestConn
        '
        Me.txtTestConn.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtTestConn.Location = New System.Drawing.Point(157, 53)
        Me.txtTestConn.Name = "txtTestConn"
        Me.txtTestConn.ReadOnly = True
        Me.txtTestConn.Size = New System.Drawing.Size(493, 20)
        Me.txtTestConn.TabIndex = 11
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(184, 131)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(105, 23)
        Me.BtnClear.TabIndex = 12
        Me.BtnClear.Text = "Limpiar campos"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(241, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(486, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "EJEMPLO: Data Source=.;Initial Catalog=TuDB;Persist Security Info=True;User ID=sa" & _
    ";Password=xxxx"
        '
        'CbxTablaDestino
        '
        Me.CbxTablaDestino.FormattingEnabled = True
        Me.CbxTablaDestino.Location = New System.Drawing.Point(157, 78)
        Me.CbxTablaDestino.Name = "CbxTablaDestino"
        Me.CbxTablaDestino.Size = New System.Drawing.Size(493, 21)
        Me.CbxTablaDestino.TabIndex = 14
        '
        'CargaMasivaDesdeExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(739, 547)
        Me.Controls.Add(Me.CbxTablaDestino)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.txtTestConn)
        Me.Controls.Add(Me.BtnConn)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCadena)
        Me.Controls.Add(Me.Log)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TxtArchivo)
        Me.Name = "CargaMasivaDesdeExcel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtArchivo As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Log As System.Windows.Forms.ListBox
    Friend WithEvents txtCadena As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BtnConn As System.Windows.Forms.Button
    Friend WithEvents txtTestConn As System.Windows.Forms.TextBox
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CbxTablaDestino As System.Windows.Forms.ComboBox

End Class
