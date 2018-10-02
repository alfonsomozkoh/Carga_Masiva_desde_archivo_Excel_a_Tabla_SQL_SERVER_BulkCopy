
Imports System.Configuration
Imports System.Configuration.ConfigurationManager
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Transactions

Public Class CargaMasivaDesdeExcel

    Private Sub BtnConn_Click(sender As System.Object, e As System.EventArgs) Handles BtnConn.Click
        Dim TxtCadenaConexion As String = txtCadena.Text.ToString
        Dim NuevaConexion As String = String.Empty
        Dim Reader As SqlDataReader = Nothing
        AppSettings.Set("MiConexion", TxtCadenaConexion)
        NuevaConexion = AppSettings.Get("MiConexion")

        Try
            If txtCadena.Text = String.Empty Then
                txtTestConn.Text = ("Cadena vacia") : txtTestConn.BackColor = Color.Red : txtTestConn.ForeColor = Color.White
            Else
                Dim TestCon As New SqlConnection(NuevaConexion)
                TestCon.Open()
                Log.Items.Add("Conexión exitosa " & TestCon.State.ToString & ProductName)
                txtTestConn.Text = ("Conexión exitosa " & TestCon.State.ToString & " " & ProductName) : txtTestConn.BackColor = Color.Green : txtTestConn.ForeColor = Color.White
                Button1.Enabled = True
                Dim cmd As New SqlCommand("SELECT * FROM sysobjects WHERE xtype = 'U'", TestCon)
                Try
                    Reader = cmd.ExecuteReader()
                    CbxTablaDestino.Items.Clear()
                    While (Reader.Read())
                        CbxTablaDestino.Items.Add(Reader.GetString(0))
                    End While
                Catch ex As Exception
                    Log.Items.Add("Error al llenar tabla: " & ex.Message.ToString)
                Finally
                    If Reader Is Nothing Then
                        Reader.Close()
                    End If
                End Try

                TestCon.Close()
            End If
        Catch ex As SqlException
            txtTestConn.Text = ("Error de Conexión ") : txtTestConn.BackColor = Color.Red : txtTestConn.ForeColor = Color.White
            Log.Items.Add("Error de conexión " & ex.Message.ToString)
        End Try
    End Sub

    Public Sub Clear()
        txtCadena.Clear() : txtTestConn.Clear() : TxtArchivo.Clear() : CbxTablaDestino.Items().Clear() : CbxTablaDestino.Text = String.Empty : txtTestConn.BackColor = DefaultBackColor : txtTestConn.ForeColor = Color.Black : Button1.Enabled = False : Button2.Enabled = False
    End Sub

    Private Sub BtnClear_Click(sender As System.Object, e As System.EventArgs) Handles BtnClear.Click
        Me.Clear()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Excel 2007 - 2016(*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls|Todos los archivos (*.*)|*.*"
        OpenFileDialog1.Multiselect = False
        Me.OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.CheckFileExists = True Then
            Button2.Enabled = True
        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Me.TxtArchivo.Text = Me.OpenFileDialog1.FileName
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim oArchivo As String = OpenFileDialog1.FileName
        'Path temporal para guardar archivo 
        Dim excelPath As String = Path.Combine("c:\temp", Path.GetFileName(OpenFileDialog1.FileName))
        'Informacion para extraer extension 
        Dim fInfo As New FileInfo(excelPath)
        Dim connString As String = String.Empty
        'Extension del archivo
        Dim xls03 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES"
        Dim xlsx07UP As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'"
        Dim extension As String = fInfo.Extension
        Dim TablaCaptura As String = CbxTablaDestino.Text.ToString

        Try
            If oArchivo IsNot Nothing And txtCadena IsNot String.Empty And CbxTablaDestino IsNot String.Empty Then
                System.IO.File.Copy(oArchivo, excelPath)
                Log.Items.Add("Archvivos copiados en" & excelPath.ToString)
                Select Case extension
                    Case ".xls"
                        connString = xls03
                        Exit Select
                    Case ".xlsx"
                        'Según la versión Excel 07 or higher
                        connString = xlsx07UP
                        Exit Select
                End Select
                Log.Items.Add("Extensión de archivo " & fInfo.Name.ToString & " " & fInfo.Extension.ToString)
            End If
            connString = String.Format(connString, excelPath)
            Using excel_con As New OleDbConnection(connString)
                excel_con.Open()
                Dim Hoja1 As String = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing).Rows(0)("TABLE_NAME").ToString()
                Dim dtExcelData As New DataTable()
                Using oda As New OleDbDataAdapter((Convert.ToString("SELECT * FROM [") & Hoja1) + "]", excel_con)
                    oda.Fill(dtExcelData)
                End Using
                Log.Items.Add("Columnas archivo = " & dtExcelData.Columns.Count.ToString & " Filas = " & dtExcelData.Rows.Count.ToString)
                excel_con.Close()

                Using transacXLSX As New TransactionScope()
                    Try
                        Dim conString As String = AppSettings.Get("MiConexion")
                        Using con As New SqlConnection(conString)
                            Using sqlBulkCopy As New SqlBulkCopy(con)
                                'Seleccionamos la tabla donde se ha de insertar 
                                sqlBulkCopy.DestinationTableName = TablaCaptura
                                con.Open()
                                sqlBulkCopy.WriteToServer(dtExcelData)
                                con.Close()
                                Kill(excelPath)
                                Me.Clear()
                            End Using
                        End Using
                        transacXLSX.Complete()
                        Log.Items.Add("Transacción completada " & fInfo.Name.ToString & " Insertado en " & TablaCaptura.ToString & " Con éxito ")
                    Catch ex As Exception
                        transacXLSX.Dispose()
                        'Si falla lo movemos a excel No
                        Kill(excelPath) : Log.Items.Add(" Transanccion NO completada " & ex.Message.ToString)
                    End Try
                End Using
            End Using
        Catch ex As Exception
            'Kill(excelPath)
            Log.Items.Add("Error: " & ex.Message.ToString)
        End Try
    End Sub
End Class
