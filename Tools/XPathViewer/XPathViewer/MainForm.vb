Imports System.Xml
Imports System.IO


Public Class MainForm

    Private Sub btnApp_Click(sender As Object, e As EventArgs) Handles btnApp.Click
        txtResult.Text = String.Empty
        Try
            Dim xmlDoc As XmlDocument = New XmlDocument()
            Dim nodes As XmlNodeList
            xmlDoc.LoadXml(txtSource.Text)
            nodes = xmlDoc.SelectNodes(txtXPath.Text)

            If nodes Is Nothing Then
                Return
            End If

            If nodes.Count > 1 Then
                txtResult.Text += String.Format("select {0} nodes:", nodes.Count) + vbCrLf + vbCrLf
            Else
                txtResult.Text += String.Format("select {0} node:", nodes.Count) + vbCrLf + vbCrLf
            End If

            For Each x As XmlElement In nodes
                txtResult.Text += x.OuterXml + vbCrLf
                txtResult.Text += New String("-", 80) + vbCrLf
            Next
        Catch ex As Exception
            txtResult.Text = ex.Message
        End Try


    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        Dim dialog As OpenFileDialog = New OpenFileDialog
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Using sr As New StreamReader(dialog.FileName)
                txtSource.Text = sr.ReadToEnd()
            End Using
        End If
    End Sub
End Class
