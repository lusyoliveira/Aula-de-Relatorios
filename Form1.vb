Public Class Form1
    Dim sql As String
    Dim tbClientes As ADODB.Recordset
    Dim wcpagina As Integer
    Dim wcimagem As Image
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sql = "Select * from tbclientes where nome <> '' order by nome"
        tbClientes = OpenRecordset(sql)
        If tbClientes.RecordCount = 0 Then
            MsgBox("Não existem clientes !")
            Exit Sub
        End If
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        wcpagina = 1

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim MYFONT As New Font("ARIAL", 8)
        Dim MYFONT3 As New Font("ARIAL", 10, FontStyle.Bold)
        Dim MYFONT2 As New Font("ARIAL", 12, FontStyle.Bold)
        Dim Pulou As Boolean = False

        Dim X1 As Single = e.MarginBounds.Left ' Variavel que controla a coluna
        Dim Y1 As Single = e.MarginBounds.Top  ' Variavel que controla a linha
        Dim Line As Single = MYFONT.GetHeight(e.Graphics) ' Pega o tamanho da linha a ser adicionada quando usar a myfont
        Dim Line2 As Single = MYFONT2.GetHeight(e.Graphics) ' Pega o tamanho da
        Dim Line3 As Single = MYFONT3.GetHeight(e.Graphics) ' Pega o tamanho da

        wcimagem = Image.FromFile("c:\temp\logoPioxiiPeq.jpg")
        e.Graphics.DrawImage(wcimagem, X1, Y1 - 30, 50, 50)
        e.Graphics.DrawString("Programa Exemplo", MYFONT2, Brushes.Black, X1 + 100, Y1)
        Y1 += Line2
        e.Graphics.DrawString("Relatório de Clientes", MYFONT3, Brushes.Black, X1 + 100, Y1)
        Y1 += Line3
        e.Graphics.DrawString("Matrícula", MYFONT3, Brushes.Black, X1, Y1)
        e.Graphics.DrawString("Nome", MYFONT3, Brushes.Black, X1 + 80, Y1)
        e.Graphics.DrawString("Endereço", MYFONT3, Brushes.Black, X1 + 300, Y1)
        Y1 += Line3
        e.Graphics.DrawLine(Pens.Black, X1, Y1, X1 + 600, Y1)
        While tbClientes.EOF = False
            e.Graphics.DrawString(tbClientes.Fields("codigo").Value.ToString, MYFONT, Brushes.Black, X1, Y1)
            e.Graphics.DrawString(Microsoft.VisualBasic.Left(tbClientes.Fields("nome").Value.ToString, 30), MYFONT, Brushes.Black, X1 + 80, Y1)
            e.Graphics.DrawString(tbClientes.Fields("endereco").Value.ToString, MYFONT, Brushes.Black, X1 + 300, Y1)
            Y1 += Line
            tbClientes.MoveNext()
            If Y1 >= e.MarginBounds.Bottom - 100 Then
                Pulou = True
                Exit While
            End If
        End While
        If Pulou Then
            e.HasMorePages = True
            Y1 = e.MarginBounds.Bottom
            e.Graphics.DrawLine(Pens.Black, X1, Y1, X1 + 600, Y1)
            Y1 += Line
            e.Graphics.DrawString("Página:" & wcpagina, MYFONT2, Brushes.Black, X1, Y1)
            wcpagina += 1
        End If
    End Sub

    Private Sub PrintPreviewDialog1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewDialog1.Load
    End Sub
End Class
