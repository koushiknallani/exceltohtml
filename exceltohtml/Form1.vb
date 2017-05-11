
Imports System.Text
Imports Microsoft.Office
Imports Microsoft.Office.Interop

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Select EXCEL Sheet"
        OpenFileDialog1.Filter = "Excel Worksheets|*.xls;*.xlt;*.xlm;*.xlsx|All Files|*.*"
        OpenFileDialog1.FileName = "Ecxel Sheet"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "Select Destination"
        OpenFileDialog1.Filter = "HTML|*.html|TEXT File|*.txt|All Files|*.*"
        OpenFileDialog1.FileName = ".html"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox2.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Excel Path/Sheet Name/HTML Path is NULL/Blank")
            Exit Sub

        End If
        Try
            Dim builder As New StringBuilder
            Dim slno, path, name, roll As String
            Dim output As String = ""
            'open
            Dim myexcel As New Excel.Application
            myexcel.Workbooks.Open(TextBox1.Text)

            'Extract info
            myexcel.Sheets(TextBox3.Text).activate()
            myexcel.Range("A2").Activate()
            'Dim thisrow As New excelrows

            Do
                If myexcel.ActiveCell.Value > Nothing Or myexcel.ActiveCell.Text > Nothing Then
                    slno = myexcel.ActiveCell.Value
                    myexcel.ActiveCell.Offset(0, 1).Activate()

                    path = myexcel.ActiveCell.Value
                    myexcel.ActiveCell.Offset(0, 1).Activate()

                    name = myexcel.ActiveCell.Value
                    myexcel.ActiveCell.Offset(0, 1).Activate()

                    roll = myexcel.ActiveCell.Value

                    'excelrowlist.Add(thisrow)

                    builder.Append("<tr>")
                    builder.AppendLine()
                    slno = "<td>" + slno + "</td>"

                    builder.Append(slno)
                    builder.AppendLine()
                    path = "<td><img src=""" + path + """  width=""57"" height=""55"" /></td>"

                    builder.Append(path)
                    builder.AppendLine()
                    name = "<td>" + name + "</td>"

                    builder.Append(name)
                    builder.AppendLine()
                    roll = "<td colspan=""2"">" + roll + "</td>"

                    builder.Append(roll)
                    builder.AppendLine()
                    builder.Append("</tr>")

                    builder.AppendLine()
                    Dim s As String = builder.ToString

                    output = s
                    myexcel.ActiveCell.Offset(1, -3).Activate()


                Else
                    Exit Do
                End If
            Loop

            'close
            myexcel.Workbooks.Close()
            myexcel = Nothing

            'Copy code to html 
            ' MsgBox(output)
            If System.IO.File.Exists(TextBox2.Text) = True Then

                Dim objWriter As New System.IO.StreamWriter(TextBox2.Text, True)

                objWriter.Write(output)
                objWriter.Close()
                MsgBox("Conversion Completed")

            Else

                MsgBox("Select Correct Destination File")

            End If
        Catch ex As ArgumentException
            MsgBox(ex.Message)
        Catch em As Exception
            MsgBox(em.Message)
        End Try

        


    End Sub

    
End Class
