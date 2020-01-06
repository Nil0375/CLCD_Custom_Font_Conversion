Public Class Form1
    Dim butt(7, 4) As Button
    Dim butthex(7, 4) As Byte, line As Byte = 0
    Dim linestr(7) As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        linestr(0) = "" : linestr(1) = "" : linestr(2) = "" : linestr(3) = "" : linestr(4) = "" : linestr(5) = "" : linestr(6) = ""
        For j As Byte = 0 To 7
            For i As Integer = 0 To 4
                butt(j, i) = New Button
                With butt(j, i)
                    .Cursor = Cursors.Cross
                    .BackColor = Color.White
                    .FlatStyle = FlatStyle.Flat
                    .ForeColor = Color.DarkGray
                    .Left = 75 + (35 * i)
                    .Top = 25 + (40 * j)
                    .Width = 35
                    .Height = 40
                End With
                Me.Controls.Add(butt(j, i))
                AddHandler butt(j, i).Click, AddressOf A_ClickHandler

                butt(j, i).Show()
            Next
        Next
        linestr(7) = ";"
    End Sub

    Private Sub A_ClickHandler(sender As Object, e As EventArgs) Handles Me.Click
        'Throw New NotImplementedException()
        Try
            If CType(sender, Button).BackColor = Color.Red Then
                CType(sender, Button).BackColor = Color.White
            Else
                CType(sender, Button).BackColor = Color.Red
            End If
            For j As Byte = 0 To 7
                For i As Byte = 0 To 4
                    If butt(j, i).BackColor = Color.Red Then
                        butthex(j, i) = 1
                    Else
                        butthex(j, i) = 0
                    End If
                Next
            Next
            Dim hexbuffer(7) As String
            For i As Byte = 0 To 7
                hexbuffer(i) = Convert.ToString((butthex(i, 0) * 16) + (butthex(i, 1) * 8) + (butthex(i, 2) * 4) + (butthex(i, 3) * 2) + (butthex(i, 4)), 16)
            Next
            Select Case line
                Case 0
                    linestr(0) = "{" & "0x" & hexbuffer(0) & "," & "0x" & hexbuffer(1) & "," & "0x" & hexbuffer(2) & "," & "0x" & hexbuffer(3) & "," & "0x" & hexbuffer(4) & "," & "0x" & hexbuffer(5) & "," & "0x" & hexbuffer(6) & "," & "0x" & hexbuffer(7) & "}"
                Case 1 To 6
                    linestr(line) = "," & vbCrLf & "{" & "0x" & hexbuffer(0) & "," & "0x" & hexbuffer(1) & "," & "0x" & hexbuffer(2) & "," & "0x" & hexbuffer(3) & "," & "0x" & hexbuffer(4) & "," & "0x" & hexbuffer(5) & "," & "0x" & hexbuffer(6) & "," & "0x" & hexbuffer(7) & "}"
                Case 7
                    linestr(7) = "," & vbCrLf & "{" & "0x" & hexbuffer(0) & "," & "0x" & hexbuffer(1) & "," & "0x" & hexbuffer(2) & "," & "0x" & hexbuffer(3) & "," & "0x" & hexbuffer(4) & "," & "0x" & hexbuffer(5) & "," & "0x" & hexbuffer(6) & "," & "0x" & hexbuffer(7) & "}" & linestr(7)
            End Select
            TextBox1.Text = linestr(0) & linestr(1) & linestr(2) & linestr(3) & linestr(4) & linestr(5) & linestr(6) & linestr(7)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For j As Byte = 0 To 7
            For i As Byte = 0 To 4
                butt(j, i).BackColor = Color.White
            Next
        Next
        If line = 0 Then
            linestr(0) = "{" & linestr(0)
            linestr(7) = "}" & linestr(7)
        End If
        line = line + 1
        Dim c As Boolean
        If line > 7 Then
            c = MsgBox("可字定義字型已達上限" & vbCrLf & "請先儲存檔案", 0 + 48, "Error")
            line = 0
        End If
        If c = True Then
            Dim FileNum As UInt16
            Dim fileroad As String
            Dim saveFileDialog1 As New SaveFileDialog()
            saveFileDialog1.Filter = "Text|*.txt"
            saveFileDialog1.Title = "Save an txt"
            saveFileDialog1.ShowDialog()
            If saveFileDialog1.FileName <> "" Then
                fileroad = saveFileDialog1.FileName

                FileNum = FreeFile()
                FileOpen(FileNum, fileroad, OpenMode.Output)
                PrintLine(FileNum, linestr(0) & linestr(1) & linestr(2) & linestr(3) & linestr(4) & linestr(5) & linestr(6) & linestr(7))
                FileClose(FileNum)
            End If
            For j As Byte = 0 To 7
                For i As Byte = 0 To 4
                    butt(j, i).BackColor = Color.White
                    linestr(j) = ""
                Next
            Next
            TextBox1.Text = ""
            line = 0
        End If
        TextBox1.Text = ""
        TextBox1.Text = linestr(0) & linestr(1) & linestr(2) & linestr(3) & linestr(4) & linestr(5) & linestr(6) & linestr(7)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim FileNum As UInt16
        Dim fileroad As String
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Text|*.txt"
        saveFileDialog1.Title = "Save an txt"
        saveFileDialog1.ShowDialog()
        If saveFileDialog1.FileName <> "" Then
            fileroad = saveFileDialog1.FileName

            FileNum = FreeFile()
            FileOpen(FileNum, fileroad, OpenMode.Output)
            PrintLine(FileNum, linestr(0) & linestr(1) & linestr(2) & linestr(3) & linestr(4) & linestr(5) & linestr(6) & linestr(7))
            FileClose(FileNum)
        End If
        For j As Byte = 0 To 7
            For i As Byte = 0 To 4
                butt(j, i).BackColor = Color.White
                linestr(j) = ""
            Next
        Next
        TextBox1.Text = ""
        line = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For j As Byte = 0 To 7
            For i As Byte = 0 To 4
                butt(j, i).BackColor = Color.White
            Next
        Next
        TextBox1.Text = ""
        line = 0
        For i As Integer = 0 To 7
            linestr(i) = ""
        Next
        linestr(7) = ";"
    End Sub
End Class
