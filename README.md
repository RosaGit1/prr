Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports Word = Microsoft.Office.Interop.Word

Public Class Form1
    Dim myArray(,) As Object
    Dim myArray2(,) As Object
    Dim myArray21(,) As Object

    Dim shript As String
    '

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RichTextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Label3.Text = 0

        '' Установить выбранный элемент программно
        'ComboBox1.SelectedItem = "12" ' Или использовать ComboBox1.SelectedIndex = 1
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        Dim cleanedText As String = RichTextBox1.Text.Replace(" ", "").Replace(vbCrLf, "")

        ' Подсчитываем количество символов в очищенном тексте
        Dim characterCount As Integer = cleanedText.Length

        ' Выводим количество символов в MessageBox
        Label3.Text = cleanedText.Length
        ' MessageBox.Show("Количество символов в RichTextBox (без пробелов и переводов строки): " & characterCount)
    End Sub

    Private Sub ОткрытьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьToolStripMenuItem.Click

        ' Настройка диалога выбора файла
        OpenFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        OpenFileDialog1.Title = "Выберите текстовый файл"

        ' Открыть диалог выбора файла
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Получить путь к выбранному файлу
            Dim filePath As String = OpenFileDialog1.FileName

            ' Прочитать содержимое файла и отобразить его в TextBox
            Try
                Using reader As New System.IO.StreamReader(filePath)
                    RichTextBox1.Text = reader.ReadToEnd()
                End Using
            Catch ex As Exception
                MessageBox.Show("Ошибка при чтении файла: " & ex.Message)
            End Try
        End If

    End Sub

    Private Sub ЗакрытьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗакрытьToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            ' Установить выбранный цвет как цвет фона формы
            RichTextBox1.BackColor = ColorDialog1.Color
        End If

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Открыть диалог выбора шрифта
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            ' Установить выбранный шрифт для выделенного текста
            If RichTextBox1.SelectionLength > 0 Then
                RichTextBox1.SelectionFont = FontDialog1.Font
            Else
                MessageBox.Show("Пожалуйста, выделите текст для изменения шрифта.")
            End If
        End If
    End Sub

    Private Sub СохранитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СохранитьToolStripMenuItem.Click
        SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        SaveFileDialog1.Title = "Сохранить текстовый файл"
        SaveFileDialog1.FileName = "document.txt"

        ' Открыть диалог сохранения файла
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Получить путь к файлу для сохранения
            Dim filePath As String = SaveFileDialog1.FileName

            ' Записать содержимое TextBox в файл
            Try
                Using writer As New System.IO.StreamWriter(filePath)
                    writer.Write(RichTextBox1.Text)
                End Using
                MessageBox.Show("Файл успешно сохранен!")
            Catch ex As Exception
                MessageBox.Show("Ошибка при сохранении файла: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Создаем новый экземпляр приложения Word
        Dim wordApp As New Word.Application
        ' Создаем новый документ
        Dim doc As Word.Document = wordApp.Documents.Add()

        ' Добавляем текст из RichTextBox в документ
        doc.Content.Text = RichTextBox1.Text

        ' Проверяем орфографию документа
        doc.CheckSpelling()

        ' Возвращаем исправленный текст обратно в RichTextBox
        RichTextBox1.Text = doc.Content.Text

        ' Закрываем документ без сохранения изменений
        doc.Close(False)

        ' Закрываем приложение Word
        wordApp.Quit()

        ' Освобождаем COM-объекты
        releaseObject(doc)
        releaseObject(wordApp)
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub



    Private Sub НайтиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НайтиToolStripMenuItem.Click
        ' Получаем текст для поиска из TextBox1
        Dim searchText As String = TextBox1.Text

        ' Проверяем, что текст для поиска не пустой
        If Not String.IsNullOrEmpty(searchText) Then
            ' Сброс предыдущих выделений
            RichTextBox1.SelectAll()
            RichTextBox1.SelectionBackColor = RichTextBox1.BackColor
            RichTextBox1.DeselectAll()

            ' Начальная позиция поиска
            Dim startIndex As Integer = 0

            ' Цикл поиска и выделения всех вхождений текста
            While startIndex < RichTextBox1.Text.Length
                ' Ищем текст
                startIndex = RichTextBox1.Find(searchText, startIndex, RichTextBoxFinds.None)

                ' Если текст найден, выделяем его
                If startIndex <> -1 Then
                    RichTextBox1.SelectionStart = startIndex
                    RichTextBox1.SelectionLength = searchText.Length
                    RichTextBox1.SelectionBackColor = Color.Yellow
                    ' Переход к следующему символу после найденного текста
                    startIndex += searchText.Length
                Else
                    ' Если текст не найден, выходим из цикла
                    Exit While
                End If
            End While
        Else
            MessageBox.Show("Пожалуйста, введите текст для поиска.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim colorDialog As New ColorDialog()
        If colorDialog.ShowDialog() = DialogResult.OK Then
            ' Устанавливаем выбранный цвет для выделенного текста в RichTextBox1
            RichTextBox1.SelectionColor = colorDialog.Color
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
