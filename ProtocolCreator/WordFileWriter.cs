using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ProtocolCreator
{
    internal class WordFileWriter
    {
        private string filePath;
        private double interval;
        private int recordNumber = 1;    // Счетчик записей (номер записи)
        private double voltage = 40;        // Начальное напряжение (вольты)
        private bool NoError;

        public WordFileWriter(string filePath, double interval)
        {
            this.filePath = filePath;
            this.interval = interval;
        }

        // Метод для записи данных в существующую таблицу Word файла
        public void RecordMode()
        {
            try
            {
                // Открываем существующий Word файл
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    // Находим главное тело документа
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    // Ищем таблицу в документе (предполагаем, что таблица первая)
                    Table table = body.Elements<Table>().FirstOrDefault();

                    if (table != null)
                    {
                        // Добавляем новую строку в таблицу
                        TableRow newRow = new TableRow();

                        // Создаем и заполняем 5 ячеек с выравниванием текста по центру
                        newRow.Append(CreateCenteredCell(recordNumber.ToString()));   // Номер записи
                        newRow.Append(CreateCenteredCell(voltage.ToString()));       // Напряжение
                        newRow.Append(CreateCenteredCell(""));                       // Пустая ячейка
                        newRow.Append(CreateCenteredCell(DateTime.Now.ToString("HH:mm:ss")));  // Время записи
                        newRow.Append(CreateCenteredCell(""));                       // Пустая ячейка

                        // Добавляем строку в таблицу
                        table.Append(newRow);

                        // Сохраняем изменения
                        wordDoc.MainDocumentPart.Document.Save();
                    }

                    // Увеличиваем номер записи и напряжение для следующего раза
                    recordNumber++;
                    voltage += interval;
                }

                NoError = true;
            }
            catch (Exception ex)
            {
                NoError = false;
            }
        }

        public void RecordCHF()
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    Table table = body.Elements<Table>().FirstOrDefault();

                    if (table != null)
                    {
                        // Получаем последнюю строку таблицы
                        TableRow lastRow = table.Elements<TableRow>().LastOrDefault();

                        if (lastRow != null)
                        {
                            // Получаем 5-ю ячейку в последней строке
                            TableCell fifthCell = lastRow.Elements<TableCell>().ElementAtOrDefault(4); // Индекс 4 для 5 колонки

                            if (fifthCell != null)
                            {
                                // Очищаем предыдущий контент ячейки (если есть)
                                fifthCell.RemoveAllChildren<Paragraph>();

                                // Создаем параграф с центровкой
                                Paragraph paragraph = new Paragraph(
                                    new ParagraphProperties(
                                        new Justification() { Val = JustificationValues.Center }), // Центровка текста
                                    new Run(new Text("Начало КТП")),
                                    new Break(),               // Перенос строки
                                    new Run(new Text(DateTime.Now.ToString("HH:mm:ss")))); // Добавляем текущее время

                                // Добавляем параграф в ячейку
                                fifthCell.Append(paragraph);
                            }

                            // Сохраняем изменения в документе
                            wordDoc.MainDocumentPart.Document.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с файлом: {ex.Message}");
            }
        }

        // Метод для создания ячейки с текстом, выровненным по центру
        private TableCell CreateCenteredCell(string text)
        {
            // Создаем свойства для выравнивания текста по центру
            var paragraphProperties = new ParagraphProperties
            {
                Justification = new Justification { Val = JustificationValues.Center } // Выравнивание текста по центру
            };

            var paragraph = new Paragraph(paragraphProperties);
            paragraph.Append(new Run(new Text(text)));

            // Создаем свойства ячейки
            var tableCellProperties = new TableCellProperties
            {
                TableCellVerticalAlignment = new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center } // Вертикальное выравнивание по центру
            };

            // Возвращаем ячейку с заданными свойствами
            return new TableCell(paragraph) { TableCellProperties = tableCellProperties };
        }

        public bool IsOk()
        {
            return NoError;
        }
    }
}
