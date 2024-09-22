using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolCreator
{
    internal class TemplateCreator
    {
        // Path to the Word file
        private string filePath, liquid, sample;
        private int pressure;
        private double interval;

        public TemplateCreator(string filePath,string liquid, string sample, int pressure, double interval)
        {
            this.filePath = filePath;
            this.liquid = liquid;
            this.sample = sample;
            this.pressure = pressure;
            this.interval = interval;
        }

        public void Create()
        {
            // Create a Word document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Create the main document components
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                DateTime date = DateTime.Now;
                double[] tSat = [61, 74.24, 84.04, 98.84, 109.6, 119.31];

                // Add data about the experiment (date, liquid, pressure, saturation temperature)
                //Title
                Paragraph paragraph = body.AppendChild(new Paragraph());
                // Create paragraph properties
                ParagraphProperties paragraphProperties = new ParagraphProperties();

                Justification justification = new Justification() { Val = JustificationValues.Center }; // Center the text
                paragraphProperties.Append(justification);
                paragraph.Append(paragraphProperties);

                Run run = paragraph.AppendChild(new Run());

                RunProperties runProperties = run.AppendChild(new RunProperties());
                runProperties.AppendChild(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
                runProperties.AppendChild(new FontSize() { Val = "32" }); // 32 is the font size in half points

                run.AppendChild(new Text($"Протокол измерений"));

                paragraph.AppendChild(new Break() { Type = BreakValues.TextWrapping });

                //Date, liquid
                Paragraph paragraph1 = body.AppendChild(new Paragraph());
                Run run1 = paragraph1.AppendChild(new Run());

                RunProperties runProperties1 = run1.AppendChild(new RunProperties());
                runProperties1.AppendChild(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
                runProperties1.AppendChild(new FontSize() { Val = "24" }); // 24 is the font size in half points

                run1.AppendChild(new Text($"Дата: {date.ToShortDateString()}"));
                for (int i = 0; i < 5; i++)
                {
                    run1.AppendChild(new TabChar());
                }
                run1.AppendChild(new Text($"Жидкость: {liquid}"));

                //Pressure, temp
                Paragraph paragraph2 = body.AppendChild(new Paragraph());
                Run run2 = paragraph2.AppendChild(new Run());

                RunProperties runProperties2 = run2.AppendChild(new RunProperties());
                runProperties2.AppendChild(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
                runProperties2.AppendChild(new FontSize() { Val = "24" }); // 24 is the font size in half points

                run2.AppendChild(new Text($"Давление: {pressure} атм"));
                for (int i = 0; i < 5; i++)
                {
                    run2.AppendChild(new TabChar());
                }
                run2.AppendChild(new Text($"Температура насыщения: {tSat[pressure - 1]} °C"));

                //Sample, interval
                Paragraph paragraph3 = body.AppendChild(new Paragraph());
                Run run3 = paragraph3.AppendChild(new Run());

                RunProperties runProperties3 = run3.AppendChild(new RunProperties());
                runProperties3.AppendChild(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
                runProperties3.AppendChild(new FontSize() { Val = "24" }); // 24 is the font size in half points

                run3.AppendChild(new Text($"Образец: {sample}"));
                for (int i = 0; i < 5; i++)
                {
                    run3.AppendChild(new TabChar());
                }
                run3.AppendChild(new Text($"Интервал изменения напряжения: {interval} В"));

                paragraph3.AppendChild(new Break() { Type = BreakValues.TextWrapping });

                Paragraph paragraph4 = body.AppendChild(new Paragraph());
                // Create paragraph properties
                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                Justification justification2 = new Justification() { Val = JustificationValues.Center }; // Center the text
                paragraphProperties2.Append(justification2);
                paragraph4.Append(paragraphProperties2);

                Run run4 = paragraph4.AppendChild(new Run());

                RunProperties runProperties4 = run4.AppendChild(new RunProperties());
                runProperties4.AppendChild(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
                runProperties4.AppendChild(new FontSize() { Val = "32" }); // 32 is the font size in half points

                run4.AppendChild(new Text($"Таблица измерений"));


                string[] header = ["№", "U, В", "I, А", "Время", "Комментарии"];
                // Create a table
                Table table = new Table();
                TableProperties tableProperties = new TableProperties(
                    new TableJustification() { Val = TableRowAlignmentValues.Center },
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick) }
                    )
                );
                table.AppendChild(tableProperties);

                // Create table rows
                for (int rowNumber = 0; rowNumber < 1; rowNumber++)
                {
                    TableRow row = new TableRow();

                    // Create cells for each row
                    for (int cellNumber = 0; cellNumber < 5; cellNumber++)
                    {
                        TableCell cell = new TableCell();

                        // Add text to the cell
                        Paragraph paragraph5 = new Paragraph();
                        ParagraphProperties paragraphProperties3 = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                        paragraph5.Append(paragraphProperties3);
                        Run run5 = new Run();
                        Text text = new Text($"{header[cellNumber]}");
                        run5.Append(text);
                        paragraph5.Append(run5);
                        cell.Append(paragraph5);

                        //else
                        //{
                        //    Paragraph paragraph6 = new Paragraph();
                        //    ParagraphProperties paragraphProperties4 = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                        //    paragraph6.Append(paragraphProperties4);
                        //    Run run6 = new Run();
                        //    Text text = new Text("");
                        //    run6.Append(text);
                        //    paragraph6.Append(run6);
                        //    cell.Append(paragraph6);
                        //}

                        // Add dimensions and padding to the cell
                        TableCellProperties cellProperties = new TableCellProperties(
                            new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2000" }, // Ширина ячейки (в том числе в единицах DXA)
                            new TableCellMargin(
                                new TopMargin() { Width = "0" }, // Top padding
                                new BottomMargin() { Width = "0" }, // Bottom padding
                                new LeftMargin() { Width = "100" }, // Left indent
                                new RightMargin() { Width = "100" } // Right indent
                            )
                        );
                        cell.Append(cellProperties);
                        // Add a cell to the row
                        row.Append(cell);
                    }

                    // Add a row to the table
                    table.Append(row);
                }

                // Add a table to the document
                body.Append(table);

                // Save changes
                wordDocument.Save();
            }
        }

    }
}
