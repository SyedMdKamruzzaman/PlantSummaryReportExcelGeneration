using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using Cdr = DocumentFormat.OpenXml.Drawing.ChartDrawing;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PlantSummaryExcelGeneration
{
    public class GenerateExcel
    {
        WorkbookPart workbookPart = null;
        public static string ImageFile = AppDomain.CurrentDomain.BaseDirectory + "bangla_cat_logo.png";
        public MemoryStream GetExcelMemoryStream()
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var excel = SpreadsheetDocument.Create(memoryStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
                {
                    workbookPart = excel.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    uint sheetId = 1;
                    excel.WorkbookPart.Workbook.Sheets = new Sheets();
                    Sheets sheets = excel.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                    WorkbookStylesPart stylesPart = excel.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = GenerateStyleSheet();
                    stylesPart.Stylesheet.Save();

                    int generatorCount = 3;
                    int sheetCount = ++generatorCount;
                    string[] strSheetName = new string[] { "Plant", "2213979", "2213963", "2213969" };


                    for (int i = 0; i < sheetCount; i++)
                    {
                        string relationshipId = "rId" + (i + 1).ToString();
                        WorksheetPart wSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                        string sheetName = strSheetName[i];
                        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                        sheets.Append(sheet);

                        Worksheet worksheet = new Worksheet();

                        wSheetPart.Worksheet = worksheet;

                        SheetData sheetData = new SheetData();
                        worksheet.Append(sheetData);

                        Drawing drawing1 = new Drawing() { Id = relationshipId };
                        worksheet.Append(drawing1);

                        DrawingsPart drawingsPart1 = wSheetPart.AddNewPart<DrawingsPart>(relationshipId);

                        if (sheetId == 1)
                        {
                            InsertImage(wSheetPart, drawingsPart1, 0, 1, 2, 4, new FileStream(ImageFile, FileMode.Open));

                            //saCategories[0] = "Apple";
                            // saLegend[0] = "2213979";
                            //faChartData[0, 0] = 22;
                            //saLegend[1] = "2213963";
                            //faChartData[1, 0] = 20;
                            //saLegend[2] = "2213969";
                            //faChartData[2, 0] = 18;


                            AddToCell(sheetData, 8, 2, "F", CellValues.String, "AA Yarn");
                            AddToCell(sheetData, 1, 2, "I", CellValues.String, "16 February, 2020");

                            //create a MergeCells class to hold each MergeCell
                            MergeCells mergeCells = new MergeCells();

                            //append a MergeCell to the mergeCells for each set of merged cells
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B5:C5") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B6:C6") });

                            wSheetPart.Worksheet.InsertAfter(mergeCells, wSheetPart.Worksheet.Elements<SheetData>().First());


                            AddToCell(sheetData, 9, 5, "B", CellValues.String, "Your Industry");
                            AddToCell(sheetData, 9, 5, "C", CellValues.String, "");

                            AddToCell(sheetData, 5, 6, "B", CellValues.String, "Our Energy");

                            AddToCell(sheetData, 8, 8, "F", CellValues.String, "Plant Summary Report");

                            AddToCell(sheetData, 1, 10, "B", CellValues.String, "Number of generators");
                            AddToCell(sheetData, 0, 10, "G", CellValues.Number, "3");


                            AddToCell(sheetData, 1, 11, "B", CellValues.String, "Total running hours");
                            AddToCell(sheetData, 0, 11, "G", CellValues.String, "56 hours");


                            AddToCell(sheetData, 1, 12, "B", CellValues.String, "Total unplanned generator shutdown");
                            AddToCell(sheetData, 0, 12, "G", CellValues.String, "10 hours");

                            Dictionary<string, int> data;

                            data = new Dictionary<string, int>();
                            data.Add("2213979", 22);
                            data.Add("2213963", 20);
                            data.Add("2213969", 18);

                            string chartTitle = "Running Hours";

                            BuildChart(drawingsPart1, sheetName, data, chartTitle, 13, 1, 26, 8, 0, 25);

                            data = new Dictionary<string, int>();
                            data.Add("2213979", 2);
                            data.Add("2213963", 4);
                            data.Add("2213969", 6);

                            chartTitle = "Unplanned Shutdown Hours";

                            BuildChart(drawingsPart1, sheetName, data, chartTitle, 13, 10, 26, 18, 0, 7);


                        }

                        if (sheetId != 1)
                        {
                            AddToCell(sheetData, 8, 2, "B", CellValues.String, "Generator " + sheetName);
                            AddToCell(sheetData, 0, 2, "F", CellValues.String, "16 February, 2020");


                            AddToCell(sheetData, 1, 3, "B", CellValues.String, "Model");
                            AddToCell(sheetData, 0, 3, "D", CellValues.String, "CG-170-16");
                            AddToCell(sheetData, 1, 3, "F", CellValues.String, "Running hour");
                            AddToCell(sheetData, 0, 3, "J", CellValues.String, "22 hrs");
                            AddToCell(sheetData, 1, 3, "L", CellValues.String, "Overhauling due in");
                            AddToCell(sheetData, 0, 3, "O", CellValues.String, "1877 hrs");

                            AddToCell(sheetData, 1, 4, "B", CellValues.String, "Rated Kilowatt");
                            AddToCell(sheetData, 0, 4, "D", CellValues.Number, "1500");
                            AddToCell(sheetData, 1, 4, "F", CellValues.String, "Unplanned shutdown hours");
                            AddToCell(sheetData, 0, 4, "J", CellValues.String, "2 hrs");

                            string chartTitle = "Generator Utilisation";

                            Dictionary<string, int> data1 = new Dictionary<string, int>();
                            data1.Add("Running", 92);
                            data1.Add("Stopped", 8);


                            InsertPieChartInSpreadSheet(drawingsPart1, chartTitle, data1, 6, 0, 18, 6);

                            string title = "Power generation (Kilo Watt Hour)";
                            Dictionary<string, int> data = new Dictionary<string, int>();
                            data.Add("01/01/2019", 999);
                            data.Add("02/01/2019", 983);
                            data.Add("03/01/2019", 945);
                            data.Add("04/01/2019", 975);
                            data.Add("05/01/2019", 950);
                            data.Add("06/01/2019", 964);
                            data.Add("07/01/2019", 989);
                            data.Add("08/01/2019", 973);
                            data.Add("09/01/2019", 954);
                            data.Add("10/01/2019", 957);
                            data.Add("11/01/2019", 905);
                            data.Add("12/01/2019", 946);
                            data.Add("13/01/2019", 998);
                            data.Add("14/01/2019", 937);
                            data.Add("15/01/2019", 945);
                            data.Add("16/01/2019", 975);
                            data.Add("17/01/2019", 950);
                            data.Add("18/01/2019", 932);
                            data.Add("19/01/2019", 947);
                            data.Add("20/01/2019", 921);
                            data.Add("21/01/2019", 984);
                            data.Add("22/01/2019", 932);
                            data.Add("23/01/2019", 945);
                            data.Add("24/01/2019", 946);

                            InsertBarChartInSpreadsheet(worksheet, drawingsPart1, title, data, 6, 9, 20, 20);

                            AddToCell(sheetData, 1, 19, "B", CellValues.String, "Running(%)");
                            AddToCell(sheetData, 0, 19, "D", CellValues.Number, "85.54");
                            AddToCell(sheetData, 1, 19, "E", CellValues.String, "Stopped(%)");
                            AddToCell(sheetData, 0, 19, "G", CellValues.Number, "14.46");

                            //create a MergeCells class to hold each MergeCell
                            MergeCells mergeCells = new MergeCells();

                            //append a MergeCell to the mergeCells for each set of merged cells
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B23:I23") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B25:C25") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("D25:G25") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("H25:I25") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B26:C26") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("D26:G26") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("H26:I26") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("B27:C27") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("D27:G27") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("H27:I27") });

                            wSheetPart.Worksheet.InsertAfter(mergeCells, wSheetPart.Worksheet.Elements<SheetData>().First());

                            AddToCell(sheetData, 0, 22, "K", CellValues.String, "Monthly power generation");
                            AddToCell(sheetData, 7, 23, "B", CellValues.String, "Alerts");
                            AddToCell(sheetData, 0, 24, "K", CellValues.String, "Shutdown history table");
                            AddToCell(sheetData, 10, 25, "B", CellValues.String, "TIME");
                            AddToCell(sheetData, 10, 25, "D", CellValues.String, "ALERT");
                            AddToCell(sheetData, 10, 25, "H", CellValues.String, "PRIORITY");
                            AddToCell(sheetData, 0, 25, "K", CellValues.String, "Date");
                            AddToCell(sheetData, 0, 25, "L", CellValues.String, "Shutdown Time");
                            AddToCell(sheetData, 0, 25, "M", CellValues.String, "Duration");
                            AddToCell(sheetData, 0, 25, "N", CellValues.String, "Reason");

                            AddToCell(sheetData, 5, 26, "B", CellValues.String, "2/15/2019 16:48");
                            AddToCell(sheetData, 0, 26, "D", CellValues.String, "Crankcase pressure high");
                            AddToCell(sheetData, 5, 26, "H", CellValues.String, "High");

                            AddToCell(sheetData, 5, 27, "B", CellValues.String, "2/15/2019 10:25");
                            AddToCell(sheetData, 0, 27, "D", CellValues.String, "Generator winding U temperature high");
                            AddToCell(sheetData, 5, 27, "H", CellValues.String, "High");

                        }
                        sheetId++;
                    }


                    excel.Close();
                }

                MemoryStream excelMemoryStream = new MemoryStream(memoryStream.ToArray());
                excelMemoryStream.Seek(0, SeekOrigin.Begin);
                return excelMemoryStream;

            }

            
            
        }
        public void AddToCell(SheetData sheetData, UInt32Value styleIndex, UInt32 uint32rowIndex, string strColumnName, DocumentFormat.OpenXml.EnumValue<CellValues> CellDataType, string strCellValue)
        {
            Row row = new Row() { RowIndex = uint32rowIndex };
            Cell cell = new Cell();

            cell = new Cell() { StyleIndex = styleIndex };
            cell.CellReference = strColumnName + row.RowIndex.ToString();
            cell.DataType = CellDataType;
            cell.CellValue = new CellValue(strCellValue);
            row.AppendChild(cell);

            sheetData.Append(row);
        }

        public Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(
            new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),// Index 0 - The default font.
            new Font(new Bold(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 1 - The bold font.
            new Font(new Italic(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 2 - The Italic font.
            new Font(new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 3 - The Times Roman font. with 16 size
            new Font(new Bold(), new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 4 - The Times Roman font. with 16 size
            new Font(new Bold(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } }, new FontName() { Val = "Calibri" })  // Index 5 - The bold font.

            ),
            new Fills(
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 0 - The default fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.None }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 1 - The default fill of gray 125 (required)
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Gray125 }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 2 - The yellow fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
            )
            { PatternType = PatternValues.Solid }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 3 - The Blue fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "8EA9DB" } }
            )
            { PatternType = PatternValues.Solid })
            ),
            new Borders(
            new Border( // Index 0 - The default border.
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(),
            new DiagonalBorder()),
            new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder()),
                   new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Rgb = new HexBinaryValue() { Value = "70AD47" } }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder())
            ),
            new CellFormats(
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }, // Index 0 - The default cell style. If a cell does not have a style index applied it will use this style combination instead
            new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 1 - Bold
            new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 2 - Italic
            new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 3 - Times Roman
            new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true }, // Index 4 - Yellow Fill
            new CellFormat( // Index 5 - Alignment
            new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            )
            { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // Index 6 - Border
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 7 - Alignment
             { FontId = 1, FillId = 0, BorderId = 0, ApplyAlignment = true },

             new CellFormat() { FontId = 4, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 8 - Times Roman
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 2, ApplyFont = true }, // Index 9 - Bottom Border with Color 70AD47
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 10 - Alignment
             { FontId = 5, FillId = 3, BorderId = 0, ApplyAlignment = true }


             )
            ); // return
        }


        public void InsertImage(WorksheetPart sheet1, DrawingsPart drawingsPart2, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
        {
            GenerateImageDrawing(drawingsPart2, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);

            //Adding the image
            ImagePart imagePart1 = drawingsPart2.AddNewPart<ImagePart>("image/png", "rId41");
            imagePart1.FeedData(imageStream);
        }

        private static void GenerateImageDrawing(DrawingsPart drawingsPart, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            // drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            // TwoCellAnchor twoCellAnchor1 = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());

            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";// "38100";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = startRowIndex.ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "85725"; //multiply of 9525
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = endRowIndex.ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "9525";  //multiply of 9525

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId41", CompressionState = A.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };// { X = 1257300L, Y = 762000L };
            A.Extents extents1 = new A.Extents() { Cx = 2381250L, Cy = 628650L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            //System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ImageFile);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart.WorksheetDrawing = worksheetDrawing1;
        }


        public void BuildChart(DrawingsPart dp, string sheetName, Dictionary<string, int> data, string chartTitle, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, double minAxisLimit, double maxAxisLimit)
        {
            const uint cnAxisId1 = 1;
            const uint cnAxisId2 = 2;

            const int cnDataWidth = 1;
            const int cnDataHeight = 3;
            double[,] faChartData = new double[cnDataHeight, cnDataWidth];
            string[] saCategories = new string[cnDataWidth];
            string[] saLegend = new string[cnDataHeight];

            saCategories[0] = "Apple";
            saLegend[0] = "2213979";
            faChartData[0, 0] = 22;
            saLegend[1] = "2213963";
            faChartData[1, 0] = 20;
            saLegend[2] = "2213969";
            faChartData[2, 0] = 18;

            //  string cnWorksheetName = sheetName;

            ChartPart chartPart = dp.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new C.ChartSpace();


            //chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
            { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Bar) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            uint i = 0;

            barChart.Append(new C.Overlap() { Val = -20 });
            barChart.Append(new C.GapWidth() { Val = (UInt16Value)145U });
            barChart.Append(new C.VaryColors() { Val = false });

            AddBarChartTitle(chart, chartTitle);

            string[] strColorArray = new string[] { "3055A6", "4071CA", "A8B6DE" };//, "FFAADD", "3055A6", "4071CA", "A8B6DE", "FFAADD" };

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            uint j = 0;

            foreach (string key in data.Keys)
            {
                BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(
                    new Index() { Val = new UInt32Value(i) },
                    new Order() { Val = new UInt32Value(i) },
                    //new ChartShapeProperties(new SolidFill(new RgbColorModelHex() { Val = strColorArray[j] })),
                    new SeriesText(new NumericValue() { Text = key })));


                //barChartSeries.Append(new ChartShapeProperties(new SolidFill(new RgbColorModelHex() { Val = "4071CA" })));
                // j++;

                StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
                strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue("1"));

                NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
                numLit.Append(new FormatCode("General"));
                numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append(new NumericValue(data[key].ToString()));

                C.ChartShapeProperties chartShapeProperties18 = new C.ChartShapeProperties();

                A.GradientFill gradientFill7 = new A.GradientFill() { RotateWithShape = true };

                A.GradientStopList gradientStopList7 = new A.GradientStopList();

                A.GradientStop gradientStop19 = new A.GradientStop() { Position = 0 };

                A.SchemeColor schemeColor113 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                A.Shade shade15 = new A.Shade() { Val = 65000 };
                A.SaturationModulation saturationModulation20 = new A.SaturationModulation() { Val = 103000 };
                A.LuminanceModulation luminanceModulation61 = new A.LuminanceModulation() { Val = 102000 };
                A.Tint tint14 = new A.Tint() { Val = 94000 };

                schemeColor113.Append(shade15);
                schemeColor113.Append(saturationModulation20);
                schemeColor113.Append(luminanceModulation61);
                schemeColor113.Append(tint14);

                //  gradientStop19.Append(schemeColor113);

                //HexBinaryValue hexBinaryValue = new HexBinaryValue();
                //hexBinaryValue =  strColorArray[j];
                // string color = string.Empty;
                //j = 0;
                // color = strColorArray[j];
                // j = 1;
                RgbColorModelHex rgbColorModelHex = new RgbColorModelHex() { Val = strColorArray[j] };
                // j++;
                gradientStop19.Append(rgbColorModelHex);

                A.GradientStop gradientStop20 = new A.GradientStop() { Position = 50000 };

                A.SchemeColor schemeColor114 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                A.Shade shade16 = new A.Shade() { Val = 65000 };
                A.SaturationModulation saturationModulation21 = new A.SaturationModulation() { Val = 110000 };
                A.LuminanceModulation luminanceModulation62 = new A.LuminanceModulation() { Val = 100000 };
                A.Shade shade17 = new A.Shade() { Val = 100000 };

                schemeColor114.Append(shade16);
                schemeColor114.Append(saturationModulation21);
                schemeColor114.Append(luminanceModulation62);
                schemeColor114.Append(shade17);

                //gradientStop20.Append(schemeColor114);

                // HexBinaryValue hexBinaryValue1 = new HexBinaryValue();
                // hexBinaryValue1 = strColorArray[j];

                RgbColorModelHex rgbColorModelHex1 = new RgbColorModelHex() { Val = strColorArray[j] };
                gradientStop20.Append(rgbColorModelHex1);

                A.GradientStop gradientStop21 = new A.GradientStop() { Position = 100000 };

                //A.SchemeColor schemeColor115 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                //A.Shade shade18 = new A.Shade() { Val = 65000 };
                //A.LuminanceModulation luminanceModulation63 = new A.LuminanceModulation() { Val = 99000 };
                //A.SaturationModulation saturationModulation22 = new A.SaturationModulation() { Val = 120000 };
                //A.Shade shade19 = new A.Shade() { Val = 78000 };

                //schemeColor115.Append(shade18);
                //schemeColor115.Append(luminanceModulation63);
                //schemeColor115.Append(saturationModulation22);
                //schemeColor115.Append(shade19);

                // HexBinaryValue hexBinaryValue2 = new HexBinaryValue();
                //  hexBinaryValue2 = strColorArray[j];

                RgbColorModelHex rgbColorModelHex2 = new RgbColorModelHex() { Val = strColorArray[j] };

                //gradientStop21.Append(schemeColor115);

                gradientStop21.Append(rgbColorModelHex2);

                gradientStopList7.Append(gradientStop19);
                gradientStopList7.Append(gradientStop20);
                gradientStopList7.Append(gradientStop21);


                j++;

                A.LinearGradientFill linearGradientFill7 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

                gradientFill7.Append(gradientStopList7);
                gradientFill7.Append(linearGradientFill7);

                A.Outline outline41 = new A.Outline();
                A.NoFill noFill29 = new A.NoFill();

                outline41.Append(noFill29);

                A.EffectList effectList27 = new A.EffectList();

                A.OuterShadow outerShadow5 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

                A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex() { Val = "000000" };//strColorArray[i]
                A.Alpha alpha5 = new A.Alpha() { Val = 63000 };

                rgbColorModelHex15.Append(alpha5);

                outerShadow5.Append(rgbColorModelHex15);

                effectList27.Append(outerShadow5);


                chartShapeProperties18.Append(gradientFill7);
                chartShapeProperties18.Append(outline41);
                chartShapeProperties18.Append(effectList27);
                C.InvertIfNegative invertIfNegative4 = new C.InvertIfNegative() { Val = false };

                C.DataLabels dataLabels5 = new C.DataLabels();

                C.ChartShapeProperties chartShapeProperties19 = new C.ChartShapeProperties();
                A.NoFill noFill30 = new A.NoFill();

                A.Outline outline42 = new A.Outline();
                A.NoFill noFill31 = new A.NoFill();

                outline42.Append(noFill31);
                A.EffectList effectList28 = new A.EffectList();

                chartShapeProperties19.Append(noFill30);
                chartShapeProperties19.Append(outline42);
                chartShapeProperties19.Append(effectList28);

                C.TextProperties textProperties9 = new C.TextProperties();

                A.BodyProperties bodyProperties18 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
                A.ShapeAutoFit shapeAutoFit5 = new A.ShapeAutoFit();

                bodyProperties18.Append(shapeAutoFit5);
                A.ListStyle listStyle18 = new A.ListStyle();

                A.Paragraph paragraph18 = new A.Paragraph();

                A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties();

                A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

                A.SolidFill solidFill58 = new A.SolidFill();

                A.SchemeColor schemeColor116 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
                A.LuminanceModulation luminanceModulation64 = new A.LuminanceModulation() { Val = 75000 };
                A.LuminanceOffset luminanceOffset44 = new A.LuminanceOffset() { Val = 25000 };

                schemeColor116.Append(luminanceModulation64);
                schemeColor116.Append(luminanceOffset44);

                solidFill58.Append(schemeColor116);
                A.LatinFont latinFont16 = new A.LatinFont() { Typeface = "+mn-lt" };
                A.EastAsianFont eastAsianFont16 = new A.EastAsianFont() { Typeface = "+mn-ea" };
                A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

                defaultRunProperties11.Append(solidFill58);
                defaultRunProperties11.Append(latinFont16);
                defaultRunProperties11.Append(eastAsianFont16);
                defaultRunProperties11.Append(complexScriptFont16);

                paragraphProperties14.Append(defaultRunProperties11);
                A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties() { Language = "en-US" };

                paragraph18.Append(paragraphProperties14);
                paragraph18.Append(endParagraphRunProperties10);

                textProperties9.Append(bodyProperties18);
                textProperties9.Append(listStyle18);
                textProperties9.Append(paragraph18);
                C.ShowLegendKey showLegendKey5 = new C.ShowLegendKey() { Val = false };
                C.ShowValue showValue5 = new C.ShowValue() { Val = true };
                C.ShowCategoryName showCategoryName5 = new C.ShowCategoryName() { Val = false };
                C.ShowSeriesName showSeriesName5 = new C.ShowSeriesName() { Val = false };
                C.ShowPercent showPercent5 = new C.ShowPercent() { Val = false };
                C.ShowBubbleSize showBubbleSize5 = new C.ShowBubbleSize() { Val = false };
                C.ShowLeaderLines showLeaderLines7 = new C.ShowLeaderLines() { Val = false };

                C.DLblsExtensionList dLblsExtensionList4 = new C.DLblsExtensionList();

                C.DLblsExtension dLblsExtension4 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
                dLblsExtension4.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
                C15.ShowLeaderLines showLeaderLines8 = new C15.ShowLeaderLines() { Val = true };

                C15.LeaderLines leaderLines4 = new C15.LeaderLines();

                C.ChartShapeProperties chartShapeProperties20 = new C.ChartShapeProperties();

                A.Outline outline43 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

                A.SolidFill solidFill59 = new A.SolidFill();

                A.SchemeColor schemeColor117 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
                A.LuminanceModulation luminanceModulation65 = new A.LuminanceModulation() { Val = 35000 };
                A.LuminanceOffset luminanceOffset45 = new A.LuminanceOffset() { Val = 65000 };

                schemeColor117.Append(luminanceModulation65);
                schemeColor117.Append(luminanceOffset45);

                solidFill59.Append(schemeColor117);
                A.Round round24 = new A.Round();

                outline43.Append(solidFill59);
                outline43.Append(round24);
                A.EffectList effectList29 = new A.EffectList();

                chartShapeProperties20.Append(outline43);
                chartShapeProperties20.Append(effectList29);

                leaderLines4.Append(chartShapeProperties20);

                dLblsExtension4.Append(showLeaderLines8);
                dLblsExtension4.Append(leaderLines4);

                dLblsExtensionList4.Append(dLblsExtension4);

                dataLabels5.Append(chartShapeProperties19);
                dataLabels5.Append(textProperties9);
                dataLabels5.Append(showLegendKey5);
                dataLabels5.Append(showValue5);
                dataLabels5.Append(showCategoryName5);
                dataLabels5.Append(showSeriesName5);
                dataLabels5.Append(showPercent5);
                dataLabels5.Append(showBubbleSize5);
                dataLabels5.Append(showLeaderLines7);
                dataLabels5.Append(dLblsExtensionList4);

                barChartSeries.Append(chartShapeProperties18);
                //            barChartSeries.Append(new ChartShapeProperties(
                //    new DocumentFormat.OpenXml.Drawing.SolidFill(
                //        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "A6B4DE" }
                //    )
                //));
                barChartSeries.Append(invertIfNegative4);
                barChartSeries.Append(dataLabels5);
                //barChartSeries.Append(new DataLabels(
                //    new ChartShapeProperties(new A.NoFill(),new A.Outline(),new A.EffectList())
                //    ,new C.TextProperties(new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true }
                //    , new A.ListStyle()
                //    , new A.Paragraph(new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 }, new A.EndParagraphRunProperties() { Language = "en-US" }))
                //    , new C.ShowLegendKey() { Val = false }
                //    ,new C.ShowValue() { Val = true }
                //    ,new C.ShowCategoryName() { Val = false }
                //    ,new C.ShowSeriesName() { Val = false }
                //    ,new C.ShowPercent() { Val = false }
                //    ,new C.ShowBubbleSize() { Val = false }
                //    ,new C.ShowLeaderLines() { Val = false }
                //    , new C.DLblsExtensionList(new C.DLblsExtension(new C15.ShowLeaderLines() { Val = true }, new C.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 35000 }, new A.LuminanceOffset() { Val = 65000 }) { Val = A.SchemeColorValues.Text1 }), new A.Round()) { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center }
                //                             , new A.EffectList()))
                //    { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" })));

                i++;
            }

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis.
            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = new UInt32Value(48672768U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) }));

            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                    DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }, new MaxAxisValue() { Val = maxAxisLimit }, new MinAxisValue() { Val = minAxisLimit }),
                      new Delete() { Val = false },
                     new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                     new MajorGridlines(new ChartShapeProperties(new A.Outline(new SolidFill(new RgbColorModelHex() { Val = "D9D9D9" })), new EffectList())),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                    {
                        FormatCode = new StringValue("General"),
                        SourceLinked = new BooleanValue(true)
                    }, new TickLabelPosition()
                    {
                        Val = new EnumValue<TickLabelPositionValues>
    (TickLabelPositionValues.NextTo)
                    }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                    new ChartShapeProperties(new A.Outline(new NoFill()), new EffectList()),
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            // Add the chart Legend.
            //Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
            //    new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // end of the chart content

            // The drawings part of the chart
            Xdr.GraphicFrame gf = new Xdr.GraphicFrame();
            gf.Macro = string.Empty;
            gf.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties();
            // this has to be unique within the WorksheetDrawing class of the DrawingsPart
            // Continue with a different ID for other charts and other images.
            // Yes, normal images too.
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = 2;
            // give a friendly name
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = "Chart 1";
            gf.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

            gf.Transform = new Xdr.Transform();
            gf.Transform.Offset = new A.Offset() { X = 0, Y = 0 };
            gf.Transform.Extents = new A.Extents() { Cx = 0, Cy = 0 };


            gf.Graphic = new A.Graphic();
            gf.Graphic.GraphicData = new A.GraphicData();
            gf.Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            gf.Graphic.GraphicData.Append(new C.ChartReference() { Id = dp.GetIdOfPart(chartPart) });

            Xdr.TwoCellAnchor tcanchor = new Xdr.TwoCellAnchor();
            tcanchor.FromMarker = new Xdr.FromMarker();
            tcanchor.FromMarker.RowId = new Xdr.RowId(startRowIndex.ToString());
            // no offset
            tcanchor.FromMarker.RowOffset = new Xdr.RowOffset("600075");
            tcanchor.FromMarker.ColumnId = new Xdr.ColumnId(startColumnIndex.ToString());
            // no offset
            tcanchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.ToMarker = new Xdr.ToMarker();
            tcanchor.ToMarker.RowId = new Xdr.RowId(endRowIndex.ToString());
            // no offset
            tcanchor.ToMarker.RowOffset = new Xdr.RowOffset("600075");
            tcanchor.ToMarker.ColumnId = new Xdr.ColumnId(endColumnIndex.ToString());
            // no offset
            tcanchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.Append(gf);
            tcanchor.Append(new Xdr.ClientData());

            dp.WorksheetDrawing.Append(tcanchor);


            string[] strGenerator = new string[] { "2213969", "2213963", "2213979", "2213969", "2213963", "2213979" };
            string[] strRowId = new string[] { "22", "20", "18", "18", "20", "22" };
            string[] strColumnId = new string[] { "1", "1", "1", "10", "10", "10" };
            string[] strRowOffset = new string[] { "18000", "27000", "27000", "50000", "50000", "50000" }; //{ "80000", "186266", "95036", "81643", "27214", "136071" };
            string[] strColumnOffset = new string[] { "489858", "508606", "520699", "497453", "497453", "497453" }; //{ "489858", "508606", "520699", "197453", "182032", "163285" };

            for (i = 0; i < 6; i++)
            {
                Xdr.OneCellAnchor oneCellAnchor6 = new Xdr.OneCellAnchor();

                Xdr.FromMarker fromMarker8 = new Xdr.FromMarker();
                Xdr.ColumnId columnId10 = new Xdr.ColumnId();
                columnId10.Text = strColumnId[i]; // "1";
                Xdr.ColumnOffset columnOffset10 = new Xdr.ColumnOffset();
                columnOffset10.Text = strColumnOffset[i]; // "520699"
                Xdr.RowId rowId10 = new Xdr.RowId();
                rowId10.Text = strRowId[i]; //"16";
                Xdr.RowOffset rowOffset10 = new Xdr.RowOffset();
                rowOffset10.Text = strRowOffset[i];//"95036"; // Convert.ToString(95036*(i+1));// "95036";

                fromMarker8.Append(columnId10);
                fromMarker8.Append(columnOffset10);
                fromMarker8.Append(rowId10);
                fromMarker8.Append(rowOffset10);
                //Xdr.Extent extent6 = new Xdr.Extent() { Cx = 1782535L, Cy = 204108L };

                Xdr.Extent extent6 = new Xdr.Extent() { Cx = 1782535L, Cy = 204108L };

                Xdr.Shape shape6 = new Xdr.Shape() { Macro = "", TextLink = "" };

                Xdr.NonVisualShapeProperties nonVisualShapeProperties6 = new Xdr.NonVisualShapeProperties();

                Xdr.NonVisualDrawingProperties nonVisualDrawingProperties8 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)21U, Name = "TextBox 20" };

                A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList8 = new A.NonVisualDrawingPropertiesExtensionList();

                A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension8 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

                OpenXmlUnknownElement openXmlUnknownElement9 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6230E3EC-6B4D-47AE-A392-192F8AB4F7CE}\" />");

                nonVisualDrawingPropertiesExtension8.Append(openXmlUnknownElement9);

                nonVisualDrawingPropertiesExtensionList8.Append(nonVisualDrawingPropertiesExtension8);

                nonVisualDrawingProperties8.Append(nonVisualDrawingPropertiesExtensionList8);
                Xdr.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new Xdr.NonVisualShapeDrawingProperties() { TextBox = true };

                nonVisualShapeProperties6.Append(nonVisualDrawingProperties8);
                nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);

                Xdr.ShapeProperties shapeProperties6 = new Xdr.ShapeProperties();

                A.Transform2D transform2D6 = new A.Transform2D();
                A.Offset offset8 = new A.Offset() { X = 1130299L, Y = 3276903L };
                A.Extents extents8 = new A.Extents() { Cx = 1782535L, Cy = 204108L };

                //A.Offset offset8 = new A.Offset() { X = 0, Y = 0 };
                //A.Extents extents8 = new A.Extents() { Cx = 0, Cy = 0 };

                transform2D6.Append(offset8);
                transform2D6.Append(extents8);

                A.PresetGeometry presetGeometry6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

                presetGeometry6.Append(adjustValueList6);
                A.NoFill noFill6 = new A.NoFill();

                shapeProperties6.Append(transform2D6);
                shapeProperties6.Append(presetGeometry6);
                shapeProperties6.Append(noFill6);

                Xdr.ShapeStyle shapeStyle6 = new Xdr.ShapeStyle();

                A.LineReference lineReference6 = new A.LineReference() { Index = (UInt32Value)0U };
                A.RgbColorModelPercentage rgbColorModelPercentage16 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                lineReference6.Append(rgbColorModelPercentage16);

                A.FillReference fillReference6 = new A.FillReference() { Index = (UInt32Value)0U };
                A.RgbColorModelPercentage rgbColorModelPercentage17 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                fillReference6.Append(rgbColorModelPercentage17);

                A.EffectReference effectReference6 = new A.EffectReference() { Index = (UInt32Value)0U };
                A.RgbColorModelPercentage rgbColorModelPercentage18 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                effectReference6.Append(rgbColorModelPercentage18);

                A.FontReference fontReference6 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

                fontReference6.Append(schemeColor29);

                shapeStyle6.Append(lineReference6);
                shapeStyle6.Append(fillReference6);
                shapeStyle6.Append(effectReference6);
                shapeStyle6.Append(fontReference6);

                Xdr.TextBody textBody6 = new Xdr.TextBody();

                A.BodyProperties bodyProperties6;
                // if (i == 2)
                // {
                //     bodyProperties6 = new A.BodyProperties() { VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Bottom };
                //  }
                // else
                //  {
                bodyProperties6 = new A.BodyProperties() { VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Top };
                // }

                A.NoAutoFit noAutoFit6 = new A.NoAutoFit();

                bodyProperties6.Append(noAutoFit6);
                A.ListStyle listStyle6 = new A.ListStyle();

                A.Paragraph paragraph6 = new A.Paragraph();
                A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };

                A.Run run9 = new A.Run();

                A.RunProperties runProperties9 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike };

                A.SolidFill solidFill15 = new A.SolidFill();
                A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

                solidFill15.Append(schemeColor30);
                A.EffectList effectList7 = new A.EffectList();
                A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
                A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
                A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

                runProperties9.Append(solidFill15);
                runProperties9.Append(effectList7);
                runProperties9.Append(latinFont6);
                runProperties9.Append(eastAsianFont6);
                runProperties9.Append(complexScriptFont6);
                A.Text text9 = new A.Text();
                text9.Text = strGenerator[i];

                run9.Append(runProperties9);
                run9.Append(text9);

                A.Run run10 = new A.Run();

                A.RunProperties runProperties10 = new A.RunProperties() { Language = "en-US", FontSize = 1200 };

                A.SolidFill solidFill16 = new A.SolidFill();
                A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

                solidFill16.Append(schemeColor31);

                runProperties10.Append(solidFill16);
                A.Text text10 = new A.Text();
                text10.Text = "";

                run10.Append(runProperties10);
                run10.Append(text10);

                paragraph6.Append(paragraphProperties3);
                paragraph6.Append(run9);
                paragraph6.Append(run10);

                textBody6.Append(bodyProperties6);
                textBody6.Append(listStyle6);
                textBody6.Append(paragraph6);

                shape6.Append(nonVisualShapeProperties6);
                shape6.Append(shapeProperties6);
                shape6.Append(shapeStyle6);
                shape6.Append(textBody6);
                Xdr.ClientData clientData8 = new Xdr.ClientData();

                oneCellAnchor6.Append(fromMarker8);
                oneCellAnchor6.Append(extent6);
                oneCellAnchor6.Append(shape6);
                oneCellAnchor6.Append(clientData8);
                dp.WorksheetDrawing.Append(oneCellAnchor6);
            }

            dp.WorksheetDrawing.Save();
        }

        public void InsertPieChartInSpreadSheet(DrawingsPart drawingsPart, string chartTitle, Dictionary<string, int> data, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();


            ChartSpace chartSpace = new ChartSpace();
            chartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());

            ManualLayout manualLayout1 = new ManualLayout();
            LayoutTarget layoutTarget1 = new LayoutTarget() { Val = LayoutTargetValues.Inner };
            LeftMode leftMode1 = new LeftMode() { Val = LayoutModeValues.Edge };
            TopMode topMode1 = new TopMode() { Val = LayoutModeValues.Edge };
            Left left1 = new Left() { Val = 0.5D };
            Top top1 = new Top() { Val = 0.2D };

            Width width1 = new Width() { Val = 0.95622038461448768D };
            Height height1 = new Height() { Val = 0.54928769841269842D };

            manualLayout1.Append(layoutTarget1);
            manualLayout1.Append(leftMode1);
            manualLayout1.Append(topMode1);
            manualLayout1.Append(left1);
            manualLayout1.Append(top1);
            manualLayout1.Append(width1);
            manualLayout1.Append(height1);



            layout.Append(manualLayout1);

            NoFill noFill = new NoFill();
            C.ShapeProperties shapeProperties = new C.ShapeProperties();

            DocumentFormat.OpenXml.Drawing.Outline outline15 = new DocumentFormat.OpenXml.Drawing.Outline();
            DocumentFormat.OpenXml.Drawing.SolidFill noFill17 = new DocumentFormat.OpenXml.Drawing.SolidFill();

            RgbColorModelHex schemeColor29 = new RgbColorModelHex() { Val = "FFFFFF" };

            noFill17.Append(schemeColor29);
            outline15.Append(noFill17);

            shapeProperties.Append(noFill);
            shapeProperties.Append(outline15);
            plotArea.Append(shapeProperties);


            PieChart pieChart = plotArea.AppendChild<PieChart>(new PieChart());

            PieChartSeries pieChartSeries = pieChart.AppendChild<PieChartSeries>(new PieChartSeries(
                new Index() { Val = (UInt32Value)0U },
                new Order() { Val = (UInt32Value)0U },
                new SeriesText(new NumericValue() { Text = "PieChartSeries" })));



            CategoryAxisData catAx = new CategoryAxisData();


            StringReference stringReference = new StringReference();
            StringCache stringCache = new StringCache();

            PointCount pointCount = new PointCount() { Val = (uint)data.Count };

            stringCache.Append(pointCount);

            uint i = 0;
            foreach (var key in data.Keys)
            {
                stringCache.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(key));
                i++;
            }

            stringReference.Append(stringCache);
            catAx.Append(stringReference);
            pieChartSeries.Append(catAx);



            C.Values values = new C.Values();
            NumberReference numberReference = new NumberReference();
            NumberingCache numberingCache = new NumberingCache();

            i = 0;
            foreach (var key in data.Keys)
            {
                numberingCache.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(data[key].ToString()));
                i++;
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);
            pieChartSeries.Append(values);

            AddChartTitle(chart, chartTitle);
            pieChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            pieChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });


            CategoryAxis catAx1 = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
                 new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                 new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                 new CrossingAxis() { Val = new UInt32Value(48672768U) },
                 new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                 new AutoLabeled() { Val = new BooleanValue(true) },
                 new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                 new LabelOffset() { Val = new UInt16Value((ushort)100) }));


            //  catAx1.Append(shapeProperties1);

            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                    DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                }, new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>
            (TickLabelPositionValues.NextTo)
                }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            // Add the chart Legend.
            Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Bottom) },
                new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            chartPart.ChartSpace = chartSpace;

            PositionChart(chartPart, drawingsPart, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        private static void PositionChart(ChartPart chartPart, DrawingsPart drawingsPart, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(startColumnIndex.ToString()),
                                            new ColumnOffset("581025"),
                                            new RowId(startRowIndex.ToString()),
                                            new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId(endColumnIndex.ToString()),
                new ColumnOffset("276225"),
                new RowId(endRowIndex.ToString()),
                new RowOffset("0")));

            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                    new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());
        }
        private static void AddChartTitle(DocumentFormat.OpenXml.Drawing.Charts.Chart chart, string title)
        {
            var ctitle = chart.AppendChild(new Title());
            var chartText = ctitle.AppendChild(new C.ChartText());
            var richText = chartText.AppendChild(new RichText());

            var bodyPr = richText.AppendChild(new BodyProperties());
            var lstStyle = richText.AppendChild(new ListStyle());
            var paragraph = richText.AppendChild(new Paragraph());

            var apPr = paragraph.AppendChild(new ParagraphProperties());
            apPr.AppendChild(new DefaultRunProperties());

            var run = paragraph.AppendChild(new DocumentFormat.OpenXml.Drawing.Run());
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-CA" });
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.Text() { Text = title });
            //ctitle.AppendChild(new Overlay() { Val = new BooleanValue(false) });

        }

        private static void AddBarChartTitle(DocumentFormat.OpenXml.Drawing.Charts.Chart chart, string title)
        {
            //var ctitle = chart.AppendChild(new Title());
            //var chartText = ctitle.AppendChild(new C.ChartText());
            //var richText = chartText.AppendChild(new RichText());

            //var bodyPr = richText.AppendChild(new BodyProperties());
            //var lstStyle = richText.AppendChild(new ListStyle());
            //var paragraph = richText.AppendChild(new Paragraph());

            //var apPr = paragraph.AppendChild(new ParagraphProperties());
            //apPr.AppendChild(new DefaultRunProperties());

            //var run = paragraph.AppendChild(new DocumentFormat.OpenXml.Drawing.Run());
            //run.AppendChild(new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-CA" });
            //run.AppendChild(new DocumentFormat.OpenXml.Drawing.Text() { Text = title });
            //ctitle.AppendChild(new Overlay() { Val = new BooleanValue(false) });

            C.Title title2 = new C.Title();

            C.ChartText chartText2 = new C.ChartText();

            C.RichText richText2 = new C.RichText();
            A.BodyProperties bodyProperties16 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph16 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill56 = new A.SolidFill();

            A.SchemeColor schemeColor111 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation59 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset42 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor111.Append(luminanceModulation59);
            schemeColor111.Append(luminanceOffset42);

            solidFill56.Append(schemeColor111);
            A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill56);
            defaultRunProperties9.Append(latinFont14);
            defaultRunProperties9.Append(eastAsianFont14);
            defaultRunProperties9.Append(complexScriptFont14);

            paragraphProperties12.Append(defaultRunProperties9);

            A.Run run12 = new A.Run();
            A.RunProperties runProperties12 = new A.RunProperties() { Language = "en-US" };
            A.Text text12 = new A.Text();
            text12.Text = title;

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph16.Append(paragraphProperties12);
            paragraph16.Append(run12);

            richText2.Append(bodyProperties16);
            richText2.Append(listStyle16);
            richText2.Append(paragraph16);

            chartText2.Append(richText2);
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties17 = new C.ChartShapeProperties();
            A.NoFill noFill27 = new A.NoFill();

            A.Outline outline40 = new A.Outline();
            A.NoFill noFill28 = new A.NoFill();

            outline40.Append(noFill28);
            A.EffectList effectList26 = new A.EffectList();

            chartShapeProperties17.Append(noFill27);
            chartShapeProperties17.Append(outline40);
            chartShapeProperties17.Append(effectList26);

            C.TextProperties textProperties8 = new C.TextProperties();
            A.BodyProperties bodyProperties17 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph17 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill57 = new A.SolidFill();

            A.SchemeColor schemeColor112 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation60 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset43 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor112.Append(luminanceModulation60);
            schemeColor112.Append(luminanceOffset43);

            solidFill57.Append(schemeColor112);
            A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill57);
            defaultRunProperties10.Append(latinFont15);
            defaultRunProperties10.Append(eastAsianFont15);
            defaultRunProperties10.Append(complexScriptFont15);

            paragraphProperties13.Append(defaultRunProperties10);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph17.Append(paragraphProperties13);
            paragraph17.Append(endParagraphRunProperties9);

            textProperties8.Append(bodyProperties17);
            textProperties8.Append(listStyle17);
            textProperties8.Append(paragraph17);

            title2.Append(chartText2);
            title2.Append(overlay2);
            title2.Append(chartShapeProperties17);
            title2.Append(textProperties8);

            chart.Append(title2);

        }

        // Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys
        // and corresponding integer data, creates a column chart with the text as the series and the integers as the values.
        public void InsertBarChartInSpreadsheet(Worksheet ws, DrawingsPart drawingsPart, string title, Dictionary<string, int> data, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new C.ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            C.PlotArea plotArea = chart.AppendChild<C.PlotArea>(new C.PlotArea());
            C.Layout layout = plotArea.AppendChild<C.Layout>(new C.Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
            { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            uint i = 0;

            barChart.Append(new C.Overlap() { Val = -100 });
            barChart.Append(new C.GapWidth() { Val = 219 });
            barChart.Append(new C.VaryColors() { Val = false });

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.

            C.BarChartSeries barChartSeries = barChart.AppendChild<C.BarChartSeries>(new C.BarChartSeries(new Index()
            {
                Val = (UInt32Value)0U
            },
               new Order() { Val = (UInt32Value)0U }));
            //new SeriesText(new C.NumericValue() { Text = "Test" }))) ;


            CategoryAxisData catAxData = new CategoryAxisData();


            StringReference stringReference = new StringReference();
            StringCache stringCache = new StringCache();

            PointCount pointCount = new PointCount() { Val = (uint)data.Count };

            stringCache.Append(pointCount);

            foreach (var key in data.Keys)
            {
                stringCache.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) }).Append(new C.NumericValue(key));
                i++;
            }

            stringReference.Append(stringCache);
            catAxData.Append(stringReference);
            barChartSeries.Append(catAxData);



            C.Values values = new C.Values();
            NumberReference numberReference = new NumberReference();
            NumberingCache numberingCache = new NumberingCache();

            i = 0;
            foreach (var key in data.Keys)
            {
                numberingCache.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) }).Append(new C.NumericValue(data[key].ToString()));
                i++;
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);
            barChartSeries.Append(values);

            AddChartTitle(chart, title);

            barChart.Append(new C.AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new C.AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis.

            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new C.AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),

           new Delete() { Val = false },
           new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
           new C.NumberingFormat() { FormatCode = "dd/mm/yyyy;@", SourceLinked = true },
           new MajorTickMark() { Val = TickMarkValues.Outside },
           new MinorTickMark() { Val = TickMarkValues.Cross },
           new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
           new CrossingAxis() { Val = new UInt32Value(48650112u) },
           new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
           new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
           new LabelOffset() { Val = new UInt16Value((ushort)100) },
           new ChartShapeProperties(new A.Outline(new NoFill()), new EffectList()),
           new NoMultiLevelLabels() { Val = true }
           ));


            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new C.AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax),

                }, new MaxAxisValue() { Val = 1020D }, new MinAxisValue() { Val = 840D }),
                 new Delete() { Val = false },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(new ChartShapeProperties(new A.Outline(new SolidFill(new RgbColorModelHex() { Val = "D9D9D9" })), new EffectList())),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                }, new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
                },
                 new MajorTickMark() { Val = TickMarkValues.None },
                 new MinorTickMark() { Val = TickMarkValues.None },
                new ChartShapeProperties(new A.Outline(new NoFill()), new EffectList()),
                new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));


            // Add the chart Legend.
            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // The drawings part of the chart
            Xdr.GraphicFrame gf = new Xdr.GraphicFrame();
            gf.Macro = string.Empty;
            gf.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties();
            // this has to be unique within the WorksheetDrawing class of the DrawingsPart
            // Continue with a different ID for other charts and other images.
            // Yes, normal images too.
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = 2;
            // give a friendly name
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = "Chart 1";
            gf.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

            gf.Transform = new Xdr.Transform();
            gf.Transform.Offset = new A.Offset() { X = 0, Y = 0 };
            gf.Transform.Extents = new A.Extents() { Cx = 0, Cy = 0 };


            gf.Graphic = new A.Graphic();
            gf.Graphic.GraphicData = new A.GraphicData();
            gf.Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            gf.Graphic.GraphicData.Append(new C.ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) });

            Xdr.TwoCellAnchor tcanchor = new Xdr.TwoCellAnchor();
            tcanchor.FromMarker = new Xdr.FromMarker();
            tcanchor.FromMarker.RowId = new Xdr.RowId(startRowIndex.ToString());
            // no offset
            tcanchor.FromMarker.RowOffset = new Xdr.RowOffset("0");
            tcanchor.FromMarker.ColumnId = new Xdr.ColumnId(startColumnIndex.ToString());
            // no offset
            tcanchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.ToMarker = new Xdr.ToMarker();
            tcanchor.ToMarker.RowId = new Xdr.RowId(endRowIndex.ToString());
            // no offset
            tcanchor.ToMarker.RowOffset = new Xdr.RowOffset("0");
            tcanchor.ToMarker.ColumnId = new Xdr.ColumnId(endColumnIndex.ToString());
            // no offset
            tcanchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.Append(gf);
            tcanchor.Append(new Xdr.ClientData());

            drawingsPart.WorksheetDrawing.Append(tcanchor);
            drawingsPart.WorksheetDrawing.Save();

        }

    }
}
