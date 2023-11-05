using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Drawing;

namespace OpenXml_Excel
{
    public class ExcelHelper
    {

        /// <summary>
        /// 根据列序号计算列名
        /// <code>
        /// A,B,.......X,Y,Z
        /// AA,AB.....AX,AY,AZ
        /// ZA,ZB.....ZX,ZY,ZZ
        /// AAA,AAB......AAX,AAY,AAZ
        /// .....
        /// </code>
        /// </summary>
        /// <param name="col">从开始</param>
        /// <returns></returns>
        public static string GetColName(int col)
        {
            //col从1开始
            string res = "";
            //col输入始终为非负数
            int remain = (col - 1) % 26;
            char addChar = (char)('A' + remain);
            res = addChar + res;

            int mod = (col - 1) / 26;
            while (mod >= 1)
            {
                int left = (mod - 1) % 26;
                char add = (char)('A' + left);
                res = add + res;
                mod = (mod - 1) / 26;
            }
            return res;
        }

        public static int GetNum(String col)
        {
            int res = 0;
            int mod = 1;
            for (int i = col.Length - 1; i >= 0; i--)
            {
                char now = col[i];
                int diff = now - 'A';
                res = res + (diff + 1) * mod;
                mod = mod * 26;
            }

            return res;
        }

        /// <summary>
        /// 写入Excel
        /// </summary>
        /// <param name="filePath">文件路径(d:/abc/abc.xlsx)</param>
        /// <param name="dt">数据源(第一行是表头)</param>
        public static void Write(String filePath, DataTable dt)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // 创建WrokbookPart
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                //创建Workbook,关联到workbookPart
                workbookPart.Workbook = new Workbook();

                // 通过workbookPart创建worksheetPart
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                // 创建Worksheet,关联到worksheetPart
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // 创建Sheets,关联到Workbook
                Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                // 创建一个sheet,添加到Sheets
                Sheet sheet1 = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet1);

                // 获取上面创建的SheetData
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                int rows = dt.Rows.Count;
                int columns = dt.Columns.Count;
                if (rows > 0 && columns > 0)
                {
                    for (int rowIndex = 1; rowIndex <= rows; rowIndex++)
                    {
                        Row row = new Row() { RowIndex = (uint)rowIndex };
                        for (int colIndex = 1; colIndex <= columns; colIndex++)
                        {
                            if (dt.Columns[colIndex - 1].ColumnName == "Image")
                            {
                                AddImage(worksheetPart, worksheetPart.Worksheet, drawingsPart, rowIndex, colIndex, dt.Rows[rowIndex - 1][colIndex - 1].ToString());
                                continue;
                            }

                            string colName = GetColName(colIndex);
                            string cellReference = colName + rowIndex;
                            Cell cell = new Cell() { CellReference = cellReference };
                            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                            object val = dt.Rows[rowIndex - 1][colIndex - 1];
                            if (val != null)
                            {
                                cell.CellValue = new CellValue(val.ToString() + "");
                            }
                            else
                            {
                                cell.CellValue = new CellValue("");
                            }
                            row.Append(cell);
                        }
                        sheetData.Append(row);
                    }
                }

                spreadsheetDocument.Save();
            }
        }

        private static void AddImage(WorksheetPart worksheetPart, Worksheet worksheet, DrawingsPart drawingsPart, int rowIndex, int columnIndex, String imgPath)
        {
            if (!File.Exists(imgPath))
            {
                return;
            }

            // xdr:nvPicPr
            DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties nvpp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties()
            {
                NonVisualDrawingProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties()
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(imgPath),
                },
                NonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties()
                {
                    PictureLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks()
                    {
                        NoChangeAspect = true
                    }
                }
            };


            // xdr:blipFill
            ImagePart imgp = drawingsPart.AddImagePart(GetImagePartType(imgPath), worksheetPart.GetIdOfPart(drawingsPart));
            using (FileStream fs = new FileStream(imgPath, FileMode.Open))
            {
                imgp.FeedData(fs);
            }
            DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill()
            {
                Blip = new DocumentFormat.OpenXml.Drawing.Blip()
                {
                    Embed = drawingsPart.GetIdOfPart(imgp),
                    CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print,
                },
            };
            blipFill.Append(new OfficeStyleSheetExtensionList());
            blipFill.Append(new DocumentFormat.OpenXml.Drawing.Stretch()
            {
                FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle()
            });

            DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D()
            {
                Offset = new DocumentFormat.OpenXml.Drawing.Offset() { X = 0, Y = 0 },
                Extents = GetExtents(imgPath)
            };
            DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
            sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
            sp.Transform2D = t2d;
            DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
            prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
            prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
            sp.Append(prstGeom);
            sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

            // xdr:pic
            DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture()
            {
                NonVisualPictureProperties = nvpp,
                BlipFill = blipFill,
                ShapeProperties = sp
            };

            TwoCellAnchor twoCellAnchor = new TwoCellAnchor() { EditAs = EditAsValues.TwoCell };
            twoCellAnchor.FromMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker()
            {
                ColumnId = new ColumnId(columnIndex.ToString()),
                ColumnOffset = new ColumnOffset((20000 * (columnIndex - 1)).ToString()),
                RowId = new RowId(rowIndex.ToString()),
                RowOffset = new RowOffset((20000 * (rowIndex - 1)).ToString())
            };
            twoCellAnchor.ToMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker()
            {
                ColumnId = new ColumnId(columnIndex.ToString()),
                ColumnOffset = new ColumnOffset((40000 * columnIndex).ToString()),
                RowId = new RowId(rowIndex.ToString()),
                RowOffset = new RowOffset((40000 * rowIndex).ToString())
            };
            twoCellAnchor.Append(picture);
            twoCellAnchor.Append(new ClientData());

            WorksheetDrawing wsd = new WorksheetDrawing();
            wsd.Append(twoCellAnchor);
            wsd.Save(drawingsPart);

            DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing()
            {
                Id = drawingsPart.GetIdOfPart(imgp)
            };
            worksheet.Append(drawing);
        }

        private static DocumentFormat.OpenXml.Drawing.Extents GetExtents(string imgPath)
        {
            using (Bitmap bm = new Bitmap(imgPath))
            {
                //http://en.wikipedia.org/wiki/English_Metric_Unit#DrawingML
                //http://stackoverflow.com/questions/1341930/pixel-to-centimeter
                //http://stackoverflow.com/questions/139655/how-to-convert-pixels-to-points-px-to-pt-in-net-c
                DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
                extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                return extents;
            }
        }

        private static ImagePartType GetImagePartType(string imgPath)
        {
            if (string.IsNullOrWhiteSpace(imgPath))
            {
                return ImagePartType.Jpeg;
            }

            // .jpg
            string extension = System.IO.Path.GetExtension(imgPath).ToLower();
            switch (extension)
            {
                case ".png":
                    return ImagePartType.Png;

                default:
                    return ImagePartType.Jpeg;

            }
        }
    }
}