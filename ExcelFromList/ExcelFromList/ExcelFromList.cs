using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;
using OfficeOpenXml.Drawing;

namespace ExcelFromList
{
    /// <summary>
    /// Creates a new instance of the ExcelWorkBook.
    /// </summary>
    public class ExcelWorkBook : IDisposable
    {
        private string[] columnLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ", "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ", "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", "LY", "LZ", "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", "NU", "NV", "NW", "NX", "NY", "NZ", "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OI", "OJ", "OK", "OL", "OM", "ON", "OO", "OP", "OQ", "OR", "OS", "OT", "OU", "OV", "OW", "OX", "OY", "OZ" };
        private List<Sheet> Sheets = new List<Sheet>();
        private bool FirstRow { get; set; } = true;
        private byte[] bytesArray { get; set; } = null;
        private string fullFileName { get; set; } = null;

        private class Sheet
        {
            public string SheetName { get; set; }
            public ExcelStyleConfig ExcelStyleConfig { get; set; }
            public List<object> Data { get; set; }
            public List<PropertyInfo> Columns { get
                {
                    if (Data.Count > 0)
                        return new List<PropertyInfo>(Data.First().GetType().GetProperties());
                    else
                        return new List<PropertyInfo>();
                }
            }
        }

        #region Public Methods
        /// <summary>
        /// Returns the ExcelWorkBook bytes array
        /// </summary>
        /// <returns></returns>
        public byte[] GetBytesArray()
        {
            if (!(Sheets.Count > 0))
            {
                throw new Exception("No sheets have been added. Add at least 1 sheet.");
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    foreach (Sheet sheet in Sheets)
                    {
                        ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add(sheet.SheetName);
                        sheet.ExcelStyleConfig = sheet.ExcelStyleConfig ?? new ExcelStyleConfig();
                        var rowCtrl = 2;
                        var rowIndex = 1;

                        // Generate rows
                        foreach (var record in sheet.Data)
                        {
                            var columnCounter = 0;
                            foreach (PropertyInfo colData in sheet.Columns)
                            {
                                var colsLength = sheet.Columns.Count;
                                var colIndex = columnCounter + 1;
                                var rowsLength = sheet.Data.Count;
                                var cellAddress = columnLetters[columnCounter] + rowCtrl;

                                // Header row
                                if (sheet.ExcelStyleConfig.ShowHeaders)
                                {
                                    if (rowCtrl == 2)
                                    {
                                        ExcelRange headerCell = ws.Cells[columnLetters[columnCounter] + (rowCtrl - 1)];
                                        headerCell.Value = Utils.SplitCamelCase(colData.Name);
                                        headerCell = FormatHeaderCell(headerCell, colIndex, sheet.Columns.Count, sheet);
                                    }
                                }
                                else
                                {
                                    if (FirstRow)
                                    {
                                        rowCtrl = 1;
                                        cellAddress = columnLetters[columnCounter] + rowCtrl;
                                        FirstRow = false;
                                    }
                                }

                                // Data row
                                ExcelRange dataCell = ws.Cells[cellAddress];
                                dataCell = FormatDataCell(dataCell, colData, record, colIndex, colsLength, rowsLength, rowIndex, sheet);
                                columnCounter++;
                            }
                            rowCtrl++;
                            rowIndex++;
                        }

                        ApplyConfigs(ws, sheet);

                    }

                    bytesArray = excelPackage.GetAsByteArray();

                }
            }
            catch (Exception ex)
            {
                var st = new StackTrace();
                var caller = st.GetFrame(1).GetMethod();
                if (caller.Name == "SaveAs" && caller.DeclaringType.FullName == GetType().FullName)
                    throw;
                else
                    throw ex;
            }

            return bytesArray;
        }

        /// <summary>
        /// Adds a sheet to the worksheet, will apply default style config
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="classList"></param>
        public void AddSheet<T>(string sheetName, List<T> list)
        {
            if (Utils.IsNullOrWhiteSpace(sheetName))
            {
                throw new Exception("Sheet name cannot be null");
            }
            if (list == null)
            {
                throw new Exception("Data list cannot be null");
            }

            try
            {
                if (!SheetExists(sheetName))
                {
                    Sheets.Add(new Sheet
                    {
                        SheetName = sheetName,
                        ExcelStyleConfig = null,
                        Data = list.Cast<object>().ToList()
                    });
                }
                else
                {
                    throw new Exception("A sheet with the same name already exists in the workbook.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Adds a sheet to the worksheet, will apply provided style config
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="classList"></param>
        public void AddSheet<T>(string sheetName, List<T> list, ExcelStyleConfig esc)
        {
            if (Utils.IsNullOrWhiteSpace(sheetName))
            {
                throw new Exception("Sheet name cannot be null");
            }
            if (list == null)
            {
                throw new Exception("Data list cannot be null");
            }
            if (esc == null)
            {
                throw new Exception("ExcelStyleConfig object cannot be null");
            }

            try
            {
                if (!SheetExists(sheetName))
                {
                    Sheets.Add(new Sheet
                    {
                        SheetName = sheetName,
                        ExcelStyleConfig = esc,
                        Data = list.Cast<object>().ToList()
                    });
                }
                else
                {
                    throw new Exception("A sheet with the same name already exists in the workbook.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Removes a sheet from the worksheet
        /// </summary>
        /// <param name="sheetName"></param>
        public void RemoveSheet(string sheetName)
        {
            if (Utils.IsNullOrWhiteSpace(sheetName))
            {
                throw new Exception("Sheet name cannot be null");
            }

            try
            {
                var sheetToRemove = Sheets.Where(x => x.SheetName == sheetName).FirstOrDefault();
                if (sheetToRemove != null)
                {
                    Sheets.Remove(sheetToRemove);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Removes all sheets from worksheet
        /// </summary>
        public void ClearWorkSheet()
        {
            try
            {
                Sheets = new List<Sheet>();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Checks if a specific sheet exists
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool SheetExists(string sheetName)
        {
            if (Utils.IsNullOrWhiteSpace(sheetName))
            {
                throw new Exception("Sheet name cannot be null");
            }

            try
            {
                return Sheets.Where(x => x.SheetName == sheetName).Any();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Saves the workbook to an Excel file
        /// </summary>
        public void SaveAs(string _fullFileName)
        {
            fullFileName = _fullFileName;
            try
            {
                if (File.Exists(fullFileName))
                {
                    File.Delete(fullFileName);
                    Utils.WaitForFileReady(fullFileName);
                }

                bytesArray = GetBytesArray();
                using (var file = File.OpenWrite(fullFileName))
                {
                    file.Write(bytesArray, 0, bytesArray.Length);
                }
                Utils.WaitForFileReady(fullFileName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Opens saved Excel file with OS default program
        /// </summary>
        public void Open()
        {
            try
            {
                if (!(Utils.IsNullOrWhiteSpace(fullFileName)))
                {
                    if (File.Exists(fullFileName))
                        Process.Start(fullFileName);
                    else
                        throw new FileNotFoundException("Unable to open file.", fullFileName);
                }
                else
                {
                    throw new FileNotFoundException("Unable to open file, no file name was provided or SaveAs hasn't been called.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Private Methods
        private void ApplyConfigs(ExcelWorksheet ws, Sheet sheet)
        {
            try
            {
                #region Auto Fit Columns
                if (sheet.ExcelStyleConfig.AutoFitColumns)
                {
                    ws.Cells.AutoFitColumns();
                }
                #endregion

                #region Show Grid Lines
                ws.View.ShowGridLines = sheet.ExcelStyleConfig.ShowGridLines;
                #endregion

                #region Show Headers / Freeze Panes
                var numRowsToInsert = 4;
                if (sheet.ExcelStyleConfig.ShowHeaders)
                {
                    if (sheet.ExcelStyleConfig.FreezePanes)
                    {
                        var skipRows = 0;
                        if (sheet.ExcelStyleConfig.Base64Image != null)
                        {
                            skipRows = numRowsToInsert + sheet.ExcelStyleConfig.PaddingRows;
                            if (sheet.ExcelStyleConfig.Subtitles.Length > 0)
                            {
                                if (sheet.ExcelStyleConfig.Title != null)
                                {
                                    if (sheet.ExcelStyleConfig.Subtitles.Length > numRowsToInsert - 1)
                                    {
                                        skipRows += sheet.ExcelStyleConfig.Subtitles.Length - (numRowsToInsert - 1);
                                    }
                                }
                                else
                                {
                                    if (sheet.ExcelStyleConfig.Subtitles.Length > numRowsToInsert)
                                    {
                                        skipRows += sheet.ExcelStyleConfig.Subtitles.Length - numRowsToInsert;
                                    }
                                }
                            }
                        }
                        else
                        {
                            skipRows = sheet.ExcelStyleConfig.PaddingRows;
                            if (sheet.ExcelStyleConfig.Subtitles.Length > 0)
                            {
                                skipRows += sheet.ExcelStyleConfig.Subtitles.Length;
                                if (sheet.ExcelStyleConfig.Title != null)
                                {
                                    skipRows++;
                                }
                            }
                            else
                            {
                                if (sheet.ExcelStyleConfig.Title != null)
                                {
                                    skipRows++;
                                }
                            }
                        }
                        ws.View.FreezePanes(2 + skipRows, 1);
                    }
                }
                #endregion

                #region Prepare area for Image
                var rowHeight = 18.75;
                if (sheet.ExcelStyleConfig.Base64Image != null)
                {
                    ws.InsertRow(1, numRowsToInsert);
                    for (int i = 1; i <= numRowsToInsert; i++)
                    {
                        ws.Row(i).Height = rowHeight;
                    }
                }
                #endregion

                #region Title / Subtitle
                var insRowNum = 0;
                ExcelRange titleCell;
                ExcelRange subtitleCell;
                var titlePadding = "                    ";
                var subtitlePadding = "                           ";

                if (sheet.ExcelStyleConfig.Title != null)
                {
                    if (sheet.ExcelStyleConfig.Base64Image == null) ws.InsertRow(1, 1);
                    titleCell = ws.Cells["A1"];
                    titleCell.Style.Font.Bold = true;
                    titleCell.Style.Font.Size = 14;
                    titleCell.Style.Font.Name = "Arial";
                    titleCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    titleCell.Value = sheet.ExcelStyleConfig.Base64Image != null ? titlePadding + sheet.ExcelStyleConfig.Title : sheet.ExcelStyleConfig.Title;
                    ws.Row(1).Height = rowHeight;

                    if (sheet.ExcelStyleConfig.Subtitles.Length > 0)
                    {
                        for (int i = 0; i < sheet.ExcelStyleConfig.Subtitles.Length; i++)
                        {
                            insRowNum = i + 2;
                            subtitleCell = ws.Cells["A" + insRowNum];
                            if (sheet.ExcelStyleConfig.Base64Image != null)
                            {
                                if (insRowNum > numRowsToInsert)
                                {
                                    ws.InsertRow(insRowNum, 1);
                                }
                            }
                            else
                            {
                                ws.InsertRow(insRowNum, 1);
                            }
                            subtitleCell.Style.Font.Name = "Arial";
                            subtitleCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Row(insRowNum).Height = rowHeight;
                            subtitleCell.Value = sheet.ExcelStyleConfig.Base64Image != null ? subtitlePadding + sheet.ExcelStyleConfig.Subtitles[i] : sheet.ExcelStyleConfig.Subtitles[i];
                        }
                    }
                }
                else
                {
                    if (sheet.ExcelStyleConfig.Subtitles.Length > 0)
                    {
                        for (int i = 0; i < sheet.ExcelStyleConfig.Subtitles.Length; i++)
                        {
                            insRowNum = i + 1;
                            subtitleCell = ws.Cells["A" + insRowNum];
                            if (sheet.ExcelStyleConfig.Base64Image != null)
                            {
                                if (insRowNum > numRowsToInsert)
                                {
                                    ws.InsertRow(insRowNum, 1);
                                    ws.Row(insRowNum).Height = rowHeight;
                                }
                                subtitleCell.Value = subtitlePadding + sheet.ExcelStyleConfig.Subtitles[i];
                            }
                            else
                            {
                                ws.InsertRow(insRowNum, 1);
                                subtitleCell.Value = sheet.ExcelStyleConfig.Subtitles[i];
                                ws.Row(insRowNum).Height = rowHeight;
                            }
                            subtitleCell.Style.Font.Name = "Arial";
                            subtitleCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }
                }
                #endregion

                #region Insert column/row
                if (sheet.ExcelStyleConfig.PaddingColumns > 0)
                {
                    ws.InsertColumn(1, sheet.ExcelStyleConfig.PaddingColumns);
                }

                if (sheet.ExcelStyleConfig.PaddingRows > 0)
                {
                    ws.InsertRow(1, sheet.ExcelStyleConfig.PaddingRows);
                }
                #endregion

                #region Insert Image
                if (sheet.ExcelStyleConfig.Base64Image != null)
                {
                    Image image;
                    Image resizedImage;
                    var imageBytes = Convert.FromBase64String(sheet.ExcelStyleConfig.Base64Image);
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        image = Image.FromStream(ms);
                    }

                    resizedImage = Utils.ResizeImage(image, 100);
                    ExcelPicture excelImage = ws.Drawings.AddPicture("Image", resizedImage);
                    excelImage.SetPosition(sheet.ExcelStyleConfig.PaddingRows, 0, sheet.ExcelStyleConfig.PaddingColumns, 0);
                }
                #endregion
            }
            catch (Exception)
            {
                throw;
            }
        }

        private ExcelRange FormatHeaderCell(ExcelRange headerCell, int colIndex, int colsLength, Sheet sheet)
        {
            try
            {
                headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerCell.Style.Fill.BackgroundColor.SetColor(sheet.ExcelStyleConfig.HeaderBackgroundColor);
                headerCell.Style.Font.Color.SetColor(sheet.ExcelStyleConfig.HeaderFontColor);
                headerCell.Style.Font.Bold = true;

                if (sheet.ExcelStyleConfig.HeaderBorder)
                {
                    headerCell.Style.Border.BorderAround(sheet.ExcelStyleConfig.HeaderBorderStyle, sheet.ExcelStyleConfig.HeaderBorderColor);
                }

                if (sheet.ExcelStyleConfig.HeaderBorderAround)
                {
                    if (colIndex == 1)
                    {
                        headerCell.Style.Border.Top.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Left.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Bottom.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Top.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                        headerCell.Style.Border.Left.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                        headerCell.Style.Border.Bottom.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                    }

                    if (colIndex > 0 && colIndex < colsLength)
                    {
                        headerCell.Style.Border.Top.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Bottom.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Top.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                        headerCell.Style.Border.Bottom.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                    }

                    if (colIndex == colsLength)
                    {
                        headerCell.Style.Border.Top.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Right.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Bottom.Style = sheet.ExcelStyleConfig.HeaderBorderAroundStyle;
                        headerCell.Style.Border.Top.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                        headerCell.Style.Border.Right.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                        headerCell.Style.Border.Bottom.Color.SetColor(sheet.ExcelStyleConfig.HeaderBorderAroundColor);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            return headerCell;

        }

        private ExcelRange FormatDataCell(ExcelRange dataCell, PropertyInfo colData, object record, int colIndex, int colsLength, int rowsLength, int rowIndex, Sheet sheet)
        {
            try
            {
                var propType = colData.PropertyType.FullName;
                var propValue = colData.GetValue(record, null);
                if (propType.Contains("Nullable"))
                {
                    if (propValue == null)
                    {
                        propValue = new object();
                    }
                }

                if (propValue != null)
                {
                    #region Data Type Processing
                    if (propType.Contains("DateTime"))
                    {
                        if (propValue.ToString() == "System.Object")
                        {
                            dataCell.Value = null;
                        }
                        else
                        {
                            dataCell.Value = (DateTime)propValue;
                            dataCell.Style.Numberformat.Format = sheet.ExcelStyleConfig.DateFormat;
                        }
                    }
                    else if (propType.Contains("Decimal"))
                    {
                        if (propValue.ToString() == "System.Object")
                        {
                            dataCell.Value = null;
                        }
                        else
                        {
                            dataCell.Value = (decimal)propValue;
                            dataCell.Style.Numberformat.Format = sheet.ExcelStyleConfig.DoubleFormat;
                        }
                    }
                    else if (propType.Contains("Double"))
                    {
                        if (propValue.ToString() == "System.Object")
                        {
                            dataCell.Value = null;
                        }
                        else
                        {
                            dataCell.Value = (double)propValue;
                            dataCell.Style.Numberformat.Format = sheet.ExcelStyleConfig.DoubleFormat;
                        }
                    }
                    else if (propType.Contains("Int32") || propType.Contains("Int64"))
                    {
                        if (propType.Contains("Int32"))
                        {
                            if (propValue.ToString() == "System.Object")
                            {
                                dataCell.Value = null;
                            }
                            else
                            {
                                dataCell.Value = (Int32)propValue;
                            }
                        }
                        else if (propType.Contains("Int64"))
                        {
                            if (propValue.ToString() == "System.Object")
                            {
                                dataCell.Value = null;
                            }
                            else
                            {
                                dataCell.Value = (Int64)propValue;
                            }
                        }
                        dataCell.Style.Numberformat.Format = sheet.ExcelStyleConfig.IntFormat;
                    }
                    else
                    {
                        if (propValue.ToString() == "System.Object")
                        {
                            dataCell.Value = null;
                        }
                        else
                        {
                            dataCell.Value = propValue.ToString();
                        }
                    }

                    if (sheet.ExcelStyleConfig.BackgroundColor != null)
                    {
                        dataCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        dataCell.Style.Fill.BackgroundColor.SetColor((Color)sheet.ExcelStyleConfig.BackgroundColor);
                    }

                    if (sheet.ExcelStyleConfig.FontColor != null)
                    {
                        dataCell.Style.Font.Color.SetColor((Color)sheet.ExcelStyleConfig.FontColor);
                    }
                    #endregion

                    #region Borders
                    if (sheet.ExcelStyleConfig.Border)
                    {
                        dataCell.Style.Border.BorderAround(sheet.ExcelStyleConfig.BorderStyle, sheet.ExcelStyleConfig.BorderColor);
                    }

                    if (sheet.ExcelStyleConfig.BorderAround)
                    {
                        if (!sheet.ExcelStyleConfig.ShowHeaders)
                        {
                            if (rowIndex == 1)
                            {
                                dataCell.Style.Border.Top.Style = sheet.ExcelStyleConfig.BorderAroundStyle;
                                dataCell.Style.Border.Top.Color.SetColor(sheet.ExcelStyleConfig.BorderAroundColor);
                            }
                        }

                        if (rowIndex <= rowsLength)
                        {
                            if (colIndex == 1)
                            {
                                dataCell.Style.Border.Left.Style = sheet.ExcelStyleConfig.BorderAroundStyle;
                                dataCell.Style.Border.Left.Color.SetColor(sheet.ExcelStyleConfig.BorderAroundColor);
                            }

                            if (colIndex == colsLength)
                            {
                                dataCell.Style.Border.Right.Style = sheet.ExcelStyleConfig.BorderAroundStyle;
                                dataCell.Style.Border.Right.Color.SetColor(sheet.ExcelStyleConfig.BorderAroundColor);
                            }
                        }

                        if (rowIndex == rowsLength)
                        {
                            dataCell.Style.Border.Bottom.Style = sheet.ExcelStyleConfig.BorderAroundStyle;
                            dataCell.Style.Border.Bottom.Color.SetColor(sheet.ExcelStyleConfig.BorderAroundColor);
                        }
                    }
                    #endregion

                }
            }
            catch (Exception)
            {
                throw;
            }

            return dataCell;

        }
        #endregion

        #region IDisposable
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
        #endregion

    }

    /// <summary>
    /// Creates a new instance of ExcelStyleConfig. Use this to add styles, title and subtitles.
    /// </summary>
    public class ExcelStyleConfig
    {
        #region Public Properties
        // Sheet configs
        /// <summary>
        /// Enable to show headers (taken from the property name), defaults to true
        /// <para>Camel case property names will be split on each capital letter with a space</para>
        /// </summary>
        public bool ShowHeaders { get; set; } = true;
        /// <summary>
        /// Enable to show grid lines, defaults to true
        /// </summary>
        public bool ShowGridLines { get; set; } = true;
        /// <summary>
        /// Enable to match the width of the column to the data length, defaults to true
        /// </summary>
        public bool AutoFitColumns { get; set; } = true;
        /// <summary>
        /// Enable to freeze the first row, defaults to true
        /// </summary>
        public bool FreezePanes { get; set; } = true;
        /// <summary>
        /// Gets or sets the number of columns to insert before column A, defaults to 0
        /// </summary>
        public int PaddingColumns { get; set; } = 0;
        /// <summary>
        /// Gets or sets the number of rows to insert before row 1, defaults to 0
        /// </summary>
        public int PaddingRows { get; set; } = 0;

        // Title configs
        /// <summary>
        /// Gets or sets the title of the sheet, defaults to null
        /// </summary>
        public string Title { get; set; } = null;
        /// <summary>
        /// Gets or sets the subtitles of the sheet, defaults to new string[0]
        /// </summary>
        public string[] Subtitles { get; set; } = new string[0];
        /// <summary>
        /// Gets or sets a base64 image to be placed on the sheet, defaults to null
        /// </summary>
        public string Base64Image { get; set; } = null;

        // Data type formatting
        /// <summary>
        /// Gets or sets custom Excel format string, defaults to m/d/yyyy
        /// </summary>
        public string DateFormat { get; set; } = "m/d/yyyy";
        /// <summary>
        /// Gets or sets custom Excel format string, defaults to #,##0.00_);[Red]-#,##0.00
        /// </summary>
        public string DecimalFormat { get; set; } = "#,##0.00_);[Red]-#,##0.00";
        /// <summary>
        /// Gets or sets custom Excel format string, defaults to #,##0.00_);[Red]-#,##0.00
        /// </summary>
        public string DoubleFormat { get; set; } = "#,##0.00_);[Red]-#,##0.00";
        /// <summary>
        /// Gets or sets custom Excel format string, defaults to #,##0_);[Red]-#,##0
        /// </summary>
        public string IntFormat { get; set; } = "#,##0_);[Red]-#,##0";

        // Data cell configs
        /// <summary>
        /// Gets or sets data cell font color, defaults to null
        /// </summary>
        public Color? FontColor { get; set; } = null;
        /// <summary>
        /// Gets or sets data cell background color, defaults to null
        /// </summary>
        public Color? BackgroundColor { get; set; } = null;
        /// <summary>
        /// Enable to draw a border around each data cell, defaults to false
        /// </summary>
        public bool Border { get; set; } = false;
        /// <summary>
        /// Enable to draw a border around the data range, defaults to false
        /// </summary>
        public bool BorderAround { get; set; } = false;
        /// <summary>
        /// Gets or sets the border color around each data cell, defaults to Color.Black
        /// </summary>
        public Color BorderColor { get; set; } = Color.Black;
        /// <summary>
        /// Gets or sets the border color around the data range, defaults to Color.Black
        /// </summary>
        public Color BorderAroundColor { get; set; } = Color.Black;
        /// <summary>
        /// Gets or sets the border style around each data cell, defaults to ExcelBorderStyle.Thin
        /// </summary>
        public ExcelBorderStyle BorderStyle { get; set; } = ExcelBorderStyle.Thin;
        /// <summary>
        /// Gets or sets the border style around the data range, defaults to ExcelBorderStyle.Thin
        /// </summary>
        public ExcelBorderStyle BorderAroundStyle { get; set; } = ExcelBorderStyle.Thin;

        // Header cell configs
        /// <summary>
        /// Gets or sets the header font color, defaults to Color.LightGray
        /// </summary>
        public Color HeaderFontColor { get; set; } = Color.Lavender;
        /// <summary>
        /// Gets or sets the header background color, defaults to Color.DarkSlateGray
        /// </summary>
        public Color HeaderBackgroundColor { get; set; } = Color.Teal;
        /// <summary>
        /// Enable to draw a border around each header cell, defaults to false
        /// </summary>
        public bool HeaderBorder { get; set; } = false;
        /// <summary>
        /// Enable to draw a border around the header range, defaults to false
        /// </summary>
        public bool HeaderBorderAround { get; set; } = false;
        /// <summary>
        /// Gets or sets the border color around each header cell, defaults to Color.Black
        /// </summary>
        public Color HeaderBorderColor { get; set; } = Color.Black;
        /// <summary>
        /// Gets or sets the border color around the header range, defaults to Color.Black
        /// </summary>
        public Color HeaderBorderAroundColor { get; set; } = Color.Black;
        /// <summary>
        /// Gets or sets the border style around each header cell, defaults to ExcelBorderStyle.Thin
        /// </summary>
        public ExcelBorderStyle HeaderBorderStyle { get; set; } = ExcelBorderStyle.Thin;
        /// <summary>
        /// Gets or sets the border style around the header range, defaults to ExcelBorderStyle.Thin
        /// </summary>
        public ExcelBorderStyle HeaderBorderAroundStyle { get; set; } = ExcelBorderStyle.Thin;
        #endregion
    }

}
