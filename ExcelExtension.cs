﻿using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace OfficeOpenXmlExtension
{
    public static class ExcelExtension
    {
        public static ExcelRange SetBold(this ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
            return cell;
        }

        public static ExcelRange SetItalic(this ExcelRange cell)
        {
            cell.Style.Font.Italic = true;
            return cell;
        }

        public static ExcelRange SetUnderline(this ExcelRange cell)
        {
            cell.Style.Font.UnderLine = true;
            return cell;
        }

        public static ExcelRange SetBorder(this ExcelRange cell)
        {
            cell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            return cell;
        }

        public static ExcelRange SetMerge(this ExcelRange cell, ExcelRange toCell)
        {
            cell.Worksheet.Cells[cell.Address + ":" + toCell.Address].Merge = true;
            return cell;
        }
        public static ExcelRange SetMerge(this ExcelRange cell, string toCellAddress)
        {
            cell.Worksheet.Cells[cell.Address + ":" + toCellAddress].Merge = true;
            return cell;
        }

        public static ExcelRange SetBackgroundColor(this ExcelRange cell, int red, int green, int blue)
        {
            cell.Style.Fill.SetBackground(color: Color.FromArgb(red, green, blue));
            return cell;
        }
        public static ExcelRange SetBackgroundColor(this ExcelRange cell, string hex)
        {
            cell.Style.Fill.SetBackground(color: ColorTranslator.FromHtml(hex));
            return cell;
        }
        public static ExcelRange SetBackgroundColor(this ExcelRange cell, Color color)
        {
            cell.Style.Fill.SetBackground(color: color);
            return cell;
        }

        public static ExcelRange SetFontSize(this ExcelRange cell, float size)
        {
            cell.Style.Font.Size = size;
            return cell;
        }

        public static ExcelRange SetWidth(this ExcelRange cell, double width)
        {
            cell.Worksheet.Column(cell.Start.Column).Width = width;
            return cell;
        }

        /// <summary>
        /// Style cells in header
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="hex"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelRange StyleHeader(this ExcelRange cell, string? hex = null, int? red = null, int? green = null, int? blue = null, Color? color = null, float? fontSize = null, double? width = null)
        {
            cell.SetBold().SetBorder();

            if (!string.IsNullOrEmpty(hex))
            {
                cell.SetBackgroundColor(hex);
            }
            else if (red.HasValue && green.HasValue && blue.HasValue)
            {
                cell.SetBackgroundColor(red.Value, green.Value, blue.Value);
            }
            else if (color.HasValue)
            {
                cell.SetBackgroundColor(color.Value);
            }

            if (fontSize.HasValue)
            {
                cell.SetFontSize(fontSize.Value);
            }

            if (width.HasValue)
            {
                cell.SetWidth(width.Value);
            }
            return cell;
        }

        /// <summary>
        /// Style cells in body
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="hex"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelRange StyleBody(this ExcelRange cell, string? hex = null, int? red = null, int? green = null, int? blue = null, Color? color = null, float? fontSize = null, double? width = null)
        {
            cell.SetBorder();

            if (!string.IsNullOrEmpty(hex))
            {
                cell.SetBackgroundColor(hex);
            }
            else if (red.HasValue && green.HasValue && blue.HasValue)
            {
                cell.SetBackgroundColor(red.Value, green.Value, blue.Value);
            }
            else if (color.HasValue)
            {
                cell.SetBackgroundColor(color.Value);
            }

            if (fontSize.HasValue)
            {
                cell.SetFontSize(fontSize.Value);
            }
            if (width.HasValue)
            {
                cell.SetWidth(width.Value);
            }
            return cell;
        }

        public static void RenderCells(this ExcelWorksheet sheet, IEnumerable<CellSetting> cellSettings)
        {
            foreach (var cellSetting in cellSettings)
            {
                if (string.IsNullOrEmpty(cellSetting.Address))
                {
                    throw new ArgumentNullException($"ERROR. {nameof(cellSetting.Address)} property cannot null or empty.");
                }

                var cell = sheet.Cells[cellSetting.Address];
                cell.Value = cellSetting.Value;

                if (cellSetting.IsStyleHeader.HasValue)
                {
                    cell.StyleHeader(cellSetting.BackgroundHex, cellSetting.BackgroundRed, cellSetting.BackgroundGreen, cellSetting.BackgroundBlue, cellSetting.BackgroundColor, cellSetting.FontSize, cellSetting.Width);
                    //continue;
                }

                if (cellSetting.IsStyleBody.HasValue)
                {
                    cell.StyleBody(cellSetting.BackgroundHex, cellSetting.BackgroundRed, cellSetting.BackgroundGreen, cellSetting.BackgroundBlue, cellSetting.BackgroundColor, cellSetting.FontSize, cellSetting.Width);
                    //continue;
                }

                if (cellSetting.IsBold.HasValue)
                {
                    cell.SetBold();
                }

                if (cellSetting.IsItalic.HasValue)
                {
                    cell.SetItalic();
                }

                if (cellSetting.IsUnderline.HasValue)
                {
                    cell.SetUnderline();
                }

                if (cellSetting.FontSize.HasValue)
                {
                    cell.SetFontSize(cellSetting.FontSize.Value);
                }

                if (cellSetting.IsSetBorder.HasValue)
                {
                    cell.SetBorder();
                }

                if (!string.IsNullOrEmpty(cellSetting.BackgroundHex))
                {
                    cell.SetBackgroundColor(cellSetting.BackgroundHex);
                }
                else if (cellSetting.BackgroundRed.HasValue && cellSetting.BackgroundGreen.HasValue && cellSetting.BackgroundBlue.HasValue)
                {
                    cell.SetBackgroundColor(cellSetting.BackgroundRed.Value, cellSetting.BackgroundGreen.Value, cellSetting.BackgroundBlue.Value);
                }
                else if (cellSetting.BackgroundColor.HasValue)
                {
                    cell.SetBackgroundColor(cellSetting.BackgroundColor.Value);
                }

                if (cellSetting.MergedCell != null)
                {
                    cell.SetMerge(cellSetting.MergedCell);
                }
                else if (!string.IsNullOrEmpty(cellSetting.MergedCellAddress))
                {
                    cell.SetMerge(cellSetting.MergedCellAddress);
                }

                if (cellSetting.Width.HasValue)
                {
                    cell.SetWidth(cellSetting.Width.Value);
                }
            }
        }
        /// <summary>
        /// Note: If below table have values like signatures or need to exclude rows/columns at the end, please pass endRow and endColumn paramaters.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endRow"></param>
        /// <param name="endColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="maxRow"></param>
        /// <param name="maxColumn"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="Exception"></exception>
        public static string GetTableJson(this ExcelWorksheet sheet, int startRow = 1, int startColumn = 1, int? endRow = null, int? endColumn = null, bool hasHeader = true, int maxRow = 3000, int maxColumn = 200)
        {

            var dataTable = new DataTable();

            if (startRow == 0 || startColumn == 0)
            {
                throw new ArgumentException($"ERROR. {nameof(startRow)}/{nameof(startColumn)} must be greater than or equal to 1.");
            }

            if ((endRow.HasValue && endRow.Value == 0) || (endColumn.HasValue && endColumn.Value == 0))
            {
                throw new ArgumentException($"ERROR. {nameof(startRow)}/{nameof(startColumn)} must be greater than or equal to 1.");
            }

            if (endRow is null && endColumn is null && (sheet.Dimension.End.Column == 0 || sheet.Dimension.End.Row == 0))
            {
                throw new Exception("ERROR. No data was found in excel file.");
            }

            endRow ??= sheet.Dimension.End.Row;
            endColumn ??= sheet.Dimension.End.Column;

            // Add column headers
            var columnCount = 1;
            for (var column = startColumn; column <= endColumn && column <= maxColumn; column++)
            {
                var cell = sheet.Cells[startRow, column];
                dataTable.Columns.Add(new DataColumn(hasHeader ? ClearText(cell.Value.ToString()) : $"Column {columnCount++}"));
            }

            // Add rows
            for (var row = hasHeader ? startRow + 1 : startRow; row <= endRow && row <= maxRow; row++)
            {
                var dataRow = dataTable.NewRow();
                var columnIndex = 0;
                for (var column = startColumn; column <= endColumn && column <= maxColumn; column++)
                {
                    dataRow[columnIndex++] = ClearText(sheet.Cells[row, column].Value?.ToString());
                }
                dataTable.Rows.Add(dataRow);
            }

            return JsonConvert.SerializeObject(dataTable);
        }

        private static string ClearText(string? text)
        {
            if (text is null)
            {
                return string.Empty;
            }

            return Regex.Replace(text.Trim(), "\r\n|\r|\n|\t", "");
        }

    }

    public class CellSetting
    {
        public CellSetting()
        {

        }
        public CellSetting(string address, string value)
        {
            Address = address;
            Value = value;
        }
        public CellSetting(string address, string value, float? fontSize, bool? isBold, bool? isItalic, bool? isUnderline, bool? isSetBorder, string? backgroundHex, int? backgroundRed, int? backgroundGreen, int? backgroundBlue, Color? backgroundColor, ExcelRange? mergedCell, string? mergedCellAddress, bool? isStyleHeader, bool? isStyleBody, double? width)
        {
            Address = address;
            Value = value;
            FontSize = fontSize;
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
            IsSetBorder = isSetBorder;
            BackgroundHex = backgroundHex;
            BackgroundRed = backgroundRed;
            BackgroundGreen = backgroundGreen;
            BackgroundBlue = backgroundBlue;
            BackgroundColor = backgroundColor;
            MergedCell = mergedCell;
            MergedCellAddress = mergedCellAddress;
            IsStyleHeader = isStyleHeader;
            IsStyleBody = isStyleBody;
            Width = width;
        }

        public string Address { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;

        public float? FontSize { get; set; }

        public bool? IsBold { get; set; } = null;
        public bool? IsItalic { get; set; } = null;
        public bool? IsUnderline { get; set; } = null;

        public bool? IsSetBorder { get; set; } = null;

        public string? BackgroundHex { get; set; }
        public int? BackgroundRed { get; set; }
        public int? BackgroundGreen { get; set; }
        public int? BackgroundBlue { get; set; }

        public Color? BackgroundColor { get; set; }

        public ExcelRange? MergedCell { get; set; }
        public string? MergedCellAddress { get; set; }

        public bool? IsStyleHeader { get; set; }
        public bool? IsStyleBody { get; set; }

        public double? Width { get; set; }
    }
}
