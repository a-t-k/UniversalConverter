using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace UniversalConverter.Converter.ExcelExtensions;
public class ExcelDocumentCreator
{
    public bool Save(DataTable dataTable, string fileName)
    {
        var dataSet = new DataSet();
        dataSet.Tables.Add(dataTable);
        Save(dataSet, fileName);
        dataSet.Tables.Remove(dataTable);
        return true;
    }

    public bool Save(DataSet ds, string fileName)
    {
        using var spreadsheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
        this.Save(ds, spreadsheet);
        return true;
    }

    private void Save(DataSet ds, SpreadsheetDocument spreadsheet)
    {
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.Append(new BookViews(new WorkbookView()));
        var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
        workbookStylesPart.Stylesheet = GenerateStyleSheet();
        workbookStylesPart.Stylesheet.Save();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        uint worksheetNumber = 1;
        foreach (DataTable dt in ds.Tables)
        {
            var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = worksheetNumber,
                Name = dt.TableName
            };

            sheets.Append(sheet);
            CreateWorksheet(dt, newWorksheetPart);
            worksheetNumber++;
        }

        workbookPart.Workbook.Append(new DefinedNames());
        workbookPart.Workbook.Save();
    }

    // DataType reference
    // https://learn.microsoft.com/en-us/dotnet/api/system.data.datacolumn.datatype?redirectedfrom=MSDN&view=net-7.0#System_Data_DataColumn_DataType
    private void CreateHeader(DataTable dataTable, OpenXmlWriter writer,
        (bool[] isIntegerColumn, bool[] isFloatColumn, bool[] isDateColumn, string[] columnNames, int columnsCount) dataInfo)
    {
        writer.WriteStartElement(new Row { RowIndex = 1, Height = 20, CustomHeight = true });
        for (var index = 0; index < dataInfo.columnsCount; index++)
        {
            var col = dataTable.Columns[index];
            AppendHeaderTextCell(dataInfo.columnNames[index] + "1", col.ColumnName, writer);
            if (col.DataType.FullName == null) continue;
            switch (Type.GetTypeCode(col.DataType))
            {
                case TypeCode.Int16 or TypeCode.Int32 or TypeCode.Int64:
                    dataInfo.isIntegerColumn[index] = true;
                    dataInfo.isFloatColumn[index] = false;
                    dataInfo.isDateColumn[index] = false;
                    continue;
                case TypeCode.Decimal or TypeCode.Double or TypeCode.Single:
                    dataInfo.isIntegerColumn[index] = false;
                    dataInfo.isFloatColumn[index] = true;
                    dataInfo.isDateColumn[index] = false;
                    continue;
                case TypeCode.DateTime:
                    dataInfo.isIntegerColumn[index] = false;
                    dataInfo.isFloatColumn[index] = false;
                    dataInfo.isDateColumn[index] = true;
                    continue;
                default:
                    dataInfo.isIntegerColumn[index] = false;
                    dataInfo.isFloatColumn[index] = false;
                    dataInfo.isDateColumn[index] = false;
                    continue;
            }
        }

        writer.WriteEndElement();
    }

    private (bool[] isIntegerColumn, bool[] isFloatColumn, bool[] isDateColumn, string[] columnNames, int columnsCount) GetDataInfo(DataTable dataTable)
    {
        var columnsCount = dataTable.Columns.Count;
        var isIntegerColumn = new bool[columnsCount];
        var isFloatColumn = new bool[columnsCount];
        var isDateColumn = new bool[columnsCount];
        var columnNames = new string[columnsCount];
        for (var n = 0; n < columnsCount; n++) columnNames[n] = GetExcelColumnName(n);
        return (isIntegerColumn, isFloatColumn, isDateColumn, columnNames, columnsCount);
    }

    private void CreateWorksheet(DataTable dataTable, WorksheetPart worksheetPart)
    {
        var writer = OpenXmlWriter.Create(worksheetPart, Encoding.ASCII);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new Columns());

        uint index = 1;
        foreach (DataColumn unused in dataTable.Columns)
        {
            writer.WriteElement(new Column { Min = index, Max = index, CustomWidth = true, Width = 20 });
            index++;
        }
        writer.WriteEndElement();

        writer.WriteStartElement(new SheetData());
        var dataInfo = GetDataInfo(dataTable);
        CreateHeader(dataTable, writer, dataInfo);

        uint rowIndex = 1;
        var cultureInfo = new CultureInfo("en-US");
        foreach (DataRow dataRow in dataTable.Rows)
        {
            ++rowIndex;
            writer.WriteStartElement(new Row { RowIndex = rowIndex });
            for (var columnIndex = 0; columnIndex < dataInfo.columnsCount; columnIndex++)
            {
                var columIndexValue = dataRow?.ItemArray?[columnIndex]?.ToString();
                var cellValue = columIndexValue is null ? string.Empty : RemoveNonAsciiCharacters(columIndexValue);
                var cellReference = dataInfo.columnNames[columnIndex] + rowIndex.ToString();

                switch (dataInfo.isIntegerColumn[columnIndex], dataInfo.isFloatColumn[columnIndex], dataInfo.isDateColumn[columnIndex])
                {
                    case (true, false, false) or (false, true, false):
                        var hasDecimalPlaces = dataInfo.isFloatColumn[columnIndex];
                        if (double.TryParse(cellValue, out var cellFloatValue))
                        {
                            cellValue = cellFloatValue.ToString(cultureInfo);
                            AppendNumericCell(cellReference, cellValue, hasDecimalPlaces, writer);
                        }
                        continue;
                    case (false, false, true):
                        if (DateTime.TryParse(cellValue, out var dateValue))
                        {
                            AppendDateCell(cellReference, dateValue, writer);
                        }
                        else
                        {
                            AppendTextCell(cellReference, cellValue, writer);
                        }
                        continue;
                    default:
                        AppendTextCell(cellReference, cellValue, writer);
                        continue;
                }
            }
            writer.WriteEndElement();
        }
        //  End of SheetData
        writer.WriteEndElement();
        //  End of worksheet
        writer.WriteEndElement();
        writer.Close();
    }

    // Add a text cell to the first row
    // style #3: gray background color & white text
    private void AppendHeaderTextCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
    {
        writer.WriteElement(new Cell
        {
            CellValue = new CellValue(cellStringValue),
            CellReference = cellReference,
            DataType = CellValues.String,
            StyleIndex = 3
        });
    }

    // Add a new text cell to a Row 
    private void AppendTextCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
    {
        // it is formula?.
        if (cellStringValue.StartsWith("="))
        {
            AppendFormulaCell(cellReference, cellStringValue, writer);
            return;
        }

        writer.WriteElement(new Cell
        {
            CellValue = new CellValue(cellStringValue),
            CellReference = cellReference,
            DataType = CellValues.String
        });
    }

    //  Add a new DateTime cell to a row.
    //  style #1: format for DateTime wit time -> "dd/MMM/yyyy hh:mm:ss".
    //  style #2: format for DateTime without time -> "dd/MMM/yyyy".
    private void AppendDateCell(string cellReference, DateTime dateTimeValue, OpenXmlWriter writer)
    {
        var value = dateTimeValue.ToOADate().ToString(CultureInfo.InvariantCulture);
        var dateWithTime = dateTimeValue.Date != dateTimeValue;
        writer.WriteElement(new Cell
        {
            CellValue = new CellValue(value),
            CellReference = cellReference,
            StyleIndex = UInt32Value.FromUInt32(dateWithTime ? (uint)1 : 2),
            DataType = CellValues.Number
        });
    }

    // Add a formula cell to a row 
    private void AppendFormulaCell(string cellReference, string cellStringValue, OpenXmlWriter writer)
    {
        writer.WriteElement(new Cell
        {
            CellFormula = new CellFormula(cellStringValue),
            CellReference = cellReference,
            DataType = CellValues.Number
        });
    }

    // add a new numeric excel cell to a row.
    // style #4: format for 0 decimal places.
    // style #5 format for 2 decimal places.
    private void AppendNumericCell(string cellReference, string cellStringValue, bool hasDecimalPlaces, OpenXmlWriter writer)
    {
        var cellStyle = (uint)(hasDecimalPlaces ? 5 : 4);
        writer.WriteElement(new Cell
        {
            CellValue = new CellValue(cellStringValue),
            CellReference = cellReference,
            StyleIndex = cellStyle,
            DataType = CellValues.Number,

        });
    }

    //  Remove non ASCII characters.
    private string RemoveNonAsciiCharacters(string text)
    {

        string r = "[\x00-\x08\x0B\x0C\x0E-\x1F]";
        return Regex.Replace(text, r, "", RegexOptions.Compiled);
    }

    // Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Z, AA, AB, AC... AY, AZ, BA, BB..)
    // 0 = A; 1 = B; 25 = Z; 26 = AA; 702 = AAA;
    public string GetExcelColumnName(int columnIndex)
    {
        var firstInt = columnIndex / 676;
        var secondInt = columnIndex % 676 / 26;
        if (secondInt == 0)
        {
            secondInt = 26;
            firstInt = firstInt - 1;
        }

        var thirdInt = columnIndex % 26;
        var firstChar = (char)('A' + firstInt - 1);
        var secondChar = (char)('A' + secondInt - 1);
        var thirdChar = (char)('A' + thirdInt);
        return columnIndex switch
        {
            < 26 => $"{thirdChar}",
            < 702 => $"{secondChar}{thirdChar}",
            _ => $"{firstChar}{secondChar}{thirdChar}"
        };
    }

    private Stylesheet GenerateStyleSheet()
    {
        //  If you want certain Excel cells to have a different Format, color, border, fonts, etc, then you need to define a "CellFormats" records containing 
        //  these attributes, then assign that style number to your cell.
        //
        //  For example, we'll define "Style # 3" with the attributes we'd like for our header row (Row #1) on each worksheet, where the text is a bit bigger,
        //  and is white text on a dark-gray background.
        // 
        //  NB: The NumberFormats from 0 to 163 are hardcoded in Excel (described in the following URL), and we'll define a couple of custom number formats below.
        //  https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
        //  http://lateral8.com/articles/2010/6/11/openxml-sdk-20-formatting-excel-values.aspx
        //

        // Custom number format # 164: for date with time
        uint dateTimeCustomFormat = 164;
        // Custom number format # 165: for date without time
        uint dateWithoutTimeCustomFormat = 165;

        return new Stylesheet(
            new NumberingFormats(
                new NumberingFormat()
                {
                    NumberFormatId = UInt32Value.FromUInt32(dateTimeCustomFormat),
                    FormatCode = StringValue.FromString("dd/MMM/yyyy hh:mm:ss")
                },
                new NumberingFormat()
                {
                    NumberFormatId = UInt32Value.FromUInt32(dateWithoutTimeCustomFormat),
                    FormatCode = StringValue.FromString("dd/MMM/yyyy")
                }
            ),
            new Fonts(
                // Index 0 - The default font.
                new Font(
                    new FontSize() { Val = 10 },
                    new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                    new FontName() { Val = "Arial" }),
                // Index 1 - A 12px bold font, in white.
                new Font(
                    new Bold(),
                    new FontSize() { Val = 12 },
                    new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } },
                    new FontName() { Val = "Arial" }),
                // Index 2 - An Italic font.
                new Font(
                    new Italic(),
                    new FontSize() { Val = 10 },
                    new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                    new FontName() { Val = "Times New Roman" })
            ),
            new Fills(
                // Index 0 - The default fill.
                new Fill(new PatternFill { PatternType = PatternValues.None }),

                // Index 1 - The default fill of gray 125 (required)
                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),

                // Index 2 - The yellow fill.
                new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid }),

                // Index 3 - Dark-gray fill.
                new Fill(new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FF404040" } }) { PatternType = PatternValues.Solid })
            ),
            new Borders(
                // Index 0 - The default border.
                new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
                // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                new Border(
                    new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            ),
            new CellFormats(
                // Style # 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },
                // Style # 1 - DateTimes
                new CellFormat() { NumberFormatId = 164 },
                // Style # 2 - Dates (with a blank time)
                new CellFormat() { NumberFormatId = 165 },
                // Style # 3 - Header row 
                new CellFormat(new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center })
                { FontId = 1, FillId = 3, BorderId = 0, ApplyFont = true, ApplyAlignment = true },
                // Style # 4 - Number format: #,##0
                new CellFormat() { NumberFormatId = 3, FormatId = 0, FillId = 0, BorderId = 0, FontId = 0, ApplyNumberFormat = BooleanValue.FromBoolean(true) },
                // Style # 5 - Number format: #,##0.00
                new CellFormat() { NumberFormatId = 4, FormatId = 0, FillId = 0, BorderId = 0, FontId = 0, ApplyNumberFormat = BooleanValue.FromBoolean(true) },
                // Style # 6 - Bold 
                new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },
                // Style # 7 - Italic
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },
                // Style # 8 - Times Roman
                new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },
                // Style # 9 - Yellow Fill
                new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },
                // Style # 10 - Alignment
                new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center })
                { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                // Style # 11 - Border
                new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }
            )
        );
    }
}