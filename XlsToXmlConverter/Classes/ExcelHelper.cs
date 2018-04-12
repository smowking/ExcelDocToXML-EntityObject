using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace XlsToXmlConverter.Classes
{
    public class ExcelHelper
    {
        /// <param name="filename">Excel File Path</param> 
        /// <returns></returns> 
        public DataTable ReadExcelFile(string filename)
        {
            DataTable dt = new DataTable();
            try
            {
                // Use a classe SpreadSheetDocument do Open XML SDK para abrir o abrir arquivos do excel 
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
                {
                    // Obtenha a parte da pasta de trabalho do documento de planilha 
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;


                    // Obtenha todas as planilhas no documento de planilha  
                    IEnumerable<Sheet> sheetcollection = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();


                    // Obtenha a Id de relacionamento 
                    string relationshipId = sheetcollection.First().Id.Value;


                    // Obtenha a parte da planilha1 do documento de planilha 
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);


                    // Obtenha dados no arquivo do Excel 
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();


                    if (rowcollection.Count() == 0)
                    {
                        return dt;
                    }


                    // Adicione colunas 
                    foreach (Cell cell in rowcollection.ElementAt(0))
                    {
                        dt.Columns.Add(GetValueOfCell(spreadsheetDocument, cell));
                    }

                    // Adicione linhas no DataTable 
                    foreach (Row row in rowcollection)
                    {
                        DataRow temprow = dt.NewRow();
                       
                        int columnIndex = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            // Obtenha índice de coluna da célula 
                            int cellColumnIndex = GetColumnIndex(GetColumnName(cell.CellReference));

                            if (columnIndex < cellColumnIndex)
                            {
                                do
                                {
                                    temprow[columnIndex] = string.Empty;
                                    columnIndex++;
                                }


                                while (columnIndex < cellColumnIndex);
                            }


                            temprow[columnIndex] = GetValueOfCell(spreadsheetDocument, cell);
                            columnIndex++;
                        }


                        // Adicione a linha ao DataTable 
                        // as linhas incluem a linha de cabeçalho 
                        dt.Rows.Add(temprow);
                    }
                }


                // Remove a linha de cabeçalho 
                dt.Rows.RemoveAt(0);
                return dt;
            }
            catch (IOException ex)
            {
                throw new IOException(ex.Message);
            }
        }

        /// <summary> 
        /// Converta a tabela de dados para o formato Xml 
        /// </summary> 
        /// <param name="filename">Excel File Path</param> 
        /// <returns>Xml format string</returns> 
        public string GetXML(string filename)
        {
            using (DataSet ds = new DataSet())
            {
                ds.Tables.Add(this.ReadExcelFile(filename));
                return ds.GetXml();
            }
        }

        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name of cell
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }


        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            // From right to left
            for (int position = columnName.Length - 1; position >= 0; position--)
            {
                // For letters
                if (Char.IsLetter(columnName[position]))
                {
                    columnIndex += factor * ((columnName[position] - 'A') + 1) - 1;
                    factor *= 26;
                }
            }

            return columnIndex;
        }


        /// <summary> 
        ///  Obtenha o valor da célula  
        /// </summary> 
        /// <param name="spreadsheetdocument">SpreadSheet Document Object</param> 
        /// <param name="cell">Cell Object</param> 
        /// <returns>The Value in Cell</returns> 
        private static string GetValueOfCell(SpreadsheetDocument spreadsheetdocument, Cell cell)
        {
            // Obtenha o valor na célula 
            SharedStringTablePart sharedString = spreadsheetdocument.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return string.Empty;
            }


            string cellValue = cell.CellValue.InnerText;

            // A condição de que o tipo de dados de célula é SharedString 
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
            }
            else
            {
                return cellValue;
            }
        }
    }
}
