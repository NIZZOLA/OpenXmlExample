using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FileUploadTest.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace FileUploadTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileImportController : ControllerBase
    {
        [HttpGet]
        public async Task<IActionResult> Test()
        {
            return Ok(new { test = "teste ok" });
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return Content("file not selected");

            /*
            var path = Path.Combine(
                        Directory.GetCurrentDirectory(), "wwwroot",
                        file.FileName.ToString());

            using (var stream = new FileStream(path, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            */

            var res = ImportCSV(file);

            return Ok(res);
        }

        private bool ImportCSV( IFormFile uploadedFile)
        {
            string worksheetName = "Planilha1";
            int rowCount, colCount;

            var result = string.Empty;
            using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
            {
                result = reader.ReadToEnd();
            }
            var csvLines = result.Split('\n');

            return true;
        }

        private List<ImportRowData> Import2(IFormFile uploadedFile)
        {
            List<ImportRowData> sheet = new List<ImportRowData>();

            using (var memoryStream = new MemoryStream())
            {
                uploadedFile.CopyToAsync(memoryStream).ConfigureAwait(false);

                //Open the Excel file using ClosedXML.
                using (XLWorkbook workBook = new XLWorkbook(memoryStream))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    //Loop through the Worksheet rows.

                    int rowNum = 0;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        rowNum++;
                        if (rowNum == 2)
                        {
                            if (!ValidateSheetFormat2(row))
                                return null;

                        }
                        else if (rowNum > 2)
                        {
                            ImportRowData linha = new ImportRowData();
                            int col = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                if (cell.Value != null)
                                {
                                    switch (col)
                                    {
                                        case 1:
                                            linha.Acao = cell.Value.ToString();
                                            break;
                                        case 2:
                                            linha.ID_Externo = cell.Value.ToString();
                                            break;
                                        case 3:
                                            linha.Nome = cell.Value.ToString();
                                            break;
                                        case 4:
                                            linha.DataNascimento = cell.Value.ToString();
                                            break;
                                        case 5:
                                            linha.Genero = cell.Value.ToString();
                                            break;
                                        case 6:
                                            linha.Login = cell.Value.ToString();
                                            break;
                                        case 7:
                                            linha.Senha = cell.Value.ToString();
                                            break;
                                        case 8:
                                            linha.CPF = cell.Value.ToString();
                                            break;
                                        case 9:
                                            linha.RG = cell.Value.ToString();
                                            break;
                                        case 10:
                                            linha.Email = cell.Value.ToString();
                                            break;
                                        case 11:
                                            linha.Celular = cell.Value.ToString();
                                            break;
                                        case 12:
                                            linha.ID_Externo_Responsavel = cell.Value.ToString();
                                            break;
                                        case 13:
                                            linha.Papel = cell.Value.ToString();
                                            break;
                                        case 14:
                                            linha.GroupCode = cell.Value.ToString();
                                            break;
                                        case 15:
                                            linha.GroupName = cell.Value.ToString();
                                            break;
                                        case 16:
                                            linha.GroupDescription = cell.Value.ToString();
                                            break;
                                        case 17:
                                            linha.TAG = cell.Value.ToString();
                                            break;
                                    }
                                }

                                col++;
                            }
                            sheet.Add(linha);
                        }

                    }
                }
            }

            return sheet;
        }



        private async Task<List<ImportRowData>> Import(IFormFile uploadedFile)
        {

            List<ImportRowData> sheet = new List<ImportRowData>();

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await uploadedFile.CopyToAsync(memoryStream).ConfigureAwait(false);

                    using (var package = new ExcelPackage(memoryStream))
                    {
                        var worksheet = package.Workbook.Worksheets[1]; // Tip: To access the first worksheet, try index 1, not 0

                        int rowCount = worksheet.Dimension.Rows;
                        int ColCount = worksheet.Dimension.Columns;

                        if (!ValidateSheetFormat(worksheet))
                            return null;

                        for (int row = 3; row <= rowCount; row++)
                        {
                            ImportRowData linha = new ImportRowData();
                            for (int col = 1; col <= ColCount; col++)
                            {
                                if (worksheet.Cells[row, col].Value != null)
                                {
                                    switch (col)
                                    {
                                        case 1:
                                            linha.Acao = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 2:
                                            linha.ID_Externo = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 3:
                                            linha.Nome = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 4:
                                            linha.DataNascimento = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 5:
                                            linha.Genero = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 6:
                                            linha.Login = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 7:
                                            linha.Senha = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 8:
                                            linha.CPF = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 9:
                                            linha.RG = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 10:
                                            linha.Email = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 11:
                                            linha.Celular = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 12:
                                            linha.ID_Externo_Responsavel = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 13:
                                            linha.Papel = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 14:
                                            linha.GroupCode = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 15:
                                            linha.GroupName = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 16:
                                            linha.GroupDescription = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                        case 17:
                                            linha.TAG = worksheet.Cells[row, col].Value.ToString();
                                            break;
                                    }
                                }
                            }
                            sheet.Add(linha);
                        }
                    }
                    return sheet;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private bool ValidateSheetFormat(ExcelWorksheet worksheet)
        {
            string infoline = @"Ação|ID Externo*|Nome*|Data de Nascimento*|Gênero*|Login*|Senha*|CPF|RG|Email|Celular|ID Externo Responsável|Papel*|Código Grupo*|Nome Grupo*|Descrição Grupo|TAG";
            string[] headerInfo = infoline.Split('|');

            for (int i = 1; i <= worksheet.Dimension.Columns; i++)
            {
                if (worksheet.Cells[2, i].Value == null)
                    return false;

                if (worksheet.Cells[2, i].Value.ToString() != headerInfo[i - 1])
                    return false;
            }
            return true;
        }

        private bool ValidateSheetFormat2(IXLRow row)
        {
            string infoline = @"Ação|ID Externo*|Nome*|Data de Nascimento*|Gênero*|Login*|Senha*|CPF|RG|Email|Celular|ID Externo Responsável|Papel*|Código Grupo*|Nome Grupo*|Descrição Grupo|TAG";
            string[] headerInfo = infoline.Split('|');

            int i = 1;
            foreach (IXLCell cell in row.Cells())
            {
                if (cell.Value.ToString() != headerInfo[i - 1])
                    return false;

                i++;
            }

            return true;
        }



    }
}