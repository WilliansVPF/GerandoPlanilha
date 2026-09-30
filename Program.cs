using System.Globalization;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;

var path = "Notas.csv";

var config = new CsvConfiguration(CultureInfo.InvariantCulture)
{
    HasHeaderRecord = true,
    Delimiter = ",",
    PrepareHeaderForMatch = args => args.Header.ToLower()
};

using var reader = new StreamReader(path);
using var csv = new CsvReader(reader, config);

var registros = csv.GetRecords<Aluno>().ToList();

using var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("titulo");

worksheet.Cell(1,1).Value = "ID";
worksheet.Cell(1,2).Value = "Nome";
worksheet.Cell(1,3).Value = "Nota 1";
worksheet.Cell(1,4).Value = "Nota 2";
worksheet.Cell(1,5).Value = "Nota 3";
worksheet.Cell(1,6).Value = "Nota 4";
worksheet.Cell(1,7).Value = "Media";
worksheet.Cell(1,8).Value = "Situação";

int linhaAtual = 2;

foreach (var pessoa in registros)
{
    // Console.WriteLine(pessoa.ID);
    worksheet.Cell(linhaAtual,1).Value = pessoa.Id;
    worksheet.Cell(linhaAtual,2).Value = pessoa.Nome;
    worksheet.Cell(linhaAtual,3).Value = pessoa.N1;
    worksheet.Cell(linhaAtual,4).Value = pessoa.N2;
    worksheet.Cell(linhaAtual,5).Value = pessoa.N3;
    worksheet.Cell(linhaAtual,6).Value = pessoa.N4;
    worksheet.Cell(linhaAtual,7).Value = pessoa.NM;
    if (pessoa.NM >= 6) worksheet.Cell(linhaAtual,8).Value = "Aprovado";
    else worksheet.Cell(linhaAtual,8).Value = "Reprovado";
    // worksheet.Cell(linhaAtual,7).Value = pessoa.NM;

    linhaAtual++;
}

var intervalo = worksheet.Range($"G2:G{linhaAtual -1}");
var intervaloStatus = worksheet.Range($"H2:H{linhaAtual -1}");

var regraAprovado = intervalo.AddConditionalFormat();

regraAprovado.WhenEqualOrGreaterThan(6.0); 
// regraAprovado.WhenEquals("Aprovado"); 
regraAprovado.Style.Fill.BackgroundColor = XLColor.FromHtml("#2ECC71");
regraAprovado.Style.Font.FontColor = XLColor.White;
regraAprovado.Style.Font.Bold = true;

var regraReprovado = intervalo.AddConditionalFormat();

regraReprovado.WhenLessThan(6.0); 
// regraReprovado.WhenEquals("Reprovado"); 
regraReprovado.Style.Fill.BackgroundColor = XLColor.Red;
regraReprovado.Style.Font.FontColor = XLColor.White;
regraReprovado.Style.Font.Bold = true;

var regraStatusAprovado = intervaloStatus.AddConditionalFormat();
regraStatusAprovado.WhenEqualOrLessThan("\"Aprovado\""); // Aspas duplas internas são obrigatórias para texto no Excel
// Dica: Nas versões mais recentes do ClosedXML, você também pode usar: regraStatusAprovado.WhenContains("Aprovado");

regraStatusAprovado.Style.Fill.BackgroundColor = XLColor.FromHtml("#2ECC71");
regraStatusAprovado.Style.Font.FontColor = XLColor.White;
regraStatusAprovado.Style.Font.Bold = true;

// Regra para "Reprovado"
var regraStatusReprovado = intervaloStatus.AddConditionalFormat();
regraStatusReprovado.WhenEqualOrLessThan("\"Reprovado\"");

regraStatusReprovado.Style.Fill.BackgroundColor = XLColor.Red;
regraStatusReprovado.Style.Font.FontColor = XLColor.White;
regraStatusReprovado.Style.Font.Bold = true;

worksheet.Columns().AdjustToContents();
workbook.SaveAs("Notas.xlsx");
