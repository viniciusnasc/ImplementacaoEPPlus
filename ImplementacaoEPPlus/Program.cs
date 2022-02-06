using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

// Para utilizar o EpPlus é necessário informar o tipo de licença
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
// ExcelPackage.LicenseContext = LicenseContext.Commercial;

// Em app simples podemos colocar direto no código, como acima, caso seja mais complexo, colocar no appsettings:
/*
   "EPPlus": {
        "ExcelPackage": {
            "LicenseContext": "Commercial" // Ou NonCommercial
            }
        }
*/

#region Criando Novo arquivo
using (ExcelPackage ep = new())
{
    // Propriedades do arquivo
    ep.Workbook.Properties.Author = "Vinicius";
    ep.Workbook.Properties.Title = "Estudo EPP Plus";
    ep.Workbook.Properties.Subject = "Estudo de implementação da biblioteca EPP Plus";
    ep.Workbook.Properties.Created = DateTime.Now;
    ep.Workbook.Date1904 = true;

    // Criação das planilhas
    ExcelWorksheet planilha = ep.Workbook.Worksheets.Add("Primeira Planilha");
    ExcelWorksheet segundaPlanilha = ep.Workbook.Worksheets.Add("Segunda Planilha");

    // Personalização
    planilha.Rows.Style.Font.Bold = true; // Altera todas as linhas
   // planilha.Row(1).Style.Font.Bold = true; // Negrido em toda linha 1
   // planilha.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid; //Modifica o padrão da linha, fundo solido, por exemplo
   // planilha.Row(2).Style.Fill.BackgroundColor.SetColor(Color.Black); // Alterar a cor de toda linha 2 - precisa que a patternType seja informada antes para funcionar
    segundaPlanilha.Column(1).Style.Font.Italic = true; // Italico em toda coluna A
    segundaPlanilha.Column(2).Width = 55; // Alterar largura da coluna
    segundaPlanilha.Row(1).Height = 55; // Altera tamanho da linha
    segundaPlanilha.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //Alinhamento a esquerda
   // planilha.Cells["A1"].Style.Font.Color.SetColor(Color.Red);

    // Proteção da planilha
    /*planilha.Protection.IsProtected = false; // Protege toda a planilha
    planilha.Protection.AllowInsertColumns = true; // Permite criar colunas
    planilha.Protection.SetPassword("seila"); // Atribui password para a proteção
    planilha.Protection.AllowSelectLockedCells = false; // Permissão de selecionar celulas bloqueadas
    planilha.Cells.Style.Locked = true;  // bloqueia as celulas
    planilha.Cells.Style.Hidden = true; // Oculta o conteudo da celula*/

    // Criando uma tabela
    planilha.Tables.Add(planilha.Cells["A1:D4"], "novatabela");
    planilha.Tables[0].ShowFilter = false;
    planilha.Tables[0].TableStyle = OfficeOpenXml.Table.TableStyles.Dark7;
    planilha.Tables[0].TableBorderStyle.BorderAround(ExcelBorderStyle.Dotted);
    planilha.Tables[0].Columns[0].Name = "Nome";
    planilha.Tables[0].Columns[1].Name = "Idade";
    planilha.Tables[0].Columns[2].Name = "Nacionalidade";
    planilha.Tables[0].Columns[3].Name = "Naturalidade";

    // Inclusão de valores nas celulas
    segundaPlanilha.Cells["A1"].Value = "Eis aqui uma tentativa de inclusão de formula:";

    segundaPlanilha.Cells["A2"].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
    segundaPlanilha.Cells["A2"].Value = 5;

    segundaPlanilha.Cells["A3"].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
    segundaPlanilha.Cells["A3"].Value = 4;

    segundaPlanilha.Cells["B1"].Formula = "=A2+A3";
    segundaPlanilha.Cells["B1"].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";

    // Salvando o arquivo
    FileInfo fi = new(@"C:\Users\Vini\Desktop\Nova pasta\EppPlusTeste.xlsx");
    ep.SaveAs(fi);
}
#endregion

/*
#region Abrindo um arquivo já existente
FileInfo file = new(@"C:\Users\Vini\Desktop\Nova pasta\EppPlusTeste.xlsx");
using(ExcelPackage excel = new(file))
{
    //Selecionando uma planilha
    // Por Index
    ExcelWorksheet primeiraPlanilha = excel.Workbook.Worksheets[0];
    //Por nome
    ExcelWorksheet planilhaPrimeira = excel.Workbook.Worksheets["Primeira Planilha"];
    // Por Linq
    ExcelWorksheet segundaPlanilha = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Segunda Planilha");

    // Pegando o valor de uma celula
    string valorA1PrimeiraPlanilha = primeiraPlanilha.Cells["C1"].Value.ToString();
    string valorA1SegPlanilha = segundaPlanilha.Cells["A2"].Value.ToString();
    string valorB1SegundaPlanilha = segundaPlanilha.Cells["B1"].Value.ToString();
}
#endregion*/