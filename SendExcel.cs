using ClosedXML.Excel;
using System.Diagnostics;
namespace RPACorreios
{
    public class SendExcel
    {
        public void retornoTabela(string logradouro, string bairro, string localidade, string CEP)
        {
            using (var workbok = new XLWorkbook())
            {
                var worksheet = workbok.Worksheets.Add("Planilha1");

                worksheet.Cell("A1").Value = "Logradouro/Nome";
                worksheet.Cell("B1").Value = "Bairro/Distrito";
                worksheet.Cell("C1").Value = "Localidade/UF";
                worksheet.Cell("D1").Value = "CEP";

                worksheet.Cell("A2").Value = logradouro;
                worksheet.Cell("B2").Value = bairro;
                worksheet.Cell("C2").Value = localidade;
                worksheet.Cell("D2").Value = CEP;

                workbok.SaveAs(@"c:\temp\RPACorreios.xlsx");
            }

            Process.Start(new ProcessStartInfo(@"c:\temp\RPACorreios.xlsx") { UseShellExecute = true }); //Esse comando chama o programa padrão que o computador utiliza para abrir arquivos xlsx

        }
    }
}
