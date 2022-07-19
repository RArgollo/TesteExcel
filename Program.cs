using ClosedXML.Excel;
using System.Linq;
namespace TesteExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var xls = new XLWorkbook(@"C:\Users\rafae\Documents\dívidas mensais.xlsx");
            var planilha = xls.Worksheets.First();
            var totalLinhas = planilha.Rows().Count();
            var cartao = double.Parse(planilha.Cell("D5").Value.ToString());
            var seguro = double.Parse(planilha.Cell("D6").Value.ToString());
            Console.WriteLine($"{cartao} e {seguro}");
        }
    }
}
