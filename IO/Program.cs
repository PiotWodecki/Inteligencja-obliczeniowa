using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IO
{
    class Program
    {
        static void Main(string[] args)
        {
            GeneticAlgorithm ga = new GeneticAlgorithm();
            ga.Unprotect();
            ga.ReadOrderColumn();
            ga.GetElementsFromListToArray();
            ga.SwapRandomGenesxTimes(100, 2);
            ga.PopulateColumn();
            ga.PopulatePopulation();
            //tu będziemy zaczynać pętle
            for (int i = 0; i < 10; i++)
            {
                ga.SelectionRouletteWheel();
                ga.MakingChilds();
            }
            //tu będziemy kończyć
            ga.WriteArrayToColumnBestOrder(3, 69, 34);//tu uzupelniamy poczatkowy i koncowy wierszy oraz kolumne, tam gdzie chcemy wpisac najlepsza kolejnosc
            ga.WriteBestSumToCell(71, 34); //tu uzupelniamy wiersz i kolumne, tam gdzie chcemy wpisac najlepszy wynik
            ga.CloseWithSaveAs(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs8.xls");
            Console.WriteLine(ga.minSum.ToString() + "\n" + ga.minSumOrder.ToString());
            Console.ReadKey();
            ga.excelQuit();
        }
    }
}
