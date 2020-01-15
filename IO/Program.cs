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
            Excel excel = new Excel(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs3.xls", 1);
            ga.Unprotect();
            ga.ReadOrderColumn();
            ga.GetElementsFromListToArray();
            excel.ReadCellAsInt(69, 26);
            ga.SwapRandomGenesxTimes(100, 2);
            ga.PopulateColumn();
            ga.PopulatePopulation();
            ga.SelectionRouletteWheel();
            //ga.WriteToColumn();
            ga.CloseWithSaveAs(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs4.xls");
        }
    }
}
