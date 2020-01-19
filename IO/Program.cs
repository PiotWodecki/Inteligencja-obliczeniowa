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
            //w tym pliku znajduje się algorytm genetyczny z selekcją ruletki oraz krzyżówką jednopunktową
            //Żeby wykonać algorytm genetyczny należy odkomentować tworzenie obiektu ga oraz wywołanie jego metod

            //GeneticAlgorithm ga = new GeneticAlgorithm();
            //    ga.Unprotect();
            //    ga.ReadOrderColumn();
            //    ga.GetElementsFromListToArray();
            //    ga.SwapRandomGenesxTimes(100, 2);
            //    ga.PopulateColumn();
            //    ga.PopulatePopulation();
            //    //tu będziemy zaczynać pętle
            //    for (int i = 0; i < 10; i++)
            //    {
            //        ga.SelectionRouletteWheel();
            //        ga.MakingChilds();
            //    }
            //    //tu będziemy kończyć
            //    ga.WriteArrayToColumnBestOrder(3, 69, 34);//tu uzupelniamy poczatkowy i koncowy wierszy oraz kolumne, tam gdzie chcemy wpisac najlepsza kolejnosc
            //    ga.WriteBestSumToCell(71, 34); //tu uzupelniamy wiersz i kolumne, tam gdzie chcemy wpisac najlepszy wynik
            //    ga.CloseWithSaveAs(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs8.xls"); //tutaj wpisujemy ścieżkę do miejsca gdzie chcemy zapisać gotowy plik
            //    Console.WriteLine(ga.minSum.ToString());
            //    Console.ReadKey();
            //    ga.excelQuit();

            //Żeby wykonać algorytm genetyczny należy odkomentować tworzenie obiektu neh oraz wywołanie jego metod
            NEH neh = new NEH();
            neh.Unprotect();
            neh.ReadSortedOrderColumn();
            neh.ChoosingFirstElements();
            neh.PopulateUszeregowanie();
            neh.CloseWithSaveAs(@"E:\studia\V semestr\IO\Sprawozdanie2\NEH\checking"); //tutaj wpisujemy ścieżkę do miejsca gdzie chcemy zapisać gotowy plik
            neh.excelQuit();
            Console.ReadKey();
        }
    }
}
