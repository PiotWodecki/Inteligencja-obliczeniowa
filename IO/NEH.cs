using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IO
{
    class NEH
    {
        //sortujemy plik w excelu!!!! plik wejściowy jest już posortowany!

        Excel excel = new Excel(@"E:\studia\V semestr\IO\Dane_sprawko2\Dane_S2_200_20DOBRA_WSPINACZKA.xlsm"); //tutaj należy wpisać ścieżkę do pliku który chcemy otworzyć

        public int kolumnazUszeregowaniem { get; set; }
        public int startingRow { get; set; }
        public int endingRow { get; set; }
        public List<int> uszeregowanie { get; set; }
        public ArrayList posortowaneUszeregowanie { get; set; }
        public int rowOfSumExcel { get; set; }
        public int colOfSumExcel { get; set; }
        public int tmpSum { get; set; }
        public int currentRow { get; set; }
        public int iloscZadan { get; set; }

        public NEH()
        {
            kolumnazUszeregowaniem = 24; //kolumna w której wpisujemy uszeregowanie (tabela środkowa)
            startingRow = 2; //wiersz w którym jest pierwsza wartość (nie nagłówek)
            endingRow = 201; //wiersz w którym kończą się wartośći
            uszeregowanie = new List<int>();
            posortowaneUszeregowanie = new ArrayList();
            colOfSumExcel = 67; //kolumna z której zczytujemy sumy (ostatnia tabela)
            tmpSum = 0;
            currentRow = 3;
            iloscZadan = 50; //ilość zadań w pliku
        }

        public void ReadSortedOrderColumn() //obsługa excela
        {
            excel.ReadColumnToArray(startingRow, endingRow, kolumnazUszeregowaniem, posortowaneUszeregowanie);
        }


        public void CloseExcel() //obsługa excela
        {
            excel.Save();
            excel.Close();
        }

        public void CloseWithSaveAs(string path)//obsługa excela
        {
            excel.SaveAs(path);
            excel.Close();
        }

        public void Unprotect() //obsługa excela
        {
            excel.UnprotectSheet();
        }


        public void excelQuit() //obsługa excela
        {
            excel.Quit();
        }

        public void ReadSumFromExcel() //obsługa excela
        {
            tmpSum = excel.ReadCellAsInt(currentRow, colOfSumExcel);
        }

        public int ReadCell(int row,int col) //obsługa excela
        {
           return excel.ReadCellAsInt(row, col);
        }

        public void WriteToCell(int row, int col, int value) //obsługa excela
        {
            excel.WriteToCellInt(row, col, value);
        }


        public void ChoosingFirstElements() //w tej metodzie sprawdzamy które ułożenie dwóch pierwszych elementów jest najlepsze
        {
            uszeregowanie.Add(Convert.ToInt32(posortowaneUszeregowanie[0]));
            uszeregowanie.Add(Convert.ToInt32(posortowaneUszeregowanie[1]));

            ReadSumFromExcel();

            int tmp;

            WriteToCell(startingRow, kolumnazUszeregowaniem, uszeregowanie[1]);
            WriteToCell(startingRow+1, kolumnazUszeregowaniem, uszeregowanie[0]);

            if (ReadCell(startingRow + 1, colOfSumExcel) < tmpSum)
            {
                tmpSum = ReadCell(startingRow + 1, colOfSumExcel);

                tmp = uszeregowanie[0];
                uszeregowanie[0] = uszeregowanie[1];
                uszeregowanie[1] = tmp;

            }
            else
            {
                WriteToCell(startingRow, kolumnazUszeregowaniem, uszeregowanie[0]);
                WriteToCell(startingRow, kolumnazUszeregowaniem, uszeregowanie[1]);
            }
            currentRow++;
        }

        public void PopulateUszeregowanie() //w tej metodzie sprawdzamy uszeregowanie od elementu 3 do końca
        {
            for (int k = 3; k < iloscZadan; k++)
            {
                Dictionary<int, int> dictionaryOfRowInsertAndSum = new Dictionary<int, int>();
                for (int i = 0; i < uszeregowanie.Count + 1; i++)
                {
                    uszeregowanie.Insert(i, Convert.ToInt32(posortowaneUszeregowanie[uszeregowanie.Count]));
                    for (int j = 0; j < uszeregowanie.Count; j++)
                    {
                        WriteToCell(startingRow + j, kolumnazUszeregowaniem, uszeregowanie[j]);
                    }
                    tmpSum = ReadCell(currentRow, colOfSumExcel);
                    dictionaryOfRowInsertAndSum.Add(i, tmpSum);
                    uszeregowanie.RemoveAt(i);
                }
                int minSum = dictionaryOfRowInsertAndSum.OrderBy(x => x.Value).First().Value;
                uszeregowanie.Insert(dictionaryOfRowInsertAndSum.OrderBy(x => x.Value).First().Key, Convert.ToInt32(posortowaneUszeregowanie[uszeregowanie.Count]));
                currentRow++;
            }
        }
    }
}
