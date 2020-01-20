using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IO
{
    class GeneticAlgorithm
    {
        public ArrayList uszeregowanie { get; set; }
        public int startingRow { get; set; }
        public int endingRow { get; set; }
        public int kolumnazUszeregowaniem { get; set; }
        public int sizeOfPopulation { get; set; }
        public int rowOfSumExcel { get; set; }
        public int colOfSumExcel { get; set; }
        public int[] uszeregowanieInt { get; set; }
        public int[] selectedParent1 { get; set; }
        public int[] selectedParent2 { get; set; }
        public int leastSum { get; set; }
        public double probabilityOfMutation { get; set; }
        public int[] orderOfLeastSum { get; set; }
        public int[] minSumOrder { get; set; }
        public int minSum { get; set; }
        public Dictionary<int[], int> population { get; set; }
        Excel excel = new Excel(@"E:\studia\V semestr\IO\Dane_sprawko2\Dane_S2_100_20DOBRA_WSPINACZKAAutomatycznie-odzyskany", "Arkusz1"); //tutaj wpisujemy ścieżkę do pliku z przygotowanym arkuszem pod algorytm genetyczny oraz numer arkusza (pierwsza ma indeks 1)
        //public delegate int PopulateColumnReadSum(int start, int end, int column, int[] ar, int arCount);
        Random random = new Random();

        public GeneticAlgorithm()
        {
            kolumnazUszeregowaniem = 24; //kolumna z której pobieramy nasze uszeregowanie
            startingRow = 2; //wiersz w którym zaczynają się nasze dane(nie nagłówek)
            endingRow = 101; //wiersz w którym kończą się nasze dane
            sizeOfPopulation = 10; //wielkość populacji
            uszeregowanie = new ArrayList();
            population = new Dictionary<int[], int>();
            rowOfSumExcel = 101; //wiersz z którego pobieramy naszą sumę (wynik)
            colOfSumExcel = 67; //kolumna z której pobieramy naszą sumę (wynik)
            minSum = 999999;
            probabilityOfMutation = 0.1; //prawdopodobieństwo mutacji
        }

        public void ReadOrderColumn() //metoda do obsługi excela
        {
            excel.ReadColumnToArray(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanie);
        }

        public void GetElementsFromListToArray() //metoda do obsługi excela
        {
            uszeregowanieInt = new int[uszeregowanie.Count];
            for (int i = 0; i < uszeregowanie.Count; i++)
            {
                uszeregowanieInt[i] = Convert.ToInt32(uszeregowanie[i]);
            }
        }

        public void CloseExcel() //metoda do obsługi excela
        {
            excel.Save();
            excel.Close();
        }

        public void CloseWithSaveAs(string path) //metoda do obsługi excela
        {
            excel.SaveAs(path);
            excel.Close();
        }

        public void Unprotect() //metoda do obsługi excela
        {
            excel.UnprotectSheet();
        }

        public void SwapRandomGenes(int startingRow, int endingRow, int col) //metodą do zamiany genów  //starting and ending row in excel
        {
            Random rnd = new Random();
            int random1;
            int random2;
            int tmp;

            do
            {
                random1 = rnd.Next(0, endingRow -2); //startingRow + 1, endingRow
                random2 = rnd.Next(0, endingRow -2); //startingRow + 1, endingRow
            } while (random1 != random2);

            tmp = uszeregowanieInt[random1];
            uszeregowanieInt[random1] = uszeregowanieInt[random2];
            uszeregowanieInt[random2] = tmp;
        }

        public int[] SwapRandomGenesxTimes(int xTimes, int seed) //metoda do zaminy genów x razy //starting and ending row in excel
        {
            Random rnd = new Random();
            //Random rnd2 = new Random(seed + 67);
            int random1;
            int random2;
            int tmp;
            int[] temp = new int[uszeregowanieInt.Length];
            temp = (int[])uszeregowanieInt.Clone();

            for (int i = 0; i < xTimes; i++)
            {
                do
                {
                    random1 = rnd.Next(0, endingRow - 2);
                   // Thread.Sleep(100); //random działa na podstawie zegara - usypiając program na 100ms wylosowane liczby powinny być bardziej rozbieżne
                    random2 = rnd.Next(0, endingRow - 2); 
                } while (random1 == random2);

                tmp = temp[random1];
                temp[random1] = temp[random2];
                temp[random2] = tmp;

                tmp = uszeregowanieInt[random1]; //tutaj rpzekazujemy przez wartość nie przez referencje
                uszeregowanieInt[random1] = uszeregowanieInt[random2];
                uszeregowanieInt[random2] = tmp;
            }
            return temp;
        }

        public void PopulatePopulation() //tworzenie pierwszej losowej populacji
        {
            for (int i = 0; i < sizeOfPopulation; i++)
            {
                SwapRandomGenesxTimes(100, i); 
                excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt.ToList(), uszeregowanieInt.Count());
                population.Add(SwapRandomGenesxTimes(20, i).ToArray(), excel.ReadCellAsInt(rowOfSumExcel, colOfSumExcel));
            }

            if (minSum > population.OrderBy(x => x.Value).First().Value)
            {
                minSum = population.OrderBy(x => x.Value).First().Value;
                minSumOrder = population.OrderBy(x => x.Value).First().Key;
            }
        }

        public void PopulateColumn() //wypełnianie kolumny
        {
            excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt.ToList(), uszeregowanieInt.Length);
        }

        public void SelectionRouletteWheel() //ruletka
        {
            int sumOfFitness = 0;
            double probability = 0;
            double sumOfProbabilities = 0;
            int numberInPopulation = 0;
            selectedParent1 = new int[uszeregowanieInt.Length];
            selectedParent2 = new int[uszeregowanieInt.Length];
            List<double> probabilities = new List<double>();
            List<double> cumulatedProbabilities = new List<double>();
            Dictionary<int[], Dictionary<int, double>> populationWithProbabilities = new Dictionary<int[], Dictionary<int, double>>();

            foreach (var individual in population)
            {
                sumOfFitness += individual.Value;
            }

            foreach (var individual in population) //prawdopodobieństwa
            {
                probability = (sumOfProbabilities + ((double)individual.Value / sumOfFitness));
                probabilities.Add(probability);
            }

            probabilities.Sort();

            for (int i = 0; i < probabilities.Count; i++) //kumulacja
            {
                sumOfProbabilities += probabilities[i];
                cumulatedProbabilities.Add(sumOfProbabilities);
            }

            var sortedDictionary = population.OrderByDescending(key => key.Value);

            for (int i = 0; i < population.Count; i++)
            {
                var tempDict = new Dictionary<int, double>();
                tempDict.Add(sortedDictionary.ElementAt(i).Value, cumulatedProbabilities[i]);
                populationWithProbabilities.Add(population.ElementAt(i).Key, tempDict);
            }

            for (int i = 0; i < sizeOfPopulation / 2; i++) 
            {
                Random rnd = new Random();
                double number = rnd.NextDouble();
                for (int individual = 0; individual < (sizeOfPopulation - 2); individual++)
                {
                    var helper = cumulatedProbabilities[individual]; //inaczej nie da się wyrzucić probabilities[i] z funckji niżej
                    var variable = populationWithProbabilities.ElementAt(individual).Value.TryGetValue(sortedDictionary.ElementAt(individual).Value, out helper);
                    var helper2 = cumulatedProbabilities[individual + 1];
                    var variable2 = populationWithProbabilities.ElementAt(individual + 1).Value.TryGetValue(sortedDictionary.ElementAt(individual + 1).Value, out helper2);
                    if (number >= helper && number < helper2)
                    {
                        for (int ite = 0; ite < uszeregowanieInt.Length; ite++)
                        {
                            selectedParent1[ite] = Convert.ToInt32(populationWithProbabilities.ElementAt(individual).Key[ite]);
                        }
                        population.Remove(population.ElementAt(numberInPopulation).Key);
                        population.Add(selectedParent1.ToArray(), excel.WriteToColumnFromArrayIntsAndReadSum(startingRow, endingRow, kolumnazUszeregowaniem, colOfSumExcel, selectedParent1.ToList(), selectedParent1.Length));
                        //minSum = population.OrderBy(x => x.Value).First().Value;
                        //minSumOrder = population.OrderBy(x => x.Value).First().Key;
                        numberInPopulation++;
                        break;
                    }
                }
            }

            for (int i = 0; i < sizeOfPopulation / 2; i++)
            {
                Random rnd = new Random();
                double number = rnd.NextDouble();
                Random rnd2 = new Random(i + 27);
                double number2 = rnd2.NextDouble();
                for (int individual = 0; individual < population.Count - 2; individual++)
                {
                    var helper = cumulatedProbabilities[individual]; //inaczej nie da się wyrzucić probabilities[i] z funckji niżej
                    var variable = populationWithProbabilities.ElementAt(individual).Value.TryGetValue(sortedDictionary.ElementAt(individual).Value, out helper);
                    var helper2 = cumulatedProbabilities[individual + 1];
                    var variable2 = populationWithProbabilities.ElementAt(individual + 1).Value.TryGetValue(sortedDictionary.ElementAt(individual + 1).Value, out helper2);
                    if (number2 >= helper && number < helper2)
                    {
                        for (int ite = 0; ite < uszeregowanieInt.Length; ite++)
                        {
                            selectedParent2[ite] = populationWithProbabilities.ElementAt(individual).Key[ite];

                        }
                        population.Remove(population.ElementAt(numberInPopulation).Key);
                        population.Add(selectedParent2.ToArray(), excel.WriteToColumnFromArrayIntsAndReadSum(startingRow, endingRow, kolumnazUszeregowaniem, colOfSumExcel, selectedParent2.ToList(), selectedParent1.Length));
                        if (minSum > population.OrderBy(x => x.Value).First().Value)
                        {
                            minSum = population.OrderBy(x => x.Value).First().Value;
                            minSumOrder = population.OrderBy(x => x.Value).First().Key;
                        }
                        numberInPopulation++;
                        break;
                    }
                }
            }
        }

        public void MakingChilds() //tworzenie dzieci
        {
            int currentIndex = 0;
            int[] tmpChild;
            int[] tmpChild2;
            List<int> valuesFromParent2;
            List<int> valuesFromParent1;
            int crossoverPoint;

            do
            {
                tmpChild = new int[uszeregowanieInt.Length];
                tmpChild2 = new int[uszeregowanieInt.Length];
                valuesFromParent2 = new List<int>();
                valuesFromParent1 = new List<int>();
                crossoverPoint = random.Next(startingRow + 1, endingRow - 3);

                //dziecko1
                for (int i = 0; i <= crossoverPoint; i++)
                {
                    tmpChild[i] = population.ElementAt(currentIndex).Key[i];
                }

                for (int ite = 0; ite < population.ElementAt(currentIndex).Key.Count(); ite++)
                {
                    bool a = Array.Exists(tmpChild, element => element == population.ElementAt(currentIndex + 1).Key[ite]);
                    //if (tmpChild.(population.ElementAt(currentIndex + 1).Key[ite]))
                    if(!a)
                    {
                        valuesFromParent2.Add(population.ElementAt(currentIndex + 1).Key[ite]);
                    }
                        
                }

                int j = 0;
                for (int i = crossoverPoint + 1; i < endingRow - 1; i++)
                {
                    tmpChild[i] = valuesFromParent2[j];
                    j++;
                }

                if(random.NextDouble() < probabilityOfMutation) //mutacja dziecka pierwszego
                {
                    int tmpRand = random.Next(0, endingRow - 2);
                    int tmpRand2 = random.Next(0, endingRow - 2);

                    int tmp = tmpChild[tmpRand];
                    tmpChild[tmpRand] = tmpChild[tmpRand2];
                    tmpChild[tmpRand2] = tmp;
                }

                population.Remove(population.ElementAt(currentIndex).Key);
                population.Add(tmpChild.ToArray(), excel.WriteToColumnFromArrayIntsAndReadSum(startingRow, endingRow, kolumnazUszeregowaniem, colOfSumExcel, tmpChild.ToList(), tmpChild.Length));

                //dziecko2
                for (int i = 0; i <= crossoverPoint; i++)
                {
                    tmpChild2[i] = population.ElementAt(currentIndex + 1).Key[i];
                }

                for (int ite = 0; ite < population.ElementAt(currentIndex + 1).Key.Count(); ite++)
                {
                    int i = 0;
                    if (!tmpChild2.Contains(population.ElementAt(currentIndex).Key[ite]))
                    {
                        valuesFromParent1.Insert(i, population.ElementAt(currentIndex).Key[ite]);
                        i++;
                    }
                }

                j = 0;
                for (int i = crossoverPoint + 1; i <= endingRow - 2; i++)
                {
                    tmpChild2[i] = valuesFromParent1[j];
                    j++;
                }

                if (random.NextDouble() < probabilityOfMutation) //mutacja dziecka drugiego
                {
                    int tmpRand = random.Next(startingRow, endingRow - 3);
                    int tmpRand2 = random.Next(startingRow, endingRow - 3);

                    int tmp = tmpChild2[tmpRand];
                    tmpChild2[tmpRand] = tmpChild2[tmpRand2];
                    tmpChild2[tmpRand2] = tmp;
                }

                population.Remove(population.ElementAt(currentIndex+1).Key);
                population.Add(tmpChild2.ToArray(), excel.WriteToColumnFromArrayIntsAndReadSum(startingRow, endingRow, kolumnazUszeregowaniem, colOfSumExcel, tmpChild2.ToList(), tmpChild2.Length));

                currentIndex = currentIndex + 2;
            } while (currentIndex < 10);

            if (minSum > population.OrderBy(x => x.Value).First().Value)
            {
                minSum = population.OrderBy(x => x.Value).First().Value;
                minSumOrder = population.OrderBy(x => x.Value).First().Key;
            }

        }
       
        
        public void WriteArrayToColumnBestOrder(int startRow, int endRow, int col) //wpisanie do kolumny najlpeszego ułożenia
        {
            excel.WriteToColumnFromArrayInts(startRow, endRow, col, minSumOrder.ToList(), minSumOrder.Count());
        }

        public void WriteBestSumToCell(int row, int col) //wpisanie najlepszej sumy
        {
            excel.WriteToCellInt(row, col, minSum);
        }

        public void excelQuit() //metoda do obsługi excela
        {
            excel.Quit();
        }
    }
}
