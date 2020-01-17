using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        Excel excel = new Excel(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs5.xls", 1);
        public delegate int PopulateColumnReadSum(int start, int end, int column, int[] ar, int arCount);
        Random random = new Random();

        public GeneticAlgorithm()
        {
            kolumnazUszeregowaniem = 10;
            startingRow = 3;
            endingRow = 69;
            sizeOfPopulation = 10;
            uszeregowanie = new ArrayList();
            population = new Dictionary<int[], int>();
            rowOfSumExcel = 69;
            colOfSumExcel = 26;
            minSum = 999999;
            probabilityOfMutation = 0.1;
        }

        public void ReadOrderColumn()
        {
            excel.ReadColumnToArray(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanie);
        }

        public void GetElementsFromListToArray()
        {
            uszeregowanieInt = new int[uszeregowanie.Count];
            for (int i = 0; i < uszeregowanie.Count; i++)
            {
                uszeregowanieInt[i] = Convert.ToInt32(uszeregowanie[i]);
            }
        }

        //public void changeArray()
        //{
        //    for(int i =0; i< uszeregowanie.Count; i++)
        //    {
        //        int zmiana = Convert.ToInt32(uszeregowanie[i]) + 1;
        //        uszeregowanie[i]=0;
        //    }
        //}

        //public void WriteToColumn()
        //{
        //    changeArray();
        //    excel.WriteToColumnFromArray(startingRow, endingRow, kolumnazUszeregowaniem,  uszeregowanie);
        //}

        public void CloseExcel()
        {
            excel.Save();
            excel.Close();
        }

        public void CloseWithSaveAs(string path)
        {
            excel.SaveAs(path);
            excel.Close();
        }

        public void Unprotect()
        {
            excel.UnprotectSheet();
        }

        public void SwapRandomGenes(int startingRow, int endingRow, int col) //starting and ending row in excel
        {
            Random rnd = new Random();
            int random1;
            int random2;
            int tmp;

            do
            {
                random1 = rnd.Next(startingRow + 1, endingRow);
                random2 = rnd.Next(startingRow + 1, endingRow);
            } while (random1 != random2);

            tmp = uszeregowanieInt[random1];
            uszeregowanieInt[random1] = uszeregowanieInt[random2];
            uszeregowanieInt[random2] = tmp;
        }

        public int[] SwapRandomGenesxTimes(int xTimes, int seed) //starting and ending row in excel
        {
            Random rnd = new Random(seed);
            Random rnd2 = new Random(seed + 67);
            int random1;
            int random2;
            int tmp;
            int[] temp = new int[uszeregowanieInt.Length];
            temp = (int[])uszeregowanieInt.Clone();

            for (int i = 0; i < xTimes; i++)
            {
                do
                {
                    random1 = rnd.Next(startingRow - 3, endingRow - 4);
                    random2 = rnd2.Next(startingRow - 3, endingRow - 4);
                } while (random1 == random2);

                tmp = temp[random1];
                temp[random1] = temp[random2];
                temp[random2] = tmp;

                tmp = uszeregowanieInt[random1];
                uszeregowanieInt[random1] = uszeregowanieInt[random2];
                uszeregowanieInt[random2] = tmp;
            }
            return temp;
        }

        public void PopulatePopulation()
        {
            for (int i = 0; i < sizeOfPopulation; i++)
            {
                SwapRandomGenesxTimes(100, i);
                excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt.ToList(), uszeregowanieInt.Count());
                excel.ReadCellAsInt(rowOfSumExcel, colOfSumExcel);
                population.Add(SwapRandomGenesxTimes(20, i).ToArray(), excel.ReadCellAsInt(rowOfSumExcel, colOfSumExcel));
            }
        }

        public void PopulateColumn()
        {
            excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt.ToList(), uszeregowanieInt.Length);
        }

        public void SelectionRouletteWheel()
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
            //int[] selectedParent1 = new int[uszeregowanieInt.Length];
            //int[] selectedParent2 = new int[uszeregowanieInt.Length];
            //PopulateColumnReadSum pcrs;
            //PopulateColumnReadSum pcrs1= new PopulateColumnReadSum(excel.WriteToColumnFromArrayInts);
            //PopulateColumnReadSum pcrs2;



            foreach (var individual in population)
            {
                sumOfFitness += individual.Value;
            }

            foreach (var individual in population)
            {
                probability = (sumOfProbabilities + ((double)individual.Value / sumOfFitness));
                //sumOfProbabilities += probability;
                probabilities.Add(probability);
            }

            //var sortedDictionary = new Dictionary<int, double>();

            probabilities.Sort();

            for (int i = 0; i < probabilities.Count; i++)
            {
                sumOfProbabilities += probabilities[i];
                cumulatedProbabilities.Add(sumOfProbabilities);
            }
            //foreach(KeyValuePair<int[], int> individual in population.OrderByDescending(key => key.Value)
            //{
            //    sortedDictionary.Add(individual.Value);
            //}
            var sortedDictionary = population.OrderByDescending(key => key.Value);
            //var sortedDictionary = population.Values.ToList();
            //sortedDictionary.Sort();

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
                        minSum = population.OrderBy(x => x.Value).First().Value;
                        minSumOrder = population.OrderBy(x => x.Value).First().Key;
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

        public void MakingChilds()
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
                    int i = 0;
                    if (!tmpChild.Contains(population.ElementAt(currentIndex + 1).Key[ite]))
                    {
                        valuesFromParent2.Insert(i, population.ElementAt(currentIndex + 1).Key[ite]);
                        i++;
                    }
                }

                int j = 0;
                for (int i = crossoverPoint + 1; i < endingRow - 2; i++)
                {
                    tmpChild[i] = valuesFromParent2[j];
                    j++;
                }

                if(random.NextDouble() < probabilityOfMutation)
                {
                    int tmpRand = random.Next(startingRow, endingRow - 3);
                    int tmpRand2 = random.Next(startingRow, endingRow - 3);

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
                for (int i = crossoverPoint + 1; i < endingRow - 2; i++)
                {
                    tmpChild2[i] = valuesFromParent1[j];
                    j++;
                }

                if (random.NextDouble() < probabilityOfMutation)
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
       
        

        public void excelQuit()
        {
            excel.Quit();
        }
        





    }
}
