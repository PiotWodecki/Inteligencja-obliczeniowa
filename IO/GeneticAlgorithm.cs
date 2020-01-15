﻿using System;
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
        public Dictionary<int[], int> population { get; set; }
        Excel excel = new Excel(@"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs3.xls", 1);

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
                excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt, uszeregowanieInt.Count());
                excel.ReadCellAsInt(rowOfSumExcel, colOfSumExcel);
                population.Add(SwapRandomGenesxTimes(20, i).ToArray(), excel.ReadCellAsInt(rowOfSumExcel, colOfSumExcel));
            }
        }

        public void PopulateColumn()
        {
            excel.WriteToColumnFromArrayInts(startingRow, endingRow, kolumnazUszeregowaniem, uszeregowanieInt, uszeregowanieInt.Length);
        }

        public void SelectionRouletteWheel()
        {
            int sumOfFitness = 0;
            double probability = 0;
            double sumOfProbabilities = 0;
            List<double> probabilities = new List<double>();
            List<double> cumulatedProbabilities = new List<double>();
            Dictionary<int[], Dictionary<int, double>> populationWithProbabilities = new Dictionary<int[], Dictionary<int, double>>();
            int[] selectedParent1 = new int[population.Count];
            int[] selectedParent2 = new int[population.Count];


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

            for(int i = 0; i < probabilities.Count; i++)
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

            for(int i=0; i< population.Count; i++)
            {
                var tempDict = new Dictionary<int, double>();
                tempDict.Add(sortedDictionary.ElementAt(i).Value, cumulatedProbabilities[i]);
                populationWithProbabilities.Add(population.ElementAt(i).Key, tempDict);
            }

            for(int i=0; i<sizeOfPopulation; i++)
            {         
                    Random rnd = new Random();
                    double number = rnd.NextDouble();
                    for(int individual=0; i<population.Count-2; individual++)
                    {
                        var helper = cumulatedProbabilities[individual]; //inaczej nie da się wyrzucić probabilities[i] z funckji niżej
                        var variable = populationWithProbabilities.ElementAt(i).Value.TryGetValue(sortedDictionary.ElementAt(i).Value, out helper);
                        var helper2 = cumulatedProbabilities[individual + 1];
                        var variable2 = populationWithProbabilities.ElementAt(i).Value.TryGetValue(sortedDictionary.ElementAt(i).Value, out helper2);
                        if (number >= helper && number < helper2)
                        {
                            for(int ite = 0; ite< populationWithProbabilities.Count; ite++)
                            {
                                selectedParent1[ite] = populationWithProbabilities.ElementAt(individual).Key[ite];
                            }
                        }
                    }
            }



        }

        





    }
}
