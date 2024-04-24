﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using ExcelDataReader;
using System.Text.RegularExpressions;

namespace PlagiarismValidation
{
    class Program
    {
        static void Main(string[] args)
        {

            Tuple<string, string, float, int>[] edges = ReadFromExcelFile();
            //var edges = new Tuple<string, string, float>[numberOfedges];

            //foreach (var edge in edges)
            //{
            //    // Access individual elements of the tuple
            //    string element1 = edge.Item1;
            //    string element2 = edge.Item2;
            //    float element3 = edge.Item3;
            //    int element4 = edge.Item4;

            //    // Print the elements
            //    Console.WriteLine($"Doc 1: {element1}, Doc 2: {element2}, Percentage: {element3}, Lines Matched: {element4}");
            //}

        }

        public static Tuple<string, string, int, int, int>[] ReadFromExcelFile()
        {
            string inputfilePath = "D:\\Uni Related\\Algorithms\\Project\\MATERIALS\\[3] Plagiarism Validation\\Algorithm-Project\\PlagiarismValidation\\trial input.xlsx";
            int numberOfEdges;
            Tuple<string, string, int, int, int>[] edges;

            using (var stream = File.Open(inputfilePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = null;

                reader = ExcelReaderFactory.CreateReader(stream);

                DataSet resultDataSet = reader.AsDataSet();

                DataTable table = resultDataSet.Tables[0];

                string column1 = "";
                string column2 = "";
                string column3 = "";

                string[] similarityPercentageArr;
                string similarityPercentage;
                string mergedSimilarityPercentage;
                Match matchPercentage;
                float percentageDoc1 , percentageDoc2;
                Regex percentageRegex = new Regex(@"\(\d+%\)");

                int linesMatched;

                numberOfEdges = table.Rows.Count - 1;
                Console.WriteLine(numberOfEdges);
                edges = new Tuple<string, string, int, int, int>[numberOfEdges];

                for (int i = 1; i < table.Rows.Count; i++)
                {

                    DataRow row = table.Rows[i];

                    column1 = row[0].ToString();

                    // Retrieving the similarity percentage of each document in column 1
                    int indexOfbrac1 = column1.LastIndexOf('(');
                    string firstPath = column1.Substring(0, indexOfbrac1);
                    int firstPercntage = int.Parse(column1.Substring(indexOfbrac1 + 1, column1.Length - (indexOfbrac1 + 3)));

                    column2 = row[1].ToString();

                    // Retrieving the similarity percentage of each document in column 2
                    int indexOfbrac2 = column2.LastIndexOf('(');
                    string secondPath = column2.Substring(0, indexOfbrac2);
                    int secondPercntage = int.Parse(column2.Substring(indexOfbrac2 + 1, column2.Length - (indexOfbrac2 + 3)));

                    // Retrieving the Lines Matched in column 3
                    column3 = row[2].ToString();
                    linesMatched = Convert.ToInt32(column3);

                    if(percentageDoc1 >= percentageDoc2)
                    {
                        percentageToBeRecorded = percentageDoc1;
                    }
                    else
                    {
                        percentageToBeRecorded = percentageDoc2;
                    }
                    edges[i - 1] = new Tuple<string, string, float, int>(documentPath1, documentPath2, percentageToBeRecorded, linesMatched);
                }

                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }

            }
            return edges;
        }

        public static void MakeDic(Tuple<string, string, float, int>[] edges, Dictionary<string, List<Tuple<string, float, int>>> elements)
        {
            foreach (var edge in edges)
            {
                if (elements.ContainsKey(edge.Item1))
                {
                    elements[edge.Item1].Add(Tuple.Create(edge.Item2, edge.Item3, edge.Item4));
                }
                else
                {
                    elements[edge.Item1] = new List<Tuple<string, float, int>>() { Tuple.Create(edge.Item2, edge.Item3, edge.Item4) };
                }
                if (elements.ContainsKey(edge.Item2))
                {
                    elements[edge.Item2].Add(Tuple.Create(edge.Item1, edge.Item3, edge.Item4));
                }
                else
                {
                    elements[edge.Item2] = new List<Tuple<string, float, int>>() { Tuple.Create(edge.Item1, edge.Item3, edge.Item4) };
                }
            }
        }

        public static void Output()
        {

        }
    }
}
