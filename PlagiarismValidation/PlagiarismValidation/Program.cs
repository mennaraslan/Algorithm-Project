using System;
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
            //string file_path;
            
            Tuple<string, string, int, int, int>[] edges = ReadFromExcelFile();
            Dictionary<string, List<Tuple<string, float, float, int>>> elements = new Dictionary<string, List<Tuple<string, int, int, int>>>();
            ConstructingTheGraph(edges, elements);

        }

        public static Tuple<string, string, int,int, int>[] ReadFromExcelFile()
        {
            string inputfilePath = "D:\\Uni Related\\Algorithms\\Project\\MATERIALS\\[3] Plagiarism Validation\\Algorithm-Project\\PlagiarismValidation\\trial input.xlsx";
            int numberOfEdges;
            //var edges = new Tuple<string, string, float>[numberOfedges];
            Tuple<string, string, float, float, int> [] edges;

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
                int percentageDoc1, percentageDoc2;
                Regex percentageRegex = new Regex(@"\(\d+%\)");

                //string[] documentPathArr;
                string documentPath1 = "";
                string documentPath2 = "";
                Regex documentPathRegex = new Regex(@"~(\(\d+%\))");
                //Match matchDocumentPath;

                int linesMatched;

                numberOfEdges = table.Rows.Count - 1;
                Console.WriteLine(numberOfEdges);
                edges = new Tuple<string, string, int, int, int>[numberOfEdges];
                int percentageToBeRecorded1, percentageToBeRecorded2;

                for (int i = 1; i < table.Rows.Count; i++)
                {

                    DataRow row = table.Rows[i];

                    column1 = row[0].ToString();

                    // Retrieving the similarity percentage of each document in column 1
                    matchPercentage = percentageRegex.Match(column1); // Output --> (percentage%)
                    similarityPercentage = matchPercentage.Value;
                    similarityPercentageArr = similarityPercentage.Split('(', '%');
                    mergedSimilarityPercentage = string.Join("", similarityPercentageArr);
                    mergedSimilarityPercentage = mergedSimilarityPercentage.Replace(")", "");
                    mergedSimilarityPercentage = Regex.Replace(mergedSimilarityPercentage, @"\D", "");
                    //percentageDoc1 = Convert.ToInt32(mergedSimilarityPercentage);
                    percentageDoc1 = int.Parse(mergedSimilarityPercentage);
                    //Console.WriteLine(percentageDoc1);

                    // Retrieving the Document's Path in column 1
                    documentPath1 = column1.Replace(similarityPercentage, "");
                    //Console.WriteLine(documentPath1);

                    column2 = row[1].ToString();

                    // Retrieving the similarity percentage of each document in column 2
                    matchPercentage = percentageRegex.Match(column2);
                    similarityPercentage = matchPercentage.Value;
                    similarityPercentageArr = similarityPercentage.Split('(', '%');
                    mergedSimilarityPercentage = string.Join("", similarityPercentageArr);
                    mergedSimilarityPercentage = mergedSimilarityPercentage.Replace(")", "");
                    mergedSimilarityPercentage = Regex.Replace(mergedSimilarityPercentage, @"\D", "");
                    //percentageDoc2 = Convert.ToInt32(mergedSimilarityPercentage);
                    percentageDoc2 = int.Parse(mergedSimilarityPercentage);
                    //Console.WriteLine(percentageDoc2);

                    // Retrieving the Document's Path in column 2
                    documentPath2 = column2.Replace(similarityPercentage, "");
                    //Console.WriteLine(documentPath2);

                    // Retrieving the Lines Matched in column 3
                    column3 = row[2].ToString();
                    linesMatched = Convert.ToInt32(column3);
                    //Console.WriteLine(linesMatched);
                    percentageToBeRecorded1 = percentageDoc1;
                    percentageToBeRecorded2 = percentageDoc2;
                    /*if (percentageDoc1 >= percentageDoc2)
                    {
                        percentageToBeRecorded = percentageDoc1;
                    }
                    else
                    {
                        percentageToBeRecorded = percentageDoc2;
                    }*/
                    edges[i - 1] = new Tuple<string, string, int,int, int>(documentPath1, documentPath2, percentageToBeRecorded1,percentageToBeRecorded2, linesMatched);
                }

                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }

            }
            return edges;
        }

        public static void ConstructingTheGraph(Tuple<string, string, float, float, int>[] edges, Dictionary<string, List<Tuple<string, float, float, int>>> elements)
        {

            //the first float number is for percentage of doc 1 to doc 2 (form the first vertex to the second vertex) (edge item 3)
            //the second float number is for percentage of doc 2 to doc 1 (form the second vertex to the first vertex) (edge item 4)
            foreach (var edge in edges)
            {
                if (elements.ContainsKey(edge.Item1))
                {
                    elements[edge.Item1].Add(Tuple.Create(edge.Item2, edge.Item3, edge.Item4, edge.Item5));
                }
                else
                {
                    elements[edge.Item1] = new List<Tuple<string, float, float, int>>() { Tuple.Create(edge.Item2, edge.Item3, edge.Item4, edge.Item5) };
                }
                if (elements.ContainsKey(edge.Item2))
                {
                    elements[edge.Item2].Add(Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5));
                }
                else
                {
                    elements[edge.Item2] = new List<Tuple<string, float, float, int>>() { Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5) };
                }
            }
        }

        public static void Output()
        {

        }
    }
}
