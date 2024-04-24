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

            Tuple<string, string, int, int, int>[] edges = ReadFromExcelFile();
            //var edges = new Tuple<string, string, float>[numberOfedges];

            foreach (var edge in edges)
            {
                // Access individual elements of the tuple
                string element1 = edge.Item1;
                string element2 = edge.Item2;
                float element3 = edge.Item3;
                int element4 = edge.Item4;

                // Print the elements
                Console.WriteLine($"Doc 1: {element1}, Doc 2: {element2}, Percentage: {element3}, Lines Matched: {element4}");
            }

        }

        public static Tuple<string, string, int, int, int>[] ReadFromExcelFile()
        {
            string inputfilePath = "D:\\Uni Related\\Algorithms\\Project\\MATERIALS\\[3] Plagiarism Validation\\Algorithm-Project\\PlagiarismValidation\\trial input.xlsx";
            int numberOfEdges;
            //var edges = new Tuple<string, string, float>[numberOfedges];
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

                //string[] similarityPercentageArr;
                //string similarityPercentage;
                //string mergedSimilarityPercentage;
                //Match matchPercentage;
                //float percentageDoc1 , percentageDoc2;
                //Regex percentageRegex = new Regex(@"\(\d+%\)");

                //string[] documentPathArr;
                //string documentPath1 = "";
                //string documentPath2 = "";
                //Regex documentPathRegex = new Regex(@"~(\(\d+%\))");
                //Match matchDocumentPath;

                int linesMatched;

                numberOfEdges = table.Rows.Count - 1;
                Console.WriteLine(numberOfEdges);
                edges = new Tuple<string, string, int, int, int>[numberOfEdges];
                //float percentageToBeRecorded;


                //static Tuple<string, string, float, int>[] ExtractTableData(List<string> tableHtml, int n)
                //{
                //    Tuple<string, string, float, int>[] data = new Tuple<string, string, float, int>[n];
                //    int indx = 0;
                //    for (int i = 0; i < tableHtml.Count(); i += 3)
                //    {
                //        string item1 = tableHtml[i];
                //        int indexOfbrac1 = item1.LastIndexOf('(');
                //        string firstPath = item1.Substring(0, indexOfbrac1);
                //        float firstPercntage = int.Parse(item1.Substring(indexOfbrac1 + 1, item1.Length - (indexOfbrac1 + 3)));

                //        string item2 = tableHtml[i + 1];
                //        int indexOfbrac2 = item2.LastIndexOf('(');
                //        string secondPath = item2.Substring(0, indexOfbrac2);
                //        float secondPercntage = int.Parse(item2.Substring(indexOfbrac2 + 1, item2.Length - (indexOfbrac2 + 3)));

                //        int linesMatched = int.Parse(tableHtml[i + 2]);

                //        float finalPercentage = 0;
                //        if (firstPercntage >= secondPercntage)
                //        {
                //            finalPercentage = firstPercntage;
                //        }
                //        else
                //        {
                //            finalPercentage = secondPercntage;
                //        }

                //        data[indx] = Tuple.Create(firstPath, secondPath, finalPercentage, linesMatched);
                //        indx++;
                //    }
                //    return data;
                //}
                for (int i = 1; i < table.Rows.Count; i++)
                {
                        
                    DataRow row = table.Rows[i];

                    column1 = row[0].ToString();

                    // Retrieving the similarity percentage of each document in column 1
                    //matchPercentage = percentageRegex.Match(column1); // Output --> (percentage%)
                    int indexOfbrac1 = column1.LastIndexOf('(');
                    string firstPath = column1.Substring(0, indexOfbrac1);
                    int firstPercntage = int.Parse(column1.Substring(indexOfbrac1 + 1, column1.Length - (indexOfbrac1 + 3)));
                    //similarityPercentage = matchPercentage.Value;
                    //similarityPercentageArr = similarityPercentage.Split('(', '%');
                    //mergedSimilarityPercentage = string.Join("", similarityPercentageArr);
                    //mergedSimilarityPercentage = mergedSimilarityPercentage.Replace(")", "");
                    //mergedSimilarityPercentage = Regex.Replace(mergedSimilarityPercentage, @"\D", "");
                    //percentageDoc1 = Convert.ToInt32(mergedSimilarityPercentage);
                    //percentageDoc1 = float.Parse(mergedSimilarityPercentage);
                    //Console.WriteLine(percentageDoc1);

                    // Retrieving the Document's Path in column 1
                    //documentPath1 = column1.Replace(similarityPercentage, "");
                    //Console.WriteLine(documentPath1);

                    column2 = row[1].ToString();

                    // Retrieving the similarity percentage of each document in column 2
                    //matchPercentage = percentageRegex.Match(column2);
                    //similarityPercentage = matchPercentage.Value;
                    //similarityPercentageArr = similarityPercentage.Split('(', '%');
                    //mergedSimilarityPercentage = string.Join("", similarityPercentageArr);
                    //mergedSimilarityPercentage = mergedSimilarityPercentage.Replace(")", "");
                    //mergedSimilarityPercentage = Regex.Replace(mergedSimilarityPercentage, @"\D", "");
                    //percentageDoc2 = Convert.ToInt32(mergedSimilarityPercentage);
                    int indexOfbrac2 = column2.LastIndexOf('(');
                    string secondPath = column2.Substring(0, indexOfbrac2);
                    int secondPercntage = int.Parse(column2.Substring(indexOfbrac2 + 1, column2.Length - (indexOfbrac2 + 3)));
                    //percentageDoc2 = float.Parse(mergedSimilarityPercentage);
                    //Console.WriteLine(percentageDoc2);

                    // Retrieving the Document's Path in column 2
                    //documentPath2 = column2.Replace(similarityPercentage, "");
                    //Console.WriteLine(documentPath2);

                    // Retrieving the Lines Matched in column 3
                    column3 = row[2].ToString();
                    linesMatched = Convert.ToInt32(column3);
                    //Console.WriteLine(linesMatched);

                    //if(percentageDoc1 >= percentageDoc2)
                    //{
                    //    percentageToBeRecorded = percentageDoc1;
                    //}
                    //else
                    //{
                    //    percentageToBeRecorded = percentageDoc2;
                    //}
                    edges[i - 1] = new Tuple<string, string, int, int, int>(firstPath, secondPath, firstPercntage, secondPercntage, linesMatched);
                }

                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }

            }
            return edges;
        }
    }
}
