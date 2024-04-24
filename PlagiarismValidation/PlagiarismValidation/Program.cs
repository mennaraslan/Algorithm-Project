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

            Dictionary < string, List < Tuple < string, int , int, int>>> elements = new Dictionary<string, List<Tuple<string, int, int, int>>>();
            Dictionary<string, List<Tuple<string, int , int>>> graph = new Dictionary<string, List<Tuple<string, int, int>>>();
            Dictionary<string, int> colored_vertices = new Dictionary<string, int>();
            Dictionary<string, List<string>> componentsLst = new Dictionary<string, List<string>>();
            ConstructingTheGraph(edges, elements,graph);
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
            int numberOfEdges = 0;
            float maxScore = 0;

            foreach (var vertex in elements) // V  
            {
                maxScore = 0;

                if (colored_vertices[vertex.Key] == 0) 
                {
                    List<string> component = new List<string>();

                    BFS(vertex.Key, ref elements, ref colored_vertices, ref component, ref numberOfEdges, ref maxScore); // V / 2 + E
                    componentsLst.Add(vertex.Key, component);
                }
            }

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

        public static void BFS(string vertex, ref Dictionary<string, List<Tuple<string, int, int, int>>> graphDictionary, ref Dictionary<string, int> colored_vertices, ref List<string> component, ref int numberOfEdges, ref float maxScore)
        {
            colored_vertices[vertex] = 1;
            Queue<string> bfsQueue = new Queue<string>();

            float avgScore = 0;
            bfsQueue.Enqueue(vertex);
            component.Add(vertex);

            numberOfEdges = 0;
            while (bfsQueue.Count != 0)
            {
                string newVertex = bfsQueue.Dequeue();
                List<Tuple<string, int, int, int>> adjacencyList = graphDictionary[newVertex];
                foreach (var vertexTuple in adjacencyList)
                {
                    numberOfEdges++;
                    if (colored_vertices[vertexTuple.Item1] == 0)  // White
                    {

                        colored_vertices[vertexTuple.Item1] = 1; // Gray
                        component.Add(vertexTuple.Item1);
                        avgScore += vertexTuple.Item2;
                        bfsQueue.Enqueue(vertexTuple.Item1);

                    }

                    else if (colored_vertices[vertexTuple.Item1] == 1)
                    {

                        avgScore += vertexTuple.Item2;
                    }
                }

                colored_vertices[newVertex] = 2; // Black

            }
            maxScore = avgScore / numberOfEdges;

        }
        public static void ConstructingTheGraph(Tuple<string, string, int, int, int>[] edges, Dictionary<string, List<Tuple<string, int, int, int>>> elements, Dictionary<string, List<Tuple<string, int, int>>> graph)
        {
            int maximum = 0;
            //the first float number is for percentage of doc 1 to doc 2 (form the first vertex to the second vertex) (edge item 3)
            //the second float number is for percentage of doc 2 to doc 1 (form the second vertex to the first vertex) (edge item 4)
            foreach (var edge in edges)
            {
                maximum = Math.Max(edge.Item3, edge.Item4);
                if (elements.ContainsKey(edge.Item1))
                {
                    elements[edge.Item1].Add(Tuple.Create(edge.Item2, edge.Item3, edge.Item4, edge.Item5));
                    graph[edge.Item1].Add(Tuple.Create(edge.Item2, maximum, edge.Item5));
                }
                else
                {
                    elements[edge.Item1] = new List<Tuple<string, int, int, int>>() { Tuple.Create(edge.Item2, edge.Item3, edge.Item4, edge.Item5) };
                    graph[edge.Item1] = new List<Tuple<string, int, int>>() { Tuple.Create(edge.Item2, maximum, edge.Item5) };
                }
                if (elements.ContainsKey(edge.Item2))
                {
                    elements[edge.Item2].Add(Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5));
                    graph[edge.Item2].Add(Tuple.Create(edge.Item1, maximum, edge.Item5));
                }
                else
                {
                    elements[edge.Item2] = new List<Tuple<string, int, int, int>>() { Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5) };
                    graph[edge.Item2] = new List<Tuple<string, int, int>>() { Tuple.Create(edge.Item1, maximum, edge.Item5) };
                }
                
            }
        }

        public static void Output()
        {

        }
    }
}

