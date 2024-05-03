using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using ExcelDataReader;
using System.Text.RegularExpressions;
using System.ComponentModel;
using OfficeOpenXml;

namespace PlagiarismValidation
{
    class Program
    {
        static void Main(string[] args)
        {


            Tuple<string, string, int, int, int>[] edges = ReadFromExcelFile();

            Dictionary<string, List<Tuple<string, int, int , int>>> elements = new Dictionary<string, List<Tuple<string, int, int, int>>>(); // edges with two values
            Dictionary<string, int> colored_vertices = new Dictionary<string, int>();//for BFS
            Dictionary<string, List<string>> componentsLst = new Dictionary<string, List<string>>(); //groups
            Dictionary<string, float> firstVandAvg = new Dictionary<string, float>(); // statistics
           
            List<Dictionary<KeyValuePair<string,string>,Tuple<int,int>>> Components = new List<Dictionary<KeyValuePair<string, string>, Tuple<int, int>>>();
            List<Dictionary<KeyValuePair<string, string>, Tuple<int,int>>> refinedGroups = new List<Dictionary<KeyValuePair<string, string>, Tuple<int, int>>>();

            ConstructingTheGraph(edges, elements, colored_vertices);

            int numberOfEdges = 0;
            float componentAVG = 0;

            foreach (var vertex in elements) // V  
            {
                componentAVG = 0;

                if (colored_vertices[vertex.Key] == 0) 
                {
                    List<string> component = new List<string>();
                    Dictionary<KeyValuePair<string, string>, Tuple<int,int>> edges_of_components = new Dictionary<KeyValuePair<string, string>, Tuple<int, int>>();
                    BFS(vertex.Key, ref elements, ref colored_vertices, ref component, ref numberOfEdges, ref componentAVG , ref edges_of_components); // V / 2 + E
                    componentsLst.Add(vertex.Key, component);
                    Components.Add(edges_of_components);
                    firstVandAvg.Add(vertex.Key, componentAVG);
                }
            }
            foreach(Dictionary<KeyValuePair<string, string>, Tuple<int,int>> component in Components)
            {
                //Dictionary<KeyValuePair<string, string>, int> sortedcomponent = component.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value); //nlog(n)
                Dictionary<KeyValuePair<string, string>, Tuple<int, int>> refinedcompnent = new Dictionary<KeyValuePair<string, string>, Tuple<int, int>>();
                Kruskal(component, ref refinedcompnent);
                refinedGroups.Add(refinedcompnent);
            }
            foreach (var group in refinedGroups)
            {
                foreach (var keyValuePair in group)
                {
                    var key = keyValuePair.Key;
                    var value = keyValuePair.Value;
                    Console.WriteLine($"Key: {key.Key}, Value: {key.Value}, Int Value: {value}");
                }
            }

            OutPut_Of_Stat(ref firstVandAvg, ref componentsLst);

        }

        public static void BFS(string vertex, ref Dictionary<string, List<Tuple<string, int, int, int>>> graphDictionary, ref Dictionary<string, int> colored_vertices, ref List<string> component, ref int numberOfEdges, ref float componentAVG , ref Dictionary<KeyValuePair<string, string>, Tuple<int,int>> edges_of_components)
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
                        KeyValuePair<string, string> edge = new KeyValuePair<string, string>(newVertex, vertexTuple.Item1);
                        if (vertexTuple.Item2 > vertexTuple.Item3)
                        {
                            Tuple<int, int> tuple = new Tuple<int, int>(vertexTuple.Item2, vertexTuple.Item4);
                            edges_of_components[edge] = tuple;
                        }
                        else
                        {
                            Tuple<int, int> tuple = new Tuple<int, int>(vertexTuple.Item3, vertexTuple.Item4);
                            edges_of_components[edge] = tuple;
                        }
                        avgScore += vertexTuple.Item2 + vertexTuple.Item3;
                        bfsQueue.Enqueue(vertexTuple.Item1);

                    }

                    else if (colored_vertices[vertexTuple.Item1] == 1)
                    {
                        KeyValuePair<string, string> edge = new KeyValuePair<string, string>(newVertex, vertexTuple.Item1);
                        if (vertexTuple.Item2 > vertexTuple.Item3)
                        {
                            Tuple<int, int> tuple = new Tuple<int, int>(vertexTuple.Item2, vertexTuple.Item4);
                            edges_of_components[edge] = tuple;
                        }
                        else
                        {
                            Tuple<int, int> tuple = new Tuple<int, int>(vertexTuple.Item3, vertexTuple.Item4);
                            edges_of_components[edge] = tuple;
                        }
                        avgScore += vertexTuple.Item2 + vertexTuple.Item3;
                    }
                }

                colored_vertices[newVertex] = 2; // Black

            }
            componentAVG = avgScore / numberOfEdges;
            componentAVG = (float)Math.Round(componentAVG, 1);

        }
        public static Tuple<string, string, int, int, int>[] ReadFromExcelFile()
        {
            string inputfilePath = "D:\\Uni Related\\Algorithms\\Project\\MATERIALS\\[3] Plagiarism Validation\\Algorithm-Project\\PlagiarismValidation\\Test Cases\\Sample\\6-Input.xlsx";
            //string inputfilePath = "F:\\Year 3 2nd term\\Analysis and Design of Algorithm\\Project\\Algorithm-Project\\PlagiarismValidation\\Test Cases\\Sample\\5-Input.xlsx";
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
                //Console.WriteLine(numberOfEdges);
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

        public static void ConstructingTheGraph(Tuple<string, string, int, int, int>[] edges, Dictionary<string, List<Tuple<string, int, int, int>>> elements, Dictionary<string, int> colored_vertices)
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
                }
                else
                {
                    elements[edge.Item1] = new List<Tuple<string, int, int, int>>() { Tuple.Create(edge.Item2, edge.Item3, edge.Item4, edge.Item5) };
                    colored_vertices[edge.Item1] = 0;
                }
                if (elements.ContainsKey(edge.Item2))
                {
                    elements[edge.Item2].Add(Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5));
                }
                else
                {
                    elements[edge.Item2] = new List<Tuple<string, int, int, int>>() { Tuple.Create(edge.Item1, edge.Item3, edge.Item4, edge.Item5) };
                    colored_vertices[edge.Item2] = 0;
                }

            }
        }
        public static void Kruskal(Dictionary<KeyValuePair<string, string>, Tuple<int,int>> component,ref Dictionary<KeyValuePair<string, string>,Tuple<int,int>> refinedGroups)
        {
            Dictionary<string,int> enumForVertices = new Dictionary<string,int>();
            int count = 0;
            foreach (KeyValuePair<string, string> edge in component.Keys)
            {
                if(!enumForVertices.ContainsKey(edge.Key))
                {
                    enumForVertices.Add(edge.Key, count);
                    count++;
                }
                if (!enumForVertices.ContainsKey(edge.Value))
                {
                    enumForVertices[edge.Value] = count;
                    count++;
                }
            }
            SetsWithArray set_for_Kruskal = new SetsWithArray(count);
            for (int i = 0; i < count; i++)
            {
                set_for_Kruskal.Make_Set(i);
            }
            Dictionary<KeyValuePair<string, string>, Tuple<int, int>> sortedcomponent = component.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            foreach(var edge in sortedcomponent)
            {
                if (set_for_Kruskal.Find_Set(enumForVertices[edge.Key.Key]) != set_for_Kruskal.Find_Set(enumForVertices[edge.Key.Value]))
                {
                    
                    refinedGroups[edge.Key] = sortedcomponent[edge.Key];
                    set_for_Kruskal.Union_Set(enumForVertices[edge.Key.Key], enumForVertices[edge.Key.Value]);
                }
            }
        }
        public static void OutPut_Of_MST()
        {

        }
        public static void OutPut_Of_Stat(ref Dictionary<string, float> firstVandAvg, ref Dictionary<string, List<string>> componentsLst)
        {
            ExcelPackage excelPackage = new ExcelPackage();

            ExcelWorksheet statisticsSheet = excelPackage.Workbook.Worksheets.Add("Statistics 1");

            statisticsSheet.Cells["A1"].Value = "Component Index";
            statisticsSheet.Cells["B1"].Value = "Vertices";
            statisticsSheet.Cells["C1"].Value = "Average Similarity";
            statisticsSheet.Cells["D1"].Value = "Component Count";

            Dictionary<string, float> sortedFirstVandAvg = firstVandAvg.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            int i = 0;
            int counter = 1;
            // O(Components log(Components) + componentItems Log(componentItems))
            
            foreach (var vertex in sortedFirstVandAvg) // --> no of component --> worst case V/2 -- Best Case --> 1 time 
            {
                statisticsSheet.Cells[i + 1, 1].Value = counter;
                statisticsSheet.Cells[i + 1, 3].Value = vertex.Value;
                List<string> component = componentsLst[vertex.Key];

                component.Sort(); // O(vlogv)
                // +d
                List<int> componentItemsList = new List<int>();
                Regex digitsRegex = new Regex("\\d+");
                string componentItems = "";
                // matchPercentage = percentageRegex.Match(column1);

                foreach (var item in component) // O(V)
                {
                    Match digitsRegexMatch = digitsRegex.Match(item);
                    //Console.WriteLine(digitsRegexMatch.Value);
                    componentItemsList.Add(Convert.ToInt32(digitsRegexMatch.Value));
                    
                    //componentItems = componentItems + digitsRegexMatch.Value + ",";
                    //
                }
                componentItemsList.Sort(); // O(vlogv)
                foreach (var item in componentItemsList) // O(V)
                {
                    componentItems = componentItems + item.ToString() + ",";
                }
                //componentItemsList.Sort();
                componentItems.Remove(componentItems.Length - 1);
                statisticsSheet.Cells[i + 1, 2].Value = componentItems;
                statisticsSheet.Cells[i + 1, 4].Value = component.Count;
                i++;
                counter++;
            }

            string outputFilePath = @"D:\Uni Related\Algorithms\Project\MATERIALS\[3] Plagiarism Validation\Algorithm-Project\PlagiarismValidation\Output\File.xlsx";
            excelPackage.SaveAs(new System.IO.FileInfo(outputFilePath));


        }
    }
}
