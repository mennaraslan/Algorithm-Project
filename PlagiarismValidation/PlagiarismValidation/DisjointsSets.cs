using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlagiarismValidation
{
    internal class DisjointsSets
    {
        private int[] parent;
        private int[] rank;

        public DisjointsSets(int size)
        {
            parent = new int[size];
            rank = new int[size];
            for (int i = 0; i < size; i++)
            {
                parent[i] = i;
                rank[i] = 0;
            }
        }

        public int Find(int i)
        {
            if (parent[i] != i)
            {
                parent[i] = Find(parent[i]);
            }
            return parent[i];
        }

        public void Union(int x, int y)
        {
            int xRoot = Find(x);
            int yRoot = Find(y);

            if (xRoot == yRoot) return;

            if (rank[xRoot] < rank[yRoot])
            {
                parent[xRoot] = yRoot;
            }
            else if (rank[xRoot] > rank[yRoot])
            {
                parent[yRoot] = xRoot;
            }
            else
            {
                parent[yRoot] = xRoot;
                rank[xRoot]++;
            }
        }
    }
}
