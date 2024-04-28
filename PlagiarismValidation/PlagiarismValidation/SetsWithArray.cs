using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace PlagiarismValidation
{
    internal class SetsWithArray
    {
        private int[] members;
        public SetsWithArray(int count)
        {
            this.members = new int[count];
        }

        public void Make_Set(int vertex_index)
        {
            members[vertex_index] = vertex_index;
        }

        public int Find_Set(int vertex_index)
        {
            return members[vertex_index];
        }

        public void Union_Set(int u, int v)
        {
            int cluster = this.members[u];
            for (int i = 0; i < this.members.Length; i++)
            {
                if (this.members[i] == cluster)
                {
                    this.members[i] = this.members[v];
                }
            }
        }
    }
}
