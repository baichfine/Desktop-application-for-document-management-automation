using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace programma
{
    class Shablon
    {
        string[] Bookmarks { get; set; }
        string[] NewMarks { get; set; }

        public Shablon(string[] bookmarks, string[] newMarks)
        {
            this.Bookmarks = bookmarks;
            this.NewMarks = newMarks;
        }
        
    }
}
