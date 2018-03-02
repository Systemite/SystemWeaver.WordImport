using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace SystemWeaver.WordImport.ViewModel
{
    public class SwStyle
    {
        public SwStyle(Style style, int level)
        {
            Name = style.NameLocal;
            Level = level;
        }
        public SwStyle(string name, int level)
        {
            Name = name;
            Level = level;
        }
        public string Name { get; set; }
        public int Level { get; set; }
    }
}
