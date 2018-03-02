using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SystemWeaver.WordImport.Common
{
    public class WordItem
    {
        public WordItem()
        {
            DescriptionParagraphs = new List<SwParagraph>();
        }
        public SwParagraph MainParagraph { get; set; }
        public List<SwParagraph> DescriptionParagraphs { get; set; }
    }
}
