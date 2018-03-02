using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using SystemWeaver.Common;
using SystemWeaver.WordImport.ViewModel;

namespace SystemWeaver.WordImport.Common
{
    public class ReadWord
    {
        public ReadWord(string sFileName, IswItem currentItem, ThreadedWindowWrapper wrap, out bool failedWithLoadDLL)
        {
            failedWithLoadDLL = false;
            Application word = new Application();
            Document doc = null;
            int progress = 0;

            object fileName = sFileName;
            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            object readOnly = true;
            try
            {
                doc = word.Documents.Open(ref fileName,
                    ref missing, ref readOnly, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

                String read = string.Empty;
                List<SwParagraph> data = new List<SwParagraph>();
                data = GetParagraphs(doc, currentItem, wrap);
                Paragraphs = data;

                List<SwStyle> styles = new List<SwStyle>();
                foreach (var p in Paragraphs)
                {
                    var name = p.Style.Name;
                    foreach (var s in styles)
                    {
                        if (s.Name.Equals(p.Style.Name))
                            name = "";
                    }
                    if (name.Length > 0)
                        styles.Add(new SwStyle(name, (int)p.OutlineLevel));
                }
                StylesInUse = styles;
                //List<SwStyle> styles = new List<SwStyle>();
                //foreach (var p in Paragraphs)
                //{
                //    var name = p.Style.Name;
                //    foreach (var s in styles)
                //    {
                //        if (s.Name.Equals(p.Style.Name))
                //            name = "";
                //    }
                //    if (name.Length > 0)
                //        styles.Add(new SwStyle(name));
                //}
                //StylesInUse = styles;
            }
            catch (DllNotFoundException)
            {
                failedWithLoadDLL = true;
                return;
            }
            catch (Exception ex)
            {
                if (doc != null)
                    ((_Document)doc).Close();
                ((_Application)word).Quit();
                System.Diagnostics.Debug.WriteLine(ex.Message);
                throw;
            }
            ((_Document)doc).Close();
            ((_Application)word).Quit();
            wrap.SetProgress(progress.ToString()); progress += 20;
            wrap.SetProgress(progress.ToString()); progress += 20;
        }

        /// <summary>
        /// Reads rtf-file into a RichTextBox. Worked with small documents.
        /// In larger documents, Paragraphs and Blocks did not match. Couldn't match the rtf-data with the position in the document.
        /// </summary>
        /// <param name="rtfFileName"></param>
        /// <returns></returns>
        private static IEnumerable<string> RtfFileToParagraphs(string rtfFileName)
        {
            var result = new List<string>();
            var tb = new System.Windows.Controls.RichTextBox();

            using (var fs = new FileStream(rtfFileName, FileMode.Open))
            {
                tb.Selection.Load(fs, System.Windows.DataFormats.Rtf);
            }
            int inl = 0;
            foreach (var source in tb.Document.Blocks)
            {
                inl++;
                using (var stream = new MemoryStream())
                {
                    var sourceRange = new System.Windows.Documents.TextRange(source.ContentStart, source.ContentEnd);
                    sourceRange.Save(stream, System.Windows.DataFormats.Rtf);
                    var rtf = Encoding.Default.GetString(stream.ToArray());

                    if (!string.IsNullOrEmpty(rtf))
                    {
                        //Add new line at end of paragraph.
                        rtf = rtf.Substring(0, rtf.LastIndexOf('}')) + @"\line" + @"}";
                        result.Add(rtf);
                    }
                    else
                        result.Add(@"{\line}");
                }
            }

            return result;
        }
        /// <summary>
        /// Use paragraph sign for finding paragraphs. Instead of doc.Paragraphs. Much faster in large documents.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private List<SwParagraph> GetParagraphs(Document doc, IswItem currentItem, ThreadedWindowWrapper wrap)
        {
            List<SwParagraph> retParagraphs = new List<SwParagraph>();
            Range range = doc.Content;
            int totalLength = range.End;
            Range temprange;
            Find find = range.Find;
            find.Text = "^p";
            find.ClearFormatting();
            object missing = Type.Missing;
            try
            {
                find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);
                int start = 0;
                while (range.Find.Found)
                {
                    if (range.Start < start)
                        break;
                    temprange = doc.Range(start, range.End);
                    retParagraphs.Add(new SwParagraph(temprange, currentItem));
                    wrap.SetProgress((range.End * 100 / totalLength).ToString());
                    start = range.End;// -1;
                    range.Find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                }
            }
            catch (DllNotFoundException)
            {
                throw;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Exception: " + ex.Message + "  " + range.Start.ToString());
                throw;
            }
            return retParagraphs;
        }

        public List<SwParagraph> Paragraphs { get; set; }
        public List<SwStyle> StylesInUse { get; set; }
        public List<SwStyle> BodyStylesInUse { get; set; }
    }
}
