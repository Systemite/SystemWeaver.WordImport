using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.Word;
using SystemWeaver.Common;
using SystemWeaver.WordImport.ViewModel;

namespace SystemWeaver.WordImport.Common
{
    public class SwParagraph
    {
        int _retryCount = 0;
 
        public SwParagraph(Range range, IswItem currentItem)
        {
            Text = range.Text.Trim();
            Style = new SwStyle((Microsoft.Office.Interop.Word.Style)range.Paragraphs[1].get_Style(), (int)range.Paragraphs[1].OutlineLevel);
            OutlineLevel = range.Paragraphs[1].OutlineLevel;
            ClipboardCopyRtf(range, currentItem);
        }

        private void ClipboardCopyRtf(Range range, IswItem currentItem)
        {
            try
            {
                InternalClipboardCopyRtf(range);
            }
            catch (DllNotFoundException ex)
            {
                System.Windows.MessageBox.Show("Exception: " + ex.Message + "  " + range.Start.ToString());
                throw;
            }
            catch (Exception ex)
            {
                if (_retryCount < 5)
                {
                    _retryCount++;
                    System.Threading.Thread.Sleep(10);
                    ClipboardCopyRtf(range, currentItem);
                }
                else
                {
                    System.Windows.MessageBox.Show("Exception: " + ex.Message + "  " + range.Start.ToString());
                    RtfData = "";
                    if (!ex.Message.Contains("size of the rvfz cannot exceed") && !ex.Message.Contains("OpenClipboard Failed"))
                        throw;
                    //Write error message to description
                }
            }

        }

        private void InternalClipboardCopyRtf(Range range)
        {
            Clipboard.Clear();
            range.Select();
            range.Copy();
            ClipboardAsync clAsync = new ClipboardAsync();
            if(clAsync.ContainsText(TextDataFormat.Rtf))
            {
                RtfData = clAsync.GetText(TextDataFormat.Rtf);
                ParagraphData = SWUtility.RtfToRvfz(RtfData);
            }
        }
        public string Text { get; set; }
        public SwStyle Style { get; set; }
        public byte[] ParagraphData { get; set; }
        public string RtfData { get; set; }
        public WdOutlineLevel OutlineLevel { get; set; }
    }
}
