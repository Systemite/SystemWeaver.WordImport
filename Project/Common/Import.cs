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
    public class Import
    {
        private const string DocumentItemSid = "SSRP";
        private const string ParagraphItemSid = "SSSE";
        private const string TitelPartSid = "IRRS";
        private const string SubSectionSid = "ISSE";
        private string _logTxt = "";

        public Import(IswItem topItem, List<SwParagraph> paragraphs, List<SwStyle> styles, List<SwStyle> description, string hostName, ThreadedWindowWrapper wrap)
        {
            LogToDescription("Importing from Office started at " + Convert.ToString(DateTime.Now));
            var currentItem = topItem;
            List<WordItem> wordItems = CreateWordItems(paragraphs, description);

            double progress = 0;
            double step = (double)100 / (double)wordItems.Count;
            wrap.SetProgress(Convert.ToInt32(progress).ToString()); progress += step;

            int handledItems = 0;
            WdOutlineLevel currentOutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            WdOutlineLevel lastOutlineLevel = 0;
            Stack<IswItem> parentItems = new Stack<IswItem>();
            parentItems.Push(topItem);
            foreach(WordItem wordItem in wordItems)
            {
                if (wordItem.MainParagraph == null)
                {
                    topItem.Description = SWDescription.MakeDescription(SWUtility.RtfToRvfz(ConcatenateParagraphsRtf(wordItem.DescriptionParagraphs))); 
                    continue;
                }
                currentOutlineLevel = wordItem.MainParagraph.OutlineLevel;

                if (currentOutlineLevel > lastOutlineLevel) // Sublevel -> Add to the stack
                {
                    currentItem = WriteItem(parentItems.Peek(), wordItem);
                    parentItems.Push(currentItem);
                }

                else // Higher level -> Remove items from the stack 
                {
                    for (WdOutlineLevel i = currentOutlineLevel; (i <= lastOutlineLevel && parentItems.Count > 1); i++)
                        parentItems.Pop();
                    currentItem = WriteItem(parentItems.Peek(), wordItem);
                    parentItems.Push(currentItem);
                }

                lastOutlineLevel = currentOutlineLevel;

                wrap.SetProgress(Convert.ToInt32(progress).ToString()); progress += step;
                wrap.SetStatus(string.Format("Imported {0} paragraphs of {1}", ++handledItems, wordItems.Count));
            }
            LogToDescription("Script terminated gracefully at " + Convert.ToString(DateTime.Now));
            WriteLogToTopItem(topItem);
        }

        private List<WordItem> CreateWordItems(List<SwParagraph> paragraphs, List<SwStyle> description)
        {
            List<WordItem> wordItems = new List<WordItem>();
            WordItem currentWordItem = null;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                SwParagraph tmpParagraph = paragraphs[i];
                if ((from d in description where d.Name.Equals(tmpParagraph.Style.Name) select d).Count() > 0)
                {
                    if (currentWordItem == null) { 
                        currentWordItem = new WordItem();
                        wordItems.Add(currentWordItem);
                    }
                    currentWordItem.DescriptionParagraphs.Add(tmpParagraph);
                }
                else
                {
                    currentWordItem = new WordItem();
                    currentWordItem.MainParagraph = tmpParagraph;
                    wordItems.Add(currentWordItem);
                }
            }
            return wordItems;
        }

        private IswItem WriteItem(IswItem parentItem, WordItem wordItem)
        {
            IswItem newItem = parentItem.HomeLibrary.CreateItem(ParagraphItemSid, wordItem.MainParagraph.Text);
            WriteDescription(parentItem, wordItem, newItem);
            return newItem;
        }

        private void WriteDescription(IswItem parentItem, WordItem wordItem, IswItem newItem)
        {
            try
            {
                try
                {
                    newItem.Description = SWDescription.MakeDescription(SWUtility.RtfToRvfz(ConcatenateParagraphsRtf(wordItem.DescriptionParagraphs)));
                }
                catch (Exception ex)
                {
                    LogToDescription(String.Format("Couldn't write description to {0} ({1}) as {2}", newItem.Name, newItem.HandleStr, newItem.swItemType.SID));
                    MessageBox.Show(ex.Message);
                }
                if (parentItem.IsSID(DocumentItemSid))
                    parentItem.AddPart(TitelPartSid, newItem);
                else
                    parentItem.AddPart(SubSectionSid, newItem);

                //LogToDescription(String.Format("Created {0} ({1}) as {2}", newItem.Name, newItem.HandleStr, newItem.swItemType.SID));
                //LogToDescription(String.Format("Added {0} ({1}) to {2} at {3} ", newItem.Name, newItem.HandleStr, parentItem.Name, Convert.ToString(DateTime.Now)));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                LogToDescription(String.Format("Exception on part to item {0}: {1}", parentItem.Name, ex.Message));
             
            }
        }

        private IswItem CreateWordItem(IswItem currentItem, SwParagraph paragraph, string createSID,/* ref string oldVersionID, */string descriptionRtf)
        {
            IswItem newItem;

            newItem = currentItem.HomeLibrary.CreateItem(createSID, paragraph.Text);
            try
            {
                newItem.Description = SWDescription.MakeDescription(SWUtility.RtfToRvfz(descriptionRtf));
            }
            catch (Exception ex)
            {
                LogToDescription(String.Format("Couldn't write description to {0} ({1}) as {2}", newItem.Name, newItem.HandleStr, newItem.swItemType.SID));
                MessageBox.Show(ex.Message);
            }
            return newItem;
        }

        public static string GetMachineName(string hostName)
        {
            //"url:swap://sys5:1201/"	
            string hostPattern = "url:swap://";
            int firstPos = hostName.IndexOf(hostPattern);
            if (firstPos >= 0)
            {
                int secondPos = hostName.Substring(firstPos + hostPattern.Length).IndexOf(":");
                if (secondPos < 0)
                    secondPos = hostName.Substring(firstPos + hostPattern.Length).IndexOf("/");
                if (secondPos > 0)
                    return hostName.Substring(firstPos + hostPattern.Length).Substring(0, secondPos);
            }
            return "";
        }

        private static string ConcatenateParagraphsRtf(List<SwParagraph> descriptionParagraphs)
        {
            if (descriptionParagraphs.Count == 0)
                return "";
            string finalRtf = descriptionParagraphs.First().RtfData;
            foreach (var item in descriptionParagraphs.Skip(1))
            {
                finalRtf = ConcatenareRtf(finalRtf, item.RtfData);
            }
            return finalRtf;
        }

        private static string ConcatenareRtf(string first, string second)
        {
            if (first == null)
                return second;
            if (second == null)
                return first;
  
            int destFirst = first.LastIndexOf("}");
            if (destFirst < 0 && first.Length > 0)
                return first;
            int destSecond = second.IndexOf("{");
            if (destSecond < 0)
                return first;
            if (first.Length == 0)
                return second;
            return first.Substring(0, destFirst) 
                + second.Substring(destSecond + 1);
        }

        private void LogToDescription(string log)
        {
            _logTxt = _logTxt + log + Environment.NewLine;
        }

        private void WriteLogToTopItem(IswItem topItem)
        {
            topItem.Description = SWDescription.MakeDescription(SWUtility.PlainTextToRvfz(_logTxt));
        }
    }
}
