using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace CriticMarkup
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Export_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            StringBuilder sb = new StringBuilder();
            int currIdx = doc.Content.Start;

            //first process revisions
            foreach (Word.Revision rev in doc.Revisions)
            {
                //Get text up to this point
                sb.Append(doc.Range(currIdx, rev.Range.Start).Text);

                //insert the markup
                switch (rev.Type)
                {
                    case Word.WdRevisionType.wdRevisionInsert:
                        sb.Append("{++");
                        sb.Append(rev.Range.Text);
                        sb.Append("++}");
                        break;
                    case Word.WdRevisionType.wdRevisionDelete:
                        sb.Append("{--");
                        sb.Append(rev.Range.Text);
                        sb.Append("--}");
                        break;
                    case Word.WdRevisionType.wdRevisionReplace:
                        MessageBox.Show("Found the rare wdRevisionReplace type! I don't know how to convert this right now, so I have to skip the change. Please send this document to the plugin developers so they can handle this properly!");
                        break;
                        sb.Append("{~~");
                        MessageBox.Show("Found replacement. Here's the text it's giving me: " + rev.Range.Text);
                        sb.Append("~~}");
                        break;
                }

                currIdx = rev.Range.End;
            }
            sb.Append(doc.Range(currIdx, doc.Content.End).Text);
            String text = sb.ToString();

            //now scan for substitutions (BROKEN)
            //String re_sub_id = @"\{\+\+(.*?)\+\+\}\{\-\-(.*?)\-\-\}";
            //text = Regex.Replace(text, re_sub_id, @"{~~$2~>$1~~}", RegexOptions.Singleline);
            //String re_sub_di = @"\{\-\-(.*?)\-\-\}\{\+\+(.*?)\+\+\}";
            //text = Regex.Replace(text, re_sub_di, @"{~~$1~>$2~~}", RegexOptions.Singleline);

            Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add();
            Word.Paragraph pgraph;

            //Intro text
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.Text = text;

            //now process comments, wrapping things fully somehow
        }

        private void btn_Import_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document origdoc = Globals.ThisAddIn.Application.ActiveDocument;
            String intext = origdoc.Range(origdoc.Content.Start, origdoc.Content.End).Text;
            Word.Document doc = Globals.ThisAddIn.Application.Documents.Add();
            doc.Range(doc.Content.Start, doc.Content.End).Text = intext;

            doc.ActiveWindow.View.ShowInsertionsAndDeletions = false;
            doc.TrackRevisions = false;

            //expand substitutions
            String pattern = @"\{\~\~(.*?)\~\>(.*?)\~\~\}";
            String replacement = @"{--$1--}{++$2++}";
            doc.Range(doc.Content.Start, doc.Content.End).Text = Regex.Replace(doc.Range(doc.Content.Start, doc.Content.End).Text, pattern, replacement);

            //insertions
            Regex insertion = new Regex(@"\{\+\+(.*?)\+\+\}");
            MatchCollection matches = insertion.Matches(doc.Range(doc.Content.Start, doc.Content.End).Text);
            Match[] insertions = matches.Cast<Match>().Reverse().ToArray();
            foreach (Match m in insertions)
            {
                String newtext = m.Groups[1].Value;
                doc.Range(m.Index, m.Index + m.Length).Delete();
                doc.TrackRevisions = true;
                doc.Range(m.Index, m.Index).Text = newtext;
                doc.TrackRevisions = false;
            }

            //deletions
            Regex deletion = new Regex(@"\{\-\-(.*?)\-\-\}");
            matches = deletion.Matches(doc.Range(doc.Content.Start, doc.Content.End).Text);
            Match[] deletions = matches.Cast<Match>().Reverse().ToArray();
            foreach (Match m in deletions)
            {
                String deltext = m.Groups[1].Value;
                doc.Range(m.Index + m.Length - 3, m.Index + m.Length).Delete();
                doc.Range(m.Index, m.Index + 3).Delete();
                doc.TrackRevisions = true;
                doc.Range(m.Index, m.Index + deltext.Length).Delete();
                doc.TrackRevisions = false;
            }
            doc.ActiveWindow.View.ShowInsertionsAndDeletions = true;
        }
    }
}
