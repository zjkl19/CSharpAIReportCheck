﻿using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using CSharpAIReportCheck.IRepository;
using CSharpAIReportCheck.Models;
//using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace CSharpAIReportCheck.Repository
{
    class AsposeAIReportCheck : IAIReportCheck
    {
        public List<ReportError> reportError = new List<ReportError>();
        public List<ReportWarnning> reportWarnning = new List<ReportWarnning>();
        readonly Document _doc;
        string _originalWholeText;

        public AsposeAIReportCheck(Document doc)
        {
            _doc = doc;
            _originalWholeText = doc.Range.Text;
        }

        public void _FindUnitError()
        {
            MatchCollection matches;
            var regex = new Regex(@"([0-9]Km/h)");
            try
            {
                matches = regex.Matches(_originalWholeText);

                if (matches.Count != 0)
                {
                    foreach (Match m in matches)
                    {
                        reportError.Add(new ReportError(ErrorNumber.CMA, "正文"+m.Index.ToString(), "应为km/h"));
                    }

                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        ReplacingCallback = new ReplaceEvaluatorFindAndHighlight(),
                        Direction = FindReplaceDirection.Forward
                    };
                    _doc.Range.Replace(regex, "", options);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void _FindSpecificationsError()
        {
            //规范
            var Specifications = new string[]
            {
                "《公路桥梁荷载试验规程》（JTG/T J21-01-2015）",
                "《混凝土结构现场检测技术标准》（GB/T 50784-2013）",
                "《城市桥梁设计规范》（CJJ 11-2011）",
            };
            double similarity;    //相似度
            //获取word文档中的第一个表格
            var table0 = _doc.GetChildNodes(NodeType.Table, true)[1] as Table;

            //var table0 = ai.GetOverViewTable();
            var cell = table0.Rows[4].Cells[1];
            string[] splitArray = cell.GetText().Split('\r');    //用GetText()的方法来获取cell中的值
            foreach (var s in splitArray)
            {
                s.Replace("\a", ""); s.Replace("\r", "");
                s.Replace("(", "（"); s.Replace(")", "）");
                var s1=Regex.Replace(s, @"(.+)《", "《");
                foreach (var sp in Specifications)
                {
                    similarity = Levenshtein(@s1, @sp);
                    if (similarity>0.85&& similarity<1)
                    {
                        reportError.Add(new ReportError(ErrorNumber.Description,"汇总表格中主要检测检验依据", "应为"+sp));
                        break;
                    }
                }
            };
        }

        public void _FindNotExplainComponentNo()
        {
            MatchCollection matches;
            var regex = new Regex(@"(构件编号说明)");
            try
            {
                matches = regex.Matches(_originalWholeText);
                if (matches.Count == 0)
                {
                    reportWarnning.Add(new ReportWarnning(WarnningNumber.NotClearInfo, "/", "构件编号未进行说明"));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void _GenerateResultReport()
        {
            var doc = new Document();

            // DocumentBuilder provides members to easily add content to a document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            if (reportError.Count != 0)
            {
                builder.Writeln("错误列表");
                int i = 1;    //计数
                builder.StartTable();
                builder.InsertCell(); builder.Write("序号");
                builder.InsertCell(); builder.Write("编号");
                builder.InsertCell(); builder.Write("名称");
                builder.InsertCell(); builder.Write("位置");
                builder.InsertCell(); builder.Write("说明");
                builder.EndRow();

                foreach (var e in reportError)
                {
                    builder.InsertCell(); builder.Write(i.ToString());
                    builder.InsertCell(); builder.Write(((int)e.No).ToString());
                    builder.InsertCell(); builder.Write(e.Name);
                    builder.InsertCell(); builder.Write(e.Position);
                    builder.InsertCell(); builder.Write(e.Description);
                    builder.EndRow();
                }
                builder.EndTable();
            }

            if (reportWarnning.Count != 0)
            {
                builder.Writeln("警告列表");
                int i = 1;    //计数
                builder.StartTable();
                builder.InsertCell(); builder.Write("序号");
                builder.InsertCell(); builder.Write("编号");
                builder.InsertCell(); builder.Write("名称");
                builder.InsertCell(); builder.Write("位置");
                builder.InsertCell(); builder.Write("说明");
                builder.EndRow();

                foreach (var w in reportWarnning)
                {
                    builder.InsertCell(); builder.Write(i.ToString());
                    builder.InsertCell(); builder.Write(((int)w.No).ToString());
                    builder.InsertCell(); builder.Write(w.Name);
                    builder.InsertCell(); builder.Write(w.Position);
                    builder.InsertCell(); builder.Write(w.Description);
                    builder.EndRow();
                }
                builder.EndTable();
            }

            // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
            // Aspose.Words supports saving any document in many more formats.
            doc.Save("校核结果.docx");
        }

        public void CheckReport()
        {
            _FindUnitError();
            _FindNotExplainComponentNo();
            _FindSpecificationsError();
            _GenerateResultReport();
            _doc.Save("标出错误或警告的报告.doc");
        }

        /// <summary>
        /// 字符串相似度计算
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        public static double Levenshtein(string str1, string str2)
        {
            //计算两个字符串的长度。  
            int len1 = str1.Length;
            int len2 = str2.Length;
            //建立上面说的数组，比字符长度大一个空间  
            int[,] dif = new int[len1 + 1, len2 + 1];
            //赋初值，步骤B。  
            for (int a = 0; a <= len1; a++)
            {
                dif[a, 0] = a;
            }
            for (int a = 0; a <= len2; a++)
            {
                dif[0, a] = a;
            }
            //计算两个字符是否一样，计算左上的值  
            int temp;
            for (int i = 1; i <= len1; i++)
            {
                for (int j = 1; j <= len2; j++)
                {
                    if (str1[i - 1] == str2[j - 1])
                    {
                        temp = 0;
                    }
                    else
                    {
                        temp = 1;
                    }
                    //取三个值中最小的  
                    dif[i, j] = Math.Min(Math.Min(dif[i - 1, j - 1] + temp, dif[i, j - 1] + 1), dif[i - 1, j] + 1);
                }
            }
            //Console.WriteLine("字符串\"" + str1 + "\"与\"" + str2 + "\"的比较");
            //取数组右下角的值，同样不同位置代表不同字符串的比较  
            //Console.WriteLine("差异步骤：" + dif[len1, len2]);
            //计算相似度  
            double similarity = 1 - (double)dif[len1, len2] / Math.Max(str1.Length, str2.Length);
            //Console.WriteLine("相似度：" + similarity + " 越接近1越相似");
            return similarity;
        }


        public static void Run()
        {
            // ExStart:FindAndHighlight
            // The path to the documents directory.

            string fileName = "TestFile.doc";

            Document doc = new Document(fileName);

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceEvaluatorFindAndHighlight();
            options.Direction = FindReplaceDirection.Forward;

            // We want the "your document" phrase to be highlighted.
            Regex regex = new Regex("your document", RegexOptions.IgnoreCase);
            doc.Range.Replace(regex, "", options);

            // Save the output document.
            doc.Save("TestFile.doc");
            // ExEnd:FindAndHighlight
        }
        // ExStart:ReplaceEvaluatorFindAndHighlight
        private class ReplaceEvaluatorFindAndHighlight : IReplacingCallback
        {
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting.
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    // Select the next Run node. 
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                // Now highlight all runs in the sequence.
                foreach (Run run in runs)
                    run.Font.HighlightColor = System.Drawing.Color.Red;

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }
        }
        // ExEnd:ReplaceEvaluatorFindAndHighlight
        // ExStart:SplitRun
        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
        // ExEnd:SplitRun
    }
}
