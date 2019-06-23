using System;
using CSharpAIReportCheck.IRepository;
using CSharpAIReportCheck.Repository;
using Ninject;
using Aspose.Words;
//using Microsoft.Office.Interop.Word;
namespace CSharpAIReportCheck
{
    class Program
    {

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver());

            var a = kernel.Get<IAIReportCheck>();

            string fileName = "glz.doc";

            var doc = new Document(fileName);
            var ai = new AsposeAIReportCheck(doc);

            //获取word文档中的第一个表格
            //var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Table;

            //var table0 = ai.GetOverViewTable();
            //var cell = table0.Rows[4].Cells[1];
            //string cbfbm = cell.GetText();

            //string[] splitArray = cbfbm.Split('\r');
            ////用GetText()的方法来获取cell中的值


            //foreach (var s in splitArray)
            //{
            //    s.Replace("\a", "");
            //    s.Replace("\r", "");
            //    Console.WriteLine(s);
            //    Levenshtein(@"《城市桥梁设计规范》（CJJ 11-2011）", @s);
            //};

            ai.CheckReport();

            //if(ai.reportError!=null)
            //{
            //    foreach(var e in ai.reportError)
            //    {
            //        Console.WriteLine(e.No);
            //        Console.WriteLine(e.Name);
            //    }

            //}
            
        }


    }
}
