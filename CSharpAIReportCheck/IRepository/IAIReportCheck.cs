//using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpAIReportCheck.IRepository
{
    interface IAIReportCheck
    {
        void _FindUnitError();
        void _FindSpecificationsError();
        void _FindNotExplainComponentNo();

        //Table GetOverViewTable();
    }
}
