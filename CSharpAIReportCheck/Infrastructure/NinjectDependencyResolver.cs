using System;
using System.Collections.Generic;
using System.Text;
using CSharpAIReportCheck.IRepository;
using CSharpAIReportCheck.Repository;

namespace CSharpAIReportCheck.Infrastructure
{
    public class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        private Aspose.Words.Document _d;
        public NinjectDependencyResolver(Aspose.Words.Document d)
        {
            _d = d;
        }
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _d);
        }
    }
}
