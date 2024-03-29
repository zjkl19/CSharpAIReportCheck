﻿using System;
using System.Collections.Generic;
using System.Text;
using CSharpAIReportCheck.IRepository;
using CSharpAIReportCheck.Repository;

namespace CSharpAIReportCheck.Infrastructure
{
    public class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>();
        }
    }
}
