using Autofac;
using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn
{
    public partial class ThisAddIn
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private IContainer _container;

        private void InternalStartup()
            => _logger.OnEntryCall(() =>
            {
                var builder = new ContainerBuilder();
                builder
                    .RegisterAssemblyTypes(typeof(IComponent).Assembly)
                    .AssignableTo<IComponent>()
                    .AsImplementedInterfaces()
                    .SingleInstance();

                _container = builder.Build();
                _container
                    .Resolve<IEnumerable<ISelfRegisteredComponent>>()
                    .ForEach(_ => _.Initialize(this.Application));
            });
    }
}
