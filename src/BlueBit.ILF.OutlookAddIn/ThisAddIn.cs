using Autofac;
using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace BlueBit.ILF.OutlookAddIn
{
    public partial class ThisAddIn
    {
        private static readonly Logger _logger;
        private static readonly IContainer _container;

        static ThisAddIn()
        {
            _logger = LogManager.GetCurrentClassLogger();
            _container = CreateContainer();
        }

        private static IContainer CreateContainer()
            => _logger.OnEntryCall(() =>
            {
                var builder = new ContainerBuilder();
                var assembly = typeof(ThisAddIn).Assembly;
                builder
                    .RegisterAssemblyTypes(assembly)
                    .AssignableTo<IComponent>()
                    .AsImplementedInterfaces()
                    .SingleInstance();
                builder
                    .RegisterAssemblyTypes(assembly)
                    .AssignableTo<IRibbonExtensibility>()
                    .AsImplementedInterfaces()
                    .SingleInstance();

                return builder.Build();
            });

        private void InternalStartup()
            => _logger.OnEntryCall(() =>
                _container
                    .Resolve<IEnumerable<ISelfRegisteredComponent>>()
                    .ForEach(_ => _.Initialize(this.Application))
            );

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
            => _logger.OnEntryCall(_container.Resolve<IRibbonExtensibility>);
    }
}
