using Autofac;
using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using Microsoft.Office.Core;
using MoreLinq;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;

namespace BlueBit.ILF.OutlookAddIn

{
    public partial class ThisAddIn

    {
        private static readonly Logger _logger;

        private static readonly IContainer _container;

        static ThisAddIn()

        {
            _logger = LogManager.GetCurrentClassLogger();
            _logger.Info($"### BUILD DATE: {GetLinkerTime()} ###");
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

        private static DateTime GetLinkerTime(TimeZoneInfo target = null)
        {
            var filePath = typeof(ThisAddIn).Assembly.Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            var buffer = new byte[2048];
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                stream.Read(buffer, 0, 2048);
            var offset = BitConverter.ToInt32(buffer, c_PeHeaderOffset);
            var secondsSince1970 = BitConverter.ToInt32(buffer, offset + c_LinkerTimestampOffset);
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var linkTimeUtc = epoch.AddSeconds(secondsSince1970);
            var tz = target ?? TimeZoneInfo.Local;
            var localTime = TimeZoneInfo.ConvertTimeFromUtc(linkTimeUtc, tz);
            return localTime;
        }
    }
}
