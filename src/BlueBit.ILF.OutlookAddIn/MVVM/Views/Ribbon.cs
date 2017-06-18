using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Views
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IReadOnlyDictionary<CommandID, ICommandComponent> _cmds;
        private Office.IRibbonUI _ribbon;

        public Ribbon(IEnumerable<ICommandComponent> cmds)
        {
            _cmds = cmds.ToDictionary(_ => _.ID);
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        public string GetCustomUI(string ribbonID) 
            => _logger.OnEntryCall(GetResourceText);

        public void OnSetDefaultCalendars(Office.IRibbonControl control)
            => _logger.OnEntryCall(() =>
            {
                _cmds[CommandID.SetDefaultCalendars].Execute();

            });


        private static string GetResourceText() => GetResourceText(typeof(Ribbon).FullName + ".xml");
        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                        if (resourceReader != null)
                            return resourceReader.ReadToEnd();

            return null;
        }
    }
}
