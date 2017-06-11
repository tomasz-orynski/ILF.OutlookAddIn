using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    interface IComponent
    {

    }

    interface ISelfRegisteredComponent : IComponent
    {
        void Initialize(Application app);
    }
}
