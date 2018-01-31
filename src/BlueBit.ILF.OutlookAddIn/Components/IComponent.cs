using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public interface IComponent
    {
    }

    public interface IInitializedComponent : IComponent
    {
        void Initialize();
    }

    public interface IInitializedAppComponent : IComponent
    {
        void Initialize(Outlook.Application app);
    }


    public enum CommandID : byte
    {
        SetDefaultCalendars = 1,
    }

    public interface ICommandComponent : IComponent
    {
        CommandID ID { get; }
        void Execute();
    }
}
