using Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public interface IComponent
    {
    }

    public interface ISelfRegisteredComponent : IComponent
    {
        void Initialize(Application app);
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
