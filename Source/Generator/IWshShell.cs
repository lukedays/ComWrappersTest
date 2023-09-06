using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ComWrappersTest.Source.Generator;

[GeneratedComInterface]
[Guid("F935DC21-1CF0-11D0-ADB9-00C04FD58A0B")] // The IID for IWshShell (COMView.exe -> TypeLib -> IWshShell interface)
public partial interface IWshShell
{
    // These entries need to be in the exact VTable order
    // Count starts after the IUnknown methods (QueryInterface, AddRef, Release) -> (0, 1, 2)
    // The order is the same as displayed in COMView.exe
    void GetTypeInfoCount(); // Slot 3
    void GetTypeInfo(); // Slot 4
    void GetIDsOfNames(); // Slot 5
    void Invoke(); // Slot 6
    void SpecialFolders(); // Slot 7

    void get_Environment(); // Slot 8

    int Run( // Slot 9
        [MarshalAs(UnmanagedType.BStr)] string Command,
        [MarshalUsing(typeof(IntToVariant)), Optional] int WindowStyle,
        [MarshalUsing(typeof(BoolToVariant)), Optional] bool WaitOnReturn
    );

    void Popup(); // Slot 10

    void CreateShortcut(); // Slot 11

    [return: MarshalAs(UnmanagedType.BStr)] // Slot 12
    string ExpandEnvironmentStrings([MarshalAs(UnmanagedType.BStr)] string Src);
}
