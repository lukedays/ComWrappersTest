using System.Runtime.InteropServices;

namespace ComWrappersTest.Source.Manual;

[Guid("F935DC21-1CF0-11D0-ADB9-00C04FD58A0B")] // The IID for IWshShell (COMView.exe -> TypeLib -> IWshShell interface)
interface IWshShell
{
    int Run(string Command, int WindowStyle, bool WaitOnReturn);
    string ExpandEnvironmentStrings(string Src);
}
