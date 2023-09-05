namespace ComWrappersTest.Source;

interface IWshShell
{
    public static Guid IID = new("F935DC21-1CF0-11D0-ADB9-00C04FD58A0B");

    int Run(string Command, int WindowStyle, bool WaitOnReturn);
    string ExpandEnvironmentStrings(string Src);
}
