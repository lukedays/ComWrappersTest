using System.Runtime.InteropServices.Marshalling;
using static ComWrappersTest.Source.Constants;

namespace ComWrappersTest.Source.Generator;

[CustomMarshaller(typeof(int), MarshalMode.Default, typeof(IntToVariant))]
public static class IntToVariant
{
    public static nint ConvertToUnmanaged(int managed)
    {
        var variant = new VARIANT { vt = VARENUM.VT_I4, lVal = managed };

        return VariantToPtr(variant);
    }

    public static int ConvertToManaged(nint unmanaged)
    {
        throw new NotImplementedException();
    }
}
