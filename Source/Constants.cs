using System.Runtime.InteropServices;

namespace ComWrappersTest.Source;

public class Constants
{
    public enum HRESULT : int
    {
        S_OK = 0
    }

    public enum CLSCTX : uint
    {
        CLSCTX_INPROC_SERVER = 0x1,
        CLSCTX_INPROC_HANDLER = 0x2,
        CLSCTX_LOCAL_SERVER = 0x4,
        CLSCTX_INPROC_SERVER16 = 0x8,
        CLSCTX_REMOTE_SERVER = 0x10,
        CLSCTX_INPROC_HANDLER16 = 0x20,
        CLSCTX_RESERVED1 = 0x40,
        CLSCTX_RESERVED2 = 0x80,
        CLSCTX_RESERVED3 = 0x100,
        CLSCTX_RESERVED4 = 0x200,
        CLSCTX_NO_CODE_DOWNLOAD = 0x400,
        CLSCTX_RESERVED5 = 0x800,
        CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
        CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
        CLSCTX_NO_FAILURE_LOG = 0x4000,
        CLSCTX_DISABLE_AAA = 0x8000,
        CLSCTX_ENABLE_AAA = 0x10000,
        CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000,
        CLSCTX_ACTIVATE_X86_SERVER = 0x40000,
        CLSCTX_ACTIVATE_32_BIT_SERVER = CLSCTX_ACTIVATE_X86_SERVER,
        CLSCTX_ACTIVATE_64_BIT_SERVER = 0x80000,
        CLSCTX_ENABLE_CLOAKING = 0x100000,
        CLSCTX_APPCONTAINER = 0x400000,
        CLSCTX_ACTIVATE_AAA_AS_IU = 0x800000,
        CLSCTX_RESERVED6 = 0x1000000,
        CLSCTX_ACTIVATE_ARM32_SERVER = 0x2000000,
        CLSCTX_ALLOW_LOWER_TRUST_REGISTRATION = 0x4000000,
        CLSCTX_PS_DLL = 0x80000000
    };

    [StructLayout(LayoutKind.Explicit)]
    public struct VARIANT
    {
        [FieldOffset(0)]
        public VARENUM vt;

        [FieldOffset(2)]
        public ushort wReserved1;

        [FieldOffset(4)]
        public ushort wReserved2;

        [FieldOffset(6)]
        public ushort wReserved3;

        [FieldOffset(8)]
        public long llVal;

        [FieldOffset(8)]
        public int lVal;

        [FieldOffset(8)]
        public byte bVal;

        [FieldOffset(8)]
        public short iVal;

        [FieldOffset(8)]
        public float fltVal;

        [FieldOffset(8)]
        public double dblVal;

        [FieldOffset(8)]
        public VARIANT_BOOL boolVal;

        [FieldOffset(8)]
        public VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;

        [FieldOffset(8)]
        public int scode;

        [FieldOffset(8)]
        public decimal cyVal;

        [FieldOffset(8)]
        public double date;

        [FieldOffset(8)]
        public nint bstrVal;

        [FieldOffset(8)]
        public nint punkVal;

        [FieldOffset(8)]
        public nint pdispVal;

        [FieldOffset(8)]
        public nint parray;

        [FieldOffset(8)]
        public nint pbVal;

        [FieldOffset(8)]
        public nint piVal;

        [FieldOffset(8)]
        public nint plVal;

        [FieldOffset(8)]
        public nint pllVal;

        [FieldOffset(8)]
        public nint pfltVal;

        [FieldOffset(8)]
        public nint pdblVal;

        [FieldOffset(8)]
        public nint pboolVal;

        [FieldOffset(8)]
        public nint __OBSOLETE__VARIANT_PBOOL;

        [FieldOffset(8)]
        public nint pscode;

        [FieldOffset(8)]
        public nint pcyVal;

        [FieldOffset(8)]
        public nint pdate;

        [FieldOffset(8)]
        public nint pbstrVal;

        [FieldOffset(8)]
        public nint ppunkVal;

        [FieldOffset(8)]
        public nint ppdispVal;

        [FieldOffset(8)]
        public nint pparray;

        [FieldOffset(8)]
        public nint pvarVal;

        [FieldOffset(8)]
        public nint byref;

        [FieldOffset(8)]
        public sbyte cVal;

        [FieldOffset(8)]
        public ushort uiVal;

        [FieldOffset(8)]
        public uint ulVal;

        [FieldOffset(8)]
        public ulong ullVal;

        [FieldOffset(8)]
        public int intVal;

        [FieldOffset(8)]
        public uint uintVal;
    }

    public enum VARIANT_BOOL : short
    {
        VARIANT_FALSE = 0,
        VARIANT_TRUE = -1
    }

    public enum VARENUM : ushort
    {
        VT_EMPTY = 0,
        VT_NULL = 1,
        VT_I2 = 2,
        VT_I4 = 3,
        VT_R4 = 4,
        VT_R8 = 5,
        VT_CY = 6,
        VT_DATE = 7,
        VT_BSTR = 8,
        VT_DISPATCH = 9,
        VT_ERROR = 10,
        VT_BOOL = 11,
        VT_VARIANT = 12,
        VT_UNKNOWN = 13,
        VT_DECIMAL = 14,
        VT_I1 = 16,
        VT_UI1 = 17,
        VT_UI2 = 18,
        VT_UI4 = 19,
        VT_I8 = 20,
        VT_UI8 = 21,
        VT_INT = 22,
        VT_UINT = 23,
        VT_VOID = 24,
        VT_HRESULT = 25,
        VT_PTR = 26,
        VT_SAFEARRAY = 27,
        VT_CARRAY = 28,
        VT_USERDEFINED = 29,
        VT_LPSTR = 30,
        VT_LPWSTR = 31,
        VT_RECORD = 36,
        VT_INT_PTR = 37,
        VT_UINT_PTR = 38,
        VT_FILETIME = 64,
        VT_BLOB = 65,
        VT_STREAM = 66,
        VT_STORAGE = 67,
        VT_STREAMED_OBJECT = 68,
        VT_STORED_OBJECT = 69,
        VT_BLOB_OBJECT = 70,
        VT_CF = 71,
        VT_CLSID = 72,
        VT_VERSIONED_STREAM = 73,
        VT_BSTR_BLOB = 0xfff,
        VT_VECTOR = 0x1000,
        VT_ARRAY = 0x2000,
        VT_BYREF = 0x4000,
        VT_RESERVED = 0x8000,
        VT_X = 0xffff,
        VT_Y = 0xfff,
        VT_TYPEMASK = 0xfff
    };
}
