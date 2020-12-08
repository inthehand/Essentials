using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using REFSENSOR_CATEGORY_ID = System.Guid;
using REFSENSOR_ID = System.Guid;
using REFSENSOR_TYPE_ID = System.Guid;

namespace Xamarin.Essentials
{
#pragma warning disable SA1124 // Do not use regions
    #region NativeMethods

    static class PropVariantNativeMethods
    {
        [DllImport("Ole32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        internal static extern void PropVariantClear([In, Out] PropVariant pvar);

        [DllImport("OleAut32.dll", PreserveSig = true, ExactSpelling = true)] // psa is actually returned, not hresult
        internal static extern IntPtr SafeArrayCreateVector(ushort vt, int lowerBound, uint cElems);

        [DllImport("OleAut32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        internal static extern IntPtr SafeArrayAccessData(IntPtr psa);

        [DllImport("OleAut32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        internal static extern HResult SafeArrayUnaccessData(IntPtr psa);

        [DllImport("OleAut32.dll", PreserveSig = true, ExactSpelling = true)] // retuns uint32
        internal static extern uint SafeArrayGetDim(IntPtr psa);

        [DllImport("OleAut32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        internal static extern int SafeArrayGetLBound(IntPtr psa, uint nDim);

        [DllImport("OleAut32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        internal static extern int SafeArrayGetUBound(IntPtr psa, uint nDim);

        // This decl for SafeArrayGetElement is only valid for cDims==1!
        [DllImport("OleAut32.dll", PreserveSig = false, ExactSpelling = true)] // returns hresult
        [return: MarshalAs(UnmanagedType.IUnknown)]
        internal static extern object SafeArrayGetElement(IntPtr psa, ref int rgIndices);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromPropVariantVectorElem([In] PropVariant propvarIn, uint iElem, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromFileTime([In] ref System.Runtime.InteropServices.ComTypes.FILETIME pftIn, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        internal static extern int PropVariantGetElementCount([In] PropVariant propVar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetBooleanElem([In] PropVariant propVar, [In] uint iElem, [Out, MarshalAs(UnmanagedType.Bool)] out bool pfVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetInt16Elem([In] PropVariant propVar, [In] uint iElem, [Out] out short pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetUInt16Elem([In] PropVariant propVar, [In] uint iElem, [Out] out ushort pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetInt32Elem([In] PropVariant propVar, [In] uint iElem, [Out] out int pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetUInt32Elem([In] PropVariant propVar, [In] uint iElem, [Out] out uint pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetInt64Elem([In] PropVariant propVar, [In] uint iElem, [Out] out long pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetUInt64Elem([In] PropVariant propVar, [In] uint iElem, [Out] out ulong pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetDoubleElem([In] PropVariant propVar, [In] uint iElem, [Out] out double pnVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetFileTimeElem([In] PropVariant propVar, [In] uint iElem, [Out, MarshalAs(UnmanagedType.Struct)] out System.Runtime.InteropServices.ComTypes.FILETIME pftVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void PropVariantGetStringElem([In] PropVariant propVar, [In] uint iElem, [MarshalAs(UnmanagedType.LPWStr)] ref string ppszVal);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromBooleanVector([In, MarshalAs(UnmanagedType.LPArray)] bool[] prgf, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromInt16Vector([In, Out] short[] prgn, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromUInt16Vector([In, Out] ushort[] prgn, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromInt32Vector([In, Out] int[] prgn, uint cElems, [Out] PropVariant propVar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromUInt32Vector([In, Out] uint[] prgn, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromInt64Vector([In, Out] long[] prgn, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromUInt64Vector([In, Out] ulong[] prgn, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromDoubleVector([In, Out] double[] prgn, uint cElems, [Out] PropVariant propvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromFileTimeVector([In, Out] System.Runtime.InteropServices.ComTypes.FILETIME[] prgft, uint cElems, [Out] PropVariant ppropvar);

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true, PreserveSig = false, ExactSpelling = true)]
        internal static extern void InitPropVariantFromStringVector([In, Out] string[] prgsz, uint cElems, [Out] PropVariant ppropvar);
    }

    public static class SensorNativeMethods
    {
        [DllImport("kernel32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SystemTimeToFileTime(
            ref SystemTime lpSystemTime,
            out System.Runtime.InteropServices.ComTypes.FILETIME lpFileTime);
    }

    #endregion

    #region Constants

    public enum HResult
    {
        Ok = 0x0000,
        InvalidArguments = unchecked((int)0x80070057),
        OutOfMemory = unchecked((int)0x8007000E),
        NoInterface = unchecked((int)0x80004002),
        Fail = unchecked((int)0x80004005),
        ElementNotFound = unchecked((int)0x80070490),
        TypeElementNotFound = unchecked((int)0x8002802B),
        NoObject = unchecked((int)0x800401E5),
        Win32ErrorCanceled = 1223,
        Canceled = unchecked((int)0x800704C7),
        ResourceInUse = unchecked((int)0x800700AA),
        AccessDenied = unchecked((int)0x80030005)
    }

    public static class CoreErrorHelper
    {
        const int facilityWin32 = 7;

        public const int Ignored = (int)HResult.Ok;

        public static int HResultFromWin32(int win32ErrorCode)
        {
            if (win32ErrorCode > 0)
            {
                win32ErrorCode =
                    (int)(((uint)win32ErrorCode & 0x0000FFFF) | (facilityWin32 << 16) | 0x80000000);
            }
            return win32ErrorCode;
        }

        public static bool Succeeded(int result)
        {
            return result >= 0;
        }

        public static bool Succeeded(HResult result)
        {
            return Succeeded((int)result);
        }

        public static bool Failed(HResult result)
        {
            return !Succeeded(result);
        }

        public static bool Failed(int result)
        {
            return !Succeeded(result);
        }

        public static bool Matches(int result, int win32ErrorCode)
        {
            return result == HResultFromWin32(win32ErrorCode);
        }
    }

    public static class EventInterestTypes
    {
        public static readonly Guid DataUpdated = new Guid(0x2ED0F2A4, 0x0087, 0x41D3, 0x87, 0xDB, 0x67, 0x73, 0x37, 0x0B, 0x3C, 0x88);
        public static readonly Guid StateChanged = new Guid(0xBFD96016, 0x6BD7, 0x4560, 0xAD, 0x34, 0xF2, 0xF6, 0x60, 0x7E, 0x8F, 0x81);
    }

    public enum SensorConnectionType
    {
        Invalid = -1,
        Integrated = 0,
        Attached = 1,
        External = 2
    }

    public enum SensorState
    {
        Removed = -1,
        Ready = 0,
        NotAvailable = 1,
        NoData = 2,
        Initializing = 3,
        AccessDenied = 4,
        Error = 5
    }

    public static class SensorCategories
    {
        public static readonly Guid All = new Guid(0xC317C286, 0xC468, 0x4288, 0x99, 0x75, 0xD4, 0xC4, 0x58, 0x7C, 0x44, 0x2C);

        public static readonly Guid Location = new Guid(0xBFA794E4, 0xF964, 0x4FDB, 0x90, 0xF6, 0x51, 0x5, 0x6B, 0xFE, 0x4B, 0x44);

        public static readonly Guid Environmental = new Guid(0x323439AA, 0x7F66, 0x492B, 0xBA, 0xC, 0x73, 0xE9, 0xAA, 0xA, 0x65, 0xD5);

        public static readonly Guid Motion = new Guid(0xCD09DAF1, 0x3B2E, 0x4C3D, 0xB5, 0x98, 0xB5, 0xE5, 0xFF, 0x93, 0xFD, 0x46);

        public static readonly Guid Orientation = new Guid(0x9E6C04B6, 0x96FE, 0x4954, 0xB7, 0x26, 0x68, 0x68, 0x2A, 0x47, 0x3F, 0x69);

        public static readonly Guid Mechanical = new Guid(0x8D131D68, 0x8EF7, 0x4656, 0x80, 0xB5, 0xCC, 0xCB, 0xD9, 0x37, 0x91, 0xC5);

        public static readonly Guid Electrical = new Guid(0xFB73FCD8, 0xFC4A, 0x483C, 0xAC, 0x58, 0x27, 0xB6, 0x91, 0xC6, 0xBE, 0xFF);

        public static readonly Guid Biometric = new Guid(0xCA19690F, 0xA2C7, 0x477D, 0xA9, 0x9E, 0x99, 0xEC, 0x6E, 0x2B, 0x56, 0x48);

        public static readonly Guid Light = new Guid(0x17A665C0, 0x9063, 0x4216, 0xB2, 0x02, 0x5C, 0x7A, 0x25, 0x5E, 0x18, 0xCE);

        public static readonly Guid Scanner = new Guid(0xB000E77E, 0xF5B5, 0x420F, 0x81, 0x5D, 0x2, 0x70, 0xA7, 0x26, 0xF2, 0x70);
    }

    #endregion

    #region ManagedComponents

    [StructLayout(LayoutKind.Explicit)]
    public sealed class PropVariant : IDisposable
    {
        static Array CrackSingleDimSafeArray(IntPtr psa)
        {
            var cDims = PropVariantNativeMethods.SafeArrayGetDim(psa);
            if (cDims != 1)
                throw new ArgumentException();

            var lBound = PropVariantNativeMethods.SafeArrayGetLBound(psa, 1U);
            var uBound = PropVariantNativeMethods.SafeArrayGetUBound(psa, 1U);

            var n = uBound - lBound + 1; // uBound is inclusive

            var array = new object[n];
            for (var i = lBound; i <= uBound; ++i)
            {
                array[i] = PropVariantNativeMethods.SafeArrayGetElement(psa, ref i);
            }

            return array;
        }

        static Dictionary<Type, Action<PropVariant, Array, uint>> vectorActions = null;

        static Dictionary<Type, Action<PropVariant, Array, uint>> GenerateVectorActions()
        {
            var cache = new Dictionary<Type, Action<PropVariant, Array, uint>>();

            cache.Add(typeof(short), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetInt16Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(ushort), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetUInt16Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(int), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetInt32Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(uint), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetUInt32Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(long), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetInt64Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(ulong), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetUInt64Elem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(DateTime), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetFileTimeElem(pv, i, out var val);

                var fileTime = GetFileTimeAsLong(ref val);

                array.SetValue(DateTime.FromFileTime(fileTime), i);
            });

            cache.Add(typeof(bool), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetBooleanElem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(double), (pv, array, i) =>
            {
                PropVariantNativeMethods.PropVariantGetDoubleElem(pv, i, out var val);
                array.SetValue(val, i);
            });

            cache.Add(typeof(float), (pv, array, i) => // float
            {
                var val = new float[1];
                Marshal.Copy(pv.ptr2, val, (int)i, 1);
                array.SetValue(val[0], (int)i);
            });

            cache.Add(typeof(decimal), (pv, array, i) =>
            {
                var val = new int[4];
                for (var a = 0; a < val.Length; a++)
                {
                    val[a] = Marshal.ReadInt32(pv.ptr2, ((int)i * sizeof(decimal)) + (a * sizeof(int)));
                }
                array.SetValue(new decimal(val), i);
            });

            cache.Add(typeof(string), (pv, array, i) =>
            {
                var val = string.Empty;
                PropVariantNativeMethods.PropVariantGetStringElem(pv, i, ref val);
                array.SetValue(val, i);
            });

            return cache;
        }

        public static PropVariant FromObject(object value)
        {
            if (value == null)
            {
                return new PropVariant();
            }
            else
            {
                var func = GetDynamicConstructor(value.GetType());
                return func(value);
            }
        }

        // A dictionary and lock to contain compiled expression trees for constructors
        static Dictionary<Type, Func<object, PropVariant>> cache = new Dictionary<Type, Func<object, PropVariant>>();
        static object padlock = new object();

        // Retrieves a cached constructor expression.
        // If no constructor has been cached, it attempts to find/add it.  If it cannot be found
        // an exception is thrown.
        // This method looks for a public constructor with the same parameter type as the object.
        static Func<object, PropVariant> GetDynamicConstructor(Type type)
        {
            lock (padlock)
            {
                // initial check, if action is found, return it
                if (!cache.TryGetValue(type, out var action))
                {
                    // iterates through all constructors
                    var constructor = typeof(PropVariant)
                        .GetConstructor(new Type[] { type });

                    if (constructor == null)
                    { // if the method was not found, throw.
                        throw new ArgumentException();
                    }
                    else
                    {
                        // create parameters to action
                        var arg = Expression.Parameter(typeof(object), "arg");

                        // create an expression to invoke the constructor with an argument cast to the correct type
                        var create = Expression.New(constructor, Expression.Convert(arg, type));

                        // compiles expression into an action delegate
                        action = Expression.Lambda<Func<object, PropVariant>>(create, arg).Compile();
                        cache.Add(type, action);
                    }
                }
                return action;
            }
        }

        [FieldOffset(0)]
        decimal @decimal;

        // This is actually a VarEnum value, but the VarEnum type
        // requires 4 bytes instead of the expected 2.
        [FieldOffset(0)]
        ushort valueType;

        // In order to allow x64 compat, we need to allow for
        // expansion of the IntPtr. However, the BLOB struct
        // uses a 4-byte int, followed by an IntPtr, so
        // although the valueData field catches most pointer values,
        // we need an additional 4-bytes to get the BLOB
        // pointer. The valueDataExt field provides this, as well as
        // the last 4-bytes of an 8-byte value on 32-bit
        // architectures.
        [FieldOffset(12)]
        IntPtr ptr2;
        [FieldOffset(8)]
        IntPtr ptr;
        [FieldOffset(8)]
        int int32;
        [FieldOffset(8)]
        uint uint32;
        [FieldOffset(8)]
        byte @byte;
        [FieldOffset(8)]
        sbyte @sbyte;
        [FieldOffset(8)]
        short @short;
        [FieldOffset(8)]
        ushort @ushort;
        [FieldOffset(8)]
        long @long;
        [FieldOffset(8)]
        ulong @ulong;
        [FieldOffset(8)]
        double @double;
        [FieldOffset(8)]
        float @float;

        public PropVariant()
        {
            // left empty
        }

        public PropVariant(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            valueType = (ushort)VarEnum.VT_LPWSTR;
            ptr = Marshal.StringToCoTaskMemUni(value);
        }

        public PropVariant(string[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromStringVector(value, (uint)value.Length, this);
        }

        public PropVariant(bool[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromBooleanVector(value, (uint)value.Length, this);
        }

        public PropVariant(short[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromInt16Vector(value, (uint)value.Length, this);
        }

        /// <summary>
        /// Set a short vector
        /// </summary>
        public PropVariant(ushort[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromUInt16Vector(value, (uint)value.Length, this);
        }

        /// <summary>
        /// Set an int vector
        /// </summary>
        public PropVariant(int[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromInt32Vector(value, (uint)value.Length, this);
        }

        /// <summary>
        /// Set an uint vector
        /// </summary>
        public PropVariant(uint[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromUInt32Vector(value, (uint)value.Length, this);
        }

        public PropVariant(long[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromInt64Vector(value, (uint)value.Length, this);
        }

        public PropVariant(ulong[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromUInt64Vector(value, (uint)value.Length, this);
        }

        public PropVariant(double[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            PropVariantNativeMethods.InitPropVariantFromDoubleVector(value, (uint)value.Length, this);
        }

        public PropVariant(DateTime[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }
            var fileTimeArr =
                new System.Runtime.InteropServices.ComTypes.FILETIME[value.Length];

            for (var i = 0; i < value.Length; i++)
            {
                fileTimeArr[i] = DateTimeToFileTime(value[i]);
            }

            PropVariantNativeMethods.InitPropVariantFromFileTimeVector(fileTimeArr, (uint)fileTimeArr.Length, this);
        }

        public PropVariant(bool value)
        {
            valueType = (ushort)VarEnum.VT_BOOL;
            int32 = (value == true) ? -1 : 0;
        }

        public PropVariant(DateTime value)
        {
            valueType = (ushort)VarEnum.VT_FILETIME;

            var ft = DateTimeToFileTime(value);
            PropVariantNativeMethods.InitPropVariantFromFileTime(ref ft, this);
        }

        public PropVariant(byte value)
        {
            valueType = (ushort)VarEnum.VT_UI1;
            @byte = value;
        }

        public PropVariant(sbyte value)
        {
            valueType = (ushort)VarEnum.VT_I1;
            @sbyte = value;
        }

        public PropVariant(short value)
        {
            valueType = (ushort)VarEnum.VT_I2;
            @short = value;
        }

        public PropVariant(ushort value)
        {
            valueType = (ushort)VarEnum.VT_UI2;
            @ushort = value;
        }

        public PropVariant(int value)
        {
            valueType = (ushort)VarEnum.VT_I4;
            int32 = value;
        }

        public PropVariant(uint value)
        {
            valueType = (ushort)VarEnum.VT_UI4;
            uint32 = value;
        }

        public PropVariant(decimal value)
        {
            @decimal = value;

            // It is critical that the value type be set after the decimal value, because they overlap.
            // If valuetype is written first, its value will be lost when _decimal is written.
            valueType = (ushort)VarEnum.VT_DECIMAL;
        }

        public PropVariant(decimal[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            valueType = (ushort)(VarEnum.VT_DECIMAL | VarEnum.VT_VECTOR);
            int32 = value.Length;

            // allocate required memory for array with 128bit elements
            ptr2 = Marshal.AllocCoTaskMem(value.Length * sizeof(decimal));
            for (var i = 0; i < value.Length; i++)
            {
                var bits = decimal.GetBits(value[i]);
                Marshal.Copy(bits, 0, ptr2, bits.Length);
            }
        }

        public PropVariant(float value)
        {
            valueType = (ushort)VarEnum.VT_R4;

            @float = value;
        }

        public PropVariant(float[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            valueType = (ushort)(VarEnum.VT_R4 | VarEnum.VT_VECTOR);
            int32 = value.Length;

            ptr2 = Marshal.AllocCoTaskMem(value.Length * sizeof(float));

            Marshal.Copy(value, 0, ptr2, value.Length);
        }

        public PropVariant(long value)
        {
            @long = value;
            valueType = (ushort)VarEnum.VT_I8;
        }

        public PropVariant(ulong value)
        {
            valueType = (ushort)VarEnum.VT_UI8;
            @ulong = value;
        }

        public PropVariant(double value)
        {
            valueType = (ushort)VarEnum.VT_R8;
            @double = value;
        }

        internal void SetIUnknown(object value)
        {
            valueType = (ushort)VarEnum.VT_UNKNOWN;
            ptr = Marshal.GetIUnknownForObject(value);
        }

        internal void SetSafeArray(Array array)
        {
            if (array == null)
            {
                throw new ArgumentNullException("array");
            }
            const ushort vtUnknown = 13;
            var psa = PropVariantNativeMethods.SafeArrayCreateVector(vtUnknown, 0, (uint)array.Length);

            var pvData = PropVariantNativeMethods.SafeArrayAccessData(psa);
            try
            {
                for (var i = 0; i < array.Length; ++i)
                {
                    var obj = array.GetValue(i);
                    var punk = (obj != null) ? Marshal.GetIUnknownForObject(obj) : IntPtr.Zero;
                    Marshal.WriteIntPtr(pvData, i * IntPtr.Size, punk);
                }
            }
            finally
            {
                PropVariantNativeMethods.SafeArrayUnaccessData(psa);
            }

            valueType = (ushort)VarEnum.VT_ARRAY | (ushort)VarEnum.VT_UNKNOWN;
            ptr = psa;
        }

        public VarEnum VarType
        {
            get { return (VarEnum)valueType; }
            set { valueType = (ushort)value; }
        }

        public bool IsNullOrEmpty => valueType == (ushort)VarEnum.VT_EMPTY || valueType == (ushort)VarEnum.VT_NULL;

        public object Value
        {
            get
            {
                switch ((VarEnum)valueType)
                {
                    case VarEnum.VT_I1:
                        return @sbyte;
                    case VarEnum.VT_UI1:
                        return @byte;
                    case VarEnum.VT_I2:
                        return @short;
                    case VarEnum.VT_UI2:
                        return @ushort;
                    case VarEnum.VT_I4:
                    case VarEnum.VT_INT:
                        return int32;
                    case VarEnum.VT_UI4:
                    case VarEnum.VT_UINT:
                        return uint32;
                    case VarEnum.VT_I8:
                        return @long;
                    case VarEnum.VT_UI8:
                        return @ulong;
                    case VarEnum.VT_R4:
                        return @float;
                    case VarEnum.VT_R8:
                        return @double;
                    case VarEnum.VT_BOOL:
                        return int32 == -1;
                    case VarEnum.VT_ERROR:
                        return @long;
                    case VarEnum.VT_CY:
                        return @decimal;
                    case VarEnum.VT_DATE:
                        return DateTime.FromOADate(@double);
                    case VarEnum.VT_FILETIME:
                        return DateTime.FromFileTime(@long);
                    case VarEnum.VT_BSTR:
                        return Marshal.PtrToStringBSTR(ptr);
                    case VarEnum.VT_BLOB:
                        return GetBlobData();
                    case VarEnum.VT_LPSTR:
                        return Marshal.PtrToStringAnsi(ptr);
                    case VarEnum.VT_LPWSTR:
                        return Marshal.PtrToStringUni(ptr);
                    case VarEnum.VT_UNKNOWN:
                        return Marshal.GetObjectForIUnknown(ptr);
                    case VarEnum.VT_DISPATCH:
                        return Marshal.GetObjectForIUnknown(ptr);
                    case VarEnum.VT_DECIMAL:
                        return @decimal;
                    case VarEnum.VT_ARRAY | VarEnum.VT_UNKNOWN:
                        return CrackSingleDimSafeArray(ptr);
                    case VarEnum.VT_VECTOR | VarEnum.VT_LPWSTR:
                        return GetVector<string>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_I2:
                        return GetVector<short>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_UI2:
                        return GetVector<ushort>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_I4:
                        return GetVector<int>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_UI4:
                        return GetVector<uint>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_I8:
                        return GetVector<long>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_UI8:
                        return GetVector<ulong>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_R4:
                        return GetVector<float>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_R8:
                        return GetVector<double>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_BOOL:
                        return GetVector<bool>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_FILETIME:
                        return GetVector<DateTime>();
                    case VarEnum.VT_VECTOR | VarEnum.VT_DECIMAL:
                        return GetVector<decimal>();
                    default:
                        // if the value cannot be marshaled
                        return null;
                }
            }
        }

        static long GetFileTimeAsLong(ref System.Runtime.InteropServices.ComTypes.FILETIME val)
        {
            return (((long)val.dwHighDateTime) << 32) + val.dwLowDateTime;
        }

        static System.Runtime.InteropServices.ComTypes.FILETIME DateTimeToFileTime(DateTime value)
        {
            var hFT = value.ToFileTime();
            var ft =
                new System.Runtime.InteropServices.ComTypes.FILETIME
                {
                    dwLowDateTime = (int)(hFT & 0xFFFFFFFF),
                    dwHighDateTime = (int)(hFT >> 32)
                };
            return ft;
        }

        object GetBlobData()
        {
            var blobData = new byte[int32];

            var pBlobData = ptr2;
            Marshal.Copy(pBlobData, blobData, 0, int32);

            return blobData;
        }

        Array GetVector<T>()
        {
            var count = PropVariantNativeMethods.PropVariantGetElementCount(this);
            if (count <= 0)
            {
                return null;
            }

            lock (padlock)
            {
                if (vectorActions == null)
                {
                    vectorActions = GenerateVectorActions();
                }
            }

            if (!vectorActions.TryGetValue(typeof(T), out var action))
            {
                throw new InvalidCastException();
            }

            Array array = new T[count];
            for (uint i = 0; i < count; i++)
            {
                action(this, array, i);
            }

            return array;
        }

        public void Dispose()
        {
            PropVariantNativeMethods.PropVariantClear(this);

            GC.SuppressFinalize(this);
        }

        ~PropVariant()
        {
            Dispose();
        }

        public override string ToString()
        {
            return $"{Value}: {VarType}";
        }
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey : IEquatable<PropertyKey>
    {
        REFSENSOR_TYPE_ID formatId;
        int propertyId;

        public REFSENSOR_TYPE_ID FormatId => formatId;

        public int PropertyId => propertyId;

        public PropertyKey(REFSENSOR_TYPE_ID formatId, int propertyId)
        {
            this.formatId = formatId;
            this.propertyId = propertyId;
        }

        public PropertyKey(string formatId, int propertyId)
        {
            this.formatId = new REFSENSOR_TYPE_ID(formatId);
            this.propertyId = propertyId;
        }

        public bool Equals(PropertyKey other)
        {
            return other.Equals((object)this);
        }

        public override int GetHashCode()
        {
            return formatId.GetHashCode() ^ propertyId;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is PropertyKey))
                return false;

            var other = (PropertyKey)obj;
            return other.formatId.Equals(formatId) && (other.propertyId == propertyId);
        }

        public static bool operator ==(PropertyKey propKey1, PropertyKey propKey2)
        {
            return propKey1.Equals(propKey2);
        }

        public static bool operator !=(PropertyKey propKey1, PropertyKey propKey2)
        {
            return !propKey1.Equals(propKey2);
        }

        public override string ToString()
        {
            return $"FormatId: {FormatId}, PropertyId: {PropertyId}";
        }
    }

    public static class SensorPropertyKeys
    {
        public static readonly REFSENSOR_TYPE_ID SensorPropertyCommonGuid = new REFSENSOR_TYPE_ID(0X7F8383EC, 0XD3EC, 0X495C, 0XA8, 0XCF, 0XB8, 0XBB, 0XE8, 0X5C, 0X29, 0X20);

        public static readonly PropertyKey SensorPropertyType = new PropertyKey(SensorPropertyCommonGuid, 2);

        public static readonly PropertyKey SensorPropertyState = new PropertyKey(SensorPropertyCommonGuid, 3);

        public static readonly PropertyKey SensorPropertySamplingRate = new PropertyKey(SensorPropertyCommonGuid, 4);

        public static readonly PropertyKey SensorPropertyPersistentUniqueId = new PropertyKey(SensorPropertyCommonGuid, 5);

        public static readonly PropertyKey SensorPropertyManufacturer = new PropertyKey(SensorPropertyCommonGuid, 6);

        public static readonly PropertyKey SensorPropertyModel = new PropertyKey(SensorPropertyCommonGuid, 7);

        public static readonly PropertyKey SensorPropertySerialNumber = new PropertyKey(SensorPropertyCommonGuid, 8);

        public static readonly PropertyKey SensorPropertyFriendlyName = new PropertyKey(SensorPropertyCommonGuid, 9);

        public static readonly PropertyKey SensorPropertyDescription = new PropertyKey(SensorPropertyCommonGuid, 10);

        public static readonly PropertyKey SensorPropertyConnectionType = new PropertyKey(SensorPropertyCommonGuid, 11);

        public static readonly PropertyKey SensorPropertyMinReportInterval = new PropertyKey(SensorPropertyCommonGuid, 12);

        public static readonly PropertyKey SensorPropertyCurrentReportInterval = new PropertyKey(SensorPropertyCommonGuid, 13);

        public static readonly PropertyKey SensorPropertyChangeSensitivity = new PropertyKey(SensorPropertyCommonGuid, 14);

        public static readonly PropertyKey SensorPropertyDeviceId = new PropertyKey(SensorPropertyCommonGuid, 15);

        public static readonly PropertyKey SensorPropertyAccuracy = new PropertyKey(SensorPropertyCommonGuid, 16);

        public static readonly PropertyKey SensorPropertyResolution = new PropertyKey(SensorPropertyCommonGuid, 17);

        public static readonly PropertyKey SensorPropertyLightResponseCurve = new PropertyKey(SensorPropertyCommonGuid, 18);

        public static readonly PropertyKey SensorDataTypeTimestamp = new PropertyKey(new REFSENSOR_TYPE_ID(0XDB5E0CF2, 0XCF1F, 0X4C18, 0XB4, 0X6C, 0XD8, 0X60, 0X11, 0XD6, 0X21, 0X50), 2);

        public static readonly PropertyKey SensorDataTypeLatitudeDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 2);

        public static readonly PropertyKey SensorDataTypeLongitudeDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 3);

        public static readonly PropertyKey SensorDataTypeAltitudeSeaLevelMeters = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 4);

        public static readonly PropertyKey SensorDataTypeAltitudeEllipsoidMeters = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 5);

        public static readonly PropertyKey SensorDataTypeSpeedKnots = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 6);

        public static readonly PropertyKey SensorDataTypeTrueHeadingDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 7);

        public static readonly PropertyKey SensorDataTypeMagneticHeadingDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 8);

        public static readonly PropertyKey SensorDataTypeMagneticVariation = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 9);

        public static readonly PropertyKey SensorDataTypeFixQuality = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 10);

        public static readonly PropertyKey SensorDataTypeFixType = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 11);

        public static readonly PropertyKey SensorDataTypePositionDilutionOfPrecision = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 12);

        public static readonly PropertyKey SensorDataTypeHorizontalDilutionOfPrecision = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 13);

        public static readonly PropertyKey SensorDataTypeVerticalDilutionOfPrecision = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 14);

        public static readonly PropertyKey SensorDataTypeSatellitesUsedCount = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 15);

        public static readonly PropertyKey SensorDataTypeSatellitesUsedPrns = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 16);

        public static readonly PropertyKey SensorDataTypeSatellitesInView = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 17);

        public static readonly PropertyKey SensorDataTypeSatellitesInViewPrns = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 18);

        public static readonly PropertyKey SensorDataTypeSatellitesInViewElevation = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 19);

        public static readonly PropertyKey SensorDataTypeSatellitesInViewAzimuth = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 20);

        public static readonly PropertyKey SensorDataTypeSatellitesInViewStnRatio = new PropertyKey(new REFSENSOR_TYPE_ID(0X055C74D8, 0XCA6F, 0X47D6, 0X95, 0XC6, 0X1E, 0XD3, 0X63, 0X7A, 0X0F, 0XF4), 21);

        public static readonly PropertyKey SensorDataTypeTemperatureCelsius = new PropertyKey(new REFSENSOR_TYPE_ID(0X8B0AA2F1, 0X2D57, 0X42EE, 0X8C, 0XC0, 0X4D, 0X27, 0X62, 0X2B, 0X46, 0XC4), 2);

        public static readonly PropertyKey SensorDataTypeAccelerationXG = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 2);

        public static readonly PropertyKey SensorDataTypeAccelerationYG = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 3);

        public static readonly PropertyKey SensorDataTypeAccelerationZG = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 4);

        public static readonly PropertyKey SensorDataTypeAngularAccelerationXDegreesPerSecond = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 5);

        public static readonly PropertyKey SensorDataTypeAngularAccelerationYDegreesPerSecond = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 6);

        public static readonly PropertyKey SensorDataTypeAngularAccelerationZDegreesPerSecond = new PropertyKey(new REFSENSOR_TYPE_ID(0X3F8A69A2, 0X7C5, 0X4E48, 0XA9, 0X65, 0XCD, 0X79, 0X7A, 0XAB, 0X56, 0XD5), 7);

        public static readonly PropertyKey SensorDataTypeAngleXDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 2);

        public static readonly PropertyKey SensorDataTypeAngleYDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 3);

        public static readonly PropertyKey SensorDataTypeAngleZDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 4);

        public static readonly PropertyKey SensorDataTypeMagneticHeadingXDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 5);

        public static readonly PropertyKey SensorDataTypeMagneticHeadingYDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 6);

        public static readonly PropertyKey SensorDataTypeMagneticHeadingZDegrees = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 7);

        public static readonly PropertyKey SensorDataTypeDistanceXMeters = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 8);

        public static readonly PropertyKey SensorDataTypeDistanceYMeters = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 9);

        public static readonly PropertyKey SensorDataTypeDistanceZMeters = new PropertyKey(new REFSENSOR_TYPE_ID(0XC2FB0F5F, 0XE2D2, 0X4C78, 0XBC, 0XD0, 0X35, 0X2A, 0X95, 0X82, 0X81, 0X9D), 10);

        public static readonly PropertyKey SensorDataTypeBooleanSwitchState = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 2);

        public static readonly PropertyKey SensorDataTypeMultivalueSwitchState = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 3);

        public static readonly PropertyKey SensorDataTypeBooleanSwitchArrayState = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 10);

        public static readonly PropertyKey SensorDataTypeForceNewtons = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 4);

        public static readonly PropertyKey SensorDataTypeWeightKilograms = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 5);

        public static readonly PropertyKey SensorDataTypePressurePascal = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 6);

        public static readonly PropertyKey SensorDataTypeAtmosphericPressureBar = new PropertyKey(new REFSENSOR_TYPE_ID("8B0AA2F1-2D57-42EE-8CC0-4D27622B46C4"), 4);

        public static readonly PropertyKey SensorDataTypeStrain = new PropertyKey(new REFSENSOR_TYPE_ID(0X38564A7C, 0XF2F2, 0X49BB, 0X9B, 0X2B, 0XBA, 0X60, 0XF6, 0X6A, 0X58, 0XDF), 7);

        public static readonly PropertyKey SensorDataTypeHumanPresence = new PropertyKey(new REFSENSOR_TYPE_ID(0X2299288A, 0X6D9E, 0X4B0B, 0XB7, 0XEC, 0X35, 0X28, 0XF8, 0X9E, 0X40, 0XAF), 2);

        public static readonly PropertyKey SensorDataTypeHumanProximity = new PropertyKey(new REFSENSOR_TYPE_ID(0X2299288A, 0X6D9E, 0X4B0B, 0XB7, 0XEC, 0X35, 0X28, 0XF8, 0X9E, 0X40, 0XAF), 3);

        public static readonly PropertyKey SensorDataTypeLightLux = new PropertyKey(new REFSENSOR_TYPE_ID(0XE4C77CE2, 0XDCB7, 0X46E9, 0X84, 0X39, 0X4F, 0XEC, 0X54, 0X88, 0X33, 0XA6), 2);

        public static readonly PropertyKey SensorDataTypeRfidTag40Bit = new PropertyKey(new REFSENSOR_TYPE_ID(0XD7A59A3C, 0X3421, 0X44AB, 0X8D, 0X3A, 0X9D, 0XE8, 0XAB, 0X6C, 0X4C, 0XAE), 2);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SystemTime
    {
        internal ushort Year;
        internal ushort Month;
        internal ushort DayOfWeek;
        internal ushort Day;
        internal ushort Hour;
        internal ushort Minute;
        internal ushort Second;
        internal ushort Millisecond;

        public DateTime DateTime => new DateTime(Year, Month, Day, Hour, Minute, Second, Millisecond);

        public static implicit operator DateTime(SystemTime systemTime)
        {
            return systemTime.DateTime;
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:D2}/{1:D2}/{2:D4}, {3:D2}:{4:D2}:{5:D2}.{6}", Month, Day, Year, Hour, Minute, Second, Millisecond);
        }
    }

    public enum NativeSensorState
    {
        Ready = 0,
        NotAvailable = 1,
        NoData,
        Initializing,
        AccessDenied,
        Error
    }

    public class UnknownSensor : Sensor
    {
    }

    public class SensorList<TSensor> : IList<TSensor>
        where TSensor : Sensor
    {
        readonly List<TSensor> sensorList = new List<TSensor>();

        public int IndexOf(TSensor item)
        {
            return sensorList.IndexOf(item);
        }

        public void Insert(int index, TSensor item)
        {
            sensorList.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            sensorList.RemoveAt(index);
        }

        public TSensor this[int index]
        {
            get
            {
                return sensorList[index];
            }

            set
            {
                sensorList[index] = value;
            }
        }

        public void Add(TSensor item)
        {
            sensorList.Add(item);
        }

        public void Clear()
        {
            sensorList.Clear();
        }

        public bool Contains(TSensor item)
        {
            return sensorList.Contains(item);
        }

        public void CopyTo(TSensor[] array, int arrayIndex)
        {
            sensorList.CopyTo(array, arrayIndex);
        }

        public int Count => sensorList.Count;

        public bool IsReadOnly => (sensorList as ICollection<TSensor>).IsReadOnly;

        public bool Remove(TSensor item)
        {
            return sensorList.Remove(item);
        }

        public IEnumerator<TSensor> GetEnumerator()
        {
            return sensorList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return sensorList.GetEnumerator();
        }
    }

    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SensorDescriptionAttribute : Attribute
    {
        Guid sensorType;

        public SensorDescriptionAttribute(string sensorType)
        {
            this.sensorType = new REFSENSOR_TYPE_ID(sensorType);
        }

        public string SensorType => sensorType.ToString();

        public Guid SensorTypeGuid => sensorType;
    }

    public enum SensorAvailabilityChange
    {
        Addition,
        Removal
    }

    public class SensorsChangedEventArgs : EventArgs
    {
        public SensorAvailabilityChange Change
        {
            get;
            set;
        }

        public Guid SensorId
        {
            get;
            set;
        }

        internal SensorsChangedEventArgs(Guid sensorId, SensorAvailabilityChange change)
        {
            SensorId = sensorId;
            Change = change;
        }
    }

    public delegate void SensorsChangedEventHandler(SensorsChangedEventArgs e);

    public static class SensorManager
    {
        public static SensorList<Sensor> GetAllSensors()
        {
            return GetSensorsByCategoryId(SensorCategories.All);
        }

        public static SensorList<Sensor> GetSensorsByCategoryId(Guid category)
        {
            var hr = sensorManager.GetSensorsByCategory(category, out var sensorCollection);
            if (hr == (int)HResult.ElementNotFound)
            {
                return new SensorList<Sensor>();
            }

            return NativeSensorCollectionToSensorCollection<Sensor>(sensorCollection);
        }

        public static SensorList<Sensor> GetSensorsByTypeId(Guid typeId)
        {
            var hr = sensorManager.GetSensorsByType(typeId, out var sensorCollection);
            if (hr == (int)HResult.ElementNotFound)
            {
                return new SensorList<Sensor>();
            }
            return NativeSensorCollectionToSensorCollection<Sensor>(sensorCollection);
        }

        public static SensorList<T> GetSensorsByTypeId<T>()
            where T : Sensor
        {
            var attrs = typeof(T).GetCustomAttributes(typeof(SensorDescriptionAttribute), true);
            if (attrs != null && attrs.Length > 0)
            {
                var sda = attrs[0] as SensorDescriptionAttribute;

                var hr = sensorManager.GetSensorsByType(sda.SensorTypeGuid, out var nativeSensorCollection);
                if (hr == (int)HResult.ElementNotFound)
                {
                    return new SensorList<T>();
                }
                return NativeSensorCollectionToSensorCollection<T>(nativeSensorCollection);
            }

            return new SensorList<T>();
        }

        public static T GetSensorBySensorId<T>(Guid sensorId)
            where T : Sensor
        {
            var hr = sensorManager.GetSensorByID(sensorId, out var nativeSensor);
            if (hr == (int)HResult.ElementNotFound)
            {
                throw new InvalidOperationException("Sensor with specific ID not found");
            }

            if (nativeSensor != null)
            {
                return GetSensorWrapperInstance<T>(nativeSensor);
            }

            return null;
        }

        public static void RequestPermission(IntPtr parentWindowHandle, bool modal, SensorList<Sensor> sensors)
        {
            if (sensors == null || sensors.Count == 0)
            {
                throw new ArgumentException(nameof(sensors));
            }

            ISensorCollection sensorCollection = new SensorCollection();

            foreach (var sensor in sensors)
            {
                sensorCollection.Add(sensor.InternalObject);
            }

            sensorManager.RequestPermissions(parentWindowHandle, sensorCollection, modal);
        }

        public static event SensorsChangedEventHandler SensorsChanged;

        static NativeSensorManager sensorManager = new NativeSensorManager();

        static NativeSensorManagerEventSink sensorManagerEventSink = new NativeSensorManagerEventSink();

        static Dictionary<Guid, SensorTypeData> guidToSensorDescr = new Dictionary<Guid, SensorTypeData>();

        static Dictionary<Type, Guid> sensorTypeToGuid = new Dictionary<Type, Guid>();

        static SensorManager()
        {
            BuildSensorTypeMap();
            Thread.MemoryBarrier();
            sensorManager.SetEventSink(sensorManagerEventSink);
        }

        internal static SensorList<TS> NativeSensorCollectionToSensorCollection<TS>(ISensorCollection nativeCollection)
            where TS : Sensor
        {
            var sensors = new SensorList<TS>();

            if (nativeCollection != null)
            {
                nativeCollection.GetCount(out var sensorCount);

                for (uint i = 0; i < sensorCount; i++)
                {
                    nativeCollection.GetAt(i, out var iSensor);
                    var sensor = GetSensorWrapperInstance<TS>(iSensor);
                    if (sensor != null)
                    {
                        sensor.InternalObject = iSensor;
                        sensors.Add(sensor);
                    }
                }
            }

            return sensors;
        }

        internal static void OnSensorsChanged(Guid sensorId, SensorAvailabilityChange change)
        {
            if (SensorsChanged != null)
            {
                SensorsChanged.Invoke(new SensorsChangedEventArgs(sensorId, change));
            }
        }

        static void BuildSensorTypeMap()
        {
            var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (var asm in loadedAssemblies)
            {
                try
                {
                    var exportedTypes = asm.GetExportedTypes();
                    foreach (var t in exportedTypes)
                    {
                        if (t.IsSubclassOf(typeof(Sensor)) && t.IsPublic && !t.IsAbstract && !t.IsGenericType)
                        {
                            var attrs = t.GetCustomAttributes(typeof(SensorDescriptionAttribute), true);
                            if (attrs != null && attrs.Length > 0)
                            {
                                var sda = (SensorDescriptionAttribute)attrs[0];
                                var stm = new SensorTypeData(t, sda);

                                guidToSensorDescr.Add(sda.SensorTypeGuid, stm);
                                sensorTypeToGuid.Add(t, sda.SensorTypeGuid);
                            }
                        }
                    }
                }
                catch (System.NotSupportedException)
                {
                    // GetExportedTypes can throw this if dynamic assemblies are loaded
                    // via Reflection.Emit.
                }
                catch (System.IO.FileNotFoundException)
                {
                    // GetExportedTypes can throw this if a loaded asembly is not in the
                    // current directory or path.
                }
            }
        }

        static TS GetSensorWrapperInstance<TS>(ISensor nativeISensor)
            where TS : Sensor
        {
            nativeISensor.GetType(out var sensorTypeGuid);

            var sensorClassType =
                guidToSensorDescr.TryGetValue(sensorTypeGuid, out var stm) ? stm.SensorType : typeof(UnknownSensor);

            try
            {
                var sensor = (TS)Activator.CreateInstance(sensorClassType);
                sensor.InternalObject = nativeISensor;
                return sensor;
            }
            catch (InvalidCastException)
            {
                return null;
            }
        }
    }

    public struct SensorTypeData
    {
        Type sensorType;
        SensorDescriptionAttribute sda;

        public SensorTypeData(Type sensorClassType, SensorDescriptionAttribute sda)
        {
            sensorType = sensorClassType;
            this.sda = sda;
        }

        public Type SensorType => sensorType;

        public SensorDescriptionAttribute Attr => sda;
    }

    public class NativeSensorManagerEventSink : ISensorManagerEvents
    {
        public void OnSensorEnter(ISensor nativeSensor, NativeSensorState state)
        {
            if (state == NativeSensorState.Ready)
            {
                var hr = nativeSensor.GetID(out var sensorId);
                if (hr == (int)HResult.Ok)
                {
                    SensorManager.OnSensorsChanged(sensorId, SensorAvailabilityChange.Addition);
                }
            }
        }
    }

    public class SensorData : IDictionary<Guid, IList<object>>
    {
        internal static SensorData FromNativeReport(ISensor iSensor, ISensorDataReport iReport)
        {
            var data = new SensorData();

            iSensor.GetSupportedDataFields(out var keyCollection);
            iReport.GetSensorValues(keyCollection, out var valuesCollection);

            keyCollection.GetCount(out var items);
            for (uint index = 0; index < items; index++)
            {
                using (var propValue = new PropVariant())
                {
                    keyCollection.GetAt(index, out var key);
                    valuesCollection.GetValue(ref key, propValue);

                    if (data.ContainsKey(key.FormatId))
                    {
                        data[key.FormatId].Add(propValue.Value);
                    }
                    else
                    {
                        data.Add(key.FormatId, new List<object> { propValue.Value });
                    }
                }
            }

            if (keyCollection != null)
            {
                Marshal.ReleaseComObject(keyCollection);
                keyCollection = null;
            }

            if (valuesCollection != null)
            {
                Marshal.ReleaseComObject(valuesCollection);
                valuesCollection = null;
            }
            return data;
        }

        Dictionary<Guid, IList<object>> sensorDataDictionary = new Dictionary<Guid, IList<object>>();

        public void Add(Guid key, IList<object> value)
        {
            sensorDataDictionary.Add(key, value);
        }

        public bool ContainsKey(Guid key)
        {
            return sensorDataDictionary.ContainsKey(key);
        }

        public ICollection<Guid> Keys => sensorDataDictionary.Keys;

        public bool Remove(Guid key)
        {
            return sensorDataDictionary.Remove(key);
        }

        public bool TryGetValue(Guid key, out IList<object> value)
        {
            return sensorDataDictionary.TryGetValue(key, out value);
        }

        public ICollection<IList<object>> Values => sensorDataDictionary.Values;

        public IList<object> this[Guid key]
        {
            get
            {
                return sensorDataDictionary[key];
            }

            set
            {
                sensorDataDictionary[key] = value;
            }
        }

        public void Add(KeyValuePair<Guid, IList<object>> item)
        {
            var c = sensorDataDictionary as ICollection<KeyValuePair<Guid, IList<object>>>;
            c.Add(item);
        }

        public void Clear()
        {
            var c = sensorDataDictionary as ICollection<KeyValuePair<Guid, IList<object>>>;
            c.Clear();
        }

        public bool Contains(KeyValuePair<Guid, IList<object>> item)
        {
            var c = sensorDataDictionary as ICollection<KeyValuePair<Guid, IList<object>>>;
            return c.Contains(item);
        }

        public void CopyTo(KeyValuePair<Guid, IList<object>>[] array, int arrayIndex)
        {
            var c = sensorDataDictionary as ICollection<KeyValuePair<Guid, IList<object>>>;
            c.CopyTo(array, arrayIndex);
        }

        public int Count => sensorDataDictionary.Count;

        public bool IsReadOnly => true;

        public bool Remove(KeyValuePair<Guid, IList<object>> item)
        {
            var c = sensorDataDictionary as ICollection<KeyValuePair<Guid, IList<object>>>;
            return c.Remove(item);
        }

        public IEnumerator<KeyValuePair<Guid, IList<object>>> GetEnumerator()
        {
            return sensorDataDictionary as IEnumerator<KeyValuePair<Guid, IList<object>>>;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return sensorDataDictionary as System.Collections.IEnumerator;
        }
    }

    public class SensorReport
    {
        public DateTime TimeStamp => timeStamp;

        public SensorData Values => sensorData;

        public Sensor Source => originator;

        SensorData sensorData;
        Sensor originator;
        DateTime timeStamp = default;

        internal static SensorReport FromNativeReport(Sensor originator, ISensorDataReport iReport)
        {
            iReport.GetTimestamp(out var systemTimeStamp);
            SensorNativeMethods.SystemTimeToFileTime(ref systemTimeStamp, out var ftTimeStamp);
            var lTimeStamp = (((long)ftTimeStamp.dwHighDateTime) << 32) + (long)ftTimeStamp.dwLowDateTime;
            var timeStamp = DateTime.FromFileTime(lTimeStamp);

            var sensorReport = new SensorReport();
            sensorReport.originator = originator;
            sensorReport.timeStamp = timeStamp;
            sensorReport.sensorData = SensorData.FromNativeReport(originator.InternalObject, iReport);

            return sensorReport;
        }
    }

    public delegate void DataReportChangedEventHandler(Sensor sender, EventArgs e);

    public delegate void StateChangedEventHandler(Sensor sender, EventArgs e);

    public class Sensor : ISensorEvents
    {
        static IntPtr IncrementIntPtr(IntPtr source, int increment)
        {
            if (IntPtr.Size == 8)
            {
                var p = source.ToInt64();
                p += increment;
                return new IntPtr(p);
            }
            else if (IntPtr.Size == 4)
            {
                var p = source.ToInt32();
                p += increment;
                return new IntPtr(p);
            }
            else
            {
                throw new InvalidOperationException();
            }
        }

        public event DataReportChangedEventHandler DataReportChanged;

        public event StateChangedEventHandler StateChanged;

        public SensorReport DataReport { get; set; }

        public REFSENSOR_TYPE_ID? SensorId
        {
            get
            {
                if (sensorId == null)
                {
                    var hr = nativeISensor.GetID(out var id);
                    if (hr == 0)
                    {
                        sensorId = id;
                    }
                }
                return sensorId;
            }
        }

        REFSENSOR_TYPE_ID? sensorId;

        public REFSENSOR_TYPE_ID? CategoryId
        {
            get
            {
                if (categoryId == null)
                {
                    var hr = nativeISensor.GetCategory(out var id);
                    if (hr == 0)
                    {
                        categoryId = id;
                    }
                }

                return categoryId;
            }
        }

        REFSENSOR_TYPE_ID? categoryId;

        public REFSENSOR_TYPE_ID? TypeId
        {
            get
            {
                if (typeId == null)
                {
                    var hr = nativeISensor.GetType(out var id);
                    if (hr == 0)
                        typeId = id;
                }

                return typeId;
            }
        }

        REFSENSOR_TYPE_ID? typeId;

        public string FriendlyName
        {
            get
            {
                if (friendlyName == null)
                {
                    var hr = nativeISensor.GetFriendlyName(out var name);
                    if (hr == 0)
                        friendlyName = name;
                }
                return friendlyName;
            }
        }

        string friendlyName;

        public SensorState State
        {
            get
            {
                nativeISensor.GetState(out var state);
                return (SensorState)state;
            }
        }

        public uint ReportInterval
        {
            get
            {
                return (uint)GetProperty(SensorPropertyKeys.SensorPropertyCurrentReportInterval);
            }

            set
            {
                SetProperties(new DataFieldInfo[] { new DataFieldInfo(SensorPropertyKeys.SensorPropertyCurrentReportInterval, value) });
            }
        }

        public uint MinimumReportInterval => (uint)GetProperty(SensorPropertyKeys.SensorPropertyMinReportInterval);

        public string Manufacturer
        {
            get
            {
                if (manufacturer == null)
                {
                    manufacturer = (string)GetProperty(SensorPropertyKeys.SensorPropertyManufacturer);
                }
                return manufacturer;
            }
        }

        string manufacturer;

        public string Model
        {
            get
            {
                if (model == null)
                {
                    model = (string)GetProperty(SensorPropertyKeys.SensorPropertyModel);
                }
                return model;
            }
        }

        string model;

        public string SerialNumber
        {
            get
            {
                if (serialNumber == null)
                {
                    serialNumber = (string)GetProperty(SensorPropertyKeys.SensorPropertySerialNumber);
                }
                return serialNumber;
            }
        }

        string serialNumber;

        public string Description
        {
            get
            {
                if (description == null)
                {
                    description = (string)GetProperty(SensorPropertyKeys.SensorPropertyDescription);
                }

                return description;
            }
        }

        string description;

        public SensorConnectionType? ConnectionType
        {
            get
            {
                if (connectionType == null)
                {
                    connectionType = (SensorConnectionType)GetProperty(SensorPropertyKeys.SensorPropertyConnectionType);
                }
                return connectionType;
            }
        }

        SensorConnectionType? connectionType;

        public string DevicePath
        {
            get
            {
                if (devicePath == null)
                {
                    devicePath = (string)GetProperty(SensorPropertyKeys.SensorPropertyDeviceId);
                }

                return devicePath;
            }
        }

        string devicePath;

        public bool AutoUpdateDataReport
        {
            get
            {
                return IsEventInterestSet(EventInterestTypes.DataUpdated);
            }

            set
            {
                if (value)
                    SetEventInterest(EventInterestTypes.DataUpdated);
                else
                    ClearEventInterest(EventInterestTypes.DataUpdated);
            }
        }

        public bool TryUpdateData()
        {
            var hr = InternalUpdateData();
            return hr == 0;
        }

        public void UpdateData()
        {
            var hr = InternalUpdateData();
            if (hr != 0)
            {
                throw new InvalidOperationException("Sensor not found", Marshal.GetExceptionForHR((int)hr));
            }
        }

        internal int InternalUpdateData()
        {
            var hr = nativeISensor.GetData(out var iReport);
            if (hr == 0)
            {
                try
                {
                    DataReport = SensorReport.FromNativeReport(this, iReport);
                    if (DataReportChanged != null)
                    {
                        DataReportChanged.Invoke(this, EventArgs.Empty);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(iReport);
                }
            }
            return hr;
        }

        public override string ToString()
        {
            return $"Sensor: {SensorId}, TypeId: {TypeId}, CategoryId: {CategoryId}, FriendlyName: {FriendlyName}";
        }

        public object GetProperty(PropertyKey propKey)
        {
            using (var pv = new PropVariant())
            {
                var hr = nativeISensor.GetProperty(ref propKey, pv);
                if (hr != 0)
                {
                    var e = Marshal.GetExceptionForHR((int)hr);
                    throw e;
                }
                return pv.Value;
            }
        }

        public object GetProperty(int propIndex)
        {
            var propKey = new PropertyKey(SensorPropertyKeys.SensorPropertyCommonGuid, propIndex);
            return GetProperty(propKey);
        }

        public IDictionary<PropertyKey, object> GetProperties(PropertyKey[] propKeys)
        {
            if (propKeys == null || propKeys.Length == 0)
            {
                throw new ArgumentException(nameof(propKeys));
            }

            IPortableDeviceKeyCollection keyCollection = new PortableDeviceKeyCollection();
            try
            {
                for (var i = 0; i < propKeys.Length; i++)
                {
                    var propKey = propKeys[i];
                    keyCollection.Add(ref propKey);
                }

                var data = new Dictionary<PropertyKey, object>();
                var hr = nativeISensor.GetProperties(keyCollection, out var valuesCollection);
                if (CoreErrorHelper.Succeeded(hr) && valuesCollection != null)
                {
                    try
                    {
                        uint count = 0;
                        valuesCollection.GetCount(ref count);

                        for (uint i = 0; i < count; i++)
                        {
                            var propKey = default(PropertyKey);
                            using (var propVal = new PropVariant())
                            {
                                valuesCollection.GetAt(i, ref propKey, propVal);
                                data.Add(propKey, propVal.Value);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(valuesCollection);
                        valuesCollection = null;
                    }
                }

                return data;
            }
            finally
            {
                Marshal.ReleaseComObject(keyCollection);
                keyCollection = null;
            }
        }

        public IList<PropertyKey> GetSupportedProperties()
        {
            if (nativeISensor == null)
            {
                throw new InvalidOperationException();
            }

            var list = new List<PropertyKey>();
            var hr = nativeISensor.GetSupportedDataFields(out var collection);
            if (hr == 0)
            {
                try
                {
                    collection.GetCount(out var elements);
                    if (elements == 0)
                    {
                        return null;
                    }

                    for (uint element = 0; element < elements; element++)
                    {
                        hr = collection.GetAt(element, out var key);
                        if (hr == 0)
                        {
                            list.Add(key);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(collection);
                    collection = null;
                }
            }
            return list;
        }

        public object[] GetProperties(params int[] propIndexes)
        {
            if (propIndexes == null || propIndexes.Length == 0)
            {
                throw new ArgumentNullException("propIndexes");
            }

            IPortableDeviceKeyCollection keyCollection = new PortableDeviceKeyCollection();
            try
            {
                var propKeyToIdx = new Dictionary<PropertyKey, int>();

                for (var i = 0; i < propIndexes.Length; i++)
                {
                    var propKey = new PropertyKey(TypeId.Value, propIndexes[i]);
                    keyCollection.Add(ref propKey);
                    propKeyToIdx.Add(propKey, i);
                }

                var data = new object[propIndexes.Length];
                var hr = nativeISensor.GetProperties(keyCollection, out var valuesCollection);
                if (hr == 0)
                {
                    try
                    {
                        if (valuesCollection == null)
                        {
                            return data;
                        }

                        uint count = 0;
                        valuesCollection.GetCount(ref count);

                        for (uint i = 0; i < count; i++)
                        {
                            var propKey = default(PropertyKey);
                            using (var propVal = new PropVariant())
                            {
                                valuesCollection.GetAt(i, ref propKey, propVal);

                                var idx = propKeyToIdx[propKey];
                                data[idx] = propVal.Value;
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(valuesCollection);
                        valuesCollection = null;
                    }
                }
                return data;
            }
            finally
            {
                Marshal.ReleaseComObject(keyCollection);
            }
        }

        public IDictionary<PropertyKey, object> SetProperties(DataFieldInfo[] data)
        {
            if (data == null || data.Length == 0)
            {
                throw new ArgumentException();
            }

            IPortableDeviceValues pdv = new PortableDeviceValues();

            for (var i = 0; i < data.Length; i++)
            {
                var propKey = data[i].Key;
                var value = data[i].Value;
                if (value == null)
                {
                    throw new ArgumentException(nameof(data));
                }

                try
                {
                    // new PropVariant will throw an ArgumentException if the value can
                    // not be converted to an appropriate PropVariant.
                    using (var pv = PropVariant.FromObject(value))
                    {
                        pdv.SetValue(ref propKey, pv);
                    }
                }
                catch (ArgumentException)
                {
                    byte[] buffer;
                    if (value is REFSENSOR_TYPE_ID)
                    {
                        var guid = (REFSENSOR_TYPE_ID)value;
                        pdv.SetGuidValue(ref propKey, ref guid);
                    }
                    else if ((buffer = value as byte[]) != null)
                    {
                        pdv.SetBufferValue(ref propKey, buffer, (uint)buffer.Length);
                    }
                    else
                    {
                        pdv.SetIUnknownValue(ref propKey, value);
                    }
                }
            }

            var results = new Dictionary<PropertyKey, object>();
            var hr = nativeISensor.SetProperties(pdv, out var pdv2);
            if (hr == 0)
            {
                try
                {
                    uint count = 0;
                    pdv2.GetCount(ref count);

                    for (uint i = 0; i < count; i++)
                    {
                        var propKey = default(PropertyKey);
                        using (var propVal = new PropVariant())
                        {
                            pdv2.GetAt(i, ref propKey, propVal);
                            results.Add(propKey, propVal.Value);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(pdv2);
                    pdv2 = null;
                }
            }

            return results;
        }

        protected virtual void Initialize()
        {
        }

        void ISensorEvents.OnStateChanged(ISensor sensor, NativeSensorState state)
        {
            if (StateChanged != null)
            {
                StateChanged.Invoke(this, EventArgs.Empty);
            }
        }

        void ISensorEvents.OnDataUpdated(ISensor sensor, ISensorDataReport newData)
        {
            DataReport = SensorReport.FromNativeReport(this, newData);
            if (DataReportChanged != null)
            {
                DataReportChanged.Invoke(this, EventArgs.Empty);
            }
        }

        void ISensorEvents.OnEvent(ISensor sensor, REFSENSOR_TYPE_ID eventID, ISensorDataReport newData)
        {
        }

        void ISensorEvents.OnLeave(REFSENSOR_TYPE_ID sensorIdArgs)
        {
            SensorManager.OnSensorsChanged(sensorIdArgs, SensorAvailabilityChange.Removal);
        }

        ISensor nativeISensor;

        internal ISensor InternalObject
        {
            get
            {
                return nativeISensor;
            }

            set
            {
                nativeISensor = value;
                SetEventInterest(EventInterestTypes.StateChanged);
                nativeISensor.SetEventSink(this);
                Initialize();
            }
        }

        protected void SetEventInterest(REFSENSOR_TYPE_ID eventType)
        {
            if (nativeISensor == null)
            {
                throw new InvalidOperationException();
            }

            var interestingEvents = GetInterestingEvents();

            if (interestingEvents.Any(g => g == eventType))
            {
                return;
            }

            var interestCount = interestingEvents.Length;

            var newEventInterest = new REFSENSOR_TYPE_ID[interestCount + 1];
            interestingEvents.CopyTo(newEventInterest, 0);
            newEventInterest[interestCount] = eventType;

            var hr = nativeISensor.SetEventInterest(newEventInterest, (uint)(interestCount + 1));
            if (hr != 0)
            {
                throw Marshal.GetExceptionForHR((int)hr);
            }
        }

        protected void ClearEventInterest(REFSENSOR_TYPE_ID eventType)
        {
            if (nativeISensor == null)
            {
                throw new InvalidOperationException();
            }

            if (IsEventInterestSet(eventType))
            {
                var interestingEvents = GetInterestingEvents();
                var interestCount = interestingEvents.Length;

                var newEventInterest = new REFSENSOR_TYPE_ID[interestCount - 1];

                var eventIndex = 0;
                foreach (var g in interestingEvents)
                {
                    if (g != eventType)
                    {
                        newEventInterest[eventIndex] = g;
                        eventIndex++;
                    }
                }

                nativeISensor.SetEventInterest(newEventInterest, (uint)(interestCount - 1));
            }
        }

        protected bool IsEventInterestSet(REFSENSOR_TYPE_ID eventType)
        {
            if (nativeISensor == null)
            {
                throw new InvalidOperationException();
            }

            return GetInterestingEvents()
                .Any(g => g.CompareTo(eventType) == 0);
        }

        REFSENSOR_TYPE_ID[] GetInterestingEvents()
        {
            nativeISensor.GetEventInterest(out var values, out var interestCount);
            var interestingEvents = new REFSENSOR_TYPE_ID[interestCount];
            for (var index = 0; index < interestCount; index++)
            {
                interestingEvents[index] = (REFSENSOR_TYPE_ID)Marshal.PtrToStructure(values, typeof(REFSENSOR_TYPE_ID));
                values = IncrementIntPtr(values, Marshal.SizeOf(typeof(REFSENSOR_TYPE_ID)));
            }
            return interestingEvents;
        }
    }

    public struct DataFieldInfo : IEquatable<DataFieldInfo>
    {
        readonly PropertyKey propKey;
        readonly object value;

        public DataFieldInfo(PropertyKey propKey, object value)
        {
            this.propKey = propKey;
            this.value = value;
        }

        public PropertyKey Key => propKey;

        public object Value => value;

        public override int GetHashCode()
        {
            var valHashCode = value != null ? value.GetHashCode() : 0;
            return propKey.GetHashCode() ^ valHashCode;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            if (!(obj is DataFieldInfo))
            {
                return false;
            }

            var other = (DataFieldInfo)obj;
            return value.Equals(other.value) && propKey.Equals(other.propKey);
        }

        public bool Equals(DataFieldInfo other)
        {
            return value.Equals(other.value) && propKey.Equals(other.propKey);
        }

        public static bool operator ==(DataFieldInfo first, DataFieldInfo second)
        {
            return first.Equals(second);
        }

        public static bool operator !=(DataFieldInfo first, DataFieldInfo second)
        {
            return !first.Equals(second);
        }
    }

    #endregion

    #region COM

    [ComImport]
    [Guid("0AB9DF9B-C4B5-4796-8898-0470706A2E1D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISensorDataReport
    {
        void GetTimestamp(out SystemTime timeStamp);

        void GetSensorValue(
            [In] ref PropertyKey propKey,
            [Out] PropVariant propValue);

        void GetSensorValues(
            [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceKeyCollection keys,
            [Out, MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues values);
    }

    [ComImport]
    [Guid("DADA2357-E0AD-492E-98DB-DD61C53BA353")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPortableDeviceKeyCollection
    {
        void GetCount(out uint pcElems);

        [PreserveSig]
        int GetAt([In] uint dwIndex, out PropertyKey pKey);

        void Add([In] ref PropertyKey key);

        void Clear();

        void RemoveAt([In] uint dwIndex);
    }

    [ComImport]
    [Guid("6848F6F2-3155-4F86-B6F5-263EEEAB3143")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPortableDeviceValues
    {
        void GetCount([In] ref uint pcelt);

        void GetAt([In] uint index, [In, Out] ref PropertyKey pKey, [In, Out] PropVariant pValue);

        void SetValue([In] ref PropertyKey key, [In] PropVariant pValue);

        void GetValue([In] ref PropertyKey key, [Out] PropVariant pValue);

        void SetStringValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.LPWStr)] string value);

        void GetStringValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.LPWStr)] out string pValue);

        void SetUnsignedIntegerValue([In] ref PropertyKey key, [In] uint value);

        void GetUnsignedIntegerValue([In] ref PropertyKey key, out uint pValue);

        void SetSignedIntegerValue([In] ref PropertyKey key, [In] int value);

        void GetSignedIntegerValue([In] ref PropertyKey key, out int pValue);

        void SetUnsignedLargeIntegerValue([In] ref PropertyKey key, [In] ulong value);

        void GetUnsignedLargeIntegerValue([In] ref PropertyKey key, out ulong pValue);

        void SetSignedLargeIntegerValue([In] ref PropertyKey key, [In] long value);

        void GetSignedLargeIntegerValue([In] ref PropertyKey key, out long pValue);

        void SetFloatValue([In] ref PropertyKey key, [In] float value);

        void GetFloatValue([In] ref PropertyKey key, out float pValue);

        void SetErrorValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Error)] int value);

        void GetErrorValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Error)] out int pValue);

        void SetKeyValue([In] ref PropertyKey key, [In] ref PropertyKey value);

        void GetKeyValue([In] ref PropertyKey key, out PropertyKey pValue);

        void SetBoolValue([In] ref PropertyKey key, [In] int value);

        void GetBoolValue([In] ref PropertyKey key, out int pValue);

        void SetIUnknownValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.IUnknown)] object pValue);

        void GetIUnknownValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.IUnknown)] out object ppValue);

        void SetGuidValue([In] ref PropertyKey key, [In] ref REFSENSOR_TYPE_ID value);

        void GetGuidValue([In] ref PropertyKey key, out REFSENSOR_TYPE_ID pValue);

        void SetBufferValue([In] ref PropertyKey key, [In] byte[] pValue, [In] uint cbValue);

        void GetBufferValue([In] ref PropertyKey key, [Out] IntPtr ppValue, out uint pcbValue);

        void SetnativeIPortableDeviceValuesValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValues pValue);

        void GetnativeIPortableDeviceValuesValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues ppValue);

        void SetIPortableDevicePropVariantCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDevicePropVariantCollection pValue);

        void GetIPortableDevicePropVariantCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDevicePropVariantCollection ppValue);

        void SetIPortableDeviceKeyCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceKeyCollection pValue);

        void GetIPortableDeviceKeyCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceKeyCollection ppValue);

        void SetnativeIPortableDeviceValuesCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValuesCollection pValue);

        void GetnativeIPortableDeviceValuesCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValuesCollection ppValue);

        void RemoveValue([In] ref PropertyKey key);

        void CopyValuesFromPropertyStore([In, MarshalAs(UnmanagedType.Interface)] IPropertyStore pStore);

        void CopyValuesToPropertyStore([In, MarshalAs(UnmanagedType.Interface)] IPropertyStore pStore);

        void Clear();
    }

    [ComImport]
    [Guid("6E3F2D79-4E07-48C4-8208-D8C2E5AF4A99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPortableDeviceValuesCollection
    {
        void GetCount([In] ref uint pcElems);

        void GetAt([In] uint dwIndex, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues ppValues);

        void Add([In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValues pValues);

        void Clear();

        void RemoveAt([In] uint dwIndex);
    }

    [ComImport]
    [Guid("89B2E422-4F1B-4316-BCEF-A44AFEA83EB3")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPortableDevicePropVariantCollection
    {
        void GetCount([In] ref uint pcElems);

        void GetAt([In] uint dwIndex, [Out] PropVariant pValue);

        void Add([In] PropVariant pValue);

        void GetType(out ushort pvt);

        void ChangeType([In] ushort vt);

        void Clear();

        void RemoveAt([In] uint dwIndex);
    }

    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        void GetCount(out uint cProps);

        void GetAt([In] uint iProp, out PropertyKey pKey);

        void GetValue([In] ref PropertyKey key, [Out] PropVariant pv);

        void SetValue([In] ref PropertyKey key, [In] PropVariant propvar);

        void Commit();
    }

    [ComImport]
    [Guid("DE2D022D-2480-43BE-97F0-D1FA2CF98F4F")]
    [ClassInterface(ClassInterfaceType.None)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    public class PortableDeviceKeyCollection : IPortableDeviceKeyCollection
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetCount(out uint pcElems);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern int GetAt([In] uint dwIndex, out PropertyKey pKey);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Add([In] ref PropertyKey key);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Clear();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void RemoveAt([In] uint dwIndex);
    }

    [ComImport]
    [Guid("0C15D503-D017-47CE-9016-7B3F978721CC")]
    [ClassInterface(ClassInterfaceType.None)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    public class PortableDeviceValues : IPortableDeviceValues
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetCount([In] ref uint pcelt);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetAt([In] uint index, [In, Out] ref PropertyKey pKey, [In, Out] PropVariant pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetValue([In] ref PropertyKey key, [In] PropVariant pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetValue([In] ref PropertyKey key, [Out] PropVariant pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetStringValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.LPWStr)] string value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetStringValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.LPWStr)] out string pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetUnsignedIntegerValue([In] ref PropertyKey key, [In] uint value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetUnsignedIntegerValue([In] ref PropertyKey key, out uint pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetSignedIntegerValue([In] ref PropertyKey key, [In] int value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetSignedIntegerValue([In] ref PropertyKey key, out int pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetUnsignedLargeIntegerValue([In] ref PropertyKey key, [In] ulong value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetUnsignedLargeIntegerValue([In] ref PropertyKey key, out ulong pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetSignedLargeIntegerValue([In] ref PropertyKey key, [In] long value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetSignedLargeIntegerValue([In] ref PropertyKey key, out long pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetFloatValue([In] ref PropertyKey key, [In] float value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetFloatValue([In] ref PropertyKey key, out float pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetErrorValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Error)] int value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetErrorValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Error)] out int pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetKeyValue([In] ref PropertyKey key, [In] ref PropertyKey value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetKeyValue([In] ref PropertyKey key, out PropertyKey pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetBoolValue([In] ref PropertyKey key, [In] int value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetBoolValue([In] ref PropertyKey key, out int pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetIUnknownValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.IUnknown)] object pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetIUnknownValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.IUnknown)] out object ppValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetGuidValue([In] ref PropertyKey key, [In] ref REFSENSOR_TYPE_ID value);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetGuidValue([In] ref PropertyKey key, out REFSENSOR_TYPE_ID pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetBufferValue([In] ref PropertyKey key, [In] byte[] pValue, [In] uint cbValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetBufferValue([In] ref PropertyKey key, [Out] IntPtr ppValue, out uint pcbValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetnativeIPortableDeviceValuesValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValues pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetnativeIPortableDeviceValuesValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues ppValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetIPortableDevicePropVariantCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDevicePropVariantCollection pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetIPortableDevicePropVariantCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDevicePropVariantCollection ppValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetIPortableDeviceKeyCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceKeyCollection pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetIPortableDeviceKeyCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceKeyCollection ppValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetnativeIPortableDeviceValuesCollectionValue([In] ref PropertyKey key, [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValuesCollection pValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetnativeIPortableDeviceValuesCollectionValue([In] ref PropertyKey key, [MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValuesCollection ppValue);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void RemoveValue([In] ref PropertyKey key);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void CopyValuesFromPropertyStore([In, MarshalAs(UnmanagedType.Interface)] IPropertyStore pStore);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void CopyValuesToPropertyStore([In, MarshalAs(UnmanagedType.Interface)] IPropertyStore pStore);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Clear();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("23571E11-E545-4DD8-A337-B89BF44B10DF")]
    public interface ISensorCollection
    {
        void GetAt([In] uint ulIndex, [Out] out ISensor ppSensor);

        void GetCount([Out] out uint pCount);

        void Add([In, MarshalAs(UnmanagedType.IUnknown)] ISensor pSensor);

        void Remove([In] ISensor pSensor);

        void RemoveByID([In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID sensorID);

        void Clear();
    }

    [ComImport]
    [Guid("79C43ADB-A429-469F-AA39-2F2B74B75937")]
    [ClassInterface(ClassInterfaceType.None)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    public class SensorCollection : ISensorCollection
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetAt([In] uint index, [MarshalAs(UnmanagedType.Interface)] out ISensor ppSensor);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void GetCount(out uint pCount);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Add([MarshalAs(UnmanagedType.Interface)] ISensor pSensor);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Remove([MarshalAs(UnmanagedType.Interface)] ISensor pSensor);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void RemoveByID([In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID sensorId);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void Clear();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("9B3B0B86-266A-4AAD-B21F-FDE5501001B7")]
    public interface ISensorManagerEvents
    {
        void OnSensorEnter(
            [In, MarshalAs(UnmanagedType.Interface)] ISensor pSensor,
            [In, MarshalAs(UnmanagedType.U4)] NativeSensorState state);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("5D8DCC91-4641-47E7-B7C3-B74F48A6C391")]
    public interface ISensorEvents
    {
        void OnStateChanged(
            [In, MarshalAs(UnmanagedType.Interface)] ISensor sensor,
            [In, MarshalAs(UnmanagedType.U4)] NativeSensorState state);

        void OnDataUpdated(
            [In, MarshalAs(UnmanagedType.Interface)] ISensor sensor,
            [In, MarshalAs(UnmanagedType.Interface)] ISensorDataReport newData);

        void OnEvent(
            [In, MarshalAs(UnmanagedType.Interface)] ISensor sensor,
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID eventID,
            [In, MarshalAs(UnmanagedType.Interface)] ISensorDataReport newData);

        void OnLeave([In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID sensorID);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("5FA08F80-2657-458E-AF75-46F73FA6AC5C")]
    public interface ISensor
    {
        [PreserveSig]
        int GetID(out REFSENSOR_TYPE_ID id);

        [PreserveSig]
        int GetCategory(out REFSENSOR_TYPE_ID sensorCategory);

        [PreserveSig]
        int GetType(out REFSENSOR_TYPE_ID sensorType);

        [PreserveSig]
        int GetFriendlyName([Out, MarshalAs(UnmanagedType.BStr)] out string friendlyName);

        [PreserveSig]
        int GetProperty([In] ref PropertyKey key, [Out] PropVariant property);

        [PreserveSig]
        int GetProperties(
            [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceKeyCollection keys,
            [Out, MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues properties);

        [PreserveSig]
        int GetSupportedDataFields(
            [Out, MarshalAs(UnmanagedType.Interface)] out IPortableDeviceKeyCollection dataFields);

        [PreserveSig]
        int SetProperties(
            [In, MarshalAs(UnmanagedType.Interface)] IPortableDeviceValues properties,
            [Out, MarshalAs(UnmanagedType.Interface)] out IPortableDeviceValues results);

        void SupportsDataField(
            [In] PropertyKey key,
            [Out, MarshalAs(UnmanagedType.VariantBool)] out bool isSupported);

        void GetState([Out, MarshalAs(UnmanagedType.U4)] out NativeSensorState state);

        [PreserveSig]
        int GetData([Out, MarshalAs(UnmanagedType.Interface)] out ISensorDataReport dataReport);

        void SupportsEvent(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID eventGuid,
            [Out, MarshalAs(UnmanagedType.VariantBool)] out bool isSupported);

        void GetEventInterest(
            out IntPtr pValues, [Out] out uint count);

        [PreserveSig]
        int SetEventInterest([In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] REFSENSOR_TYPE_ID[] pValues, [In] uint count);

        void SetEventSink([In, MarshalAs(UnmanagedType.Interface)] ISensorEvents pEvents);
    }

    [ComImport]
    [Guid("77A1C827-FCD2-4689-8915-9D613CC5FA3E")]
    [ClassInterface(ClassInterfaceType.None)]
    public class NativeSensorManager : InativeISensorManager
    {
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern int GetSensorsByCategory(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_CATEGORY_ID sensorCategory,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensorCollection ppSensorsFound);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern int GetSensorsByType(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID sensorType,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensorCollection ppSensorsFound);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern int GetSensorByID(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_ID sensorID,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensor ppSensor);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void SetEventSink(
            [Out, MarshalAs(UnmanagedType.Interface)] ISensorManagerEvents pEvents);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public virtual extern void RequestPermissions(
            IntPtr hParent,
            [In, MarshalAs(UnmanagedType.Interface)] ISensorCollection pSensors,
            [In, MarshalAs(UnmanagedType.Bool)] bool modal);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("BD77DB67-45A8-42DC-8D00-6DCF15F8377A")]
    public interface InativeISensorManager
    {
        [PreserveSig]
        int GetSensorsByCategory(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_CATEGORY_ID sensorCategory,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensorCollection ppSensorsFound);

        [PreserveSig]
        int GetSensorsByType(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_TYPE_ID sensorType,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensorCollection ppSensorsFound);

        [PreserveSig]
        int GetSensorByID(
            [In, MarshalAs(UnmanagedType.LPStruct)] REFSENSOR_ID sensorID,
            [Out, MarshalAs(UnmanagedType.Interface)] out ISensor ppSensor);

        void SetEventSink(
            [In, MarshalAs(UnmanagedType.Interface)] ISensorManagerEvents pEvents);

        void RequestPermissions(
            [In] IntPtr parent,
            [In, MarshalAs(UnmanagedType.Interface)] ISensorCollection sensors,
            [In, MarshalAs(UnmanagedType.Bool)] bool modal);
    }

    #endregion

    public partial struct BarometerData
    {
        public BarometerData(SensorReport report)
        {
            if (report == null)
            {
                throw new ArgumentNullException("report");
            }

            if (report.Values != null &&
                report.Values.ContainsKey(SensorPropertyKeys.SensorDataTypeAtmosphericPressureBar.FormatId))
            {
                PressureInHectopascals =
                    ((float)report.Values[SensorPropertyKeys.SensorDataTypeAtmosphericPressureBar.FormatId][0]) / 1000f; // Convert bar to hPa
            }
            else
            {
                PressureInHectopascals = -1;
            }
        }
    }

    public static partial class Barometer
    {
        [SensorDescription("0E903829-FF8A-4A93-97DF-3DCBDE402288")]
        public class BarometerSensor : Sensor
        {
            public BarometerData CurrentPressure => new BarometerData(DataReport);
        }

        internal static bool IsSupported
        {
            get
            {
                var x = SensorManager.GetSensorsByTypeId<BarometerSensor>());
                return x.Any();
            }
        }

        internal static void PlatformStart(SensorSpeed sensorSpeed)
        {
            var barometerList = SensorManager.GetSensorsByTypeId<BarometerSensor>().First();
            barometerList.DataReportChanged += BarometerList_DataReportChanged;
        }

        static void BarometerList_DataReportChanged(Sensor sender, EventArgs e)
        {
            OnChanged(((BarometerSensor)sender).CurrentPressure);
        }

        internal static void PlatformStop()
        {
            var barometerList = SensorManager.GetSensorsByTypeId<BarometerSensor>().First();
            barometerList.DataReportChanged -= BarometerList_DataReportChanged;
        }
    }
}
