using System.Runtime.InteropServices;

namespace vb2net
{
	// There is System.Runtime.InteropServices.ComTypes.IEnumVARIANT, but
	// ComType's IEnumVARIANT cannot return IDispatch object at Next method, so
	// alternative IEnumVARIANT is necessary to return IDispatch object
	[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020404-0000-0000-C000-000000000046")]
	interface IEnumVARIANT
	{
		[PreserveSig]
		[return: MarshalAs(UnmanagedType.I4)]
		int Next([In, MarshalAs(UnmanagedType.U4)] int celt, [In] IntPtr rgVar, [In] IntPtr pCeltFetched);
		void Skip([In, MarshalAs(UnmanagedType.U4)] int celt);
		void Reset();
		IntPtr Clone();
	}

	[StructLayout(LayoutKind.Sequential)]
	struct DummyVariant
	{
		public short vt;
		public short wReserved1;
		public short wReserved2;
		public short wReserved3;
		public IntPtr val1;
		public IntPtr val2;

		public static int SizeofVariant = Marshal.SizeOf(typeof(DummyVariant));
	}

	class EnumeratorWrapper : IEnumVARIANT
	{
		public readonly System.Collections.IEnumerator wrapped;
		public EnumeratorWrapper(System.Collections.IEnumerator wrapped)
		{ this.wrapped = wrapped; }

		[PreserveSig]
		int IEnumVARIANT.Next(int celt, IntPtr rgVar, IntPtr pCeltFetched)
		{
			try
			{
				var list = new List<object?>(celt);
				while (celt-- > 0)
				{
					if (!wrapped.MoveNext())
						break;
					list.Add(wrapped.Current);
				}
				foreach (var o in list)
				{
					DispatchImpl.OutputToVariant(o, rgVar, o != null ? o.GetType() : typeof(object));
					rgVar = (IntPtr)((long)rgVar + DummyVariant.SizeofVariant);
				}
				if (pCeltFetched != IntPtr.Zero)
				{
					Marshal.Copy(new int[] { list.Count }, 0, pCeltFetched, 1);
				}
				if (list.Count == 0)
					return 1;
				return 0;
			}
			catch (Exception e)
			{
				return Marshal.GetHRForException(e);
			}
		}

		void IEnumVARIANT.Skip(int celt)
		{
			while (celt-- > 0)
				wrapped.MoveNext();
		}

		void IEnumVARIANT.Reset()
		{
			wrapped.Reset();
		}

		IntPtr IEnumVARIANT.Clone()
		{
			var newObj = new EnumeratorWrapper(wrapped);
			var instance = newObj as IEnumVARIANT;
			return Marshal.GetComInterfaceForObject(instance, typeof(IEnumVARIANT));
		}
	}
}
