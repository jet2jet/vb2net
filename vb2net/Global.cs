using System.Reflection;
using System.Runtime.InteropServices;

namespace vb2net
{
	public class Global
    {
		// C++ signature: HRESULT LoadAssembly(LPCWSTR assemblyNamePtr, IDispatch** pResult);
		[UnmanagedCallersOnly]
		public static int LoadAssembly(IntPtr assemblyNamePtr, IntPtr pResult)
		{
			if (assemblyNamePtr == IntPtr.Zero || pResult == IntPtr.Zero)
				return HResults.E_POINTER;
			try
			{
				var assemblyName = Marshal.PtrToStringUni(assemblyNamePtr);
				if (assemblyName == null)
				{
					return HResults.E_POINTER;
				}
				var assembly = Assembly.Load(assemblyName);
				var dispImpl = new DispatchImpl(assembly);
				var disp = dispImpl as IDispatch;
				var dispPtr = Marshal.GetComInterfaceForObject(disp, typeof(IDispatch));
				Marshal.Copy(new IntPtr[] { dispPtr }, 0, pResult, 1);
				return 0;
			}
			catch (Exception e)
			{
				return Marshal.GetHRForException(e);
			}
		}

		// C++ signature: HRESULT LoadAssemblyFromFile(LPCWSTR fileNamePtr, IDispatch** pResult);
		[UnmanagedCallersOnly]
		public static int LoadAssemblyFromFile(IntPtr fileNamePtr, IntPtr pResult)
		{
			if (fileNamePtr == IntPtr.Zero || pResult == IntPtr.Zero)
				return HResults.E_POINTER;
			try
			{
				var fileName = Marshal.PtrToStringUni(fileNamePtr);
				if (fileName == null)
				{
					return HResults.E_POINTER;
				}
				var assembly = Assembly.LoadFile(fileName);
				var dispImpl = new DispatchImpl(assembly);
				var disp = dispImpl as IDispatch;
				var dispPtr = Marshal.GetComInterfaceForObject(disp, typeof(IDispatch));
				Marshal.Copy(new IntPtr[] { dispPtr }, 0, pResult, 1);
				return 0;
			}
			catch (Exception e)
			{
				return Marshal.GetHRForException(e);
			}
		}
	}
}
