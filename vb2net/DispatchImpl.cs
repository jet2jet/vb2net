using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace vb2net
{
	[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
	interface IDispatch
	{
		int GetTypeInfoCount();
		[return: MarshalAs(UnmanagedType.Interface)]
		object GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
		[PreserveSig]
		int GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] IntPtr[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames, [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] int[] rgDispId);
		[PreserveSig]
		int Invoke([In] int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In] short wFlags, [In] IntPtr pDispParams, [In] IntPtr pVarResult, [In] IntPtr pExcepInfo, [In] IntPtr puArgErr);
	}

	struct MemberData
	{
		public string name;
		public MemberInfo info;
	}

	class DispatchImpl : IDispatch
	{
		public readonly object wrapped;
		private readonly List<MemberData> members;

		private const short VT_DISPATCH = 9;
		private const short VT_UNKNOWN = 13;

		private const int DISPID_UNKNOWN = -1;
		private const int DISPID_PROPERTYPUT = -3;
		private const int DISPID_NEWENUM = -4;

		private const int INVOKE_FUNC = 1;
		private const int INVOKE_PROPERTYGET = 2;
		private const int INVOKE_PROPERTYPUT = 4;
		private const int INVOKE_PROPERTYPUTREF = 8;

		public DispatchImpl(object wrapped)
		{
			this.wrapped = wrapped;
			members = new List<MemberData>();
			var type = wrapped.GetType();
			var nameCounts = new Dictionary<string, int>();
			foreach (var member in type.GetMembers())
			{
				var name = member.Name.ToLower();
				if (nameCounts.TryGetValue(name, out int count))
				{
					++count;
					nameCounts[name] = count;
					name += "_" + count.ToString();
				}
				else
				{
					nameCounts[name] = 1;
				}
				members.Add(new MemberData { name = name, info = member });
			}
		}

		int IDispatch.GetTypeInfoCount()
		{
			return 0;
		}

		object IDispatch.GetTypeInfo(int iTInfo, int lcid)
		{
			throw new NotImplementedException();
		}

		int IDispatch.GetIDsOfNames(ref Guid riid, IntPtr[] rgszNames, int cNames, int lcid, int[] rgDispId)
		{
			if (cNames < 1)
				return HResults.DISP_E_MEMBERNOTFOUND;
			var name = Marshal.PtrToStringUni(rgszNames[0])?.ToLower();
			for (var j = 0; j < members.Count; ++j)
			{
				var member = members[j];
				if (member.name == name)
				{
					rgDispId[0] = j + 1;
					ParameterInfo[] piArray;
					var method = member.info as MethodInfo;
					if (method != null)
					{
						piArray = method.GetParameters();
					}
					else
					{
						var prop = member.info as PropertyInfo;
						if (prop != null)
						{
							piArray = prop.GetIndexParameters();
						}
						else
						{
							piArray = Array.Empty<ParameterInfo>();
						}
					}
					var paramNotFound = false;
					for (var i = 1; i < cNames; ++i)
					{
						var paramName = Marshal.PtrToStringUni(rgszNames[i])?.ToLower();
						var found = false;
						for (var k = 0; k < piArray.Length; ++k)
						{
							var param = piArray[k];
							if (param.Name?.ToLower() == paramName)
							{
								rgDispId[i] = k + 1;
								found = true;
								break;
							}
						}
						if (!found)
						{
							paramNotFound = true;
							rgDispId[i] = DISPID_UNKNOWN;
						}
					}
					if (paramNotFound)
					{
						return HResults.DISP_E_UNKNOWNNAME;
					}
					return 0;
				}
			}
			return HResults.DISP_E_MEMBERNOTFOUND;
		}

		int IDispatch.Invoke(int dispIdMember, ref Guid riid, int lcid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr)
		{
			MemberData member;
			if (dispIdMember == DISPID_NEWENUM)
			{
				var enumerable = wrapped as System.Collections.IEnumerable;
				if (enumerable == null)
				{
					return HResults.DISP_E_MEMBERNOTFOUND;
				}
				var enumerator = new EnumeratorWrapper(enumerable.GetEnumerator());
				var pUnk = Marshal.GetIUnknownForObject(enumerator);
				Marshal.GetNativeVariantForObject(pUnk, pVarResult);
				short[] vt = { VT_UNKNOWN };
				Marshal.Copy(vt, 0, pVarResult, 1);
				return 0;
			}
			else if (dispIdMember < 1 || dispIdMember > members.Count)
				return HResults.DISP_E_MEMBERNOTFOUND;
			member = members[dispIdMember - 1];
			var method = member.info as MethodInfo;
			var property = member.info as PropertyInfo;
			var dispParams = Marshal.PtrToStructure<DISPPARAMS>(pDispParams)!;
			ParameterInfo[]? piArray = null;
			if (method != null)
			{
				piArray = method.GetParameters();
				if ((wFlags & INVOKE_FUNC) == 0)
				{
					return HResults.DISP_E_MEMBERNOTFOUND;
				}
				var requiredParamCount = piArray.Length;
				for (var i = requiredParamCount - 1; i >= 0; i--)
				{
					if (piArray[i].IsOptional)
						--requiredParamCount;
				}
				if (dispParams.cArgs < requiredParamCount || dispParams.cArgs > piArray.Length)
				{
					return HResults.DISP_E_BADPARAMCOUNT;
				}
			}
			else if (property != null)
			{
				piArray = property.GetIndexParameters();
				if ((wFlags & (INVOKE_PROPERTYGET | INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF)) == 0)
				{
					return HResults.DISP_E_MEMBERNOTFOUND;
				}
				var paramCount = piArray.Length;
				if ((wFlags & (INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF)) != 0)
					++paramCount;
				var requiredParamCount = paramCount;
				for (var i = piArray.Length - 1; i >= 0; i--)
				{
					if (piArray[i].IsOptional)
						--requiredParamCount;
				}
				if (dispParams.cArgs < requiredParamCount || dispParams.cArgs > paramCount)
				{
					return HResults.DISP_E_BADPARAMCOUNT;
				}
			}
			List<object?> args = new(piArray?.Length ?? 0);
			if (piArray != null && dispParams.cArgs > 0)
			{
				for (var i = 0; i < piArray.Length; ++i)
					args.Add(null);
				var dispidNamedArgs = new int[dispParams.cNamedArgs];
				if (dispParams.cNamedArgs != 0)
				{
					Marshal.Copy(dispParams.rgdispidNamedArgs, dispidNamedArgs, 0, dispParams.cNamedArgs);
				}
				var iNamedArgs = dispParams.cNamedArgs - 1;
				var vargs = Marshal.GetObjectsForNativeVariants(dispParams.rgvarg, dispParams.cArgs);
				var errNamedArg = -1;
				for (var i = dispParams.cArgs - 1; i >= 0; --i)
				{
					var newIndex = i;
					if (iNamedArgs >= 0)
					{
						var dispid = dispidNamedArgs[iNamedArgs];
						if (dispid >= 1 && dispid <= piArray.Length)
						{
							newIndex = dispid - 1;
						}
						else if (dispid == DISPID_PROPERTYPUT)
						{
							newIndex = i;
						}
						else
						{
							errNamedArg = iNamedArgs;
						}
						--iNamedArgs;
					}
					var arg = UnwrapObject(vargs[i], piArray[newIndex].ParameterType);
					args[newIndex] = arg;
				}
				if (errNamedArg >= 0)
				{
					if (puArgErr != IntPtr.Zero)
						Marshal.Copy(new int[] { errNamedArg }, 0, puArgErr, 1);
					return HResults.DISP_E_PARAMNOTFOUND;
				}
			}
			try
			{
				if (method != null)
				{
					var result = method.Invoke(wrapped, args.ToArray());
					OutputToVariant(result, pVarResult, method.ReturnType);
				}
				else if (property != null)
				{
					if ((wFlags & (INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF)) != 0)
					{
						property.SetValue(wrapped, args[0], args.GetRange(1, args.Count - 1).ToArray());
					}
					else
					{
						var result = property.GetValue(wrapped, args.ToArray());
						OutputToVariant(result, pVarResult, property.PropertyType);
					}
				}
				return 0;
			}
			catch (Exception e)
			{
				return ExceptionToExcepInfo(e, pExcepInfo);
			}
		}

		private static object? UnwrapObject(object? value, Type? typeHint)
		{
			if (value == null)
			{
				return null;
			}
			if (IsNativeValue(value.GetType()))
			{
				if (typeHint != null)
				{
					var o = Convert.ChangeType(value, typeHint);
					if (o != null)
						return o;
				}
				return value;
			}
			var arr = value as object?[];
			if (arr != null)
			{
				var elemType = typeHint?.GetElementType();
				var r = Array.ConvertAll(arr, (o) => UnwrapObject(o, elemType));
				if (elemType != null)
				{
					var newArray = Array.CreateInstance(elemType, r.Length);
					Array.Copy(r, newArray, newArray.Length);
					return newArray;
				}
				return r;
			}
			var enumeratorWrapper = value as EnumeratorWrapper;
			if (enumeratorWrapper != null)
			{
				return enumeratorWrapper.wrapped;
			}
			var disp = value as DispatchImpl;
			if (disp != null)
			{
				return disp.wrapped;
			}
			return null;
		}

		public static void OutputToVariant(object? value, IntPtr outVariant, Type typeHint)
		{
			if (IsNativeValue(typeHint))
				Marshal.GetNativeVariantForObject(value, outVariant);
			else if (value == null)
			{
				Marshal.GetNativeVariantForObject(null, outVariant);
				short[] vt = { VT_DISPATCH };
				Marshal.Copy(vt, 0, outVariant, 1);
			}
			else if (typeHint == typeof(IntPtr) || typeHint == typeof(UIntPtr))
				Marshal.GetNativeVariantForObject((long)value, outVariant);
			else
			{
				var disp = new DispatchImpl(value) as IDispatch;
				Marshal.GetNativeVariantForObject(Marshal.GetComInterfaceForObject(disp, typeof(IDispatch)), outVariant);
				short[] vt = { VT_DISPATCH };
				Marshal.Copy(vt, 0, outVariant, 1);
			}
		}

		private static bool IsNativeValue(Type type)
		{
			return type == typeof(bool) ||
				type == typeof(char) ||
				type == typeof(sbyte) ||
				type == typeof(short) ||
				type == typeof(int) ||
				type == typeof(long) ||
				type == typeof(byte) ||
				type == typeof(ushort) ||
				type == typeof(uint) ||
				type == typeof(ulong) ||
				type == typeof(float) ||
				type == typeof(double) ||
				type == typeof(decimal) ||
				type == typeof(string) ||
				type == typeof(DateTime);
		}

		private static int ExceptionToExcepInfo(Exception e, IntPtr pExcepInfo)
		{
			while (e.InnerException != null)
			{
				e = e.InnerException;
			}
			var hr = Marshal.GetHRForException(e);
			if (pExcepInfo != IntPtr.Zero)
			{
				EXCEPINFO ex = new()
				{
					bstrDescription = e.Message,
					bstrSource = e.Source ?? "",
					bstrHelpFile = e.HelpLink ?? "",
					dwHelpContext = 0,
					scode = hr
				};
				Marshal.StructureToPtr(ex, pExcepInfo, false);
			}
			return hr;
		}
	}
}
