package com4e.core.impl;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.swt.SWT;
import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.DISPPARAMS;
import org.eclipse.swt.internal.ole.win32.EXCEPINFO;
import org.eclipse.swt.internal.ole.win32.GUID;
import org.eclipse.swt.internal.ole.win32.IDispatch;
import org.eclipse.swt.internal.ole.win32.IEnumVARIANT;
import org.eclipse.swt.internal.ole.win32.IUnknown;
import org.eclipse.swt.internal.ole.win32.VARIANT;
import org.eclipse.swt.internal.win32.OS;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.Variant;

import com4e.core.api.ComIDispatch;
import com4e.core.api.ComIUnknown;

@SuppressWarnings("restriction")
public class VariantUtils {
	public static void setVariantData(Variant v, int /* long */pData) {
		if (pData == 0)
			OLE.error(OLE.ERROR_INVALID_ARGUMENT);

		// TODO - use VARIANT structure
		short[] dataType = new short[1];
		COM.MoveMemory(dataType, pData, 2);

		setPrivateField(v, "type", dataType[0]);

		if ((v.getType() & COM.VT_BYREF) == COM.VT_BYREF) {
			int /* long */[] newByRefPtr = new int /* long */[1];
			OS.MoveMemory(newByRefPtr, pData + 8, OS.PTR_SIZEOF);
			v.setByRef(newByRefPtr[0]);
			return;
		}

		switch (v.getType()) {
		case COM.VT_EMPTY:
		case COM.VT_NULL:
			break;
		case COM.VT_BOOL:
			short[] newBooleanData = new short[1];
			COM.MoveMemory(newBooleanData, pData + 8, 2);
			setPrivateField(v, "booleanData",
					(newBooleanData[0] != COM.VARIANT_FALSE));
			break;
		case COM.VT_I1:
			byte[] newByteData = new byte[1];
			COM.MoveMemory(newByteData, pData + 8, 1);
			setPrivateField(v, "byteData", newByteData[0]);
			break;
		case COM.VT_I2:
			short[] newShortData = new short[1];
			COM.MoveMemory(newShortData, pData + 8, 2);
			setPrivateField(v, "shortData", newShortData[0]);
			break;
		case COM.VT_UI2:
			char[] newCharData = new char[1];
			COM.MoveMemory(newCharData, pData + 8, 2);
			setPrivateField(v, "charData", newCharData[0]);
			break;
		case COM.VT_I4:
			int[] newIntData = new int[1];
			OS.MoveMemory(newIntData, pData + 8, 4);
			setPrivateField(v, "intData", newIntData[0]);
			break;
		case COM.VT_I8:
			long[] newLongData = new long[1];
			OS.MoveMemory(newLongData, pData + 8, 8);
			setPrivateField(v, "longData", newLongData[0]);
			break;
		case COM.VT_R4:
			float[] newFloatData = new float[1];
			COM.MoveMemory(newFloatData, pData + 8, 4);
			setPrivateField(v, "floatData", newFloatData[0]);
			break;
		case COM.VT_R8:
			double[] newDoubleData = new double[1];
			COM.MoveMemory(newDoubleData, pData + 8, 8);
			setPrivateField(v, "doubleData", newDoubleData[0]);
			break;
		case COM.VT_DISPATCH: {
			int /* long */[] ppvObject = new int /* long */[1];
			OS.MoveMemory(ppvObject, pData + 8, OS.PTR_SIZEOF);
			if (ppvObject[0] == 0) {
				setPrivateField(v, "type", COM.VT_EMPTY);
				break;
			}
			setPrivateField(v, "dispatchData", new IDispatch(ppvObject[0]));
			v.getDispatch().AddRef();
			break;
		}
		case COM.VT_UNKNOWN: {
			int /* long */[] ppvObject = new int /* long */[1];
			OS.MoveMemory(ppvObject, pData + 8, OS.PTR_SIZEOF);
			if (ppvObject[0] == 0) {
				setPrivateField(v, "type", COM.VT_EMPTY);
				break;
			}
			setPrivateField(v, "unknownData", new IUnknown(ppvObject[0]));
			v.getUnknown().AddRef();
			break;
		}
		case COM.VT_BSTR:
			// get the address of the memory in which the string resides
			int /* long */[] hMem = new int /* long */[1];
			OS.MoveMemory(hMem, pData + 8, OS.PTR_SIZEOF);
			if (hMem[0] == 0) {
				setPrivateField(v, "type", COM.VT_EMPTY);
				break;
			}
			// Get the size of the string from the OS - the size is expressed in
			// number
			// of bytes - each unicode character is 2 bytes.
			int size = COM.SysStringByteLen(hMem[0]);
			if (size > 0) {
				// get the unicode character array from the global memory and
				// create a String
				char[] buffer = new char[(size + 1) / 2]; // add one to avoid
															// rounding errors
				COM.MoveMemory(buffer, hMem[0], size);
				setPrivateField(v, "stringData", new String(buffer));
			} else {
				setPrivateField(v, "stringData", ""); //$NON-NLS-1$
			}
			break;

		default:
			// try coercing it into one of the known forms
			int /* long */newPData = OS.GlobalAlloc(OS.GMEM_FIXED
					| OS.GMEM_ZEROINIT, Variant.sizeof);
			if (COM.VariantChangeType(newPData, pData, (short) 0, COM.VT_R4) == COM.S_OK) {
				setVariantData(v, newPData);
			} else if (COM.VariantChangeType(newPData, pData, (short) 0,
					COM.VT_I4) == COM.S_OK) {
				setVariantData(v, newPData);
			} else if (COM.VariantChangeType(newPData, pData, (short) 0,
					COM.VT_BSTR) == COM.S_OK) {
				setVariantData(v, newPData);
			}
			COM.VariantClear(newPData);
			OS.GlobalFree(newPData);
			break;
		}
	}

	public static void getVariantData(Variant variant, int /* long */pData) {
		if (pData == 0)
			OLE.error(OLE.ERROR_OUT_OF_MEMORY);

		COM.VariantInit(pData);

		if ((variant.getType() & COM.VT_BYREF) == COM.VT_BYREF) {
			// TODO - use VARIANT structure
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8,
					new int /* long */[] { variant.getByRef() }, OS.PTR_SIZEOF);
			return;
		}

		switch (variant.getType()) {
		case COM.VT_EMPTY:
		case COM.VT_NULL:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			break;
		case COM.VT_BOOL:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8,
					new short[] { (variant.getBoolean()) ? COM.VARIANT_TRUE
							: COM.VARIANT_FALSE }, 2);
			break;
		case COM.VT_I1:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new byte[] { variant.getByte() }, 1);
			break;
		case COM.VT_I2:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new short[] { variant.getShort() }, 2);
			break;
		case COM.VT_UI2:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new char[] { variant.getChar() }, 2);
			break;
		case COM.VT_I4:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new int[] { variant.getInt() }, 4);
			break;
		case COM.VT_I8:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new long[] { variant.getLong() }, 8);
			break;
		case COM.VT_R4:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new float[] { variant.getFloat() }, 4);
			break;
		case COM.VT_R8:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new double[] { variant.getDouble() }, 8);
			break;
		case COM.VT_DISPATCH:
			variant.getDispatch().AddRef();
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new int /* long */[] { variant
					.getDispatch().getAddress() }, OS.PTR_SIZEOF);
			break;
		case COM.VT_UNKNOWN:
			variant.getUnknown().AddRef();
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			COM.MoveMemory(pData + 8, new int /* long */[] { variant.getUnknown()
					.getAddress() }, OS.PTR_SIZEOF);
			break;
		case COM.VT_BSTR:
			COM.MoveMemory(pData, new short[] { variant.getType() }, 2);
			char[] data = (variant.getString()).toCharArray();
			int /* long */ptr = COM.SysAllocString(data);
			COM
					.MoveMemory(pData + 8, new int /* long */[] { ptr },
							OS.PTR_SIZEOF);
			break;

		default:
			OLE.error(SWT.ERROR_NOT_IMPLEMENTED);
		}
	}

	private static void setPrivateField(Variant v, String fieldName,
			Object value) {
		// Urggh...
		try {
			Field fType = v.getClass().getDeclaredField(fieldName);
			fType.setAccessible(true);
			fType.set(v, value);
		} catch (NoSuchFieldException e) {
			throw new RuntimeException(e);
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e);
		}
	}

	public static Object getVariantValue(Variant v) {

		switch (v.getType()) {
		case COM.VT_EMPTY:
		case COM.VT_NULL:
			return null;
		case COM.VT_BOOL:
			return v.getBoolean();
		case COM.VT_I1:
			return v.getByte();
		case COM.VT_I2:
			return v.getShort();
		case COM.VT_UI2:
			return v.getChar();
		case COM.VT_I4:
			return v.getInt();
		case COM.VT_I8:
			return v.getLong();
		case COM.VT_R4:
			return v.getFloat();
		case COM.VT_R8:
			return v.getDouble();
		case COM.VT_DISPATCH:
			v.getDispatch().AddRef(); // TODO ?
			return v.getDispatch();
		case COM.VT_UNKNOWN:
			v.getUnknown().AddRef(); // TODO ?
			return v.getUnknown().getAddress();
		case COM.VT_BSTR:
			return v.getString();

		default:
			throw new UnsupportedOperationException(
					"Type de variant non supporté : " + v.getType());
		}
	}

	public static final int DISPID_NEWENUM = -4;

	public static Object[] getArrayvalues(IDispatch dispatch) {

		// Variant vLength = getProperty(dispatch, "length");
		// int length = vLength.getInt();

		List<Object> values = new ArrayList<Object>();

		Variant vUnknown = getProperty(dispatch, DISPID_NEWENUM);
		IUnknown unknown = vUnknown.getUnknown();

		int /* long */[] ppvObject = new int /* long */[1];
		int code = unknown.QueryInterface(COM.IIDIEnumVARIANT, ppvObject);

		if (code == COM.E_NOINTERFACE) {
			System.err.println("Dommage...");
		}

		if (ppvObject[0] == 0) {
			System.err.println("IEnumVARIANT not supported!!");
			throw new IllegalArgumentException(
					"Ce IDispatch n'est pas un tableau - il ne fournit pas la propriété DISPID_NEWENUM");
		}

		IEnumVARIANT ienumvariant = new IEnumVARIANT(ppvObject[0]);

		int /* long */pVarResultAddress = 0;
		pVarResultAddress = OS.GlobalAlloc(OS.GMEM_FIXED | OS.GMEM_ZEROINIT,
				VARIANT.sizeof);
		int[] cEltFetched = new int[1];
		// Variant v = Variant.win32_new(pVarResultAddress);
		Variant v = new Variant();

		while (COM.S_OK == (code = ienumvariant.Next(1, pVarResultAddress,
				cEltFetched))) {

			if (pVarResultAddress != 0) {
				VariantUtils.setVariantData(v, pVarResultAddress);

				values.add(VariantUtils.getVariantValue(v));
			}
		}
		if (pVarResultAddress != 0) {
			COM.VariantClear(pVarResultAddress);
			OS.GlobalFree(pVarResultAddress);
		}
		ienumvariant.Release();
		// COM.MoveMemory(pceltFetched, cEltFetched, 4);

		Object[] ret = values.toArray();
		return ret;
	}

	public static Variant convert(Object o, Class<?> targetClass) {
		if (null == targetClass) {
			throw new IllegalArgumentException("null targetClass");
		}
		if (null == o || targetClass.equals(void.class)) {
			return null;
		}
		if (targetClass.equals(int.class)) {
			return new Variant((Integer) o);
		}
		if (targetClass.equals(long.class)) {
			return new Variant((Long) o);
		}
		if (targetClass.equals(String.class)) {
			return new Variant((String) o);
		}
		if (targetClass.equals(boolean.class)) {
			return new Variant((Boolean) o);
		}
		if (targetClass.equals(byte.class)) {
			return new Variant((Byte) o);
		}
		if (targetClass.equals(char.class)) {
			return new Variant((Character) o);
		}
		if (targetClass.equals(double.class)) {
			return new Variant((Double) o);
		}
		if (targetClass.equals(float.class)) {
			return new Variant((Float) o);
		}
		if (targetClass.equals(short.class)) {
			return new Variant((Short) o);
		}
		if (ComIDispatch.class.isAssignableFrom(targetClass)) {
			return new Variant(new IDispatch(((ComIDispatch) o).getAddress()));
		}
		if (ComIUnknown.class.isAssignableFrom(targetClass)) {
			return new Variant(new IUnknown(((ComIUnknown) o).getAddress()));
		}
		if (IDispatch.class.isAssignableFrom(targetClass)) {
			return new Variant((IDispatch) o);
		}
		if (IUnknown.class.isAssignableFrom(targetClass)) {
			return new Variant((IUnknown) o);
		}

		throw new IllegalArgumentException("Type de Retour non supporté : "
				+ targetClass);
	}

	public static Object convert(Variant variant, Class<?> targetClass) {
		if (null == targetClass) {
			throw new IllegalArgumentException("null targetClass");
		}
		if (null == variant || variant.getType() == COM.VT_EMPTY
				|| variant.getType() == COM.VT_NULL) {
			return null;
		}
		if (targetClass.equals(int.class)) {
			return variant.getInt();
		}
		if (targetClass.equals(long.class)) {
			return variant.getLong();
		}
		if (targetClass.equals(String.class)) {
			return variant.getString();
		}
		if (targetClass.equals(boolean.class)) {
			return variant.getBoolean();
		}
		if (targetClass.equals(byte.class)) {
			return variant.getByte();
		}
		if (targetClass.equals(char.class)) {
			return variant.getChar();
		}
		if (targetClass.equals(double.class)) {
			return variant.getDouble();
		}
		if (targetClass.equals(float.class)) {
			return variant.getFloat();
		}
		if (targetClass.equals(short.class)) {
			return variant.getShort();
		}
		if (targetClass.equals(IUnknown.class)) {
			return variant.getUnknown();
		}
		if (targetClass.equals(IDispatch.class)) {
			return variant.getDispatch();
		}
		if (targetClass.isArray()) {
			Object[] oRet = getArrayvalues(variant.getDispatch());
			if (null == oRet) {
				return null;
			}

			Object[] ret = (Object[]) Array.newInstance(targetClass
					.getComponentType(), oRet.length);
			for (int i = 0; i < oRet.length; i++) {
				Object o = oRet[i];
				if (!targetClass.getComponentType().isInstance(o)) {
					throw new IllegalArgumentException("Element #" + i
							+ " is not an instance of "
							+ targetClass.getComponentType());
				}
				ret[i] = o;
			}
			return ret;
		}
		throw new IllegalArgumentException("Type de Parametre non supporté : "
				+ targetClass);
	}

	public static int invoke(IDispatch objIDispatch, int dispIdMember,
			int wFlags, Variant[] rgvarg, int[] rgdispidNamedArgs,
			Variant pVarResult) {

		// get the IDispatch interface for the control
		if (objIDispatch == null)
			return COM.E_FAIL;

		// create a DISPPARAMS structure for the input parameters
		DISPPARAMS pDispParams = new DISPPARAMS();
		// store arguments in rgvarg
		if (rgvarg != null && rgvarg.length > 0) {
			pDispParams.cArgs = rgvarg.length;
			pDispParams.rgvarg = OS.GlobalAlloc(COM.GMEM_FIXED
					| COM.GMEM_ZEROINIT, VARIANT.sizeof * rgvarg.length);
			int offset = 0;
			for (int i = rgvarg.length - 1; i >= 0; i--) {
				VariantUtils.getVariantData(rgvarg[i], pDispParams.rgvarg
						+ offset);
				offset += VARIANT.sizeof;
			}
		}

		// if arguments have ids, store the ids in rgdispidNamedArgs
		if (rgdispidNamedArgs != null && rgdispidNamedArgs.length > 0) {
			pDispParams.cNamedArgs = rgdispidNamedArgs.length;
			pDispParams.rgdispidNamedArgs = OS.GlobalAlloc(COM.GMEM_FIXED
					| COM.GMEM_ZEROINIT, 4 * rgdispidNamedArgs.length);
			int offset = 0;
			for (int i = rgdispidNamedArgs.length; i > 0; i--) {
				COM.MoveMemory(pDispParams.rgdispidNamedArgs + offset,
						new int[] { rgdispidNamedArgs[i - 1] }, 4);
				offset += 4;
			}
		}

		// invoke the method
		EXCEPINFO excepInfo = new EXCEPINFO();
		int[] pArgErr = new int[1];
		int /* long */pVarResultAddress = 0;
		if (pVarResult != null)
			pVarResultAddress = OS.GlobalAlloc(
					OS.GMEM_FIXED | OS.GMEM_ZEROINIT, VARIANT.sizeof);
		int result = objIDispatch.Invoke(dispIdMember, new GUID(),
				COM.LOCALE_USER_DEFAULT, wFlags, pDispParams,
				pVarResultAddress, excepInfo, pArgErr);

		if (pVarResultAddress != 0) {
			VariantUtils.setVariantData(pVarResult, pVarResultAddress);
			COM.VariantClear(pVarResultAddress);
			OS.GlobalFree(pVarResultAddress);
		}

		// free the Dispparams resources
		if (pDispParams.rgdispidNamedArgs != 0) {
			OS.GlobalFree(pDispParams.rgdispidNamedArgs);
		}
		if (pDispParams.rgvarg != 0) {
			int offset = 0;
			for (int i = 0, length = rgvarg.length; i < length; i++) {
				COM.VariantClear(pDispParams.rgvarg + offset);
				offset += VARIANT.sizeof;
			}
			OS.GlobalFree(pDispParams.rgvarg);
		}

		// save error string and cleanup EXCEPINFO
		manageExcepinfo(result, excepInfo);

		return result;
	}

	/**
	 * Traduire un EXCEPINFO en Exception Runtime si hResult n'est pas OK.
	 * @param hResult
	 * @param excepInfo
	 */
	public static void manageExcepinfo(int hResult, EXCEPINFO excepInfo) {

		String exceptionDescription = null;
		if (hResult == COM.S_OK) {
			exceptionDescription = "No Error"; //$NON-NLS-1$
			return;
		}

		// extract exception info
		if (hResult == COM.DISP_E_EXCEPTION) {
			if (excepInfo.bstrDescription != 0) {
				int size = COM.SysStringByteLen(excepInfo.bstrDescription);
				char[] buffer = new char[(size + 1) / 2];
				COM.MoveMemory(buffer, excepInfo.bstrDescription, size);
				exceptionDescription = new String(buffer);
			} else {
				exceptionDescription = "OLE Automation Error Exception "; //$NON-NLS-1$
				if (excepInfo.wCode != 0) {
					exceptionDescription += "code = " + excepInfo.wCode; //$NON-NLS-1$
				} else if (excepInfo.scode != 0) {
					exceptionDescription += "code = " + excepInfo.scode; //$NON-NLS-1$
				}
			}
		} else {
			exceptionDescription = "OLE Automation Error HResult : " + hResult; //$NON-NLS-1$
		}

		// cleanup EXCEPINFO struct
		if (excepInfo.bstrDescription != 0)
			COM.SysFreeString(excepInfo.bstrDescription);
		if (excepInfo.bstrHelpFile != 0)
			COM.SysFreeString(excepInfo.bstrHelpFile);
		if (excepInfo.bstrSource != 0)
			COM.SysFreeString(excepInfo.bstrSource);

		throw new RuntimeException(exceptionDescription);
	}

	private static Variant getProperty(IDispatch dispatch, int dispIdMember) {
		Variant pVarResult = new Variant();
		int result = invoke(dispatch, dispIdMember, COM.DISPATCH_PROPERTYGET,
				null, null, pVarResult);
		return (result == OLE.S_OK) ? pVarResult : null;
	}

	public static Variant getProperty(IDispatch dispatch, String propertyName) {
		int[] rgDispId = new int[1];
		int dispIdMember = rgDispId[0];
		dispatch.GetIDsOfNames(null, new String[] { propertyName }, 1, 0,
				rgDispId);
		return getProperty(dispatch, dispIdMember);
	}

}
