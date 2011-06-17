package com4e.core.impl;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.UUID;

import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.COMObject;
import org.eclipse.swt.internal.ole.win32.GUID;
import org.eclipse.swt.internal.win32.OS;

import com4e.core.annotations.ComMethod;
import com4e.core.annotations.DispMethod;
import com4e.core.api.ComIUnknown;
import com4e.core.api.IComObject;

@SuppressWarnings({ "restriction"})
public class ComUnknown extends ComObject implements ComIUnknown, IComObject {
	public ComUnknown() {
		this(ComUnknown.class);
	}

	public ComUnknown(Class<?> c) {
		super(/* implementer, */c);
	}

	// protected static int[] concatenateArgCounts(int[] a,
	// int[] b) {
	// int[] ret = new int[a.length+b.length];
	// System.arraycopy(a, 0, ret, 0, a.length);
	// System.arraycopy(b, 0, ret, a.length, b.length);
	// return ret;
	// }

	// (){
	// public int /* long */method0(int /* long */[] args) {
	// return QueryInterface(args[0], args[1]);
	// }
	//
	// public int /* long */method1(int /* long */[] args) {
	// return AddRef();
	// }
	//
	// public int /* long */method2(int /* long */[] args) {
	// return Release();
	// }


	@SuppressWarnings("serial")
	protected final Set<GUID> implementedInterfaces = new HashSet<GUID>() {
		{
			add(COM.IIDIUnknown);
		}
	};

	int refCount = 0;

	@ComMethod(index = 1)
	@DispMethod(name = "AddRef", id = 2)
	public int AddRef() {
		refCount++;
		return refCount;
	}

	@ComMethod(index = 2)
	@DispMethod(name = "Release", id = 3)
	public int Release() {
		refCount--;

		if (refCount == 0) {
			dispose();
		}
		return refCount;
	}

	@ComMethod(index = 0)
	@DispMethod(name = "QueryInterface", id = 1)
	public int QueryInterface(int /* long */riid, int /* long */ppvObject) {

		if (ppvObject == 0) {
			return COM.E_POINTER;
		}

		if (riid == 0)
			return COM.E_NOINTERFACE;
		GUID requestedGuid = new GUID();
		COM.MoveMemory(requestedGuid, riid, GUID.sizeof);

		for (GUID clsid : implementedInterfaces) {
			if (COM.IsEqualGUID(requestedGuid, clsid)) {
				System.out.println("MATCH! : " + _display(requestedGuid) + " "
						+ _display(clsid));
				COM.MoveMemory(ppvObject, new int /* long */[] { getAddress() },
						OS.PTR_SIZEOF);
				AddRef();
				System.out.println("INTERFACE SUPPORTED: "
						+ _debugTranslateClsid(requestedGuid));
				return COM.S_OK;
			}

			// System.err.println("NO MATCH : " + _display(requestedGuid) + " "
			// + _display(clsid));
		}
		COM.MoveMemory(ppvObject, new int /* long */[] { 0 }, OS.PTR_SIZEOF);

		// System.err.println("INTERFACE NOT SUPPORTED: " +
		// _debugTranslateClsid(requestedGuid));

		return COM.E_NOINTERFACE;
	}

	private String _display(GUID guid) {
		String ret = "{";
		ret += Integer.toHexString(guid.Data1);
		ret += "-";
		ret += Integer.toHexString(guid.Data2);
		ret += "-";
		ret += Integer.toHexString(guid.Data3);
		ret += "-";
		ret += guid.Data4;
		ret += "}";

		return ret;
	}

	public static UUID fromBytes(byte[] bytesOriginal) {
		byte[] bytes = new byte[16];
		// Reverse the first 4 bytes
		bytes[0] = bytesOriginal[3];
		bytes[1] = bytesOriginal[2];
		bytes[2] = bytesOriginal[1];
		bytes[3] = bytesOriginal[0];
		// Reverse 6th and 7th
		bytes[4] = bytesOriginal[5];
		bytes[5] = bytesOriginal[4];
		// Reverse 8th and 9th
		bytes[6] = bytesOriginal[7];
		bytes[7] = bytesOriginal[6];
		// Copy the rest straight up
		for (int i = 8; i < 16; i++) {
			bytes[i] = bytesOriginal[i];
		}
		return toUUID(bytes);
	}

	public static byte[] toBytes(UUID uuid) {
		byte[] bytesOriginal = asByteArray(uuid);
		byte[] bytes = new byte[16];
		// Reverse the first 4 bytes
		bytes[0] = bytesOriginal[3];
		bytes[1] = bytesOriginal[2];
		bytes[2] = bytesOriginal[1];
		bytes[3] = bytesOriginal[0];
		// Reverse 6th and 7th
		bytes[4] = bytesOriginal[5];
		bytes[5] = bytesOriginal[4];
		// Reverse 8th and 9th
		bytes[6] = bytesOriginal[7];
		bytes[7] = bytesOriginal[6];
		// Copy the rest straight up
		for (int i = 8; i < 16; i++) {
			bytes[i] = bytesOriginal[i];
		}
		return bytes;
	}

	private static byte[] asByteArray(UUID uuid) {
		long msb = uuid.getMostSignificantBits();
		long lsb = uuid.getLeastSignificantBits();
		byte[] buffer = new byte[16];
		for (int i = 0; i < 8; i++) {
			buffer[i] = (byte) (msb >>> 8 * (7 - i));
		}
		for (int i = 8; i < 16; i++) {
			buffer[i] = (byte) (lsb >>> 8 * (7 - i));
		}
		return buffer;
	}

	private static UUID toUUID(byte[] byteArray) {
		long msb = 0;
		long lsb = 0;
		for (int i = 0; i < 8; i++)
			msb = (msb << 8) | (byteArray[i] & 0xff);
		for (int i = 8; i < 16; i++)
			lsb = (lsb << 8) | (byteArray[i] & 0xff);
		UUID result = new UUID(msb, lsb);
		return result;
	}

	private static GUID asGUID(UUID uuid) {
		byte[] bytes = toBytes(uuid);
		GUID ret = new GUID();
		ret.Data1 = (int) bytes[3];
		ret.Data1 += 8 * (int) bytes[2];
		ret.Data1 += 8 * 8 * (int) bytes[1];
		ret.Data1 += 8 * 8 * 8 * (int) bytes[0];

		ret.Data4 = new byte[8];
		System.arraycopy(bytes, 8, ret.Data4, 0, 8);

		return ret;
	}

	public static void main(String[] args) throws Exception {

		// ComUnknown moi = new ComUnknown();
		// System.out.println(moi._display(IIDFromString("{00000000-0000-0000-C000-000000000046}")));
		GUID guid0 = IIDFromString("{4657278A-411B-11d2-839A-00C04FD918D0}");
		UUID uuid = UUID.fromString("4657278A-411B-11d2-839A-00C04FD918D0");
		GUID guid1 = asGUID(uuid);
		System.out.println(guid0.Data1);
		System.out.println(guid1.Data1);
		System.out.println(guid0.Data1 == guid1.Data1);
		System.out.println(guid0.Data2 == guid1.Data2);
		System.out.println(guid0.Data3 == guid1.Data3);
		System.out.println(Arrays.equals(guid0.Data4, guid1.Data4));

		System.out.println(Arrays.toString(guid0.Data4));
		System.out.println(Arrays.toString(toBytes(uuid)));

		System.out.println(Arrays.toString(asByteArray(uuid)));
		System.out.println(UUID.nameUUIDFromBytes(asByteArray(uuid)));
		System.out.println(UUID.nameUUIDFromBytes(toBytes(uuid)));
		System.out.println(fromBytes(toBytes(uuid)));
		System.out.println(uuid.getMostSignificantBits());
	}

	protected String _debugTranslateClsid(GUID guid) {
		if (COM.IsEqualGUID(guid,
				IIDFromString("{4657278A-411B-11d2-839A-00C04FD918D0}")))
			return ("CLSID_DragDropHelper");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{4657278B-411B-11d2-839A-00C04FD918D0}")))
			return ("IID_IDropTargetHelper");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{DE5BF786-477A-11d2-839D-00C04FD918D0}")))
			return ("IID_IDragSourceHelper");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{83E07D0D-0C5F-4163-BF1A-60B274051E40}")))
			return ("IID_IDragSourceHelper2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{8AD9C840-044E-11D1-B3E9-00805F499D93}")))
			return ("IIDJavaBeansBridge");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{166B1BCA-3F9C-11CF-8075-444553540000}")))
			return ("IIDShockwaveActiveXControl");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{6BD2AEFE-7876-45e6-A6E7-3BFCDF6540AA}")))
			return ("IIDIEditorSiteTime");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{D381A1F4-2326-4f3c-AFB9-B7537DB9E238}")))
			return ("IIDIEditorSiteProperty");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{61E55B0B-2647-47c4-8C89-E736EF15D636}")))
			return ("IIDIEditorBaseProperty");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{CDD88AB9-B01D-426E-B0F0-30973E9A074B}")))
			return ("IIDIEditorSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{BEE283FE-7B42-4FF3-8232-0F07D43ABCF1}")))
			return ("IIDIEditorService");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{EFDE08C4-BE87-4B1A-BF84-15FC30207180}")))
			return ("IIDIEditorManager");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{618736E0-3C3D-11CF-810C-00AA00389B71}")))
			return ("IIDIAccessible");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{03022430-ABC4-11D0-BDE2-00AA001A1953}")))
			return ("IIDIAccessibleHandler");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0C733A8C-2A1C-11CE-ADE5-00AA0044773D}")))
			return ("IIDIAccessor");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010F-0000-0000-C000-000000000046}")))
			return ("IIDIAdviseSink");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000125-0000-0000-C000-000000000046}")))
			return ("IIDIAdviseSink2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000E-0000-0000-C000-000000000046}")))
			return ("IIDIBindCtx");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000001-0000-0000-C000-000000000046}")))
			return ("IIDIClassFactory");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B28F-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIClassFactory2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B286-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIConnectionPoint");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B284-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIConnectionPointContainer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{22F03340-547D-101B-8E65-08002B2BD119}")))
			return ("IIDICreateErrorInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020405-0000-0000-C000-000000000046}")))
			return ("IIDICreateTypeInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020406-0000-0000-C000-000000000046}")))
			return ("IIDICreateTypeLib");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000110-0000-0000-C000-000000000046}")))
			return ("IIDIDataAdviseHolder");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010E-0000-0000-C000-000000000046}")))
			return ("IIDIDataObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020400-0000-0000-C000-000000000046}")))
			return ("IIDIDispatch");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{BD3F23C0-D43E-11CF-893B-00AA00BDCE1A}")))
			return ("IIDIDocHostUIHandler");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{C4D244B0-D43E-11CF-893B-00AA00BDCE1A}")))
			return ("IIDIDocHostShowUI");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000121-0000-0000-C000-000000000046}")))
			return ("IIDIDropSource");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000122-0000-0000-C000-000000000046}")))
			return ("IIDIDropTarget");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B285-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIEnumConnectionPoints");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B287-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIEnumConnections");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000103-0000-0000-C000-000000000046}")))
			return ("IIDIEnumFORMATETC");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000102-0000-0000-C000-000000000046}")))
			return ("IIDIEnumMoniker");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000104-0000-0000-C000-000000000046}")))
			return ("IIDIEnumOLEVERB");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000105-0000-0000-C000-000000000046}")))
			return ("IIDIEnumSTATDATA");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000D-0000-0000-C000-000000000046}")))
			return ("IIDIEnumSTATSTG");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000101-0000-0000-C000-000000000046}")))
			return ("IIDIEnumString");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000100-0000-0000-C000-000000000046}")))
			return ("IIDIEnumUnknown");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020404-0000-0000-C000-000000000046}")))
			return ("IIDIEnumVARIANT");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{1CF2B120-547D-101B-8E65-08002B2BD119}")))
			return ("IIDIErrorInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{3127CA40-446E-11CE-8135-00AA004BB851}")))
			return ("IIDIErrorLog");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000019-0000-0000-C000-000000000046}")))
			return ("IIDIExternalConnection");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{BEF6E002-A874-101A-8BBA-00AA00300CAB}")))
			return ("IIDIFont");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{BEF6E003-A874-101A-8BBA-00AA00300CAB}")))
			return ("IIDIFontDisp");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{3050F613-98B5-11CF-BB82-00AA00BDCE0B}")))
			return ("IIDIHTMLDocumentEvents2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{79eac9ee-baf9-11ce-8c82-00aa004ba90b}")))
			return ("IIDIInternetSecurityManager");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000A-0000-0000-C000-000000000046}")))
			return ("IIDILockBytes");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000002-0000-0000-C000-000000000046}")))
			return ("IIDIMalloc");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000001D-0000-0000-C000-000000000046}")))
			return ("IIDIMallocSpy");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000003-0000-0000-C000-000000000046}")))
			return ("IIDIMarshal");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000016-0000-0000-C000-000000000046}")))
			return ("IIDIMessageFilter");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000F-0000-0000-C000-000000000046}")))
			return ("IIDIMoniker");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000111-0000-0000-C000-000000000046}")))
			return ("IIDIOleAdviseHolder");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000011E-0000-0000-C000-000000000046}")))
			return ("IIDIOleCache");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000128-0000-0000-C000-000000000046}")))
			return ("IIDIOleCache2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000129-0000-0000-C000-000000000046}")))
			return ("IIDIOleCacheControl");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000118-0000-0000-C000-000000000046}")))
			return ("IIDIOleClientSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B722BCCB-4E68-101B-A2BC-00AA00404770}")))
			return ("IIDIOleCommandTarget");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000011B-0000-0000-C000-000000000046}")))
			return ("IIDIOleContainer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B288-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIOleControl");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B289-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIOleControlSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B722BCC5-4E68-101B-A2BC-00AA00404770}")))
			return ("IIDIOleDocument");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B722BCC7-4E68-101B-A2BC-00AA00404770}")))
			return ("IIDIOleDocumentSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000117-0000-0000-C000-000000000046}")))
			return ("IIDIOleInPlaceActiveObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000116-0000-0000-C000-000000000046}")))
			return ("IIDIOleInPlaceFrame");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000113-0000-0000-C000-000000000046}")))
			return ("IIDIOleInPlaceObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000119-0000-0000-C000-000000000046}")))
			return ("IIDIOleInPlaceSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000115-0000-0000-C000-000000000046}")))
			return ("IIDIOleInPlaceUIWindow");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000011C-0000-0000-C000-000000000046}")))
			return ("IIDIOleItemContainer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000011D-0000-0000-C000-000000000046}")))
			return ("IIDIOleLink");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000112-0000-0000-C000-000000000046}")))
			return ("IIDIOleObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000114-0000-0000-C000-000000000046}")))
			return ("IIDIOleWindow");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000011A-0000-0000-C000-000000000046}")))
			return ("IIDIParseDisplayName");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{376BD3AA-3845-101B-84ED-08002B2EC713}")))
			return ("IIDIPerPropertyBrowsing");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010C-0000-0000-C000-000000000046}")))
			return ("IIDIPersist");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010B-0000-0000-C000-000000000046}")))
			return ("IIDIPersistFile");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{BD1AE5E0-A6AE-11CE-BD37-504200C10000}")))
			return ("IIDIPersistMemory");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{37D84F60-42CB-11CE-8135-00AA004BB851}")))
			return ("IIDIPersistPropertyBag");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010A-0000-0000-C000-000000000046}")))
			return ("IIDIPersistStorage");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000109-0000-0000-C000-000000000046}")))
			return ("IIDIPersistStream");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{7FD52380-4E07-101B-AE2D-08002B2EC713}")))
			return ("IIDIPersistStreamInit");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{7BF80980-BF32-101A-8BBB-00AA00300CAB}")))
			return ("IIDIPicture");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{7BF80981-BF32-101A-8BBB-00AA00300CAB}")))
			return ("IIDIPictureDisp");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{55272A00-42CB-11CE-8135-00AA004BB851}")))
			return ("IIDIPropertyBag");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{9BFBBC02-EFF1-101A-84ED-00AA00341D07}")))
			return ("IIDIPropertyNotifySink");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B28D-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIPropertyPage");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{01E44665-24AC-101B-84ED-08002B2EC713}")))
			return ("IIDIPropertyPage2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B28C-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIPropertyPageSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B283-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDIProvideClassInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851}")))
			return ("IIDIProvideClassInfo2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{D5F569D0-593B-101A-B569-08002B2DBF7A}")))
			return ("IIDIPSFactoryBuffer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000012-0000-0000-C000-000000000046}")))
			return ("IIDIRootStorage");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{F29F6BC0-5021-11CE-AA15-00006901293F}")))
			return ("IIDIROTData");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{D5F56B60-593B-101A-B569-08002B2DBF7A}")))
			return ("IIDIRpcChannelBuffer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{D5F56A34-593B-101A-B569-08002B2DBF7A}")))
			return ("IIDIRpcProxyBuffer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{D5F56AFC-593B-101A-B569-08002B2DBF7A}")))
			return ("IIDIRpcStubBuffer");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000126-0000-0000-C000-000000000046}")))
			return ("IIDIRunnableObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000010-0000-0000-C000-000000000046}")))
			return ("IIDIRunningObjectTable");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{742B0E01-14E6-101B-914E-00AA00300CAB}")))
			return ("IIDISimpleFrameSite");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{6d5140c1-7436-11ce-8034-00aa006009fa}")))
			return ("IIDIServiceProvider");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{B196B28B-BAB4-101A-B69C-00AA00341D07}")))
			return ("IIDISpecifyPropertyPages");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000018-0000-0000-C000-000000000046}")))
			return ("IIDIStdMarshalInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000B-0000-0000-C000-000000000046}")))
			return ("IIDIStorage");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000000C-0000-0000-C000-000000000046}")))
			return ("IIDIStream");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{DF0B3D60-548F-101B-8E65-08002B2BD119}")))
			return ("IIDISupportErrorInfo");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020403-0000-0000-C000-000000000046}")))
			return ("IIDITypeComp");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00020402-0000-0000-C000-000000000046}")))
			return ("IIDITypeLib");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000000-0000-0000-C000-000000000046}")))
			return ("IIDIUnknown");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{0000010D-0000-0000-C000-000000000046}")))
			return ("IIDIViewObject");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{00000127-0000-0000-C000-000000000046}")))
			return ("IIDIViewObject2");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{f38bc242-b950-11d1-8918-00c04fc2c836}")))
			return ("CGID_DocHostCommandHandler");
		if (COM.IsEqualGUID(guid,
				IIDFromString("{000214D0-0000-0000-C000-000000000046}")))
			return ("CGID_Explorer");
		if (COM.IsEqualGUID(guid, IID_IBEUtils))
			return "IID_IBEUtils";
		if (COM.IsEqualGUID(guid, IID_IBEEvents))
			return "IID_IBEEvents";
		if (COM.IsEqualGUID(guid, IID_IBEWorkset))
			return "IID_IBEWorkset";
		if (COM.IsEqualGUID(guid, IID_IBELauncher))
			return "IID_IBELauncher";
		if (COM.IsEqualGUID(guid, IID_IBEHistory))
			return "IID_IBEHistory";
		return "NOT TRANSLATED";
	}

	/**
	 * GUID identifiant l'interface IBEUtils (usage interne).
	 */
	public static final GUID IID_IBEUtils = COMObject
			.IIDFromString("{A831B17C-6990-447A-8EEB-61173AE78EBD}");

	/**
	 * GUID identifiant l'interface IBEEvents (envoi d'évènements vers le BE).
	 */
	public static final GUID IID_IBEEvents = COMObject
			.IIDFromString("{DBB11DC5-5693-4D3E-9CC5-A4AC4A1310CB}");

	/**
	 * GUID identifiant l'interface IBEWorkset (accès au workset depuis
	 * l'extérieur).
	 */
	public static final GUID IID_IBEWorkset = COMObject
			.IIDFromString("{2AE61736-CA81-4BF3-86E5-9A22DDF8DDAF}");

	/**
	 * GUID identifiant l'interface IBELauncher (accès au launcher).
	 */
	public static final GUID IID_IBELauncher = COMObject
			.IIDFromString("{98157DDE-76BD-4831-8719-156A61FAA454}");

	/**
	 * GUID identifiant l'interface IBEHistory (accès à l'historique des objets
	 * métier).
	 */
	public static final GUID IID_IBEHistory = COMObject
			.IIDFromString("{5516165B-A47B-4A4B-970C-2789FE23C553}");

}