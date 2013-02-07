package com.jeevaneo.com4e.automation;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;

import org.eclipse.swt.internal.ole.win32.IDispatch;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

public class OleAutomationProxy implements InvocationHandler {

	private OleAutomation auto;

	public OleAutomationProxy(OleAutomation auto) {
		super();
		this.auto = auto;
	}

	@SuppressWarnings("unchecked")
	public static <T> T instrument(OleAutomation auto, Class<? extends T> clazz) {
		OleAutomationProxy oap = new OleAutomationProxy(auto);
		T ret = (T) Proxy.newProxyInstance(clazz.getClassLoader(),
				new Class[] { clazz }, oap);
		// Main.listMembers(auto);
		return ret;
	}

	@Override
	public Object invoke(Object proxy, Method method, Object[] args)
			throws Throwable {
		if (method.getDeclaringClass().equals(IOleAutomated.class)) {
			if (method.getName().equals("getPointer")
					&& method.getReturnType().equals(Variant.class)) {
				return new Variant(auto);
			}
			if (method.getName().equals("getOleAutomation")
					&& method.getReturnType().equals(OleAutomation.class)) {
				return auto;
			}
		}

		if (method.getName().startsWith("is")
				&& (method.getReturnType().equals(Boolean.class) || method
						.getReturnType().equals(boolean.class))) {
			// GET Property
			String prop = method.getName().substring(2);
			int[] dispIds = auto.getIDsOfNames(new String[] { prop });
			if (null == dispIds) {
				throw new IllegalStateException("Property " + prop + " : "
						+ auto.getLastError());
			}
			// if (method.getParameterTypes().length == 0) {
			Variant vRet = auto.getProperty(dispIds[0],
					asVariants(args, method.getParameterTypes()));
			return castAs(vRet, method.getReturnType());
			// } else {
			// throw new
			// UnsupportedOperationException("Getting property with params is not -yet- supported! : "
			// + prop);
			// }
		}
		if (method.getName().startsWith("get")
				&& !method.getReturnType().equals(void.class)) {
			// GET Property
			String prop = method.getName().substring(3);
			int[] dispIds = auto.getIDsOfNames(new String[] { prop });
			if (null == dispIds) {
				Main.listMembers(auto);
				throw new IllegalStateException("Property " + prop + " : "
						+ auto.getLastError());
			}
			// if (method.getParameterTypes().length == 0) {
			Variant vRet = auto.getProperty(dispIds[0],
					asVariants(args, method.getParameterTypes()));
			return castAs(vRet, method.getReturnType());
			// } else {
			// throw new
			// UnsupportedOperationException("Getting property with params is not -yet- supported! : "
			// + prop);
			// }
		}
		if (method.getName().startsWith("set")
				&& method.getReturnType().equals(void.class)) {
			// SET Property
			String prop = method.getName().substring(3);
			int[] dispIds = auto.getIDsOfNames(new String[] { prop });
			if (null == dispIds) {
				throw new IllegalStateException("Property " + prop + " : "
						+ auto.getLastError());
			}
			auto.setProperty(dispIds[0],
					asVariants(args, method.getParameterTypes()));
			return null;
		}

		// this is an invocation
		String methodName = method.getName();
		int[] dispIds = auto.getIDsOfNames(new String[] { methodName });
		if (null == dispIds) {
			throw new IllegalStateException("Method " + methodName + " : "
					+ auto.getLastError());
		}
		Variant[] vArgs = asVariants(args, method.getParameterTypes());
		Variant vRet = auto.invoke(dispIds[0], vArgs);
		return castAs(vRet, method.getReturnType());
	}

	private Variant[] asVariants(Object[] args, Class<?>[] types) {
		if (null == args)
			return null;
		Variant[] vArgs = new Variant[args.length];
		for (int i = 0; i < args.length; i++) {
			vArgs[i] = asVariant(args[i], types[i]);
		}
		return vArgs;
	}

	private Variant asVariant(Object object, Class<?> class1) {
		if (null == object)
			return null;
		if (class1.equals(String.class)) {
			return new Variant((String) object);
		}
		if (class1.equals(Integer.class) || class1.equals(int.class)) {
			return new Variant((Integer) object);
		}
		if (class1.equals(Boolean.class) || class1.equals(boolean.class)) {
			return new Variant((Boolean) object);
		}
		if (class1.equals(Double.class) || class1.equals(double.class)) {
			return new Variant((Double) object);
		}
		if (class1.equals(Float.class) || class1.equals(float.class)) {
			return new Variant((Float) object);
		}
		if (class1.equals(Short.class) || class1.equals(short.class)) {
			return new Variant((Short) object);
		}
		if (class1.equals(Long.class) || class1.equals(long.class)) {
			return new Variant((Long) object);
		}
		if (class1.isAssignableFrom(OleAutomation.class)) {
			return new Variant((OleAutomation) object);
		}
		if (class1.isAssignableFrom(IDispatch.class)) {
			return new Variant((IDispatch) object);
		}
		if (class1.isAssignableFrom(IOleAutomated.class)) {
			return ((IOleAutomated) object).getPointer();
		}
		throw new UnsupportedOperationException(
				"Type de paramètre non supporté : " + class1);
	}

	private Object castAs(Variant vRet, Class<?> returnType) {
		if (null == vRet)
			return null;
		if (vRet.getType() == OLE.VT_EMPTY) {
			return null;
		}
		if (returnType.equals(String.class)) {
			return vRet.getString();
		}
		if (returnType.equals(Integer.class) || returnType.equals(int.class)) {
			return vRet.getInt();
		}
		if (returnType.equals(Boolean.class)
				|| returnType.equals(boolean.class)) {
			return vRet.getBoolean();
		}
		if (returnType.equals(Double.class) || returnType.equals(double.class)) {
			return vRet.getDouble();
		}
		if (returnType.equals(Float.class) || returnType.equals(float.class)) {
			return vRet.getFloat();
		}
		if (returnType.equals(Short.class) || returnType.equals(short.class)) {
			return vRet.getShort();
		}
		if (returnType.equals(Long.class) || returnType.equals(long.class)) {
			return vRet.getLong();
		}
		if (returnType.isAssignableFrom(OleAutomation.class)) {
			return vRet.getAutomation();
		}
		if (returnType.isAssignableFrom(IDispatch.class)) {
			return vRet.getDispatch();
		}
		if (returnType.equals(void.class)) {
			return null;
		}
		if (vRet.getType() == OLE.VT_DISPATCH) {
			Class<?> clazz = returnType;
			return Proxy.newProxyInstance(clazz.getClassLoader(),
					new Class[] { clazz },
					new OleAutomationProxy(vRet.getAutomation()));
		}
		throw new UnsupportedOperationException(
				"Type de paramètre non supporté : " + returnType);
	}

}