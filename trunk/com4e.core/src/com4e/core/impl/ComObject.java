package com4e.core.impl;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.Arrays;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.Map.Entry;

import org.eclipse.swt.internal.ole.win32.COMObject;

import com4e.core.annotations.ComMethod;
import com4e.core.annotations.DispMethod;
import com4e.core.api.IComObject;

@SuppressWarnings("restriction")
public class ComObject extends COMObject implements IComObject {

	private static int[] computeArgCounts(Class<?> c, VTable vtable) {
		NavigableMap<Integer, Method> comMethodsByVTableIndex = new TreeMap<Integer, Method>();
		for (Method m : c.getMethods()) {
			ComMethod comMethod = m.getAnnotation(ComMethod.class);
			if (null != comMethod) {
				if (comMethodsByVTableIndex.containsKey(comMethod.index())) {
					System.err
							.println("Plusieurs méthodes avec le même index dans la VTable; Ecrasement de "
									+ comMethodsByVTableIndex.get(comMethod
											.index()) + " par " + m);
				}
				comMethodsByVTableIndex.put(comMethod.index(), m);
			}
		}
		int maxIndex = comMethodsByVTableIndex.lastKey();
		if (null != vtable) {
			Method[] comMethods = new Method[maxIndex + 1];
			for (Entry<Integer, Method> entry : comMethodsByVTableIndex
					.entrySet()) {
				comMethods[entry.getKey()] = entry.getValue();
			}
			vtable.setComMethods(comMethods);
		}

		int[] ret = new int[maxIndex + 1];
		for (int i = 0; i < ret.length; i++) {
			Method m = comMethodsByVTableIndex.get(i);
			if (null == m) {
				// un trou dans la VTable!!
				System.err
						.println("WARNING - Il y a un trou dans la VTable à l'index "
								+ i + " (sera rempli par 0))");
				ret[i] = 0;
			} else {
				ret[i] = m.getParameterTypes().length;
			}
		}
		return ret;
	}

	// protected ComIUnknown control;
	protected int[] argCounts;

	final VTable vtable = new VTable();

	public ComObject(/* ComIUnknown implementer, */Class<?> c) {
		super(computeArgCounts(c, null));

		if (!c.isInstance(this)) {
			throw new IllegalArgumentException(
					"this n'est pas une instance de la classe passée : " + c);
		}
		// this.control = implementer;
		argCounts = computeArgCounts(c, vtable);
		// concatenateArgCounts(myArgCounts, argCounts);
		checkVTable();
	}

	protected void checkVTable() {
		// if (argCounts.length > exposedMethods.length) {
		// System.err
		// .printf(
		// "Warning - il y a moins d'entrées dans la VTable (%d) que de méthodes exposées (%d)!\n",
		// argCounts.length, exposedMethods.length);
		// }
		// for (int i = 0; i < Math.max(this.argCounts.length,
		// exposedMethods.length); i++) {
		// int dispId = i + 1;
		// int argCount = -1;
		// if (i < argCounts.length) {
		// argCount = argCounts[i];
		// }
		// ExposedMethod em = null;
		// if (i < exposedMethods.length) {
		// em = exposedMethods[i];
		// }
		//
		// if (null == em) {
		// if (argCount == 0) {
		// // OK
		// continue;
		// } else {
		// System.err
		// .printf(
		// "Mismatch ExposedMethods/VTable : pas de methode exposée pour le dispId=%d, mais on trouve %d params dans la VTable\n",
		// dispId, argCount);
		//
		// displayVTable();
		// displayExposedMethods();
		// }
		// } else if (em.getArgsCount() != argCount) {
		// if (i < argCounts.length) {
		// System.err
		// .printf(
		// "Mismatch ExposedMethods/VTable : Methode exposée '%s' pour le dispId=%d a %d parametre(s), mais on trouve %d params dans la VTable\n",
		// em.getExposedName(), dispId, em
		// .getArgsCount(), argCount);
		// } else {
		// System.err
		// .printf(
		// "Mismatch ExposedMethods/VTable : Methode exposée '%s' pour le dispId=%d a %d parametre(s), mais on ne la trouve pas dans la VTable, qui n'a que %d entrées.\n",
		// em.getExposedName(), dispId, em
		// .getArgsCount(), argCounts.length);
		// }
		// }
		// }
		System.err.println(getClass());
		displayVTable();
	}

	protected void displayVTable() {
		System.err.println("VTable: " + Arrays.toString(argCounts));
	}

	private IComObject comDelegate = (IComObject) Proxy.newProxyInstance(
			IComObject.class.getClassLoader(),
			new Class[] { IComObject.class }, new IComObjectInvocationHandler(
					this));

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method0(int[])
	 */
	public int method0(int[] args) {
		return comDelegate.method0(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method1(int[])
	 */
	public int method1(int[] args) {
		return comDelegate.method1(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method10(int[])
	 */
	public int method10(int[] args) {
		return comDelegate.method10(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method11(int[])
	 */
	public int method11(int[] args) {
		return comDelegate.method11(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method12(int[])
	 */
	public int method12(int[] args) {
		return comDelegate.method12(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method13(int[])
	 */
	public int method13(int[] args) {
		return comDelegate.method13(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method14(int[])
	 */
	public int method14(int[] args) {
		return comDelegate.method14(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method15(int[])
	 */
	public int method15(int[] args) {
		return comDelegate.method15(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method16(int[])
	 */
	public int method16(int[] args) {
		return comDelegate.method16(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method17(int[])
	 */
	public int method17(int[] args) {
		return comDelegate.method17(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method18(int[])
	 */
	public int method18(int[] args) {
		return comDelegate.method18(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method19(int[])
	 */
	public int method19(int[] args) {
		return comDelegate.method19(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method2(int[])
	 */
	public int method2(int[] args) {
		return comDelegate.method2(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method20(int[])
	 */
	public int method20(int[] args) {
		return comDelegate.method20(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method21(int[])
	 */
	public int method21(int[] args) {
		return comDelegate.method21(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method22(int[])
	 */
	public int method22(int[] args) {
		return comDelegate.method22(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method23(int[])
	 */
	public int method23(int[] args) {
		return comDelegate.method23(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method24(int[])
	 */
	public int method24(int[] args) {
		return comDelegate.method24(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method25(int[])
	 */
	public int method25(int[] args) {
		return comDelegate.method25(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method26(int[])
	 */
	public int method26(int[] args) {
		return comDelegate.method26(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method27(int[])
	 */
	public int method27(int[] args) {
		return comDelegate.method27(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method28(int[])
	 */
	public int method28(int[] args) {
		return comDelegate.method28(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method29(int[])
	 */
	public int method29(int[] args) {
		return comDelegate.method29(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method3(int[])
	 */
	public int method3(int[] args) {
		return comDelegate.method3(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method30(int[])
	 */
	public int method30(int[] args) {
		return comDelegate.method30(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method31(int[])
	 */
	public int method31(int[] args) {
		return comDelegate.method31(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method32(int[])
	 */
	public int method32(int[] args) {
		return comDelegate.method32(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method33(int[])
	 */
	public int method33(int[] args) {
		return comDelegate.method33(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method34(int[])
	 */
	public int method34(int[] args) {
		return comDelegate.method34(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method35(int[])
	 */
	public int method35(int[] args) {
		return comDelegate.method35(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method36(int[])
	 */
	public int method36(int[] args) {
		return comDelegate.method36(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method37(int[])
	 */
	public int method37(int[] args) {
		return comDelegate.method37(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method38(int[])
	 */
	public int method38(int[] args) {
		return comDelegate.method38(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method39(int[])
	 */
	public int method39(int[] args) {
		return comDelegate.method39(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method4(int[])
	 */
	public int method4(int[] args) {
		return comDelegate.method4(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method40(int[])
	 */
	public int method40(int[] args) {
		return comDelegate.method40(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method41(int[])
	 */
	public int method41(int[] args) {
		return comDelegate.method41(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method42(int[])
	 */
	public int method42(int[] args) {
		return comDelegate.method42(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method43(int[])
	 */
	public int method43(int[] args) {
		return comDelegate.method43(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method44(int[])
	 */
	public int method44(int[] args) {
		return comDelegate.method44(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method45(int[])
	 */
	public int method45(int[] args) {
		return comDelegate.method45(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method46(int[])
	 */
	public int method46(int[] args) {
		return comDelegate.method46(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method47(int[])
	 */
	public int method47(int[] args) {
		return comDelegate.method47(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method48(int[])
	 */
	public int method48(int[] args) {
		return comDelegate.method48(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method49(int[])
	 */
	public int method49(int[] args) {
		return comDelegate.method49(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method5(int[])
	 */
	public int method5(int[] args) {
		return comDelegate.method5(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method50(int[])
	 */
	public int method50(int[] args) {
		return comDelegate.method50(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method51(int[])
	 */
	public int method51(int[] args) {
		return comDelegate.method51(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method52(int[])
	 */
	public int method52(int[] args) {
		return comDelegate.method52(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method53(int[])
	 */
	public int method53(int[] args) {
		return comDelegate.method53(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method54(int[])
	 */
	public int method54(int[] args) {
		return comDelegate.method54(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method55(int[])
	 */
	public int method55(int[] args) {
		return comDelegate.method55(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method56(int[])
	 */
	public int method56(int[] args) {
		return comDelegate.method56(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method57(int[])
	 */
	public int method57(int[] args) {
		return comDelegate.method57(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method58(int[])
	 */
	public int method58(int[] args) {
		return comDelegate.method58(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method59(int[])
	 */
	public int method59(int[] args) {
		return comDelegate.method59(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method6(int[])
	 */
	public int method6(int[] args) {
		return comDelegate.method6(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method60(int[])
	 */
	public int method60(int[] args) {
		return comDelegate.method60(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method61(int[])
	 */
	public int method61(int[] args) {
		return comDelegate.method61(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method62(int[])
	 */
	public int method62(int[] args) {
		return comDelegate.method62(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method63(int[])
	 */
	public int method63(int[] args) {
		return comDelegate.method63(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method64(int[])
	 */
	public int method64(int[] args) {
		return comDelegate.method64(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method65(int[])
	 */
	public int method65(int[] args) {
		return comDelegate.method65(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method66(int[])
	 */
	public int method66(int[] args) {
		return comDelegate.method66(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method67(int[])
	 */
	public int method67(int[] args) {
		return comDelegate.method67(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method68(int[])
	 */
	public int method68(int[] args) {
		return comDelegate.method68(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method69(int[])
	 */
	public int method69(int[] args) {
		return comDelegate.method69(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method7(int[])
	 */
	public int method7(int[] args) {
		return comDelegate.method7(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method70(int[])
	 */
	public int method70(int[] args) {
		return comDelegate.method70(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method71(int[])
	 */
	public int method71(int[] args) {
		return comDelegate.method71(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method72(int[])
	 */
	public int method72(int[] args) {
		return comDelegate.method72(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method73(int[])
	 */
	public int method73(int[] args) {
		return comDelegate.method73(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method74(int[])
	 */
	public int method74(int[] args) {
		return comDelegate.method74(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method75(int[])
	 */
	public int method75(int[] args) {
		return comDelegate.method75(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method76(int[])
	 */
	public int method76(int[] args) {
		return comDelegate.method76(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method77(int[])
	 */
	public int method77(int[] args) {
		return comDelegate.method77(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method78(int[])
	 */
	public int method78(int[] args) {
		return comDelegate.method78(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method79(int[])
	 */
	public int method79(int[] args) {
		return comDelegate.method79(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method8(int[])
	 */
	public int method8(int[] args) {
		return comDelegate.method8(args);
	}

	/**
	 * @param args
	 * @return
	 * @see com4e.core.api.IComObject#method9(int[])
	 */
	public int method9(int[] args) {
		return comDelegate.method9(args);
	}

}

class IComObjectInvocationHandler implements InvocationHandler {
	private ComObject proxied;

	public IComObjectInvocationHandler(ComObject proxied) {
		super();
		this.proxied = proxied;
	}

	@Override
	public Object invoke(Object proxy, Method method, Object[] args)
			throws Throwable {
		if (method.getName().matches("method[0-9]+")) {
			if (null == args) {
				throw new RuntimeException("Null args array");
			}

			if (!(args[0] instanceof int[])) {
				throw new IllegalArgumentException(
						String
								.format(
										"La methode %s ne prend qu'un paramètre, de type int[] (vs %s).",
										method, args[0]));
			}

			// real args
			int[] realArgs = (int[]) args[0];
			// args must be ints and in equal number as args

			int index = Integer.parseInt(method.getName().replaceFirst(
					"method", ""));
			if (index < proxied.vtable.getComMethods().length) {
				Method realMethod = proxied.vtable.getComMethods()[index];

				// realMethod doit renvoyer un int
				if (!realMethod.getReturnType().equals(int.class)
						&& !realMethod.getReturnType().equals(void.class)) {
					throw new IllegalArgumentException(
							String
									.format(
											"La méthode %s devrait retourner un entier et non %s\n",
											realMethod, realMethod
													.getGenericReturnType()));
				}

				Class<?>[] paramTypes = realMethod.getParameterTypes();
				// il ne doit y avoir que des int, et autant que d'items dans
				// realArgs
				if (realArgs.length != paramTypes.length) {
					throw new IllegalArgumentException(
							String
									.format(
											"La methode %s a %d parametres, mais on lui en a passé %d.\n",
											method, paramTypes.length,
											args.length));
				}
				for (int i = 0; i < paramTypes.length; i++) {
					if (!paramTypes[i].equals(int.class)) {
						throw new IllegalArgumentException(
								String
										.format(
												"Le paramètre #%d de la méthode %s devrait être un int : %s\n",
												i, realMethod, paramTypes[i]));
					}
				}
				Object[] params = new Object[realArgs.length];
				for (int i = 0; i < params.length; i++) {
					params[i] = new Integer(realArgs[i]);
				}
				Object ret = realMethod.invoke(proxied, params);

				// On renvoit toujours 0 en cas de méthode void
				if (realMethod.getReturnType().equals(void.class)) {
					ret = 0;
				}
				return ret;
			} else {
				return method.invoke(proxied, args);
			}
		} else {
			return method.invoke(proxied, args);
		}
	}

}

class VTable {
	Method[] comMethods;

	public void setComMethods(Method[] comMethods) {
		this.comMethods = comMethods;
	}

	public VTable() {
		super();
	}

	public Method[] getComMethods() {
		return comMethods;
	}
}

class ExposedMethod {
	private String exposedName;
	private int argsCount;

	public String getExposedName() {
		return exposedName;
	}

	private Method method;

	public Method getMethod() {
		return method;
	}

	public ExposedMethod(Method method) {
		super();
		if (null == method) {
			throw new IllegalArgumentException("Methode nulle");
		}
		this.method = method;
		this.exposedName = method.getAnnotation(DispMethod.class).name();
		this.argsCount = method.getParameterTypes().length;
	}

	@Override
	public int hashCode() {
		return exposedName.hashCode();
	}

	@Override
	public boolean equals(Object obj) {
		return hashCode() == obj.hashCode();
	}

	@Override
	public String toString() {
		return exposedName + "@" + argsCount;
	}

	public int getArgsCount() {
		return argsCount;
	}
}
