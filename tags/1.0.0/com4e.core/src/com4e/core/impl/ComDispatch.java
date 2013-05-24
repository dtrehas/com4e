package com4e.core.impl;

import java.lang.annotation.Annotation;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.Map.Entry;

import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.DISPPARAMS;
import org.eclipse.swt.internal.win32.OS;
import org.eclipse.swt.ole.win32.Variant;

import com4e.core.annotations.ComMethod;
import com4e.core.annotations.ComNamedParam;
import com4e.core.annotations.DispMethod;
import com4e.core.api.ComIDispatch;

@SuppressWarnings("restriction")
public class ComDispatch extends ComUnknown implements ComIDispatch {

	/**
	 * ExposedMethods peut contenir des items nuls. dispId N est stocké à
	 * l'index N-1;
	 */
	private ExposedMethod[] exposedMethods;
	private ExposedMethod defaultExposedMethod = null;

	private static ExposedMethod[] initExposedMethods(Class<?> clazz) {
		NavigableMap<Integer, ExposedMethod> exposedMethodsByDispId = new TreeMap<Integer, ExposedMethod>();
		for (Method m : clazz.getMethods()) {
			DispMethod dispMethod = m.getAnnotation(DispMethod.class);
			if (null != dispMethod) {

				if (exposedMethodsByDispId.containsKey(dispMethod.id())) {
					System.err
							.println("Plusieurs méthodes avec le même dispId; Ecrasement de "
									+ exposedMethodsByDispId.get(dispMethod
											.id())
									+ " par "
									+ dispMethod.name());
				}
				exposedMethodsByDispId.put(dispMethod.id(),
						new ExposedMethod(m));
			}
		}
		int maxIndex = exposedMethodsByDispId.lastKey();
		ExposedMethod[] exposedMethods = new ExposedMethod[maxIndex + 1];
		for (Entry<Integer, ExposedMethod> entry : exposedMethodsByDispId
				.entrySet()) {
			exposedMethods[entry.getKey()] = entry.getValue();
		}

		return exposedMethods;

	}

	private ExposedMethod findDefaultExposedMethod(Class<?> clazz) {
		ExposedMethod ret = null;
		for (Method m : clazz.getMethods()) {
			DispMethod dispMethod = m.getAnnotation(DispMethod.class);
			if (null != dispMethod) {

				if (dispMethod.defaultMethod()) {
					if (null != ret) {
						System.err
								.printf(
										"La classe de dispatch %s a plusieurs méthodes par défaut - écrasement de %s par %s\n",
										getClass(), ret.getMethod(), m);
					}
					ret = new ExposedMethod(m);
				}
			}
		}
		return ret;
	}

	protected void displayExposedMethods() {
		System.err
				.println("ExposedMethods: " + Arrays.toString(exposedMethods));
	}

	public ComDispatch() {
		this(ComDispatch.class);
	}

	// private static int[] computeArgCounts() {
	// fillInExposedMethods();
	// int[] argCounts = new int[exposedMethods.size()];
	// return null;
	// }

	protected ComDispatch(Class<?> c) {
		super(c);
		setDefaultExposedMethod(findDefaultExposedMethod(getClass()));
		exposedMethods = initExposedMethods(getClass());
		implementedInterfaces.add(COM.IIDIDispatch);
	}

	/**
	 * Implémentation de IDispatch::GetIDsOfNames.
	 * 
	 * @param riid
	 *            inutilisé
	 * @param rgszNames
	 *            tableau des noms à mapper
	 * @param cNames
	 *            nombre de noms dans le tableau
	 * @param lcid
	 *            locale
	 * @param rgDispId
	 *            tableau destiné a recevoir les identifiants mappes
	 * @return S_OK en cas de réussite, DISP_E_UNKNOWNNAME si au moins un des
	 *         noms demandés n'est pas connu
	 */
	@Override
	@ComMethod(index = 5)
	@DispMethod(name = "GetIDsOfNames", id = 6)
	public int GetIDsOfNames(int riid, int rgszNames, int cNames, int lcid,
			int rgDispId) {

		if (cNames < 1) {
			throw new IllegalArgumentException(
					"cNames devrait être supérieur ou égal à 1");
		}

		// récupère le nom de la methode
		int[] pName = new int[1];
		COM.MoveMemory(pName, rgszNames, OS.PTR_SIZEOF);
		String name = COMStringToString(pName[0]);

		System.out.print("On me demande l'ID de " + name + " ... ");

		Integer methodDispId = null;
		ExposedMethod exposedMethod = null;
		// on recherche cette méthode par le nom
		for (int dispId = 0; dispId < exposedMethods.length; dispId++) {
			exposedMethod = exposedMethods[dispId];
			if (exposedMethod != null
					&& exposedMethod.getExposedName().equalsIgnoreCase(name)) {
				methodDispId = dispId;
				break;
			}
		}

		int[] dispIds = new int[cNames];
		int retour = COM.S_OK;

		if (null == methodDispId) {
			// methode non trouvée!
			for (int i = 0; i < dispIds.length; i++) {
				dispIds[i] = DISPID_UNKNOWN;
			}
			retour = DISP_E_UNKNOWNNAME;
		} else {

			dispIds[0] = methodDispId;

			System.out.println("... je vais retourner " + methodDispId);

			// si on arrive ici, c'est que la méthode est trouvée, cherchons ses
			// arguments!
			for (int i = 1; i < cNames; ++i) {
				int[] pParamName = new int[1];
				COM.MoveMemory(pParamName, rgszNames + OS.PTR_SIZEOF * i,
						OS.PTR_SIZEOF);
				String paramName = COMStringToString(pParamName[0]);

				// on cherche le paramètre nommé paramName
				for (int paramIndex = 0; paramIndex < exposedMethod.getMethod()
						.getParameterAnnotations().length; paramIndex++) {
					// on parcourt les annotations de ce paramètre pour trouver
					// son
					// nommage
					Annotation[] paramAnnotations = exposedMethod.getMethod()
							.getParameterAnnotations()[paramIndex];
					for (Annotation annotation : paramAnnotations) {
						if (annotation instanceof ComNamedParam) {
							String currentName = ((ComNamedParam) annotation)
									.name();
							if (null != currentName
									&& paramName.equalsIgnoreCase(currentName)) {
								// on l'a trouve!
								dispIds[i] = paramIndex;
								continue;
							}
						}
					}
					System.out.println("Param " + paramName + " pas trouvé!");
					// pas trouvé!
					dispIds[i] = DISPID_UNKNOWN;
					retour = DISP_E_UNKNOWNNAME;
				}
			}
		}
		// int[] pDispId = new int[1];
		// if (dispIds.isEmpty()) {
		// return DISP_E_UNKNOWNNAME;
		// }

		// /debug
		if (name.equals("0")) {
			methodDispId = 0;
		}

		System.out.println("GetIDsOfNames: " + name + "=" + methodDispId);

		if (rgDispId != 0) {
			int i = 0;
			for (int dispId : dispIds) {
				COM.MoveMemory(rgDispId + OS.PTR_SIZEOF * i,
						new int[] { dispId }, OS.PTR_SIZEOF);
				i++;
			}

		}

		return retour;
	}

	/**
	 * Transforme la chaine en chaine java.
	 * 
	 * @param address
	 *            adresse de la chaine
	 * @return chaine java (ou une chaine vide en cas de probleme)
	 */
	protected String COMStringToString(int address) {
		if (address != 0) {
			int size = COM.SysStringByteLen(address);
			char[] buffer = new char[size];
			COM.MoveMemory(buffer, address, size * 2);
			// suppression du 0 terminal
			String retour = new String(buffer, 0, size - 1);
			return retour;
		}
		return "";
	}

	@Override
	@ComMethod(index = 4)
	@DispMethod(name = "GetTypeInfo", id = 5)
	public int GetTypeInfo(int info, int lcid, int ppTInfo) {
		return COM.E_NOTIMPL;
	}

	@Override
	@ComMethod(index = 3)
	@DispMethod(name = "GetTypeInfoCount", id = 4)
	public int GetTypeInfoCount(int pctinfo) {
		return COM.E_NOTIMPL;
	}

	private String flagToString(int dwFlags) {
		String ret = "";
		if ((dwFlags & COM.DISPATCH_METHOD) == COM.DISPATCH_METHOD) {
			ret += "METHOD ";
		}
		if ((dwFlags & COM.DISPATCH_PROPERTYGET) == COM.DISPATCH_PROPERTYGET) {
			ret += "GET ";
		}
		if ((dwFlags & COM.DISPATCH_PROPERTYPUT) == COM.DISPATCH_PROPERTYPUT) {
			ret += "PUT ";
		}
		if ((dwFlags & COM.DISPATCH_PROPERTYPUTREF) == COM.DISPATCH_PROPERTYPUTREF) {
			ret += "PUTREF ";
		}
		return ret;
	}

	/**
	 * Implémentation de IDispatch::Invoke. Seule la méthode "1" est prise en
	 * charge. Cette méthode correspond à un signal "pageLoaded".
	 * 
	 * @param dispIdMember
	 *            identifiant de la méthode a invoquer
	 * @param riid
	 *            inutilisé
	 * @param lcid
	 *            locale (utilisé pour les noms)
	 * @param dwFlags
	 *            parametrage (on accepte uniquement si DISPATCH_METHOD est
	 *            active)
	 * @param pDispParams
	 *            liste des parametres
	 * @param pVarResult
	 *            adresse de la valeur de retour
	 * @param pExcepInfo
	 *            adresse de la variable recevant les informations d'exception
	 * @param pArgErr
	 *            adresse de la variable recevant l'index du premier paramètre
	 *            en erreur
	 * @return code de retour (S_OK si tout c'est bien passé).
	 */
	@Override
	@ComMethod(index = 6)
	@DispMethod(name = "Invoke", id = 7)
	public int Invoke(int dispIdMember, int riid, int lcid, int dwFlags,
			int pDispParams, int pVarResult, int pExcepInfo, int pArgErr) {

		System.out.println("FLAG:" + flagToString(dwFlags));

		System.out.println("On invoke " + dispIdMember);

		if (dispIdMember == 0) {
			if (defaultExposedMethod != null) {
				return _invoke(getDefaultExposedMethod(), new Variant[] {},
						pVarResult);
			}
			return COM.DISP_E_MEMBERNOTFOUND;
		}

		Variant variants[] = null;
		if (pDispParams != 0) {
			DISPPARAMS params = new DISPPARAMS();
			COM.MoveMemory(params, pDispParams, DISPPARAMS.sizeof);

			variants = new Variant[params.cArgs];
			for (int i = 0; i < params.cArgs; i++) {
				int pParam = params.rgvarg + i * Variant.sizeof;
				variants[params.cArgs - 1 - i] = Variant.win32_new(pParam);
			}
			System.out.printf("PARAMs: %s\n", Arrays.toString(variants));

		}

		ExposedMethod em = exposedMethods[dispIdMember]/*
														 * le dispMember
														 * commence à 1...
														 */;
		return _invoke(em, variants, pVarResult);
	}

	private int _invoke(ExposedMethod em, Variant[] vParams, int pVarResult) {
		if (null == em) {
			return COM.DISP_E_MEMBERNOTFOUND;
		}

		// if (em.getMethod().getGenericParameterTypes().length > 0) {
		// throw new RuntimeException(
		// "Only zero-arg methods are supported for the moment!");
		// }

		if (null != vParams) {
			if (vParams.length != em.getMethod().getGenericParameterTypes().length) {
				// TODO Attention aux paramètres optionnels!
				return DISP_E_BADPARAMCOUNT;
			}
		}

		Object[] params = new Object[vParams.length];
		for (int i = 0; i < params.length; i++) {
			Class<?> targetClass = em.getMethod().getParameterTypes()[i];
			Object param = VariantUtils.convert(vParams[i], targetClass);
			params[i] = param;
		}

		try {
			Object ret = em.getMethod().invoke(this, params);
			Variant vRet = VariantUtils.convert(ret, em.getMethod().getReturnType());
			if (null != vRet) {
				if (pVarResult != 0) {
					// copier le retour dans la zone de retour
					Variant.win32_copy(pVarResult, vRet);
				}
			}

		} catch (IllegalArgumentException e) {
			e.printStackTrace();
			return DISP_E_PARAMNOTFOUND;
		} catch (InvocationTargetException e) {
			e.printStackTrace();
			// TODO remplir la structure d'exception
			return COM.DISP_E_EXCEPTION;
		} catch (IllegalAccessException e) {
			e.printStackTrace();
			// TODO remplir la structure d'exception
			return COM.DISP_E_EXCEPTION;
		}

		return COM.S_OK;
	}

	/**
	 * @return the defaultExposedMethod
	 */
	public ExposedMethod getDefaultExposedMethod() {
		return defaultExposedMethod;
	}

	/**
	 * @param defaultExposedMethod
	 *            the defaultExposedMethod to set
	 */
	public void setDefaultExposedMethod(ExposedMethod defaultExposedMethod) {
		this.defaultExposedMethod = defaultExposedMethod;
	}

}
