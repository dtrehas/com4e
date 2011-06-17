package com4e.core.api;

public interface ComIDispatch extends ComIUnknown {
	/**
	 * Unknown name.
	 */
	public static final int DISP_E_UNKNOWNNAME = 0x80020006;
	/**
	 * Parameter not found.
	 */
	public static final int DISP_E_PARAMNOTFOUND = 0x80020004;
	/**
	 * Unknown dispid.
	 */
	public static final int DISPID_UNKNOWN = -1;
	/**
	 * Invalid number of parameters.
	 */
	public static final int DISP_E_BADPARAMCOUNT = 0x8002000E;

	// int GetIDsOfNames(GUID riid, String[] rgszNames, int cNames, int lcid,
	// int[] rgDispId) ;
	int GetIDsOfNames(int riid, int rgszNames, int cNames, int lcid,
			int rgDispId);

	int GetTypeInfo(int iTInfo, int lcid, int /* long */ppTInfo);

	int GetTypeInfoCount(int pctinfo);

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
	int Invoke(int dispIdMember, int riid, int lcid, int dwFlags,
			int pDispParams, int pVarResult, int pExcepInfo, int pArgErr);

}
