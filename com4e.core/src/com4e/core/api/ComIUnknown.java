package com4e.core.api;

import com4e.core.annotations.ComMethod;
import com4e.core.annotations.DispMethod;

public interface ComIUnknown extends IComObject {

	@ComMethod(index = 1)
	@DispMethod(name = "QueryInterface", id = 1)
	public int QueryInterface(int /* long */riid, int /* long */ppvObject);

	@ComMethod(index = 2)
	@DispMethod(name = "AddRef", id = 2)
	public int AddRef();

	@ComMethod(index = 3)
	@DispMethod(name = "Release", id = 3)
	public int Release();

}
