package com.jeevaneo.com4e.automation;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

public interface IOleAutomated {
	Variant getPointer();
	OleAutomation getOleAutomation();
}
