package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordField extends IOleAutomated {
	WordApplication getApplication();

	String getData();

	void setData(String s);

	WordRange getCode();

	WordRange getResult();

	boolean getShowCodes();

	void setShowCodes(boolean b);

	void update();
}