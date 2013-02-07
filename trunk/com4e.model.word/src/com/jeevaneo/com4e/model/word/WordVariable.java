package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordVariable extends IOleAutomated {
	String getName();

	String getValue();

	void setValue(String val);

	void delete();

	WordApplication getApplication();

	long getIndex();
}