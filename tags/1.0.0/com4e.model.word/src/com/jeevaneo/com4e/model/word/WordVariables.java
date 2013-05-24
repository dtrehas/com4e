package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordVariables extends IOleAutomated {
	WordVariable item(long i);

	WordVariable item(String name);

	WordVariable add(String name, String value);

	WordApplication getApplication();

	long getCount();
}