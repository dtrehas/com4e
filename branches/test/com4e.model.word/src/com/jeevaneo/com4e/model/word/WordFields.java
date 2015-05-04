package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordFields extends IOleAutomated {
	// WordField Add();
	WordField item(long i);

	long getCount();

}