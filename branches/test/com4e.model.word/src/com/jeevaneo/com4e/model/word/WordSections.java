package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordSections extends IOleAutomated {
	WordSection add();

	WordSection item(long i);

	WordApplication getApplication();

	long getCount();

	WordSection getFirst();

	WordSection getLast();

}