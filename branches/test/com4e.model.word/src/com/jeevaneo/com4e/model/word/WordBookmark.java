package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordBookmark extends IOleAutomated {
	void copy(String name);

	void delete();

	void select();

	WordApplication getApplication();

	boolean isColumn();

	boolean isEmpty();

	String getName();

	WordRange getRange();

	long getStart();

	void setStart(long s);

	long getEnd();

	void setEnd(long s);
}