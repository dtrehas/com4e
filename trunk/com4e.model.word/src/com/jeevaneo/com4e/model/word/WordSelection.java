package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordSelection extends IOleAutomated {
	WordRange getRange();
	long getStart();
	long getEnd();
	void setEnd(long end);
	void setStart(long start);
}