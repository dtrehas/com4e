package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordParagraph extends IOleAutomated {
	WordRange getRange();
	String getID();
	void setID(String id);
}