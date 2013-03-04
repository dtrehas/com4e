package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordApplication extends IOleAutomated {
	void activate();
	
	WordDocuments getDocuments();
	
	WordDocument getActiveDocument();
	
	void quit();
	
}