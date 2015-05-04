package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordDocuments extends IOleAutomated {
	WordDocument Open(String filename);
	WordDocument Add();
	WordDocument item(long i);
	long getCount();
	
}