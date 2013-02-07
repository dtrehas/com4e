package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordDocuments extends IOleAutomated {
	WordDocument Open(String filename);
}