package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordRange extends IOleAutomated {
	String getXML();

	void setText(String t);

	String getText();

	void collapse();
	
	void delete();

	void insertParagraph();

	void insertParagraphAfter();

	void insertParagraphBefore();

	void insertFile(String filename);

	void insertBefore(String text);

	void insertAfter(String text);
}