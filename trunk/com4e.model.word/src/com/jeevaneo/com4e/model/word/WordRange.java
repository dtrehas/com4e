package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordRange extends IOleAutomated {
	
	long wdCollapseEnd =0;
	long wdCollapseStart=1;
	
	String getXML();

	void setText(String t);

	String getText();
	
	long getStart();
	void setStart(long pos);
	
	long getEnd();
	void setEnd(long pos);

	void collapse();
	void collapse(long direction);
	
	void delete();

	void insertParagraph();

	void insertParagraphAfter();

	void insertParagraphBefore();

	void insertFile(String filename);

	void insertBefore(String text);

	void insertAfter(String text);
}