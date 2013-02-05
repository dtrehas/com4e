package com.jeevaneo.com4e.model.word;

public interface WordRange {
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