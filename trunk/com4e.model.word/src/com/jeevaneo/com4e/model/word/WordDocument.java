package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordDocument extends IOleAutomated {
	String getName();

	WordApplication getApplication();
	
	WordParagraphs getParagraphs();

	WordVariables getVariables();

	WordFields getFields();

	WordBookmarks getBookmarks();

	WordSections getSections();
	String getCodeName();
	WordRange getContent();

	void close();
	/**
	 * <li>wdNotSaveChanges : 0</li>
	 * <li>wdPromptToSaveChanges : -2</li>
	 * <li>wdSaveChanges : -1</li>
	 * @param option
	 */
	void close(long option); 
	void save();
	void saveAs(String filename);
	void select();
	void undo();
	WordRange range();

	void activate();
}