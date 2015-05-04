package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordBookmarks extends IOleAutomated {

	WordBookmark add(String name);

	WordBookmark add(String name, WordRange range);

	boolean exists(String name);

	WordBookmark item(String name);

	WordBookmark item(long i);

	WordApplication getApplication();

	long getCount();

	boolean getShowHidden();

	void setShowHidden(boolean show);

}