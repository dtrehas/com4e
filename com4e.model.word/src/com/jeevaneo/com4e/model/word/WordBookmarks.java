package com.jeevaneo.com4e.model.word;

public interface WordBookmarks {

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