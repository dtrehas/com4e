package com.jeevaneo.com4e.model.word;

public interface WordBookmark {
	void copy(String name);

	void delete();

	void select();

	WordApplication getApplication();

	boolean isColumn();

	boolean isEmpty();

	String getName();

	WordRange getRange();

	long getStart();

	void setStart(long s);
}