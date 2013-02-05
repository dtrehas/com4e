package com.jeevaneo.com4e.model.word;

public interface WordVariable {
	String getName();

	String getValue();

	void setValue(String val);

	void delete();

	WordApplication getApplication();

	long getIndex();
}