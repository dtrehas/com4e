package com.jeevaneo.com4e.model.word;

public interface WordVariables {
	WordVariable item(long i);

	WordVariable item(String name);

	WordVariable add(String name, String value);

	WordApplication getApplication();

	long getCount();
}