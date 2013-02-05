package com.jeevaneo.com4e.model.word;

public interface WordSections {
	WordSection add();

	WordSection item(long i);

	WordApplication getApplication();

	long getCount();

	WordSection getFirst();

	WordSection getLast();

}