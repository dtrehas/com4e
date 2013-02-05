package com.jeevaneo.com4e.model.word;

public interface WordField {
	WordApplication getApplication();

	String getData();

	void setData(String s);

	WordRange getCode();

	WordRange getResult();

	boolean getShowCodes();

	void setShowCodes(boolean b);

	void update();
}