package com.jeevaneo.com4e.model.word;

import com.jeevaneo.com4e.automation.IOleAutomated;

public interface WordParagraphs extends IOleAutomated {
	WordParagraph add();
	WordParagraph add(WordRange range);

	WordParagraph item(long i);

	long getCount();
}