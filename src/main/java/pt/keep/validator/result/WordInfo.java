package pt.keep.validator.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlRootElement(name = "word")
@XmlType(propOrder = { "wordError"})
public class WordInfo {
	private String wordError;


	@XmlElement
	public String getWordError() {
		return wordError;
	}

	public void setWordError(String wordError) {
		this.wordError = wordError;
	}
}
