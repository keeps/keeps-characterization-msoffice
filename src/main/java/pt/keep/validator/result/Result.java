package pt.keep.validator.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;



@XmlRootElement(name = "ooxmlValidationResult")
@XmlType(propOrder = { "valid","validationError","word","excel","powerpoint"})
public class Result {
	private boolean valid;
	private String validationError;
	
	private ExcelInfo excel;
	private WordInfo word;
	private PowerpointInfo powerpoint;
	
	
	@XmlElement(required=true)
	public boolean isValid() {
		return valid;
	}
	public void setValid(boolean valid) {
		this.valid = valid;
	}
	
	@XmlElement(required=false)
	public String getValidationError() {
		return validationError;
	}
	public void setValidationError(String validationError) {
		this.validationError = validationError;
	}
	
	@XmlElement
	public ExcelInfo getExcel() {
		return excel;
	}
	public void setExcel(ExcelInfo excel) {
		this.excel = excel;
	}
	
	@XmlElement
	public WordInfo getWord() {
		return word;
	}
	public void setWord(WordInfo word) {
		this.word = word;
	}
	
	@XmlElement
	public PowerpointInfo getPowerpoint() {
		return powerpoint;
	}
	public void setPowerpoint(PowerpointInfo powerpoint) {
		this.powerpoint = powerpoint;
	}
	
	
	
	
}
