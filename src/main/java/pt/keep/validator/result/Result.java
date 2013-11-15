package pt.keep.validator.result;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;



@XmlRootElement(name = "msofficeCharacterizationResult")
@XmlType(propOrder = { "valid","validationErrors","word","excel","powerpoint"})
public class Result {
	private boolean valid;
	private List<String> validationErrors;
	private ExcelInfo excel;
	private WordInfo word;
	private PowerpointInfo powerpoint;

	
	public void addError(String error){
		if(validationErrors==null){
			validationErrors=new ArrayList<String>();
		}
		if(error==null){
			error = "Null pointer exception";
		}
		if(!validationErrors.contains(error)){
			validationErrors.add(error);
		}
	}
	
	@XmlElement(required=true)
	public boolean isValid() {
		return valid;
	}
	public void setValid(boolean valid) {
		this.valid = valid;
	}
	
	@XmlElement(required=false)
	public List<String> getValidationErrors() {
		return validationErrors;
	}
	public void setValidationErrors(List<String> validationErrors) {
		this.validationErrors = validationErrors;
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
