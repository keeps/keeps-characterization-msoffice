package pt.keep.validator.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlRootElement(name = "excel")
@XmlType(propOrder = { "excelError","numberOfSheets"})
public class ExcelInfo {
	private int numberOfSheets;
	private String excelError;

	@XmlElement
	public int getNumberOfSheets() {
		return numberOfSheets;
	}

	public void setNumberOfSheets(int numberOfSheets) {
		this.numberOfSheets = numberOfSheets;
	}

	@XmlElement
	public String getExcelError() {
		return excelError;
	}

	public void setExcelError(String excelError) {
		this.excelError = excelError;
	}
	
	
	
	
	
	
	
}
