package pt.keep.validator.msoffice.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "excel")
public class ExcelInfo extends OfficeInfo{
	private int numberOfSheets;

	@XmlElement
	public int getNumberOfSheets() {
		return numberOfSheets;
	}

	public void setNumberOfSheets(int numberOfSheets) {
		this.numberOfSheets = numberOfSheets;
	}

	

	
	
	
}
