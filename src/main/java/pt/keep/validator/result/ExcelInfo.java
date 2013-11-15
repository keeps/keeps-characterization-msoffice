package pt.keep.validator.result;

import java.util.Date;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;

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
