package pt.keep.validator.result;

import java.util.Date;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;

@XmlRootElement(name = "powerpoint")
public class PowerpointInfo extends OfficeInfo {	
	private int numberOfSlides;

	@XmlElement
	public int getNumberOfSlides() {
		return numberOfSlides;
	}

	public void setNumberOfSlides(int numberOfSlides) {
		this.numberOfSlides = numberOfSlides;
	}
	
	

}
