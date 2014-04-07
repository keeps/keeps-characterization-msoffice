package pt.keep.validator.msoffice.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

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
