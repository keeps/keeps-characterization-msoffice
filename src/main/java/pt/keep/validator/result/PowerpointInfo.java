package pt.keep.validator.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlRootElement(name = "powerpoint")
@XmlType(propOrder = { "powerpointError","numberOfSlides"})
public class PowerpointInfo {
	private String powerpointError;
	
	private int numberOfSlides;

	@XmlElement
	public int getNumberOfSlides() {
		return numberOfSlides;
	}

	public void setNumberOfSlides(int numberOfSlides) {
		this.numberOfSlides = numberOfSlides;
	}

	@XmlElement
	public String getPowerpointError() {
		return powerpointError;
	}

	public void setPowerpointError(String powerpointError) {
		this.powerpointError = powerpointError;
	}
	
	
	
	

}
