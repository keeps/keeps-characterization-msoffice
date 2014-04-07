package pt.keep.validator.msoffice.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "pst")
public class PstInfo {
	private int emailCount;
	private int folderCount;
	
	
	@XmlElement
	public int getEmailCount() {
		return emailCount;
	}
	public void setEmailCount(int emailCount) {
		this.emailCount = emailCount;
	}
	
	@XmlElement
	public int getFolderCount() {
		return folderCount;
	}
	public void setFolderCount(int folderCount) {
		this.folderCount = folderCount;
	}
	
	
}
