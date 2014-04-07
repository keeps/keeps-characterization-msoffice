package pt.keep.validator.msoffice.result;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "msg")
public class MsgInfo {
	private int numberOfAttachmentFiles;
	private String conversationTopic;
	private String displayBCC;
	private String displayCC;
	private String displayFrom;
	private String displayTo;
	private String[] headers;
	private String htmlBody;
	private String textBody;
	private String messageDate;
	private String recipientEmailAddress;
	private String[] recipientEmailAddressList;
	private String rtfBody;
	private String subject;
	
	@XmlElement
	public String getTextBody() {
		return textBody;
	}
	public void setTextBody(String textBody) {
		this.textBody = textBody;
	}
	@XmlElement
	public int getNumberOfAttachmentFiles() {
		return numberOfAttachmentFiles;
	}
	public void setNumberOfAttachmentFiles(int numberOfAttachmentFiles) {
		this.numberOfAttachmentFiles = numberOfAttachmentFiles;
	}
	@XmlElement
	public String getConversationTopic() {
		return conversationTopic;
	}
	public void setConversationTopic(String conversationTopic) {
		this.conversationTopic = conversationTopic;
	}
	@XmlElement
	public String getDisplayBCC() {
		return displayBCC;
	}
	public void setDisplayBCC(String displayBCC) {
		this.displayBCC = displayBCC;
	}
	@XmlElement
	public String getDisplayCC() {
		return displayCC;
	}
	public void setDisplayCC(String displayCC) {
		this.displayCC = displayCC;
	}
	@XmlElement
	public String getDisplayFrom() {
		return displayFrom;
	}
	public void setDisplayFrom(String displayFrom) {
		this.displayFrom = displayFrom;
	}
	@XmlElement
	public String getDisplayTo() {
		return displayTo;
	}
	public void setDisplayTo(String displayTo) {
		this.displayTo = displayTo;
	}
	@XmlElement
	public String[] getHeaders() {
		return headers;
	}
	public void setHeaders(String[] headers) {
		this.headers = headers;
	}
	@XmlElement
	public String getHtmlBody() {
		return htmlBody;
	}
	public void setHtmlBody(String htmlBody) {
		this.htmlBody = htmlBody;
	}
	@XmlElement
	public String getMessageDate() {
		return messageDate;
	}
	public void setMessageDate(String messageDate) {
		this.messageDate = messageDate;
	}
	@XmlElement
	public String getRecipientEmailAddress() {
		return recipientEmailAddress;
	}
	public void setRecipientEmailAddress(String recipientEmailAddress) {
		this.recipientEmailAddress = recipientEmailAddress;
	}
	@XmlElement
	public String[] getRecipientEmailAddressList() {
		return recipientEmailAddressList;
	}
	public void setRecipientEmailAddressList(String[] recipientEmailAddressList) {
		this.recipientEmailAddressList = recipientEmailAddressList;
	}
	@XmlElement
	public String getRtfBody() {
		return rtfBody;
	}
	public void setRtfBody(String rtfBody) {
		this.rtfBody = rtfBody;
	}
	@XmlElement
	public String getSubject() {
		return subject;
	}
	public void setSubject(String subject) {
		this.subject = subject;
	}
	
	
}
