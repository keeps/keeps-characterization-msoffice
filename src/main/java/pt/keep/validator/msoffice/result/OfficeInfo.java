package pt.keep.validator.msoffice.result;

import java.util.Date;

import javax.xml.bind.annotation.XmlType;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;


@XmlType(propOrder = {"category","contentStatus","contentType","creator","created","identifier","description","keywords","lastPrinted","modified","revision","subject","title","company","presentationFormat","manager","lineCount","parCount","sectionCount","applicationName","comments","editTime","osVersion","wordCount","pageCount"})
public abstract	class OfficeInfo {
	private String category;
	private String contentStatus;
	private String contentType;
	private String creator;
	private Date created;
	private String identifier;
	private String description;
	private String keywords;
	private Date lastPrinted;
	private Date modified;
	private String revision;
	private String subject;
	private String title;
	private String company;
	private String presentationFormat;
	private String manager;
	private int lineCount;
	private int parCount;
	private int sectionCount;
	private String applicationName;
	private String comments;
	private long editTime;
	private int osVersion;
	private int wordCount;
	private int pageCount;
	public String getCategory() {
		return category;
	}
	public void setCategory(String category) {
		this.category = category;
	}
	public String getContentStatus() {
		return contentStatus;
	}
	public void setContentStatus(String contentStatus) {
		this.contentStatus = contentStatus;
	}
	public String getContentType() {
		return contentType;
	}
	public void setContentType(String contentType) {
		this.contentType = contentType;
	}
	public String getCreator() {
		return creator;
	}
	public void setCreator(String creator) {
		this.creator = creator;
	}
	public Date getCreated() {
		return created;
	}
	public void setCreated(Date created) {
		this.created = created;
	}
	public String getIdentifier() {
		return identifier;
	}
	public void setIdentifier(String identifier) {
		this.identifier = identifier;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public String getKeywords() {
		return keywords;
	}
	public void setKeywords(String keywords) {
		this.keywords = keywords;
	}
	public Date getLastPrinted() {
		return lastPrinted;
	}
	public void setLastPrinted(Date lastPrinted) {
		this.lastPrinted = lastPrinted;
	}
	public Date getModified() {
		return modified;
	}
	public void setModified(Date modified) {
		this.modified = modified;
	}
	public String getRevision() {
		return revision;
	}
	public void setRevision(String revision) {
		this.revision = revision;
	}
	public String getSubject() {
		return subject;
	}
	public void setSubject(String subject) {
		this.subject = subject;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getCompany() {
		return company;
	}
	public void setCompany(String company) {
		this.company = company;
	}
	public String getPresentationFormat() {
		return presentationFormat;
	}
	public void setPresentationFormat(String presentationFormat) {
		this.presentationFormat = presentationFormat;
	}
	public String getManager() {
		return manager;
	}
	public void setManager(String manager) {
		this.manager = manager;
	}
	public int getLineCount() {
		return lineCount;
	}
	public void setLineCount(int lineCount) {
		this.lineCount = lineCount;
	}
	public int getParCount() {
		return parCount;
	}
	public void setParCount(int parCount) {
		this.parCount = parCount;
	}
	public int getSectionCount() {
		return sectionCount;
	}
	public void setSectionCount(int sectionCount) {
		this.sectionCount = sectionCount;
	}
	public String getApplicationName() {
		return applicationName;
	}
	public void setApplicationName(String applicationName) {
		this.applicationName = applicationName;
	}
	public String getComments() {
		return comments;
	}
	public void setComments(String comments) {
		this.comments = comments;
	}
	public long getEditTime() {
		return editTime;
	}
	public void setEditTime(long editTime) {
		this.editTime = editTime;
	}
	public int getOsVersion() {
		return osVersion;
	}
	public void setOsVersion(int osVersion) {
		this.osVersion = osVersion;
	}
	public int getWordCount() {
		return wordCount;
	}
	public void setWordCount(int wordCount) {
		this.wordCount = wordCount;
	}
	public int getPageCount() {
		return pageCount;
	}
	public void setPageCount(int pageCount) {
		this.pageCount = pageCount;
	}
	
	public void addProperties(POIXMLProperties properties) {
		this.setCategory(properties.getCoreProperties().getCategory());
		this.setContentStatus(properties.getCoreProperties().getContentStatus());
		this.setContentType(properties.getCoreProperties().getContentType());
		this.setCreated(properties.getCoreProperties().getCreated());
		this.setCreator(properties.getCoreProperties().getCreator());
		this.setDescription(properties.getCoreProperties().getDescription());
		this.setIdentifier(properties.getCoreProperties().getIdentifier());
		this.setKeywords(properties.getCoreProperties().getKeywords());
		this.setLastPrinted(properties.getCoreProperties().getLastPrinted());
		this.setModified(properties.getCoreProperties().getModified());
		this.setRevision(properties.getCoreProperties().getRevision());
		this.setSubject(properties.getCoreProperties().getSubject());
		this.setTitle(properties.getCoreProperties().getTitle());		
	}

	public void addDocumentSummaryInformation(DocumentSummaryInformation dsi){
		this.setCategory(dsi.getCategory());
		this.setCompany(dsi.getCompany());
		this.setManager(dsi.getManager());
		this.setLineCount(dsi.getLineCount());
		this.setParCount(dsi.getParCount());
		this.setSectionCount(dsi.getSectionCount());
	}

	public void addSummaryInformation(SummaryInformation summaryInformation) {
		this.setApplicationName(summaryInformation.getApplicationName());
		this.setCreator(summaryInformation.getAuthor());
		this.setComments(summaryInformation.getComments());
		this.setCreated(summaryInformation.getCreateDateTime());
		this.setEditTime(summaryInformation.getEditTime());
		this.setKeywords(summaryInformation.getKeywords());
		this.setLastPrinted(summaryInformation.getLastPrinted());
		this.setOsVersion(summaryInformation.getOSVersion());
		this.setPageCount(summaryInformation.getPageCount());
		this.setSectionCount(summaryInformation.getSectionCount());
		this.setWordCount(summaryInformation.getWordCount());
		this.setRevision(summaryInformation.getRevNumber());
		this.setSubject(summaryInformation.getSubject());
		this.setTitle(summaryInformation.getTitle());		
	}
}
