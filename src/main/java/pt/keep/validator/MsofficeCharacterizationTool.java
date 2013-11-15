package pt.keep.validator;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.Vector;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Marshaller;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.GnuParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.Chunks;
import org.apache.poi.hsmf.datatypes.StringChunk;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

import pt.keep.validator.result.ExcelInfo;
import pt.keep.validator.result.MsgInfo;
import pt.keep.validator.result.PowerpointInfo;
import pt.keep.validator.result.PstInfo;
import pt.keep.validator.result.Result;
import pt.keep.validator.result.WordInfo;

/**
 * Hello world!
 *
 */
public class MsofficeCharacterizationTool 
{
	private static String version = "1.0";

	public void run(File f) {
		try {
			Result res = process(f);
			JAXBContext jaxbContext = JAXBContext.newInstance(Result.class);
			Marshaller jaxbMarshaller = jaxbContext.createMarshaller();
			jaxbMarshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
			jaxbMarshaller.marshal(res, System.out);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public Result process(File f) {
		Result res = null;
		boolean isValidDocx=false,isValidPptx=false,isValidXlsx=false,isValidDoc=false,isValidPpt=false,isValidXls=false, isValidPst=false, isValidMsg=false;
		
		XWPFDocument docx = null;
		XSSFWorkbook xlsx = null;
		XMLSlideShow pptx = null; 
		HWPFDocument doc = null;
		HSSFWorkbook xls = null;
		HSLFSlideShow ppt = null;
		PSTFile pst = null;
		MAPIMessage msg = null;
		
		List<String> errors = new ArrayList<String>();
		
		try {
			docx = new XWPFDocument(new FileInputStream(f));
			isValidDocx=true;
		}catch (Exception e) {
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		
		try {
			xlsx = new XSSFWorkbook(new FileInputStream(f));
			isValidXlsx=true;
		}catch (Exception e) {
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		
		try {
			pptx = new XMLSlideShow(new FileInputStream(f));
			isValidPptx=true;
		}catch (Exception e) {
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		try{
			doc = new HWPFDocument(new FileInputStream(f));
			isValidDoc=true;
		}catch(Exception e){
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		try{
			xls = new HSSFWorkbook(new FileInputStream(f));
			isValidXls=true;
		}catch(Exception e){
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		try{
			ppt = new HSLFSlideShow(new FileInputStream(f));
			isValidPpt=true;
		}catch(Exception e){
			String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
		}
		
		 try {
            pst = new PSTFile(f);
            isValidPst=true;
        } catch (Exception e) {
        	String error = e.getMessage();
			if(error==null){
				error = "null";
			}
			if(!errors.contains(error)){
				errors.add(error);
			}
        }
		 
		 try{
			 msg = new MAPIMessage(new FileInputStream(f));
			 isValidMsg=true;
		 }catch(Exception e){
			 String error = e.getMessage();
				if(error==null){
					error = "null";
				}
				if(!errors.contains(error)){
					errors.add(error);
				}
		 }
		if(isValidDocx){
			res = extractInfoFromValidDocx(docx);
		}else if(isValidPptx){
			res = extractInfoFromValidPptx(pptx);
		}else if(isValidXlsx){
			res = extractInfoFromValidXlsx(xlsx);
		}else if(isValidDoc){
			res = extractInfoFromValidDoc(doc);
		}else if(isValidXls){
			res = extractInfoFromValidXls(xls);
		}else if(isValidPpt){
			res = extractInfoFromValidPpt(ppt);
		}else if(isValidPst){
			res = extractInfoFromValidPst(pst);
		}else if(isValidMsg){
			res = extractInfoFromValidMsg(msg);
		}else{
			res = notValidFile(errors);
		}
		return res;
	}

	
	private Result extractInfoFromValidMsg(MAPIMessage msg) {
		Result res = new Result();
		res.setValid(true);
		try{
			
			MsgInfo info = new MsgInfo();
			info.setConversationTopic(msg.getConversationTopic());
			info.setDisplayBCC(msg.getDisplayBCC());
			info.setDisplayCC(msg.getDisplayCC());
			info.setDisplayFrom(msg.getDisplayFrom());
			info.setDisplayTo(msg.getDisplayTo());
			info.setHeaders(msg.getHeaders());
			try{
				info.setHtmlBody(msg.getHtmlBody());
			}catch(Exception e1){
				
			}
			info.setMessageDate(msg.getMessageDate().toString());
			info.setNumberOfAttachmentFiles(msg.getAttachmentFiles().length);
			info.setRecipientEmailAddress(msg.getRecipientEmailAddress());
			info.setRecipientEmailAddressList(msg.getRecipientEmailAddressList());
			info.setRtfBody(msg.getRtfBody());
			info.setTextBody(msg.getTextBody());
			info.setSubject(msg.getSubject());
			res.setMsg(info);
		}catch(Exception e){
			e.printStackTrace();
		}
		return res;

	}

	private Result extractInfoFromValidPst(PSTFile pst) {
		Result res = new Result();
		res.setValid(true);
		
		try{
			PstInfo info = new PstInfo();
			info.setEmailCount(countMails(pst.getRootFolder(),0));
			info.setFolderCount(countFolder(pst.getRootFolder(), 0));
			res.setPst(info);
		}catch(Exception e){
			e.printStackTrace();
		}
		return res;
	}

	public int countMails(PSTFolder folder, int counter) throws PSTException, java.io.IOException{
		counter+=folder.getContentCount();
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
            	counter+=countMails(childFolder,counter);
            }
        }
       return counter;
	}
	
	public int countFolder(PSTFolder folder, int counter) throws PSTException, java.io.IOException{
		counter++;
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
            	counter+=countFolder(childFolder,counter);
            }
        }
       return counter;
	}

	

	
	
	
	
	
	private Result notValidFile(List<String> errors) {
		Result res = new Result();
		res.setValid(false);
		res.setValidationErrors(errors);
		return res;
	}

	private Result extractInfoFromValidXlsx(XSSFWorkbook xlsx) {
		Result res = new Result();
		res.setValid(true);
		ExcelInfo excel = new ExcelInfo();
		excel.setNumberOfSheets(xlsx.getNumberOfSheets());
		excel.addProperties(xlsx.getProperties());
		res.setExcel(excel);
		return res;
	}

	private Result extractInfoFromValidPptx(XMLSlideShow pptx) {
		Result res = new Result();
		res.setValid(true);
		PowerpointInfo powerpoint = new PowerpointInfo();
		powerpoint.setNumberOfSlides(pptx.getSlides().length);
		powerpoint.addProperties(pptx.getProperties());
		res.setPowerpoint(powerpoint);
		return res;
	}

	private Result extractInfoFromValidDocx(XWPFDocument docx) {
		Result res = new Result();
		res.setValid(true);
		WordInfo word = new WordInfo();
		word.addProperties(docx.getProperties());
		res.setWord(word);
		return res;
	}
	
	private Result extractInfoFromValidXls(HSSFWorkbook xls) {
		Result res = new Result();
		res.setValid(true);
		ExcelInfo excel = new ExcelInfo();
		excel.setNumberOfSheets(xls.getNumberOfSheets());
		excel.addDocumentSummaryInformation(xls.getDocumentSummaryInformation());
		excel.addSummaryInformation(xls.getSummaryInformation());
		res.setExcel(excel);
		return res;
	}

	private Result extractInfoFromValidPpt(HSLFSlideShow ppt) {
		Result res = new Result();
		res.setValid(true);
		PowerpointInfo powerpoint = new PowerpointInfo();
		powerpoint.setNumberOfSlides(ppt.getDocumentSummaryInformation().getSlideCount());
		powerpoint.addDocumentSummaryInformation(ppt.getDocumentSummaryInformation());
		res.setPowerpoint(powerpoint);
		return res;
	}

	private Result extractInfoFromValidDoc(HWPFDocument doc) {
		Result res = new Result();
		res.setValid(true);
		WordInfo word = new WordInfo();
		word.addDocumentSummaryInformation(doc.getDocumentSummaryInformation());
		res.setWord(word);
		return res;
	}

	private void printHelp(Options opts) {
		HelpFormatter formatter = new HelpFormatter();
		formatter.printHelp("java -jar [jarFile]", opts);
	}
	
	private void printVersion() {
		System.out.println(version);
	}

	
    public static void main( String[] args )
    {
    	
    	Logger.getRootLogger().setLevel(Level.OFF);
    	try {
			MsofficeCharacterizationTool ov = new MsofficeCharacterizationTool();
			
			Options options = new Options();
			options.addOption("f", true, "file to analyze");
			options.addOption("v", false, "print this tool version");
			options.addOption("h", false, "print this message");

			CommandLineParser parser = new GnuParser();
			CommandLine cmd = parser.parse(options, args);

			if (cmd.hasOption("h")) {
				ov.printHelp(options);
				System.exit(0);
			}
			
			if (cmd.hasOption("v")) {
				ov.printVersion();
				System.exit(0);
			}

			if (!cmd.hasOption("f")) {
				ov.printHelp(options);
				System.exit(0);
			}

			File f = new File(cmd.getOptionValue("f"));
			if (!f.exists()) {
				System.out.println("File doesn't exist");
				System.exit(0);
			}
			ov.run(f);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
}




