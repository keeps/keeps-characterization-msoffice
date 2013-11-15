package pt.keep.validator;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import pt.keep.validator.result.ExcelInfo;
import pt.keep.validator.result.PowerpointInfo;
import pt.keep.validator.result.Result;
import pt.keep.validator.result.WordInfo;

/**
 * Hello world!
 *
 */
public class OoxmlValidator 
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
		boolean isValidDocx=false,isValidPptx=false,isValidXlsx=false,isValidDoc=false,isValidPpt=false,isValidXls=false;
		
		XWPFDocument docx = null;
		XSSFWorkbook xlsx = null;
		XMLSlideShow pptx = null; 
		HWPFDocument doc = null;
		HSSFWorkbook xls = null;
		HSLFSlideShow ppt = null;
		
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
		}else{
			res = notValidFile(errors);
		}
		return res;
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
			OoxmlValidator ov = new OoxmlValidator();
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




