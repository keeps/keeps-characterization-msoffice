package pt.keep.validator;

import java.io.File;
import java.io.FileInputStream;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Marshaller;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.GnuParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.poi.extractor.ExtractorFactory;
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
		boolean isValidDocx=false,isValidPptx=false,isValidXlsx=false;
		
		XWPFDocument docx = null;
		XSSFWorkbook xlsx = null;
		XMLSlideShow pptx = null; 
		
		try {
			docx = new XWPFDocument(new FileInputStream(f));
			isValidDocx=true;
		}catch (Exception e) {
		}
		
		try {
			xlsx = new XSSFWorkbook(new FileInputStream(f));
			isValidXlsx=true;
		}catch (Exception e) {
		}
		
		try {
			pptx = new XMLSlideShow(new FileInputStream(f));
			isValidPptx=true;
		}catch (Exception e) {
		}
		
		if(isValidDocx){
			res = extractInfoFromValidDocx(docx);
		}else if(isValidPptx){
			res = extractInfoFromValidPptx(pptx);
		}else if(isValidXlsx){
			res = extractInfoFromValidXlsx(xlsx);
		}else{
			res = notValidFile(f);
		}
		return res;
	}
	
	private Result notValidFile(File f) {
		Result res = new Result();
		res.setValid(false);
		try {
			XWPFDocument docx = new XWPFDocument(new FileInputStream(f));
		}catch (Exception e) {
			WordInfo word = new WordInfo();
			word.setWordError(e.getMessage());
			res.setWord(word);
		}
		
		try {
			XSSFWorkbook xlsx = new XSSFWorkbook(new FileInputStream(f));
		}catch (Exception e) {
			ExcelInfo excel = new ExcelInfo();
			excel.setExcelError(e.getMessage());
			res.setExcel(excel);
		}
		
		try {
			XMLSlideShow pptx = new XMLSlideShow(new FileInputStream(f));
		}catch (Exception e) {
			PowerpointInfo powerpoint = new PowerpointInfo();
			powerpoint.setPowerpointError(e.getMessage());
			res.setPowerpoint(powerpoint);
		}
		
		
		
		return res;
	}

	private Result extractInfoFromValidXlsx(XSSFWorkbook xlsx) {
		Result res = new Result();
		res.setValid(true);
		res.setValidationError(null);
		ExcelInfo excel = new ExcelInfo();
		excel.setNumberOfSheets(xlsx.getNumberOfSheets());
		res.setExcel(excel);
		return res;
	}

	private Result extractInfoFromValidPptx(XMLSlideShow pptx) {
		Result res = new Result();
		res.setValid(true);
		res.setValidationError(null);
		PowerpointInfo powerpoint = new PowerpointInfo();
		powerpoint.setNumberOfSlides(pptx.getSlides().length);
		return res;
	}

	private Result extractInfoFromValidDocx(XWPFDocument docx) {
		Result res = new Result();
		res.setValid(true);
		res.setValidationError(null);
		WordInfo word = new WordInfo();
		res.setWord(word);
		return res;
	}

	private void printHelp(Options opts) {
		HelpFormatter formatter = new HelpFormatter();
		formatter.printHelp("report", opts);
	}
	
	private void printVersion() {
		System.out.println(version);
	}

	
    public static void main( String[] args )
    {
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




