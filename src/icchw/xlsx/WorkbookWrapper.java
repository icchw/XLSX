package icchw.xlsx;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;


/**
 * using inlineStr instead of sharedStrings
 *
 */
public class WorkbookWrapper {

	public static final String STYLE_DATE = "STYLE_DATE";

	/** template Zipファイル */
	private ZipFile templateZip;

	/** templateファイルをworkbookに変換したもの */
	private XSSFWorkbook templateWb;

	/** 置換用xmlのMap */
	Map<String, File> substituteMap = new HashMap<String, File>();

	/**
	 * 物理ファイルをtemplateととして値を書き込む場合
	 */
	public WorkbookWrapper(File file) throws FileNotFoundException, IOException {
		super();
		this.templateZip = new ZipFile(file);
		this.templateWb = new XSSFWorkbook(new FileInputStream(file));
	}

	/**
	 * memory上のXSSFWorkbookをtemplateととして値を書き込む場合
	 */
	public WorkbookWrapper(XSSFWorkbook wb) throws IOException {
		super();
		this.templateWb = wb;

		File templateFile = File.createTempFile("template", "xlsx");
		templateFile.deleteOnExit();
		templateWb.write(new FileOutputStream(templateFile));
		this.templateZip = new ZipFile(templateFile);
	}


	/**
	 * templateのxlsxファイルに対して、dataを書き込んだsheetのみxmlファイルを置換して、OutputStreamに書き込む
	 * @param os
	 * @throws IOException
	 */
	public void write(OutputStream os) throws IOException {
		ZipUtil.substitute(templateZip, substituteMap, os);
	}

	/**
	 * return xml file name for the sheet
	 * @param sheetName
	 * @return
	 */
	private String getSheetXmlName(String sheetName) {
		return "sheet" + (templateWb.getSheetIndex(sheetName) + 1);
	}

	/**
	 * 当該sheetのXMLをDocument形式で取得する
	 * @param sheetName
	 * @return
	 * @throws ZipException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 */
	private Document getSheetXML(String sheetName) throws ZipException, IOException, ParserConfigurationException, SAXException {
		XSSFSheet sheet = templateWb.getSheet(sheetName);
		if (sheet == null) {
			return null;
		}
		return ZipUtil.getXmlDocument(templateZip, getEntry(sheet));
	}

	/**
	 * sheetからzip書込み用のEntryを取得する
	 * @param sheet
	 * @return
	 */
	private String getEntry(XSSFSheet sheet) {
		return sheet.getPackagePart().getPartName().getName().substring(1);
	}

	/**
	 * sheetにdatasを書き込む<br/>
	 * header行があれば残して、後続に追加する<br/>
	 * header行の次の行の書式をコピーして使用する<br/>
	 * @param sheetName
	 * @param prepareData
	 * @throws IOException
	 * @throws SAXException
	 * @throws ParserConfigurationException
	 * @throws TransformerException
	 */
	public void writeSheet(String sheetName, List<? extends XlsxWritable> datas) throws IOException, ParserConfigurationException, SAXException, TransformerException {
		Document sheetXml = getSheetXML(sheetName);
		if (sheetXml == null) {
			throw new IllegalArgumentException("No Such Sheet");
		}
		XlsxWriter.addDataToSheet(sheetXml, datas);
		substituteMap.put(getEntry(templateWb.getSheet(sheetName)), ZipUtil.createTempFileFromDocument(getSheetXmlName(sheetName), ".xml", sheetXml));
	}
}
