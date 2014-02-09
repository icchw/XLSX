package icchw.xlsx;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.util.IOUtils;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;


public class ZipUtil {

	/**
	 * @param zipFile
	 * @return
	 */
	public static Map<String, ZipEntry> getZipEntryMap(ZipFile zipFile) {
		@SuppressWarnings("unchecked")
		Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zipFile.entries();

		Map<String, ZipEntry> map = new HashMap<String, ZipEntry>();
		while (en.hasMoreElements()) {
			ZipEntry ze = en.nextElement();
			map.put(ze.getName(), ze);
		}
		return map;
	}

	/**
	 * @param zipFile
	 * @param entry
	 * @return
	 * @throws IOException
	 */
	public static InputStream getInputStream(ZipFile zipFile, String entry) throws IOException {
		Map<String, ZipEntry> map = getZipEntryMap(zipFile);
		ZipEntry zipEntry = map.get(entry);
		if (zipEntry == null) {
			return null;
		}
		return zipFile.getInputStream(zipEntry);
	}

	/**
	 * @param zipFile
	 * @param entry
	 * @return
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 */
	public static Document getXmlDocument(ZipFile zipFile, String entry) throws IOException, ParserConfigurationException, SAXException {
		InputStream is = ZipUtil.getInputStream(zipFile, entry);
		if (is == null) {
			return null;
		}
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		return builder.parse(is);
	}

	/**
	 * @param arg0
	 * @param arg1 ex:".xml"
	 * @param document
	 * @return
	 * @throws IOException
	 * @throws TransformerException
	 */
	public static File createTempFileFromDocument(String arg0, String arg1, Document document) throws IOException, TransformerException {
		TransformerFactory transFactory = TransformerFactory.newInstance();
		Transformer transformer = transFactory.newTransformer();

		DOMSource source = new DOMSource(document);
		File tempFile = File.createTempFile(arg0, arg1);
		tempFile.deleteOnExit();
		FileOutputStream os = new FileOutputStream(tempFile);
		StreamResult result = new StreamResult(os);
		transformer.transform(source, result);
		if (os != null) {
			os.close();
		}

		return tempFile;

	}


	/**
	 * substitute files of templateZip with files of the Map
	 * @param templateFile
	 * @param os
	 * @throws ZipException
	 * @throws IOException
	 */
	public static void substitute(ZipFile templateZip, Map<String, File> substituteMap, OutputStream os) throws ZipException, IOException {
		ZipOutputStream zos = new ZipOutputStream(os);

		@SuppressWarnings("unchecked")
		Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) templateZip.entries();

		Set<String> entries = substituteMap.keySet();

		// copy files of templateFile
		while (en.hasMoreElements()) {
			ZipEntry ze = en.nextElement();
			if (!entries.contains(ze.getName())) {
				zos.putNextEntry(new ZipEntry(ze.getName()));
				InputStream is = templateZip.getInputStream(ze);
				IOUtils.copy(is, zos);
				is.close();
			}
		}

		// substitute xmls
		for (String entry : entries) {
			zos.putNextEntry(new ZipEntry(entry));
			InputStream is = new FileInputStream(substituteMap.get(entry));
			IOUtils.copy(is, zos);
			is.close();
			zos.closeEntry();
		}
		zos.close();
		templateZip.close();
	}

}
