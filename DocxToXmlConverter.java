package com.lls.wec.local.document.common.helper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Base64;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.opc.internal.PackagePropertiesPart;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import lombok.extern.slf4j.Slf4j;

/**
 * docx转为Office Open XML
 * @author panxiao
 *
 */
@Slf4j
public class DocxToXmlConverter {
	
	private static Corrector corrector = new Corrector();

	public void docxToXml(String srcPath, String descPath) {
		XMLWriter writer = null;
		try {
			Document document = DocumentHelper
					.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
							+ "<?mso-application progid=\"Word.Document\"?>\r\n"
							+ "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"></pkg:package>");

			OPCPackage pkg = ZipPackage.open(srcPath);
			ArrayList<PackagePart> parts = pkg.getParts();
			for (PackagePart packagePart : parts) {
				Element partE = DocumentHelper.createElement("pkg:part");
				partE.addAttribute("pkg:contentType", packagePart.getContentType());
				partE.addAttribute("pkg:name", packagePart.getPartName().getName());
				if ("image/png".equals(packagePart.getContentType())) {
					partE.addAttribute("pkg:compression", "store");
				}
				setPart(partE, packagePart);
				document.getRootElement().add(partE);
			}
			OutputFormat format = OutputFormat.createPrettyPrint();
			writer = new XMLWriter(new FileOutputStream(descPath), format);
			writer.write(document);
		} catch (DocumentException | IOException | InvalidFormatException e) {
			log.error("docx to xml error", e);
		} finally {
			if (writer != null)
				try {
					writer.close();
				} catch (IOException e) {
					log.error("stream close error", e);
				}
		}
		corrector.correctTemplate(StringUtils.substringBeforeLast(descPath, File.separator), StringUtils.substringAfterLast(descPath, File.separator));
	}

	private void setPart(Element part, PackagePart packagePart) throws DocumentException, IOException {
		String contentType = packagePart.getContentType();

		switch (contentType) {
		case "image/png":
		case "image/x-emf":
		case "application/vnd.openxmlformats-officedocument.oleObject": {
			byte[] d = IOUtils.toByteArray(packagePart.getInputStream());
			Element data = DocumentHelper.createElement("pkg:binaryData");
			data.setText(Base64.getEncoder().encodeToString(d));
			part.add(data);
			break;
		}
		default:
			if (packagePart instanceof PackagePropertiesPart) {
				Document properties = DocumentHelper
						.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> \r\n"
								+ "<pkg:xmlData xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"> \r\n"
								+ "    <cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \r\n"
								+ "        xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \r\n"
								+ "        xmlns:dcterms=\"http://purl.org/dc/terms/\" \r\n"
								+ "        xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" \r\n"
								+ "        xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n"
								+ "        <dc:creator>lenovo</dc:creator>\r\n"
								+ "        <cp:lastModifiedBy>lenovo</cp:lastModifiedBy>\r\n"
								+ "        <cp:revision>2</cp:revision>\r\n"
								+ "        <dcterms:created xsi:type=\"dcterms:W3CDTF\">2019-03-01T08:06:00Z</dcterms:created>\r\n"
								+ "        <dcterms:modified xsi:type=\"dcterms:W3CDTF\">2019-03-01T08:06:00Z</dcterms:modified>\r\n"
								+ "    </cp:coreProperties>\r\n" + "</pkg:xmlData>");
				part.add(properties.getRootElement());
			} else {
				byte[] d = IOUtils.toByteArray(packagePart.getInputStream());
				Document source = DocumentHelper.parseText(new String(d, "UTF-8"));
				Element properties = source.getRootElement();
				Element data = DocumentHelper.createElement("pkg:xmlData");
				data.add(properties);
				part.add(data);
			}
			break;
		}
	}

	public static void main(String[] args) {
		new DocxToXmlConverter().docxToXml("d://trade_receipt_mortgage.docx", "d://trade_receipt_mortgage-b.xml");

	}
}
