package com.example.poi;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Created by SDD on 2017/4/26.
 */
@SpringBootTest
public class PoiWordToPdf {

	@Test
	public void makePdfByXcode(){
		long startTime=System.currentTimeMillis();
		String sourceFile = "D:/file/zjdgxy.docx";
		String destFile =  "D:/file/zjdgxy_out_poi.pdf";
		try {
			XWPFDocument document=new XWPFDocument(new FileInputStream(new File(sourceFile)));
			//    document.setParagraph(new Pa );
			File outFile=new File(destFile);
			outFile.getParentFile().mkdirs();
			OutputStream out=new FileOutputStream(outFile);
			//    IFontProvider fontProvider = new AbstractFontRegistry();
			PdfOptions options= PdfOptions.create();  //gb2312
			PdfConverter.getInstance().convert(document,out,options);

		} catch (  Exception e) {
			e.printStackTrace();
		}
		System.out.println("Generate ooxml.pdf with " + (System.currentTimeMillis() - startTime) + " ms.");
	}
}
