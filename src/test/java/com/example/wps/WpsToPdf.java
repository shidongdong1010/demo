package com.example.wps;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;

/**
 * Created by SDD on 2017/4/26.
 */
@SpringBootTest
public class WpsToPdf {
	private Logger logger = LoggerFactory.getLogger(this.getClass());

	// word保存为pdf格式宏，值为17
	private static final int wdFormatPDF = 17;

	public static void main(String[] args){
		long startTime=System.currentTimeMillis();
		String sourceFile = args[0];
		String destFile =  args[1];

		new WpsToPdf().doc2pdf(sourceFile, destFile);

		System.out.println("Generate pdf with " + ((System.currentTimeMillis() - startTime) / 1000) + " s.");
	}

	@Test
	public void toPdf(){
		long startTime=System.currentTimeMillis();

		String sourceFile = "D:/file/zjdgxy.docx";
		String destFile =  "D:/file/zjdgxy_out_wps.pdf";
		this.doc2pdf(sourceFile, destFile);

		System.out.println("Generate ooxml.pdf with " + ((System.currentTimeMillis() - startTime) / 1000) + " s.");
	}

	protected void doc2pdf(String srcFilePath, String pdfFilePath) {
		File wordFile = new File(srcFilePath);
		File pdfFile = new File(pdfFilePath);

		ActiveXComponent wps = null;
		ActiveXComponent doc = null;

		try {
			wps = new ActiveXComponent("KWPS.Application");

//                Dispatch docs = wps.getProperty("Documents").toDispatch();
//                Dispatch d = Dispatch.call(docs, "Open", wordFile.getAbsolutePath(), false, true).toDispatch();
//                Dispatch.call(d, "SaveAs", pdfFile.getAbsolutePath(), 17);
//                Dispatch.call(d, "Close", false);

			doc = wps.invokeGetComponent("Documents").invokeGetComponent("Open", new Variant(wordFile.getAbsolutePath()));
			try {
				doc.invoke("SaveAs", new Variant(pdfFile.getAbsolutePath()), new Variant(wdFormatPDF));
			} catch (Exception e) {
				logger.warn("生成PDF失败");
				e.printStackTrace();
			}

			/*File saveAsFile = new File("D:/file/saveasfile.doc");
			try {
				doc.invoke("SaveAs", saveAsFile.getAbsolutePath());
				logger.info("成功另存为" + saveAsFile.getAbsolutePath());
			} catch (Exception e) {
				logger.info("另存为" + saveAsFile.getAbsolutePath() + "失败");
				e.printStackTrace();
			}*/
		} finally {
			if (doc == null) {
				logger.info("打开文件 " + wordFile.getAbsolutePath() + " 失败");
			} else {
				try {
					logger.info("释放文件 " + wordFile.getAbsolutePath());
					doc.invoke("Close");
					doc.safeRelease();
				} catch (Exception e1) {
					logger.info("释放文件 " + wordFile.getAbsolutePath() + " 失败");
				}
			}

			if (wps == null) {
				logger.info("加载 WPS 控件失败");
			} else {
				try {
					logger.info("释放 WPS 控件");
					wps.invoke("Quit");
					wps.safeRelease();
				} catch (Exception e1) {
					logger.info("释放 WPS 控件失败");
				}
			}
		}
	}
}
