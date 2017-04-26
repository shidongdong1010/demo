package com.example.msoffice;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

/**
 * Created by SDD on 2017/4/26.
 */
@SpringBootTest
public class MSOfficeToPdf {

	private static final int wdFormatPDF = 17;

	@Test
	public void toPdf(){
		String sourceFile = "D:/file/zjdgxy.docx";
		String destFile =  "D:/file/zjdgxy_out_msoffice.pdf";
		this.doc2pdf(sourceFile, destFile);
	}

	protected boolean doc2pdf(String srcFilePath, String pdfFilePath) {
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			ComThread.InitSTA();
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);
			Dispatch docs = app.getProperty("Documents").toDispatch();
			doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
					new Object[] { srcFilePath,
							new Variant(false),
							new Variant(true),//是否只读
							new Variant(false),
							new Variant("pwd") },
					new int[1]).toDispatch();
			// Dispatch.put(doc, "Compatibility", false);  //兼容性检查,为特定值false不正确
			Dispatch.put(doc, "RemovePersonalInformation", false);
			Dispatch.call(doc, "ExportAsFixedFormat", pdfFilePath, wdFormatPDF); // word保存为pdf格式宏，值为17
			return true; // set flag true;
		}finally {
			if (doc != null) {
				Dispatch.call(doc, "Close", false);
			}
			if (app != null) {
				app.invoke("Quit", 0);
			}
			ComThread.Release();
		}
	}
}
