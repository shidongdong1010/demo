package com.example.wps;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.DispatchEvents;
import com.jacob.com.Variant;
import org.junit.Test;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by SDD on 2017/3/15.
 */
public class Doc2Pdf {


	@Test
	public void toPdf(){
		String sourceFile = "D:/file/zjdgxy.docx";
		String destFile =  "D:/file/zjdgxy_out_wps.pdf";
		this.convert(sourceFile, destFile);
	}

	public static Converter newConverter(String name) {
		if (name.equals("wps")) {
			return new Wps();
		} else if (name.equals("pdfcreator")) {
			return new PdfCreator();
		}
		return null;
	}

	public synchronized static boolean convert(String word, String pdf) {
		return newConverter("pdfcreator").convert(word, pdf);
	}

	public static interface Converter {

		public boolean convert(String word, String pdf);
	}

	public static class Wps implements Converter {

		public synchronized boolean convert(String word, String pdf) {
			File pdfFile = new File(pdf);
			File wordFile = new File(word);
			ActiveXComponent wps = null;
			try {
				wps = new ActiveXComponent("wps.application");
				ActiveXComponent doc = wps.invokeGetComponent("Documents").invokeGetComponent("Open", new Variant(wordFile.getAbsolutePath()));
				doc.invoke("ExportPdf", new Variant(pdfFile.getAbsolutePath()));
				doc.invoke("Close");
				doc.safeRelease();
			} catch (Exception ex) {
				Logger.getLogger(Doc2Pdf.class.getName()).log(Level.SEVERE, null, ex);
				return false;
			} catch (Error ex) {
				Logger.getLogger(Doc2Pdf.class.getName()).log(Level.SEVERE, null, ex);
				return false;
			} finally {
				if (wps != null) {
					wps.invoke("Terminate");
					wps.safeRelease();
				}
			}
			return true;
		}
	}

	public static class PdfCreator implements Converter {

		public static final int STATUS_IN_PROGRESS = 2;
		public static final int STATUS_WITH_ERRORS = 1;
		public static final int STATUS_READY = 0;
		private ActiveXComponent pdfCreator;
		private DispatchEvents dispatcher;
		private volatile int status;
		private Variant defaultPrinter;

		private void init() {
			pdfCreator = new ActiveXComponent("PDFCreator.clsPDFCreator");
			dispatcher = new DispatchEvents(pdfCreator, this);
			pdfCreator.setProperty("cVisible", new Variant(false));
			pdfCreator.invoke("cStart", new Variant[]{new Variant("/NoProcessingAtStartup"), new Variant(true)});
			setCOption("UseAutosave", 1);
			setCOption("UseAutosaveDirectory", 1);
			setCOption("AutosaveFormat", 0);  // 0 = PDF
			defaultPrinter = pdfCreator.getProperty("cDefaultPrinter");
			status = STATUS_IN_PROGRESS;
			pdfCreator.setProperty("cDefaultprinter", "PDFCreator");
			pdfCreator.invoke("cClearCache");
			pdfCreator.setProperty("cPrinterStop", false);
		}

		private void setCOption(String property, Object value) {
			Dispatch.invoke(pdfCreator, "cOption", Dispatch.Put, new Object[]{property, value}, new int[2]);
		}

		private void close() {
			if (pdfCreator != null) {
				pdfCreator.setProperty("cDefaultprinter", defaultPrinter);
				pdfCreator.invoke("cClearCache");
				pdfCreator.setProperty("cPrinterStop", true);
				pdfCreator.invoke("cClose");
				pdfCreator.safeRelease();
				pdfCreator = null;
			}
			if (dispatcher != null) {
				dispatcher.safeRelease();
				dispatcher = null;
			}
		}

		public synchronized boolean convert(String word, String pdf) {
			File pdfFile = new File(pdf);
			File wordFile = new File(word);
			try {
				init();
				setCOption("AutosaveDirectory", pdfFile.getParentFile().getAbsolutePath());
				if (pdfFile.exists()) {
					pdfFile.delete();
				}
				pdfCreator.invoke("cPrintfile", wordFile.getAbsolutePath());
				int seconds = 0;
				while (isInProcess()) {
					seconds++;
					if (seconds > 30) { // timeout
						throw new Exception("convertion timeout!");
					}
					Thread.sleep(1000);
				}
				if (isWithErrors()) return false;
				// 由于转换前设置cOption的AutosaveFilename不能保证输出的文件名与设置的相同（pdfcreator会加入/修改后缀名）
				// 所以这里让pdfcreator使用自动生成的文件名进行输出，然后在保存后将其重命名为目标文件名
				File outputFile = new File(pdfCreator.getPropertyAsString("cOutputFilename"));
				if (outputFile.exists()) {
					outputFile.renameTo(pdfFile);
				}
			} catch (InterruptedException ex) {
				Logger.getLogger(Doc2Pdf.class.getName()).log(Level.SEVERE, null, ex);
				return false;
			} catch (Exception ex) {
				Logger.getLogger(Doc2Pdf.class.getName()).log(Level.SEVERE, null, ex);
				return false;
			} catch (Error ex) {
				Logger.getLogger(Doc2Pdf.class.getName()).log(Level.SEVERE, null, ex);
				return false;
			} finally {
				close();
			}
			return true;
		}

		private boolean isInProcess() {
			return status == STATUS_IN_PROGRESS;
		}

		private boolean isWithErrors() {
			return status == STATUS_WITH_ERRORS;
		}

		// eReady event
		public void eReady(Variant[] args) {
			status = STATUS_READY;
		}

		// eError event
		public void eError(Variant[] args) {
			status = STATUS_WITH_ERRORS;
		}
	}

}
