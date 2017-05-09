package com.example.wps;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;

/**
 * Created by SDD on 2017/4/26.
 */
public class WpsConverterUtil {

	private static Logger logger = LoggerFactory.getLogger(WpsConverterUtil.class);

	// word保存为pdf格式宏，值为17
	private static final int wdFormatPDF = 17;

	public static void convert(String srcFilePath, String destFilePath) {
		long startTime=System.currentTimeMillis();

		File srcFile = new File(srcFilePath);
		File destFile = new File(destFilePath);
		if (destFile.exists()) {
			destFile.delete();
		}

		ActiveXComponent wps = null;
		ActiveXComponent doc = null;

		try {
			wps = new ActiveXComponent("KWPS.Application");

			doc = wps.invokeGetComponent("Documents").invokeGetComponent("Open", new Variant(srcFile.getAbsolutePath()));
			try {
				doc.invoke("SaveAs", new Variant(destFile.getAbsolutePath()), new Variant(wdFormatPDF));
			} catch (Exception e) {
				logger.error("生成PDF失败", e);
			}
		} finally {
			if (doc == null) {
				logger.info("打开文件 " + srcFile.getAbsolutePath() + " 失败");
			} else {
				try {
					logger.info("释放文件 " + srcFile.getAbsolutePath());
					doc.invoke("Close");
					doc.safeRelease();
				} catch (Exception e1) {
					logger.error("释放文件 " + srcFile.getAbsolutePath() + " 失败", e1);
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
					logger.error("释放 WPS 控件失败", e1);
				}
			}
		}
		logger.info("生成花费：" + ((System.currentTimeMillis() - startTime) / 1000) + "s.");
	}
}
