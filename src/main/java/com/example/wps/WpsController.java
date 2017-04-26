package com.example.wps;

import org.apache.commons.io.FileUtils;
import org.aspectj.util.FileUtil;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.io.IOException;
import java.util.UUID;

/**
 * Created by SDD on 2017/4/26.
 */
@RestController
public class WpsController {

	@Value("${file.base.path}")
	private String fileBasePaht;

	@Value("${file.wps.src.path}")
	private String fileWpsPath;
	@Value("${file.wps.pdf.path}")
	private String fileWps2PdfPath;

	@RequestMapping("/wps2pdf/{srcFileName:.+}")
	@ResponseBody
	public String toPdf(@PathVariable("srcFileName") String srcFileName) throws IOException {
		// 文件的绝对路径
		String srcAbsolutePath = fileBasePaht + File.separator + fileWpsPath + File.separator + srcFileName;

		// 输出的文件绝对路径
		String descAbsolutePath = fileBasePaht + File.separator + fileWps2PdfPath + File.separator + this.getFileNameNoEx(srcFileName);
		//String descAbsolutePath = fileBasePaht + File.separator + fileWps2PdfPath + File.separator + System.currentTimeMillis()+".pdf";

		// 转换PDF
		WpsConverterUtil.doc2pdf(srcAbsolutePath, descAbsolutePath);

		return descAbsolutePath;
	}

	/*
 * Java文件操作 获取不带扩展名的文件名
 *
 *  Created on: 2011-8-2
 *      Author: blueeagle
 */
	public static String getFileNameNoEx(String filename) {
		if ((filename != null) && (filename.length() > 0)) {
			int dot = filename.lastIndexOf('.');
			if ((dot > -1) && (dot < (filename.length()))) {
				return filename.substring(0, dot);
			}
		}
		return filename;
	}
}
