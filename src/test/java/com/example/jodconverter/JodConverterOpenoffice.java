package com.example.jodconverter;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.StreamOpenOfficeDocumentConverter;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.IOException;
import java.net.ConnectException;

/**
 * Created by SDD on 2017/3/14.
 */
@SpringBootTest
public class JodConverterOpenoffice {

	@Test
	public void test(){
		// 此处的目录应该为服务器上的目录
		//String sourceFile = "/root/work/openoffice/11_out.docx";
		//String destFile =  "/root/work/openoffice/11_out.pdf";
		String sourceFile = "D:/file/11_out.doc";
		String destFile =  "D:/file/11_out_openoffice.pdf";

		//String sourceFile = "D:/file/33.doc";
		//String destFile =  "D:/file/33.pdf";
		JodConverterOpenoffice.office2PDF(sourceFile, destFile);

	}

	/**
	 * 将Office文档转换为PDF. 运行该函数需要用到OpenOffice, OpenOffice下载地址为
	 * http://www.openoffice.org/
	 *
	 * <pre>
	 * 方法示例:
	 * String sourcePath = "F:\\office\\source.doc";
	 * String destFile = "F:\\pdf\\dest.pdf";
	 * Converter.office2PDF(sourcePath, destFile);
	 * </pre>
	 *
	 * @param sourceFile
	 *            源文件, 绝对路径. 可以是Office2003-2007全部格式的文档, Office2010的没测试. 包括.doc,
	 *            .docx, .xls, .xlsx, .ppt, .pptx等. 示例: F:\\office\\source.doc
	 * @param destFile
	 *            目标文件. 绝对路径. 示例: F:\\pdf\\dest.pdf
	 * @return 操作成功与否的提示信息. 如果返回 -1, 表示找不到源文件, 或url.properties配置错误; 如果返回 0,
	 *         则表示操作成功; 返回1, 则表示转换失败
	 */
	public static boolean office2PDF(String sourceFile, String destFile) {
		try {
			File inputFile = new File(sourceFile);
			File outputFile = new File(destFile);

			// connect to an OpenOffice.org instance running on port 8100
			OpenOfficeConnection connection = new SocketOpenOfficeConnection("192.168.129.155", 8100);
			//OpenOfficeConnection connection = new SocketOpenOfficeConnection("192.168.129.156", 8100);
			connection.connect();

			// convert
			// 本地调用
			// DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
			// 远程调用
			DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection);
			converter.convert(inputFile, outputFile);

			// close the connection
			connection.disconnect();

			return true;
		} catch (ConnectException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return false;
	}
}
