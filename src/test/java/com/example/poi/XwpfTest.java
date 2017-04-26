package com.example.poi;

import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by SDD on 2017/3/14.
 */
@SpringBootTest
public class XwpfTest {
	/**
	 * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
	 *
	 * @throws Exception
	 */
	@Test
	public void testTemplateWrite() throws Exception {
		Map<String, Object> params = new HashMap<String, Object>();
		params.put("investUserName", "全体投资人");
		params.put("investLoginName", "shidongdong");
		params.put("loanAmount", "1000000.00");
		params.put("investIdCard", "34088119890901011X");
		params.put("investBankNo", "324234234234234234234");
		params.put("loanUserName", "魏楠汽车销售有限公司");
		params.put("loanLoginName", "weinan");
		params.put("applyNo", "LC2016030700046");
		params.put("loanRate", "7");
		params.put("loanAddRate", "2");
		params.put("endDate", "2017年03月14日");

		String filePath = "C:\\Users\\SDD\\Desktop\\协议\\11.docx";
		File file = new File("C:\\Users\\SDD\\Desktop\\协议\\11_out.docx");

		InputStream is = new FileInputStream(filePath);
		XWPFDocument doc = new XWPFDocument(is);
		//替换段落里面的变量
		this.replaceInPara(doc, params);
		//替换表格里面的变量
		this.replaceInTable(doc, params);

		if(file.exists()){
			file.delete();
		}
		OutputStream os = new FileOutputStream("C:\\Users\\SDD\\Desktop\\协议\\11_out.docx");
		doc.write(os);
		this.close(os);
		this.close(is);
	}


	/**
	 * 替换段落里面的变量
	 *
	 * @param doc    要替换的文档
	 * @param params 参数
	 */

	public void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			para = iterator.next();
			this.replaceInPara(para, params);
		}
	}


	/**
	 * 替换段落里面的变量
	 * @param para 要替换的段落
	 * @param params 参数
	 */
	private void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
		List<XWPFRun> runs;
		Matcher matcher;
		if (this.matcher(para.getParagraphText()).find()) {
			runs = para.getRuns();
			System.out.println("模板：" + runs);
			for (int i=0; i<runs.size(); i++) {
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				matcher = this.matcher(runText);
				if (matcher.find()) {
					while ((matcher = this.matcher(runText)).find()) {
						runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
					}
					//直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
					//所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
					run.setText(runText, 0);
				}
			}

			System.out.println("替换后：" + runs);
		}
	}

	/**
	 * 替换表格里面的变量
	 *
	 * @param doc    要替换的文档
	 * @param params 参数
	 */
	public void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
		Iterator<XWPFTable> iterator = doc.getTablesIterator();
		XWPFTable table;
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;
		List<XWPFParagraph> paras;
		while (iterator.hasNext()) {
			table = iterator.next();
			rows = table.getRows();
			for (XWPFTableRow row : rows) {
				cells = row.getTableCells();
				for (XWPFTableCell cell : cells) {
					paras = cell.getParagraphs();
					for (XWPFParagraph para : paras) {
						this.replaceInPara(para, params);
					}
				}
			}
		}
	}

	/**
	 * 正则匹配字符串
	 *
	 * @param str
	 * @return
	 */
	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}

	/**
	 * 关闭输入流
	 *
	 * @param is
	 */
	public void close(InputStream is) {
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 关闭输出流
	 *
	 * @param os
	 */
	public void close(OutputStream os) {
		if (os != null) {
			try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
