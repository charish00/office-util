package com.lls.wec.local.document.common.helper;

import java.io.File;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;

import lombok.extern.slf4j.Slf4j;

/**
 * 模板自动纠正类，纠正${xxxx}中间会增加空白标签的问题
 * @author panxiao
 *
 */
@Slf4j
public class Corrector {

	public synchronized void correctTemplate(String templateFileDir, String templateFileName) {
		String filePath = templateFileDir + File.separator + templateFileName;

		try {
			File file = new File(filePath);
			String content = FileUtils.readFileToString(file, "UTF-8");

			int length = content.length();
			char c;
			char s1 = 0;
			char s2 = 0;
			char s3 = 0;
			StringBuilder result = new StringBuilder();
			StringBuilder build = new StringBuilder();
			for (int i = 0; i < length; i++) {
				c = content.charAt(i);
				if (s1 == '$' && s2 == '{') {
					build.append(c);
				} else {
					result.append(c);
				}
				if (s2 == '{' && c == '}') {
					s3 = '}';
					s1 = 0;
					s2 = 0;
				}
				if (s1 == '$' && c == '{') {
					s2 = '{';
					s3 = 0;
				}
				if (c == '$') {
					s1 = c;
					s2 = 0;
					s3 = 0;
				}
				if (s3 == '}') {
					s1 = 0;
					s2 = 0;
					s3 = 0;
					String target = trim(build.toString());
					result.append(target);
					build = new StringBuilder();
				}

			}
			file = new File(filePath);
			FileUtils.write(file, result.toString(), "UTF-8");
		} catch (IOException e) {
			log.error("correctTemplate error", e);
		}
	}

	private String trim(String s) {
		int length = s.length();
		StringBuilder result = new StringBuilder();
		char c;
		boolean append = true;
		boolean change = false;
		for (int i = 0; i < length; i++) {
			c = s.charAt(i);
			if (c == '<') {
				append = false;
				change = true;
			}
			if (c == '>') {
				append = true;
			} else if (append) {
				result.append(c);
			}
		}
		return replaceSpecialStr(change ? result.toString() : s);
	}
	
	public static String replaceSpecialStr(String str) {
        String repl = "";
        if (str!=null) {
            Pattern p = Pattern.compile("\\s*|\t|\r|\n");
            Matcher m = p.matcher(str);
            repl = m.replaceAll("");
        }
        return repl.trim();
    }


	public static void main(String[] args) {

		new Corrector().correctTemplate("D:/", "gnlLimitContract-JL.xml");
//		new Corrector().docxToXml("d://trade_receipt_mortgage.docx", "d://trade_receipt_mortgage-a.xml");
		
	}

}
