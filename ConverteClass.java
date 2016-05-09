package Java_Code;

import java.io.BufferedReader;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.SpringLayout.Constraints;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;

import com.opencsv.CSVReader;

public class ConverteClass {

	public static void main(String[] args) throws IOException {

		Workbook wb = new HSSFWorkbook();
		CreationHelper helper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");

		System.setProperty("file.encoding","UTF-8");
		Field charset = null;
		try {
			charset = Charset.class.getDeclaredField("defaultCharset");
			charset.setAccessible(true);
		} catch (NoSuchFieldException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		}
		try {
			charset.set(null,null);
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
		
		String[] types =new String[]{"NUMBER\\({1}[0-9]++,{1}[0-9]++\\){1}","TIMESTAMP\\({1}[0-9]++\\){1}",
				"VARCHAR[0-9]++\\({1}[0-9]++\\){1}", "INTEGER","DATE","UNIQ"}; 
		

		Matcher tableMatcher1 = null, tableMatcher2 = null, rowsMatcher=null, bodyMatcher = null, oumMatcher = null;
		Pattern tablePattern1 = null, tablePattern2 = null, rowsPattern=null, bodyPatter = null, oumPattern = null;
		String[] parts = null;
		String line;
		int r = 0;
		String table = null;
		String[] rowsCut= new String[]{"#INDEX(.*)", "#CONSTRAINT(.*)" };
		
		
		File[] files = new File("D:/PTE-PMMK/Munka/suit/dif2/dif/").listFiles();

		for (File file : files) {
			boolean firstRow = true;

			Scanner sc = new Scanner(file);

			while (sc.hasNext()) {
				line = sc.nextLine();

				if (firstRow == true) {
					String tableCash = null;
					tablePattern1 = Pattern.compile("TABLE");
					tableMatcher1 = tablePattern1.matcher(line);
					if (tableMatcher1.find()) {
						tableCash= line.replace(tableMatcher1.group(), "");
					}
					tablePattern2 = Pattern.compile("[A-Z](.*)");
					tableMatcher2 = tablePattern2.matcher(tableCash);
					if (tableMatcher2.find()) {
						table = tableCash.replace(tableMatcher2.group(), "");
					}

					firstRow = false;
				} else {
					String typeValue;
					Row row = sheet.createRow((short) r);
					line = line.replace("FIELD", "");
				
					
					for (String rowCut : rowsCut) {
					rowsPattern = Pattern.compile(rowCut);
					rowsMatcher = rowsPattern.matcher(line);
					if (rowsMatcher.find()) {
						line = line.replace(rowsMatcher.group(), "");
					}
					}
					if (!line.equals("")) {
						r++;
						row.createCell(0).setCellValue(helper.createRichTextString(table));
					}
					
					for (String type : types) {
						bodyPatter = Pattern.compile(type);
						bodyMatcher = bodyPatter.matcher(line);
						if (bodyMatcher.find()) {
							parts = line.split(type);
							typeValue = bodyMatcher.group();
							row.createCell(1).setCellValue(helper.createRichTextString(parts[0]));
							row.createCell(2).setCellValue(helper.createRichTextString(typeValue));
							row.createCell(3).setCellValue(helper.createRichTextString(parts[1]));
							break;
						} else {
							row.createCell(1).setCellValue(helper.createRichTextString(line));
						}
					}

					oumPattern = Pattern.compile("ÜOM(.*)");
					oumMatcher = oumPattern.matcher(line);
					if (oumMatcher.find()) {
						row.createCell(4).setCellValue(helper.createRichTextString(oumMatcher.group()));
					}

				}
				

			}

			FileOutputStream fileOut = new FileOutputStream(
					"D:/PTE-PMMK/Munka/suit/ExcelConverter/src/05.09_new.xls");
			wb.write(fileOut);
			fileOut.close();
		}
	}
}

