package main;

import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Main {

	public static void main(String[] args) {
		String filePath=System.getProperties().getProperty("user.dir").replace('\\', '/')+"/resources/";
		String fileName = filePath+"安信三号8.16.xls";
		processExcel(fileName);
	}

	private static void processExcel(String fileName) {
		String dateName="发生日期";
		String newFileName=fileName.substring(0,fileName.lastIndexOf("/")+1)+"new_"+fileName.substring(fileName.lastIndexOf("/")+1,fileName.length());
		File newFile = new File(newFileName);
		try {
			newFile.createNewFile();
			WritableWorkbook book = Workbook.createWorkbook(newFile);
			Sheet sheetFrom = Workbook.getWorkbook(new File(fileName)).getSheet("Sheet1");
			//copy it
			WritableSheet sheet = book.createSheet("Sheet1", 0);
			for (int i = 0; i < sheetFrom.getRows(); i++) {
				for (int j = 0; j < sheetFrom.getColumns(); j++) {
					Cell cell = sheetFrom.getCell(j, i);
					String content = cell.getContents();
					Label label = new Label(j, i, content);
					sheet.addCell(label);
				}
			}
			
			List<String> titles =getExcelTitles(sheetFrom);
			Set<String> dateSet=getDistinctDate(sheetFrom, titles ,dateName);
			
			for (String date:dateSet) {
				//循环创建sheet页
				WritableSheet ws = book.createSheet(date, 1);
				//设置每个sheet的标题
				for (int i = 0; i < titles.size(); i++) {
					Label label = new Label(i, 0, titles.get(i));
					ws.addCell(label);
				}
				int rowIndex=1;
				//遍历行
				for (int i = 1; i < sheetFrom.getRows(); i++) {
					Cell cell = sheetFrom.getCell(titles.indexOf(dateName), i);
					String content = cell.getContents().trim();
					//如果该行的日期列等于指定日期，则将该行数据添加到新的sheet中
					if(date.equals(content)){
						for (int j = 0; j < sheetFrom.getColumns(); j++) {
							Cell cell2 = sheetFrom.getCell(j, i);
							String content2 = cell2.getContents();
							Label label = new Label(j, rowIndex, content2);
							ws.addCell(label);
						}
						rowIndex++;
					}
				}
			}
			book.write();
			book.close();
			System.out.println("create finished");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static Set<String> getDistinctDate(Sheet sheetFrom, List<String> titles ,String titleName) {
		int num=titles.indexOf(titleName);
		Set<String> dateSet=new HashSet<>();
		for (int i = 1; i < sheetFrom.getRows(); i++) {
			Cell cell = sheetFrom.getCell(num, i);
			String content = cell.getContents().trim();
			dateSet.add(content);
		}
		return dateSet;
	}

	private static List<String> getExcelTitles(Sheet sheet) {
		List<String> titles = new ArrayList<>();
		//获取列名(第一行数据)
		for (int i = 0; i < sheet.getColumns(); i++) {
			Cell cell = sheet.getCell(i, 0);
			titles.add(cell.getContents());
		}
		return titles;
	}
}
