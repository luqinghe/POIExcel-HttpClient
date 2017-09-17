package excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	public static List<Map<String, String>> readWorkBook(String fromPath) throws Exception {
		// 读取文件
		InputStream is = new FileInputStream(fromPath);
		// 加载excel
		XSSFWorkbook workBook = new XSSFWorkbook(is);
		// 锁定sheet页
		XSSFSheet sheet = workBook.getSheetAt(0);
		// 定义list用于存储数据
		List<Map<String, String>> dataList = new ArrayList<Map<String, String>>();
		// 遍历excel的行
		for (int rowNum = 0 ; rowNum <= sheet.getLastRowNum(); rowNum++) {
			int cellNum = 0;
			XSSFRow row = sheet.getRow(rowNum);
			Map<String, String> dataMap = new HashMap<String, String>();
			// 第一列是准考证号
			XSSFCell zkzhCell = row.getCell(cellNum++);
			dataMap.put("zkzh", zkzhCell.toString());
			// 第二列是身份证号
			XSSFCell sfzhCell = row.getCell(cellNum++);
			dataMap.put("sfzh", sfzhCell.toString());
			
			System.out.println("第" + (rowNum + 1) + "行：准考证号(" + zkzhCell.toString()
					+ "),身份证号(" + sfzhCell.toString() + ");");
			dataList.add(dataMap);
		}
		workBook.close();
		// 关闭输入流
		is.close();
		return dataList;
	}
	
	public static void main(String[] args) throws Exception {
		List<Map<String, String>> dataList = ReadExcel.readWorkBook("C:\\Users\\qinghe\\Desktop\\test1.xlsx");
		System.out.println(dataList.size());
	}
}
