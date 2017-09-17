package excel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	
	public static void ExportWorkBook(List<Map<String, String>> dataList, String toPath) throws Exception {
		toPath = (toPath == null || "".equals(toPath)) ? "D:\\result.xlsx" : toPath;
		// 文件输出位置
		FileOutputStream out = new FileOutputStream(new File(toPath));
		// 定义excel
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		sheet.setColumnWidth(0, 20 * 256);
		sheet.setColumnWidth(1, 13 * 256);
		// 创建第一行的标题
		XSSFRow titleRow = sheet.createRow(0);
		XSSFCell zkzhTitleCell = titleRow.createCell(0);
		zkzhTitleCell.setCellValue("准考证号");
		XSSFCell sfzhTitleCell = titleRow.createCell(1);
		sfzhTitleCell.setCellValue("身份证号");
		XSSFCell xznlcjTitleCell = titleRow.createCell(2);
		xznlcjTitleCell.setCellValue("行政能力成绩");
		XSSFCell slcjTitleCell = titleRow.createCell(3);
		slcjTitleCell.setCellValue("申论成绩");
		XSSFCell zycjTitleCell = titleRow.createCell(4);
		zycjTitleCell.setCellValue("专业成绩");
		XSSFCell zcjTitleCell = titleRow.createCell(5);
		zcjTitleCell.setCellValue("总成绩");
		
		// 循环填充数据
		if (dataList != null && dataList.size() > 0) {
			for (int rowNum = 1; rowNum < dataList.size(); rowNum++) {
				Map<String, String> dataMap = dataList.get(rowNum);
				XSSFRow row = sheet.createRow(rowNum);
				XSSFCell zkzhCell = row.createCell(0);
				zkzhCell.setCellValue(dataMap.get("zkzh"));
				XSSFCell sfzhCell = row.createCell(1);
				sfzhCell.setCellValue(dataMap.get("sfzh"));
				XSSFCell xznlcj = row.createCell(2);
				xznlcj.setCellValue(dataMap.get("xznlcj"));
				XSSFCell slcj = row.createCell(3);
				slcj.setCellValue(dataMap.get("slcj"));
				XSSFCell zycj = row.createCell(4);
				zycj.setCellValue(dataMap.get("zycj"));
				XSSFCell zcj = row.createCell(5);
				zcj.setCellValue(dataMap.get("zcj"));
			}
		}
		workbook.write(out);
		out.close();
	}
	
	public static void main(String[] args) throws Exception {
		List<Map<String, String>> dataList = ReadExcel.readWorkBook("C:\\Users\\qinghe\\Desktop\\test1.xlsx");
		ExportWorkBook(dataList, null);
	}
}
