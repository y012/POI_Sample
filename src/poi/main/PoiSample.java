package poi.main;

import static poi.common.constant.ExcelConstants.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import poi.common.dao.CSVInputDao;
import poi.common.dto.CSVDto;
import poi.common.util.ExcelManager;

public class PoiSample {


	public static void main(String[] args) {

		System.out.println("開始");

		FileInputStream fis = null;
		FileOutputStream fos = null;
		Workbook wb = null;
		ExcelManager manager = new ExcelManager();
		int rowIndex = 0;
		CSVInputDao csvInDao = new CSVInputDao();
		CSVDto csvDto = new CSVDto();
		Map<Integer, List<CSVDto>> agesMap = new TreeMap<>();
		SimpleDateFormat sdFormat = new SimpleDateFormat("yyyy/MM/dd");

		csvInDao.setFile(FILE_PATH_INPUT);


		try {

			fis = new FileInputStream(FILE_PATH_TEMPLATE);
			wb = (XSSFWorkbook)WorkbookFactory.create(fis);

			wb.createFont().setFontName(FONT);
			// シート名を指定して取得
			Sheet sheet = wb.cloneSheet(wb.getSheetIndex(SHEET_NAME_TEMPLATE));

			// シート名の設定
			wb.setSheetName(wb.getSheetIndex(sheet), SHEET_NAME_ALL_DATA);

			System.out.println("createSheet:" + sheet.getSheetName());

			manager.setSheet(sheet);

			csvInDao.open();

			// 全データ出力
			while((csvDto = csvInDao.read()) != null) {
				if(rowIndex != 0){
					if(rowIndex >= 2) {
						manager.shiftRows();
					}

					int age = Integer.parseInt(csvDto.getAge());
					String comment = "電話:" + csvDto.getTelephone() + LINE_SEPARATOR
							+ "携帯:" + csvDto.getMobile() + LINE_SEPARATOR
							+ "メール:" + csvDto.getMail();
					manager.setCellComment(comment);
					manager.setCellValue(csvDto.getName());
					manager.setCellValue(csvDto.getFurigana());
					manager.setCellValue(csvDto.getSex());
					manager.setCellValue(age);
					manager.setCellValue(sdFormat.parse(csvDto.getBirthDay()));
					manager.setCellValue(csvDto.getMarriage());
					manager.setCellValue(csvDto.getBloodType());
					manager.setCellValue(csvDto.getBirthPlace());
					manager.setCellValue(csvDto.getCareer());
					// 年代取得 ex)23→20
					age = (int)(age / 10) * 10;
					if(agesMap.containsKey(age)) {
						agesMap.get(age).add(csvDto);
					}else {
						List<CSVDto> dtoList = new ArrayList<>();
						dtoList.add(csvDto);
						agesMap.put(age, dtoList);
					}

				}
				rowIndex++;
				manager.nextRow();
			}
			// 印刷設定、印刷範囲
			manager.setPrintSetup(wb.getSheetAt(wb.getSheetIndex(SHEET_NAME_TEMPLATE)).getPrintSetup());
			wb.setPrintArea(wb.getSheetIndex(sheet), manager.getPrintArea());


			// 年代毎にシート分けして出力
			List<String> sheetNames = new ArrayList<>();
			for(Map.Entry<Integer, List<CSVDto>> ageMap : agesMap.entrySet()) {
				String sheetName = String.valueOf(ageMap.getKey()) + "代";
				if((sheet = wb.getSheet(sheetName)) == null) {
					sheet = wb.cloneSheet(wb.getSheetIndex(SHEET_NAME_TEMPLATE));
					wb.setSheetName(wb.getSheetIndex(sheet), sheetName);
				}
				sheetNames.add(sheet.getSheetName());
				System.out.println("createSheet:" + sheet.getSheetName());
				manager.setSheet(sheet);
				rowIndex = 1;
				manager.nextRow();
				for(CSVDto nextDto: ageMap.getValue()) {
					if(rowIndex >= 2) {
						manager.shiftRows();
					}
					String comment = "電話:" + nextDto.getTelephone() + LINE_SEPARATOR
							+ "携帯:" + nextDto.getMobile() + LINE_SEPARATOR
							+ "メール:" + nextDto.getMail();
					manager.setCellComment(comment);
					manager.setCellValue(nextDto.getName());
					manager.setCellValue(nextDto.getFurigana());
					manager.setCellValue(nextDto.getSex());
					manager.setCellValue(Integer.parseInt(nextDto.getAge()));
					manager.setCellValue(sdFormat.parse(nextDto.getBirthDay()));
					manager.setCellValue(nextDto.getMarriage());
					manager.setCellValue(nextDto.getBloodType());
					manager.setCellValue(nextDto.getBirthPlace());
					manager.setCellValue(nextDto.getCareer());
					rowIndex++;
					manager.nextRow();
				}
				// 印刷設定、印刷範囲
				manager.setPrintSetup(wb.getSheetAt(wb.getSheetIndex(SHEET_NAME_TEMPLATE)).getPrintSetup());
				wb.setPrintArea(wb.getSheetIndex(sheet), manager.getPrintArea());
			}


			// 集計出力
			sheet = wb.cloneSheet(wb.getSheetIndex(SHEET_NAME_GRAPH_TEMPLATE));
			wb.setSheetName(wb.getSheetIndex(sheet), SHEET_NAME_AGGREGATED);

			System.out.println("createSheet:" + sheet.getSheetName());
			manager.setSheet(sheet);

			String formula = "";
			for(rowIndex = 3; 7 >= rowIndex; rowIndex++) {
				manager.setPosition(rowIndex, 2);
				for(String sheetName: sheetNames) {
					if(rowIndex == 7) {
						formula = "SUM($" + manager.getColumnAlphabet() + "$4:$"
								+ manager.getColumnAlphabet() +"$7)";
					}else {
						formula = "COUNTIF(\'" + sheetName + "\'!$I:$I,$B$" + String.valueOf(rowIndex+1) + ")";
					}
					manager.setCellFormula(formula);
				}
			}
			// 数式は設定しただけでは計算されないので、再計算
			wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
			rowIndex++;
			manager.setPosition(rowIndex, 1);
			// 棒グラフ設定
			manager.setBarChart();

			// 印刷設定、印刷範囲
			manager.setPrintSetup(wb.getSheetAt(wb.getSheetIndex(SHEET_NAME_GRAPH_TEMPLATE)).getPrintSetup());
			wb.setPrintArea(wb.getSheetIndex(sheet), manager.getPrintArea());

			// 不要なテンプレートを削除
			wb.removeSheetAt(wb.getSheetIndex(SHEET_NAME_TEMPLATE));
			wb.removeSheetAt(wb.getSheetIndex(SHEET_NAME_GRAPH_TEMPLATE));

			// ファイル出力
			fos = new FileOutputStream(FILE_PATH_OUTPUT);
			wb.write(fos);

			System.out.println("終了");

		}catch(Exception e) {
			e.printStackTrace();
		}finally {
			if(fis != null) {
				try {
					fis.close();
				}catch(IOException e) {

				}
			}
			if(fos != null) {
				try {
					fos.close();
				}catch(IOException e) {

				}
			}
			if(wb != null) {
				try {
					wb.close();
				}catch(IOException e) {

				}
			}
		}

	}
}
