package jp.example;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class Main {
	/**
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		String path = args[0];
		byte[] templateBytes = null;
		// Excelファイルを読み込み
		try (Workbook templateWorkbook = WorkbookFactory.create(new File(path))) {
			// Workbookをバイナリデータに変換
			templateBytes = convertWorkbookToByteArray(templateWorkbook);
		}

		// バイト配列からWorkbookオブジェクトを作成する
		try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(templateBytes))) {
			List<List<String>> listData = List.of(
					List.of("2023/01/01", "Alice", "20"),
					List.of("2023/01/01", "Bob", "25"),
					List.of("2023/01/01", "Charlie", "30"));
			writeDataToExcel(workbook, listData, "B3");

			// Workbookをバイナリデータに変換
			byte[] data = convertWorkbookToByteArray(workbook);

			// バイナリデータをファイルに書き込み
			String fileName = "E:\\文書\\newFile.xlsx"; // 保存するファイル名
			FileOutputStream fos = new FileOutputStream(fileName);
			fos.write(data);
			fos.close();
		}
	}
	
	/**
	 * セル書き込み処理
	 * @param workbook
	 * @param data
	 * @throws Exception
	 */
	public static void writeDataToExcel(Workbook workbook, List<List<String>> data, String cellAddress) throws Exception {
        // CellReferenceオブジェクトを作成する
        CellReference cellReference = new CellReference(cellAddress);

        // 行番号を取得する
        int rowNumber = cellReference.getRow();

        // 列番号を取得する
        int columnNumber = cellReference.getCol();
        
        writeDataToExcel(workbook, data, rowNumber, columnNumber);

	}
	
	/**
	 * セル書き込み処理
	 * @param workbook
	 * @param data
	 * @throws Exception
	 */
	public static void writeDataToExcel(Workbook workbook, List<List<String>> data, int startRow, int startColumn) throws Exception {
        
		Sheet sheet = workbook.getSheetAt(0);

		// 行とセルを作成し、データを書き込む
		for (int i = 0; i < data.size(); i++) {
			Row row = sheet.createRow(startRow + i);
			List<String> rowData = data.get(i);
			for (int j = 0; j < rowData.size(); j++) {
				Cell cell = row.createCell(j + startColumn);
				cell.setCellValue(rowData.get(j));
			}
		}
	}

	/**
	 * バイナリ書き込み処理
	 * @param workbook
	 * @return
	 * @throws IOException
	 */
	public static byte[] convertWorkbookToByteArray(Workbook workbook) throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		workbook.write(baos);
		return baos.toByteArray();
	}
}
