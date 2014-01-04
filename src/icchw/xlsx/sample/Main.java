package icchw.xlsx.sample;
import icchw.xlsx.WorkbookWrapper;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Main {

	public static void main(String[] args) {

		String sheetName = "sheet1";

		try {
			// 使用例１：物理ファイルをtemplateにする場合
			File template = new File("input/template.xlsx");
			WorkbookWrapper wr = new WorkbookWrapper(template);

			// 使用例２：メモリ上のWorkbookをtemplateにする場合
//			XSSFWorkbook wb = new XSSFWorkbook();
//			wb.createSheet(sheetName);
//			WorkbookWrapper wr = new WorkbookWrapper(wb);

			// Write xmls
			List<DataDto> dataDtos = prepareData();
			System.out.println("start");
			long startTime = (new Date()).getTime();
			wr.writeSheet(sheetName, dataDtos);
			long endTime = (new Date()).getTime();
			System.out.println("end:" + ((long)(endTime - startTime)/1000) + "s");

			// Generate zip
			FileOutputStream output = new FileOutputStream(new File("output/" + generateFileName() + ".xlsx"));
			wr.write(output);
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * テストデータ作成
	 * @return
	 */
	private static List<DataDto> prepareData() {
		List<DataDto> dataDtos = new ArrayList<DataDto>();
		for (int i=0; i<50000; i++) {
			dataDtos.add(new DataDto("こんにちは", (double)i/100, new BigDecimal("-0.5"), new Date(), Calendar.getInstance(), String.format("%1$02d", i%47+1)));
			dataDtos.add(new DataDto("", (double)i/100, new BigDecimal("0.5"), null, Calendar.getInstance(), String.format("%1$02d", i%47+1)));
		}
		return dataDtos;
	}

	/**
	 * テスト用ファイル名生成
	 * @return
	 */
	private static String generateFileName() {
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd_HHmmss");
		return sdf1.format(new Date());
	}

}
