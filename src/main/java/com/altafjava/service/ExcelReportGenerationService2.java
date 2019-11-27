package com.altafjava.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.altafjava.pojo.Counter;
import com.altafjava.pojo.Pojo;
import com.monitorjbl.xlsx.StreamingReader;

public class ExcelReportGenerationService2 {

	public static void main(String[] args) {
		int count = 0;
		try {
			String fileName = "CasesSearchReport-26-11-19.xlsx";
//			String fileName = "21-11-2019.xlsx";
			String path = "src/main/resources/files/" + fileName;
			InputStream is = new FileInputStream(new File(path));
			StreamingReader reader = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).sheetIndex(0).read(is);
			Map<String, Counter> map = new LinkedHashMap<>();
			Map<String, String> admissionDateMap = new LinkedHashMap<>();
			for (Row row : reader) {
				count++;
				if (count > 1) {
					String admissionDate = row.getCell(16).getStringCellValue();
					String hospitalDistrict = row.getCell(14).getStringCellValue();
					String hospitalName = row.getCell(12).getStringCellValue();
					String preauthDate = row.getCell(17).getStringCellValue();
					String preauthApprovedDate = row.getCell(19).getStringCellValue();
					String dischargeDate = row.getCell(24).getStringCellValue();
					String claimSubmittedDate = row.getCell(25).getStringCellValue();
					String key = hospitalDistrict + "<->" + hospitalName;
					admissionDateMap.put(key, admissionDate);
					Counter counter = map.get(key);
					if (counter == null) {
						counter = new Counter();
						map.put(key, counter);
					} else {
						int noOfPreauthDate = counter.getNoOfPreauthDate();
						int noOfPreauthApprovedDate = counter.getNoOfPreauthApprovedDate();
						int noOfDischargeDate = counter.getNoOfDischargeDate();
						int noOfClaimSubmittedDate = counter.getNoOfClaimSubmittedDate();

						if (!preauthDate.equalsIgnoreCase("NA"))
							noOfPreauthDate++;
						if (!preauthApprovedDate.equalsIgnoreCase("NA"))
							noOfPreauthApprovedDate++;
						if (!dischargeDate.equalsIgnoreCase("NA"))
							noOfDischargeDate++;
						if (!claimSubmittedDate.equalsIgnoreCase("NA"))
							noOfClaimSubmittedDate++;

						counter.setNoOfClaimSubmittedDate(noOfClaimSubmittedDate);
						counter.setNoOfDischargeDate(noOfDischargeDate);
						counter.setNoOfPreauthApprovedDate(noOfPreauthApprovedDate);
						counter.setNoOfPreauthDate(noOfPreauthDate);
						map.put(key, counter);
					}
				}
			}
			List<Pojo> pojos = new ArrayList<>();
			for (Map.Entry<String, Counter> entry : map.entrySet()) {
				Counter counter = entry.getValue();
				String key = entry.getKey();
//				System.out.println(admissionDateMap.get(key) + "  " + key + "  " + counter.getNoOfPreauthDate() + "  " + counter.getNoOfPreauthApprovedDate() + "  "
//						+ counter.getNoOfDischargeDate() + "  " + counter.getNoOfClaimSubmittedDate());
				Pojo pojo = new Pojo();
				pojo.setDate(admissionDateMap.get(key));
				String splits[] = key.split("<->");
				pojo.setHospitalDistrict(splits[0]);
				pojo.setHospitalName(splits[1]);
				pojo.setNoOfClaimSubmittedDate(counter.getNoOfClaimSubmittedDate());
				pojo.setNoOfDischargeDate(counter.getNoOfDischargeDate());
				pojo.setNoOfPreauthApprovedDate(counter.getNoOfPreauthApprovedDate());
				pojo.setNoOfPreauthDate(counter.getNoOfPreauthDate());
				pojos.add(pojo);
			}
			createExcelFile(pojos, fileName);
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.err.println("Total Count=" + count);
	}

	static private void createExcelFile(List<Pojo> pojos, String fileName) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("report");
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		Row headerRow = sheet.createRow(0);
		String[] columns = new String[] { "Date", "Hospital District", "Hospital Name", "No. Preauth initiated", "No. of Preauth Arrpoved", "No. Patient Discharge",
				"No.Claim Initiated " };
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}
		int rowNum = 1;
		for (Pojo employee : pojos) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(employee.getDate());
			row.createCell(1).setCellValue(employee.getHospitalDistrict());
			row.createCell(2).setCellValue(employee.getHospitalName());
			row.createCell(3).setCellValue(employee.getNoOfPreauthDate());
			row.createCell(4).setCellValue(employee.getNoOfPreauthApprovedDate());
			row.createCell(5).setCellValue(employee.getNoOfDischargeDate());
			row.createCell(6).setCellValue(employee.getNoOfClaimSubmittedDate());
		}
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		try {
			FileOutputStream fileOut = new FileOutputStream("src/main/resources/files/" + fileName + "-Output.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.err.println("---------------- Excel Report file created successfully ------------ ");
	}

}
