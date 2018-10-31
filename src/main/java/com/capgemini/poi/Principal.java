package com.capgemini.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Principal {

	private static final String fileName = "C:/testes.xls";

	public static void main(String[] args) throws IOException {

		try {
			FileInputStream arquivo = new FileInputStream(new File(Principal.fileName));

			@SuppressWarnings("resource")
			HSSFWorkbook workbook = new HSSFWorkbook(arquivo);

			HSSFSheet sheetAlunos = workbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheetAlunos.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				StringBuilder sb = new StringBuilder();

				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getColumnIndex() != 5) {
//								System.out.println(cell.getStringCellValue());
						sb.append(cell.getStringCellValue() + " ");
					}

					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
//								System.out.println(cell.getNumericCellValue());
						String value = String.valueOf((int) cell.getNumericCellValue());
						sb.append(value + " ");
					}

					if (cell.getColumnIndex() == 5) {
//							System.out.println(cell.getNumericCellValue());
						String objeto = new String(cell.getStringCellValue());
						sb.append(" ====== >" + 564465 + " ");
					}

				}
				System.out.println(sb);

				sb = new StringBuilder();
			}
			arquivo.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		}
	}

}
