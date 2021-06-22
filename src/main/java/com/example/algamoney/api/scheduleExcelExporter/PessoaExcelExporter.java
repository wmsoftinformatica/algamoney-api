
package com.example.algamoney.api.scheduleExcelExporter;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.algamoney.api.model.Pessoa;

// @Component
public class PessoaExcelExporter {
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	List<Pessoa> listPessoas;

	public PessoaExcelExporter(final List<Pessoa> pessoas) {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("PESSOASEXPORT");
		listPessoas = pessoas;

	}

	private void writeHeaderRow() {

		final Row row = sheet.createRow(0);

		Cell cell = row.createCell(0);
		cell.setCellValue("Codigo");

		cell = row.createCell(1);
		cell.setCellValue("Nome");

		cell = row.createCell(2);
		cell.setCellValue("Ativo");

	}

	private void writeDataRows() {
		int rowCount = 1;
		for (final Pessoa pessoa : listPessoas) {
			final Row row = sheet.createRow(rowCount++);

			Cell cell = row.createCell(0);
			cell.setCellValue(pessoa.getCodigo());

			cell = row.createCell(1);
			cell.setCellValue(pessoa.getNome());

			cell = row.createCell(2);
			cell.setCellValue(pessoa.getAtivo());

		}
	}

	public void export(final HttpServletResponse response) throws IOException {
		writeHeaderRow();
		writeDataRows();
		final ServletOutputStream outputStream = response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

	}

	// public List<ScheduleEntity> getListSchedules() {
	// return listSchedules;
	// }

}
