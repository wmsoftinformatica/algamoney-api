
package com.example.algamoney.api.exportar;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.algamoney.api.model.Categoria;

public class CategoriaExcelExporter {

	XSSFWorkbook workbook;
	XSSFSheet sheet;
	List<Categoria> listCategorias;

	public CategoriaExcelExporter(final List<Categoria> categorias) {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("CATEGORIASEXPORT");
		listCategorias = categorias;
	}

	private void writeHeaderRow() {
		// cabeçario---------------
		final Row row = sheet.createRow(0);

		Cell cell = row.createCell(0);
		cell.setCellValue("codigo");

		cell = row.createCell(1);
		cell.setCellValue("Nome");

	}

	private void writeDataRows() {
		int rowCount = 1;
		for (final Categoria categoria : listCategorias) {
			final Row row = sheet.createRow(rowCount++);

			Cell cell = row.createCell(0);
			cell.setCellValue(categoria.getCodigo());

			cell = row.createCell(1);
			cell.setCellValue(categoria.getNome());

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

}
