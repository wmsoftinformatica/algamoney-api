
package com.example.algamoney.api.exportar;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.algamoney.api.model.Lancamento;

public class LancamentoExcelExporter {
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	List<Lancamento> listLancamento;

	public LancamentoExcelExporter(final List<Lancamento> lancamento) {
		workbook = new XSSFWorkbook();

		sheet = workbook.createSheet("LANCAMENTOEXPORT");
		listLancamento = lancamento;

	}

	private void writeHeaderRow() {

		final Row row = sheet.createRow(0);

		Cell cell = row.createCell(0);
		cell.setCellValue("Codigo");

		cell = row.createCell(1);
		cell.setCellValue("Descricao");

		cell = row.createCell(2);
		cell.setCellValue("Data Vencimento");

		cell = row.createCell(3);
		cell.setCellValue("Data Pagamento");

		cell = row.createCell(4);
		cell.setCellValue("Valor");

		cell = row.createCell(5);
		cell.setCellValue("observacao");

		cell = row.createCell(6);
		cell.setCellValue("tipo");

		cell = row.createCell(7);
		cell.setCellValue("Codigo Categoria");

		cell = row.createCell(8);
		cell.setCellValue("Codigo Pessoa");

	}

	private void writeDataRows() {

		int rowCount = 1;

		for (final Lancamento lancamento : listLancamento) {
			final Row row = sheet.createRow(rowCount++);

			Cell cell = row.createCell(0);
			cell.setCellValue(lancamento.getCodigo());

			cell = row.createCell(1);
			cell.setCellValue(lancamento.getDescricao());

			cell = row.createCell(2);
			cell.setCellValue(lancamento.getDataVencimento().toString());

			cell = row.createCell(3);
			cell.setCellValue(lancamento.getDataPagamento().toString());

			cell = row.createCell(4);
			cell.setCellValue(lancamento.getValor().toString());

			cell = row.createCell(5);
			cell.setCellValue(lancamento.getObservacao());

			cell = row.createCell(6);
			cell.setCellValue(lancamento.getTipo().toString());

			cell = row.createCell(7);
			cell.setCellValue(lancamento.getCategoria().getCodigo());

			cell = row.createCell(8);
			cell.setCellValue(lancamento.getPessoa().getCodigo());

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
