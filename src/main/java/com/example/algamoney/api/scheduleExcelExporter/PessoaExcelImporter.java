
package com.example.algamoney.api.scheduleExcelExporter;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;

import com.example.algamoney.api.model.Pessoa;

public class PessoaExcelImporter {

	public void processUpload(final Iterator<Row> rowIterator, final Pessoa pessoa) throws Exception {
		Integer count = 0;
		while (rowIterator.hasNext()) {

			if (count == 0) {
				rowIterator.next();
				count++;
				continue;
			}

		}
	}

}
