
package com.example.algamoney.api.resource;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.activity.InvalidActivityException;
import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationEventPublisher;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.algamoney.api.event.RecursoCriadoEvent;
import com.example.algamoney.api.exportar.CategoriaExcelExporter;
import com.example.algamoney.api.model.Categoria;
import com.example.algamoney.api.repository.CategoriaRepository;

@RestController
@RequestMapping("/categorias")
public class CategoriaResource {

	@Autowired
	private CategoriaRepository categoriaRepository;

	@Autowired
	private ApplicationEventPublisher publisher;

	@GetMapping
	@PreAuthorize("hasAuthority('ROLE_PESQUISAR_CATEGORIA') and #oauth2.hasScope('read')")
	public List<Categoria> listar() {
		return categoriaRepository.findAll();
	}

	@GetMapping("/exportarcategoria")
	public void exportarCategoria(final HttpServletResponse response) throws IOException {
		final List<Categoria> categoria = categoriaRepository.findAll();
		response.setContentType("application/octet-strem");
		final CategoriaExcelExporter categoriaExelExporter = new CategoriaExcelExporter(categoria);

		categoriaExelExporter.export(response);

	}

	@PostMapping
	@PreAuthorize("hasAuthority('ROLE_CADASTRAR_CATEGORIA') and #oauth2.hasScope('write')")
	public ResponseEntity<Categoria> criar(@Valid @RequestBody final Categoria categoria,
			final HttpServletResponse response) {
		final Categoria categoriaSalva = categoriaRepository.save(categoria);
		publisher.publishEvent(new RecursoCriadoEvent(this, response, categoriaSalva.getCodigo()));
		return ResponseEntity.status(HttpStatus.CREATED).body(categoriaSalva);
	}

	@PostMapping("/importarcategoria")
	public void
			upLoad(@RequestParam final MultipartFile arquivo) throws Exception, InvalidActivityException, IOException {
		final Workbook workbook = WorkbookFactory.create(arquivo.getInputStream());
		final Sheet sheet = workbook.getSheetAt(0);
		final Iterator<Row> rowIterator = sheet.rowIterator();

		Integer count = 0;

		while (rowIterator.hasNext()) {
			if (count == 0) {
				rowIterator.next();
				count++;
				continue;
			}

			// count++;

			final Row row = rowIterator.next();

			final Cell codigo = row.getCell(0);
			final Cell nome = row.getCell(1);

			// String codigo2;
			// Long codigoConv;
			// codigo2 = codigo.toString();
			// codigoConv = Long.parseLong(codigo2);

			final Categoria categoria = new Categoria(nome.getStringCellValue());
			System.out.println("------------------------------------");
			System.out.println("Codigo " + codigo);
			System.out.println("Nome " + nome);
			System.out.println("------------------------------------");

			categoriaRepository.save(categoria);
		}

	}

	@GetMapping("/{codigo}")
	@PreAuthorize("hasAuthority('ROLE_PESQUISAR_CATEGORIA') and #oauth2.hasScope('read')")
	public ResponseEntity<Categoria> buscarPeloCodigo(@PathVariable final Long codigo) {
		final Categoria categoria = categoriaRepository.findOne(codigo);
		return categoria != null ? ResponseEntity.ok(categoria) : ResponseEntity.notFound().build();
	}

}
