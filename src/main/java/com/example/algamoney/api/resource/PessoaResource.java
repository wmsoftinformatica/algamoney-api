
package com.example.algamoney.api.resource;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationEventPublisher;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.algamoney.api.event.RecursoCriadoEvent;
import com.example.algamoney.api.model.Pessoa;
import com.example.algamoney.api.repository.PessoaRepository;
import com.example.algamoney.api.scheduleExcelExporter.PessoaExcelExporter;
import com.example.algamoney.api.service.PessoaService;

@RestController
@RequestMapping("/pessoas")
public class PessoaResource {

	@Autowired
	private PessoaRepository pessoaRepository;

	@Autowired
	private PessoaService pessoaService;

	@Autowired
	private ApplicationEventPublisher publisher;

	@GetMapping
	public List<Pessoa> listar() {

		return pessoaRepository.findAll();
	}

	@GetMapping("/exportarexcel")

	public void exportar(final HttpServletResponse response) throws IOException {
		final List<Pessoa> pessoas = pessoaRepository.findAll();
		response.setContentType("application/octet-stream");
		final PessoaExcelExporter pessoaExcelExporter = new PessoaExcelExporter(pessoas);

		pessoaExcelExporter.export(response);

	}

	@PostMapping("/importarexcel")
	public void
			upload(@RequestParam final MultipartFile arquivo) throws Exception, InvalidFormatException, IOException {
		final Workbook workbook = WorkbookFactory.create(arquivo.getInputStream());
		final Sheet sheet = workbook.getSheetAt(0);
		final Iterator<Row> rowIterator = sheet.rowIterator();

		new Pessoa(null, null);

		while (rowIterator.hasNext()) {
			final Row row = rowIterator.next();

			final Cell celulaCodigo = row.getCell(0);
			final Cell nome = row.getCell(1);
			System.out.println(celulaCodigo);
			System.out.println(nome);
			System.out.println("------------------------------------------------------");

		}

	}

	@PostMapping
	public ResponseEntity<Pessoa> criar(@Valid @RequestBody final Pessoa pessoa, final HttpServletResponse response) {
		final Pessoa pessoaSalva = pessoaRepository.save(pessoa);
		publisher.publishEvent(new RecursoCriadoEvent(this, response, pessoaSalva.getCodigo()));
		return ResponseEntity.status(HttpStatus.CREATED).body(pessoaSalva);
	}

	@GetMapping("/{codigo}")
	public ResponseEntity<Pessoa> buscarPeloCodigo(@PathVariable final Long codigo) {
		final Pessoa pessoa = pessoaRepository.findOne(codigo);
		return pessoa != null ? ResponseEntity.ok(pessoa) : ResponseEntity.notFound().build();
	}

	@DeleteMapping("/{codigo}")
	@ResponseStatus(HttpStatus.NO_CONTENT)
	public void remover(@PathVariable final Long codigo) {
		pessoaRepository.delete(codigo);
	}

	@PutMapping("/{codigo}")
	public ResponseEntity<Pessoa> atualizar(@PathVariable final Long codigo, @Valid @RequestBody final Pessoa pessoa) {
		final Pessoa pessoaSalva = pessoaService.atualizar(codigo, pessoa);
		return ResponseEntity.ok(pessoaSalva);
	}

	@PutMapping("/{codigo}/ativo")
	@ResponseStatus(HttpStatus.NO_CONTENT)
	public void atualizarPropriedadeAtivo(@PathVariable final Long codigo, @RequestBody final Boolean ativo) {
		pessoaService.atualizarPropriedadeAtivo(codigo, ativo);
	}

}
