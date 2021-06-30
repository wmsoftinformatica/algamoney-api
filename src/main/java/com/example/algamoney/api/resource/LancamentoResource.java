
package com.example.algamoney.api.resource;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationEventPublisher;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.algamoney.api.event.RecursoCriadoEvent;
import com.example.algamoney.api.exceptionhandler.AlgamoneyExceptionHandler.Erro;
import com.example.algamoney.api.exportar.LancamentoExcelExporter;
import com.example.algamoney.api.model.Lancamento;
import com.example.algamoney.api.model.TipoLancamento;
import com.example.algamoney.api.repository.LancamentoRepository;
import com.example.algamoney.api.repository.filter.LancamentoFilter;
import com.example.algamoney.api.repository.projection.ResumoLancamento;
import com.example.algamoney.api.service.LancamentoService;
import com.example.algamoney.api.service.exception.PessoaInexistenteOuInativaException;
import com.fasterxml.jackson.databind.exc.InvalidFormatException;

@RestController
@RequestMapping("/lancamentos")
public class LancamentoResource {

	@Autowired
	private LancamentoRepository lancamentoRepository;

	@Autowired
	private LancamentoService lancamentoService;

	@Autowired
	private ApplicationEventPublisher publisher;

	@Autowired
	private MessageSource messageSource;

	@GetMapping
	@PreAuthorize("hasAuthority('ROLE_PESQUISAR_LANCAMENTO') and #oauth2.hasScope('read')")
	public Page<Lancamento> pesquisar(final LancamentoFilter lancamentoFilter, final Pageable pageable) {
		return lancamentoRepository.filtrar(lancamentoFilter, pageable);
	}

	@GetMapping("/exportar")
	public void exportar(final HttpServletResponse response) throws IOException {
		final List<Lancamento> lancamento = lancamentoRepository.findAll();
		response.setContentType("application/octet-strem");

		final LancamentoExcelExporter lancamentoExcelExporter = new LancamentoExcelExporter(lancamento);

		lancamentoExcelExporter.export(response);

	}

	@GetMapping(params = "resumo")
	@PreAuthorize("hasAuthority('ROLE_PESQUISAR_LANCAMENTO') and #oauth2.hasScope('read')")
	public Page<ResumoLancamento> resumir(final LancamentoFilter lancamentoFilter, final Pageable pageable) {
		return lancamentoRepository.resumir(lancamentoFilter, pageable);
	}

	@GetMapping("/{codigo}")
	@PreAuthorize("hasAuthority('ROLE_PESQUISAR_LANCAMENTO') and #oauth2.hasScope('read')")
	public ResponseEntity<Lancamento> buscarPeloCodigo(@PathVariable final Long codigo) {
		final Lancamento lancamento = lancamentoRepository.findOne(codigo);
		return lancamento != null ? ResponseEntity.ok(lancamento) : ResponseEntity.notFound().build();
	}

	@PostMapping("/importar")
	public void
			upload(@RequestParam final MultipartFile arquivo) throws Exception, InvalidFormatException, IOException {

		final Workbook workbook = WorkbookFactory.create(arquivo.getInputStream());
		final Sheet sheet = workbook.getSheetAt(0);
		final Iterator<Row> rowIterator = sheet.rowIterator();

		Integer count = 0;

		Cell cellCodigo;

		String descricao;
		TipoLancamento tipo;
		Cell cellDescricao;

		Cell cellCodigoCategoria;
		String codigoCategoria;

		while (rowIterator.hasNext()) {
			final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
			LocalDate dataVencimentoConv;
			LocalDate dataPagamentoConv;
			String dataVencimento2;
			String dataPagamento2;
			String valor2;
			BigDecimal valorConv;
			if (count == 0) {
				rowIterator.next();
				count++;
				continue;
			}

			final Row row = rowIterator.next();

			cellCodigo = row.getCell(0);
			cellDescricao = row.getCell(1);
			descricao = cellDescricao.getStringCellValue();

			final Cell dataVencimento = row.getCell(2);
			final Cell dataPagamento = row.getCell(3);
			final Cell valor = row.getCell(4);
			final Cell observacao = row.getCell(5);

			final Cell celltipo = row.getCell(6);
			tipo = celltipo == null ? null : TipoLancamento.valueOf(celltipo.getStringCellValue().trim());

			cellCodigoCategoria = row.getCell(7);
			codigoCategoria = cellCodigoCategoria.getStringCellValue();

			final Cell codigoPessoa = row.getCell(8);

			System.out.println(cellCodigo);
			System.out.println(descricao);
			System.out.println(dataVencimento);
			System.out.println(dataPagamento);
			System.out.println(valor);
			System.out.println(observacao);
			System.out.println(tipo);
			System.out.println(cellCodigoCategoria);
			System.out.println(codigoPessoa);

			System.out.println("---------------------------------------------");

			dataVencimento2 = dataVencimento.toString();
			dataPagamento2 = dataPagamento.toString();

			dataVencimentoConv = LocalDate.parse(dataVencimento2, formatter);
			dataPagamentoConv = LocalDate.parse(dataPagamento2, formatter);
			valor2 = valor.toString();
			valorConv = new BigDecimal(valor2);

			// final Categoria categoria = new Categoria(codigoCategoria.getStringCellValue());

			System.out.println("Descricao-------------------------" + descricao);
			System.out.println("Data Vencimento-------------------" + dataVencimentoConv);
			System.out.println("Data Pagamento--------------------" + dataPagamentoConv);
			System.out.println("Valor-----------------------------" + valorConv);
			System.out.println("Observacao------------------------" + observacao);
			System.out.println("Codigo Categoria------------------" + codigoCategoria);

			final Lancamento lancamento = new Lancamento(descricao,
					dataVencimentoConv,
					dataPagamentoConv,
					valorConv,
					observacao.getStringCellValue(),
					tipo);

			lancamentoRepository.save(lancamento);

		}

	}

	@PostMapping
	@PreAuthorize("hasAuthority('ROLE_CADASTRAR_LANCAMENTO') and #oauth2.hasScope('write')")
	public ResponseEntity<Lancamento> criar(@Valid @RequestBody final Lancamento lancamento,
			final HttpServletResponse response) {
		final Lancamento lancamentoSalvo = lancamentoService.salvar(lancamento);
		publisher.publishEvent(new RecursoCriadoEvent(this, response, lancamentoSalvo.getCodigo()));
		return ResponseEntity.status(HttpStatus.CREATED).body(lancamentoSalvo);
	}

	@ExceptionHandler({ PessoaInexistenteOuInativaException.class })
	public ResponseEntity<Object>
			handlePessoaInexistenteOuInativaException(final PessoaInexistenteOuInativaException ex) {
		final String mensagemUsuario =
				messageSource.getMessage("pessoa.inexistente-ou-inativa", null, LocaleContextHolder.getLocale());
		final String mensagemDesenvolvedor = ex.toString();
		final List<Erro> erros = Arrays.asList(new Erro(mensagemUsuario, mensagemDesenvolvedor));
		return ResponseEntity.badRequest().body(erros);
	}

	@DeleteMapping("/{codigo}")
	@ResponseStatus(HttpStatus.NO_CONTENT)
	@PreAuthorize("hasAuthority('ROLE_REMOVER_LANCAMENTO') and #oauth2.hasScope('write')")
	public void remover(@PathVariable final Long codigo) {
		lancamentoRepository.delete(codigo);
	}

}
