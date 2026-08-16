# Regras de Verificação de PDFs — `verificador.py`

Este documento descreve, com base no código-fonte de `verificador.py`, como cada comprovante em PDF é localizado, associado a um item do Lattes (XML) e validado antes de pontuar.

---

## 1. Estrutura de arquivos esperada

Cada candidato tem uma pasta com `lattes.xml` e os PDFs comprobatórios nomeados por **seção.sequência** (ex.: `1.1.pdf`, `1.2.pdf` para artigos; `7.3.pdf`). A numeração de seção é a mesma do Anexo II do edital (também usada nas colunas do ranking geral, `ranking.xlsx` — ver `_ANEXO2_SECAO_LABELS`):

| Seção | Conteúdo |
|:-----:|----------|
| 1 | Artigos em periódicos |
| 2 | Trabalhos completos **e** resumos em anais (mesmo prefixo — o Anexo II não separa os dois) |
| 4 | Livros (publicados e organizados ficam na mesma pasta/prefixo; a distinção "publicado" (seção 3) vs. "organizado" (seção 4) é feita depois, por conteúdo do campo `TIPO`, só para a comparação declarado×sistema — ver §6) |
| 5 | Capítulos de livro |
| 6 | Bancas examinadoras (compartilha o prefixo com Orientações concluídas — que não é pontuada neste edital e por isso nunca chega a buscar PDF, evitando colisão) |
| 7 | Atuação docente — ensino superior |
| 8 | Atuação docente — ensino básico |
| 9 | Atuação profissional não docente |
| 10 | Projetos de pesquisa |
| 11 | Projetos de extensão |
| 13 | Organização de evento científico |


Seção 3 (livro publicado, isolado de organizado) e 12 (projeto de ensino) existem na numeração do Anexo II mas **não têm pasta/prefixo de PDF próprio** — a primeira porque se resolve por reclassificação de conteúdo dentro do prefixo 4 (acima), a segunda porque o sistema não extrai essa categoria do Lattes (não é um bug: categoria não implementada).

`localizar_pdf()` procura primeiro o arquivo direto na pasta do candidato (`{seq}.pdf`); se não achar, cai para uma estrutura legada em subpastas (`comprovantes/<nome_da_pasta>/{seq}.pdf`, mapeada em `PASTA_POR_CRITERIO`).

---

## 2. Extração de texto do PDF

`extrair_texto_pdf()` (linha 136):

1. Tenta extrair texto das 3 primeiras páginas com `pdfplumber` (padrão; ver §3.1 pra quando o sistema relê com mais páginas).
2. Se não vier texto (PDF escaneado/imagem), cai para OCR via `_ocr_pdf()` — `pdf2image` + `pytesseract` (idiomas português + inglês, 200 dpi). Se as bibliotecas de OCR não estiverem instaladas, retorna vazio.
3. Se a abertura do PDF falhar (arquivo corrompido, senha, etc.), retorna a string `"__ERRO_PDF__: <mensagem>"`, tratada depois como status **ERRO PDF**.
4. PDF sem texto extraível nem OCR → status **SEM TEXTO**.

---

## 3. Como um PDF é associado a um item do Lattes

O processamento é **centrado no PDF, não no item do XML** (`_processar_secao_por_pdf()`, linha 1792): o sistema percorre os PDFs encontrados na pasta e tenta achar, para cada um, o item do Lattes que melhor corresponde — não o contrário. Consequências diretas:

- **PDF sem correspondência no Lattes → REPROVADO** (mesmo que o texto seja legível).
- **Item do Lattes sem PDF correspondente → simplesmente omitido do relatório** (não aparece nem como reprovado).
- Um mesmo item do Lattes só pode ser usado por **um** PDF (`itens_usados`); depois de casado, sai da disputa para os próximos PDFs da seção.
- Entre todos os itens ainda disponíveis, vence o de **maior score de correspondência** (não o primeiro que ultrapassar o limiar) — `melhor_score_global` em `_processar_secao_por_pdf()`.

### 3.1 Cálculo do score de correspondência (`_score_match`, linha 1718)

Usado só para *achar de qual item o PDF é* (não para aprovar/reprovar o conteúdo). Testa os campos "identificadores" do tipo (`_CAMPOS_ID`: título, doi, isbn, issn, evento, orientando, candidato, entidade, editora — **autor é excluído de propósito**, pois o sobrenome do próprio pesquisador aparece em todos os seus PDFs e gera falso positivo):

| Campo | Regra |
|-------|-------|
| `doi` / `isbn` (exato) | bate → score fixo **2,0** (match definitivo, encerra a busca daquele campo) |
| `titulo` | similaridade de palavras (ver §4.1) → score entre 0 e 1 |
| outros campos de identificação (issn, evento, orientando, candidato, entidade, editora) | se o checador aprova → score **0,5** (sinal fraco) |

O maior score entre os campos testados é o score final do par (PDF, item). Um PDF só é aceito para o item de maior score se esse score for **≥ 0,60** (`LIMIAR_TITULO`) e o item ainda não tiver sido usado por outro PDF.

### 3.2 Releitura com mais páginas quando o score "quase bateu"

Dentro de `_processar_secao_por_pdf()` (linha ~1862), se o score de um par (PDF, item) cair numa janela de **0,15 abaixo do limiar** (`JANELA_QUASE_BATEU`, ou seja, entre 0,45 e 0,60), o sistema relê aquele PDF com **até 50 páginas** (`PAGINAS_RETRY`) em vez das 3 padrão e recalcula o score — só substitui se o novo score for maior. Cobre o caso de o candidato anexar o **documento inteiro** em vez de só a parte relevante (ex.: o livro todo em vez do capítulo, ou o caderno de resumos do congresso inteiro em vez de só o resumo dela), quando o título/identificador procurado está bem mais adiante do que as 3 primeiras páginas alcançam.

Se a releitura muda o texto, esse texto mais completo **passa a ser o texto "oficial"** do PDF dali em diante (verificação de campos, extração de período etc.), não só para o cálculo do score.

Esse retry só dispara nos casos "quase bateu" (custo extra evitado quando o PDF já bate ou não bate claramente nas 3 primeiras páginas).

### 3.3 Ordem de PDFs sem XML correspondente

Iniciação Científica (seção 14) e Banca de Especialização (seção 15) **não têm item correspondente no Lattes** (o XML do Lattes não tem tag própria para essas participações) — a "verificação" nesses casos é apenas a **presença do PDF na pasta com nome no padrão correto**, sem cruzamento de conteúdo (`_scan_pdfs_manuais`, linha 834).

---

## 4. Verificação de conteúdo (aprovação/reprovação)

Depois que o PDF é casado com um item, o sistema decide se **aprova** o item checando campos configuráveis por planilha, em **lógica OU**: basta **um** dos campos configurados bater para o item ser aprovado (`verificar_por_config()`, linha 1493). Os campos a checar por subcategoria vêm da coluna **"O que verificar no PDF"** de `criterios_producao.xlsx` (ex.: `titulo + doi + autor`, `titulo + isbn`, `apenas_pdf`).

**Fallback em duas camadas quando a planilha não define os campos** (`_processar_secao_por_pdf`, linha ~1828): se a célula "O que verificar no PDF" de um tipo vier **em branco** na planilha (nem toda linha preenche isso), o sistema não cai direto para o fallback genérico `["titulo"]` — primeiro tenta os padrões embutidos em `_campos_config_padrao()` para aquele tipo, e só usa `["titulo"]` se nem isso existir. Isso importa de verdade para capítulo/livro: `_campos_config_padrao()` configura `titulo + isbn` para esses tipos, e o ISBN é um identificador bem mais confiável que título quando o PDF anexado é o **livro inteiro** (às vezes centenas de páginas) — o ISBN costuma aparecer já na capa/ficha catalográfica, dentro do alcance das 3 primeiras páginas lidas, enquanto o título do capítulo em si pode estar bem mais adiante (ver também §3.2, que cobre o caso via releitura).

Se a planilha inteira não puder ser aberta, o mesmo `_campos_config_padrao()` também é usado como padrão geral do sistema.

### 4.1 Checadores atômicos (`_CHECADORES`, linha 1453)

| Campo | Função | O que verifica | Critério de aprovação |
|-------|--------|-----------------|------------------------|
| `titulo` | `_checar_titulo` | título do item Lattes aparece no texto do PDF | similaridade de palavras ≥ **60%** (`LIMIAR_TITULO`) |
| `doi` | `_checar_doi` | DOI do artigo no texto | DOI normalizado (sem prefixo `http(s)://(dx.)doi.org/`) presente no texto, case-insensitive |
| `issn` | `_checar_issn` | ISSN do periódico | ISSN (sem hífen) presente no texto |
| `autor` | `_checar_autor` | nome do pesquisador no PDF | **sobrenome** normalizado presente no texto |
| `periodico` | `_checar_periodico` | nome do periódico | similaridade de palavras ≥ **50%** (`LIMIAR_PERIODICO`) |
| `evento` | `_checar_evento` | nome do evento/congresso | similaridade de palavras ≥ 50% |
| `isbn` | `_checar_isbn` | ISBN do livro/capítulo | ISBN (sem hífen/espaço) presente no texto |
| `orientando` | `_checar_orientando` | nome do orientando | sobrenome do orientando presente no texto |
| `candidato` | `_checar_candidato` | nome do candidato (banca de graduação) | sobrenome do candidato presente no texto |
| `instituicao` | `_checar_instituicao` | nome da instituição | similaridade **específica de instituição** ≥ **75%** (`LIMIAR_INSTITUICAO`) — ver §4.2 |
| `periodo` | `_checar_periodo` | ano do item | ano presente literalmente no texto |
| `entidade` | `_checar_entidade` | nome da entidade promotora/financiadora | similaridade de palavras ≥ 50% |
| `editora` | `_checar_editora` | nome da editora | similaridade de palavras ≥ 50% |
| `apenas_pdf` | (tratado à parte em `verificar_por_config`) | nenhum — aprova só pela existência do arquivo | sempre aprovado |

**Similaridade de palavras** (`similaridade()`, linha 70): normaliza acentos/maiúsculas, ignora palavras com ≤ 3 letras e uma lista de *stopwords* comuns (artigos, preposições em PT/EN), e mede a fração das palavras relevantes do campo do Lattes que aparecem no texto do PDF.

### 4.2 Similaridade de instituição — regra mais rígida

`similaridade_instituicao()` (linha 103) usa a mesma lógica, mas remove também um conjunto adicional de termos genéricos demais para identificar uma instituição específica: tipo de instituição (universidade, faculdade, instituto...), termos como "ensino/superior/federal/estadual", e nomes de estado/UF. Isso evita, por exemplo, que "Faculdade de Ensino de Minas Gerais" seja confundida com "Universidade do Estado de Minas Gerais" só por compartilharem palavras genéricas.

Se, depois de remover esses termos, **nenhuma palavra distintiva sobrar** no nome da instituição do Lattes, a função retorna **0,0 de propósito** (nunca aprova por coincidência) — força revisão manual em vez de arriscar falso positivo.

---

## 5. Caso especial: Atuação Profissional (seções 7/8/9)

Tem verificador próprio, `verificar_atuacao()` (linha 1277), com lógica **E** (todas as condições precisam passar, ao contrário da lógica OU das demais seções):

| Condição | Regra |
|----------|-------|
| Nome | sobrenome do pesquisador aparece no PDF |
| Instituição | similaridade de instituição ≥ 75% (mesma regra do §4.2) |
| Cargo | se o Lattes informou cargo/função, similaridade de palavras ≥ 50%; se o Lattes não informou cargo, essa condição é considerada automaticamente ok (nada a conferir) |
| Período | ano de início precisa aparecer no PDF **e** (ano de fim aparece no PDF **ou**, se o vínculo não tem fim, alguma palavra indicando vínculo em curso — "atual", "presente data", "momento", "vigente" — aparece no texto) |

O item só é **APROVADO** se as 4 condições passarem simultaneamente. Como um mesmo PDF pode teoricamente bater com mais de um vínculo do candidato, o sistema testa **todos os vínculos ainda não usados** e fica com o de **maior similaridade de instituição**, não o primeiro que passar — mesma filosofia "melhor score global" usada nas demais seções.

Classificação da atividade (docência superior / básica / não docência), usada para definir a pontuação por mês, é feita por `classificar_atividade()` (linha 974) por palavra-chave no cargo ("professor"/"docente") e no nome da instituição (listas `_SUPERIOR_KW` e `_BASICO_KW`, linha 956). `_BASICO_KW` inclui, além de nomes de escola ("escola estadual/municipal", "colégio"...), termos de **administração municipal/estadual da educação** — "prefeitura (municipal)", "município de", "secretaria (municipal/estadual) de educação" — porque no Brasil professor(a) de rede municipal/estadual costuma ter, no Lattes, a Prefeitura/Secretaria em si como "instituição" (não uma escola nomeada), já que o município administra diretamente educação infantil e ensino fundamental. Sem esses termos, esse vínculo caía em "não docência" por não bater com nenhuma palavra-chave de instituição.

### 5.1 Extração do período do vínculo — prioridade da "Carteira de Trabalho Digital"

`_periodo_do_pdf()` (linha 1098) tenta, em ordem de prioridade:

1. **"Contratos de trabalho"** — padrão específico do Extrato de Outros Vínculos / Carteira de Trabalho Digital (governo, via eSocial), que lista o contrato logo após esse cabeçalho como `DD/MM/AAAA - Aberto` (vigente, conta até hoje) ou `DD/MM/AAAA - DD/MM/AAAA` (encerrado) — `_CONTRATO_TRABALHO_RE` / `_pares_contratos_trabalho()`, linha 1026. Tem **prioridade máxima**: sem esse reconhecimento específico, a busca genérica de par de datas podia pegar datas de **outra seção** do mesmo documento (ex.: "Anotações: ... Férias DD/MM/AAAA a DD/MM/AAAA"), que não têm nada a ver com a duração real do vínculo. Documento pode listar mais de um contrato (histórico de empregos); todos são extraídos e o período final cobre do menor início ao maior fim entre eles.
2. Pares de data numa janela ancorada no nome do pesquisador (evita misturar com o período de outra pessoa em documento que lista vários vínculos).
3. Pares "tendo atuado de X até Y" (participação real da pessoa, mais confiável que a duração nominal do programa/edital).
4. Padrão genérico `DD/MM/AAAA <conector> DD/MM/AAAA` em todo o texto, como último recurso.

Retorna `None` (sem período reconhecível) se nenhum padrão bater.

---

## 6. Comparação com a autodeclaração do candidato (Anexo II)

Além de aprovar/reprovar item a item, o sistema tenta ler o **Anexo II** que o próprio candidato preenche (autodeclaração de pontos por seção) e comparar, seção a seção, com o que os comprovantes sustentam — mostrado lado a lado na aba "Resumo" do relatório individual (Decl. Qtd./Decl. Pts vs. Sist. Aprov./Sist. Pts). `_extrair_declarado_anexo2()` (linha 2120) localiza o PDF do Anexo II na pasta do candidato (`_achar_anexo2`, por nome do arquivo ou, se não achar, pelo cabeçalho característico no conteúdo) e isola o bloco de texto de cada seção numerada (do "N." até o próximo "N+1.").

Como o PDF é gerado a partir de um `.docx`, a quebra de linha dentro de cada bloco é inconsistente (quantidade e pontos às vezes na mesma linha da seção, às vezes separados por texto livre — inclusive o título de cada item, que pode ter um número solto no meio, ex. um ano). Duas regras tratam isso:

- **Par adjacente "quantidade pontos" tem prioridade** sobre "pegar o último número solto do bloco": sem isso, um título como "CAMINHADA NÓRDICA PARA IDOSOS 2023" fazia o sistema capturar **2023** (o ano) como se fosse a quantidade declarada, em vez do **2** de verdade. Só cai para "último inteiro solto do bloco" se não achar esse par adjacente.
- **Sub-total em branco (seção de Artigos) tem fallback**: nem todo Anexo II preenche a célula "Sub-total" mesmo com as linhas de Qualis individuais (Primeiro/Demais autor por nível) tendo valor — quando isso acontece, o sistema soma o par quantidade/pontos de cada linha preenchida do bloco em vez de considerar a seção sem declaração nenhuma.

Essa comparação é só de **exibição/conferência manual** — não altera a pontuação do candidato, que continua vindo inteiramente do lado "sistema" (comprovantes aprovados). O mesmo agrupamento por seção usado aqui do lado "sistema" (`_pontos_por_secao()`) é o que alimenta as colunas de quesito do ranking geral (`ranking.xlsx`).

---

## 7. Status possíveis de cada item no relatório

| Status | Quando ocorre |
|--------|---------------|
| **APROVADO** | PDF casado com item do Lattes e passou na verificação de conteúdo (§4 ou, na Atuação Profissional, §5) |
| **REPROVADO** | PDF não atingiu o score mínimo de correspondência com nenhum item do Lattes (ou, na Atuação Profissional, nenhum vínculo passou nas 4 condições) |
| **ERRO PDF** | Falha ao abrir/ler o arquivo PDF |
| **SEM TEXTO** | PDF abriu, mas não foi possível extrair texto (nem via OCR) |

Somente itens **APROVADO** entram no somatório de pontos do candidato.

---

## 8. Limiares de similaridade usados (constantes do código)

| Constante | Valor | Uso |
|-----------|:-----:|-----|
| `LIMIAR_TITULO` | 0,60 | título do item vs. texto do PDF; também o score mínimo de correspondência PDF↔item |
| `LIMIAR_PERIODICO` | 0,50 | periódico, evento, cargo (atuação), entidade, editora |
| `LIMIAR_INSTITUICAO` | 0,75 | nome de instituição (mais rígido, pois nomes de instituição são curtos e cada palavra pesa muito) |

---

## 9. Observações importantes de comportamento

- **Match por DOI ou ISBN é definitivo** (score 2,0): mesmo que o título tenha baixa similaridade textual (ex.: PDF só com a capa da revista), DOI ou ISBN exatos bastam para casar o PDF com o item certo.
- **Campo "autor" nunca é usado para casar PDF a item** (só para aprovar conteúdo, via §4.1) — usá-lo para casamento faria qualquer PDF do próprio pesquisador "bater" com qualquer item, pois o sobrenome se repete em todos os documentos.
- A ordem de leitura dos PDFs dentro de uma seção segue a numeração do nome do arquivo (`1.1.pdf` antes de `1.2.pdf`), mas a numeração **não precisa bater com a ordem dos itens no XML** — o casamento é sempre por conteúdo/score, nunca por posição.
- Se a planilha de critérios não puder ser aberta, todo o sistema (pontos e campos a verificar) cai para os valores-padrão embutidos em `_criterios_padrao()` e `_campos_config_padrao()`; o mesmo padrão também cobre, tipo a tipo, uma célula em branco na coluna "O que verificar no PDF" mesmo com o resto da planilha carregando normalmente (§4).
- PDF com score "quase bateu" (§3.2) é relido com mais páginas antes de ser dado como reprovado — o retry só acontece nesse caso, não em todo PDF.
- Extração de período de vínculo empregatício reconhece o formato específico da Carteira de Trabalho Digital/eSocial ("Contratos de trabalho") com prioridade máxima sobre os padrões genéricos de data (§5.1).
