// =============================================
// DASHBOARD DE VENDAS
// Importação Excel, indicadores, gráficos, cadastro
// e monitoramento automático de alterações no Excel
// =============================================

const STORAGE_KEY = "dashboard_vendas_registros_v1";

// Conexão com o arquivo Excel local
let arquivoExcelHandle = null;
let monitoramentoExcel = null;
let ultimaModificacaoExcel = null;
let verificandoExcel = false;

let pedidos = carregarDados();
let catalogo = new Map();
let graficoMensal = null;
let graficoRegiao = null;

// Atalho para localizar elementos HTML pelo ID
const $ = (id) => document.getElementById(id);

// ---------------------------------------------
// FUNÇÕES GERAIS
// ---------------------------------------------

function moeda(valor) {
  return Number(valor || 0).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });
}

// Converte números no formato brasileiro ou numérico do Excel
function numero(valor) {
  if (typeof valor === "number") {
    return Number.isFinite(valor) ? valor : 0;
  }

  if (valor === null || valor === undefined || valor === "") {
    return 0;
  }

  let texto = String(valor).trim();

  // Ex.: "R$ 1.234,56"
  texto = texto.replace(/[^\d,.-]/g, "");

  if (texto.includes(",")) {
    texto = texto.replace(/\./g, "").replace(",", ".");
  }

  const convertido = Number(texto);
  return Number.isFinite(convertido) ? convertido : 0;
}

function carregarDados() {
  try {
    const dados = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return Array.isArray(dados) ? dados : [];
  } catch (erro) {
    console.error("Erro ao carregar dados:", erro);
    return [];
  }
}

function salvarDados() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(pedidos));
  } catch (erro) {
    console.error("Erro ao salvar dados no navegador:", erro);
  }
}

// Converte datas do Excel para AAAA-MM-DD
function converterData(valor) {
  if (!valor) return "";

  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    const ano = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, "0");
    const dia = String(valor.getDate()).padStart(2, "0");
    return `${ano}-${mes}-${dia}`;
  }

  if (typeof valor === "number" && window.XLSX) {
    const data = XLSX.SSF.parse_date_code(valor);

    if (data) {
      return `${data.y}-${String(data.m).padStart(2, "0")}-${String(data.d).padStart(2, "0")}`;
    }
  }

  const texto = String(valor).trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) {
    return texto.slice(0, 10);
  }

  const partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (partes) {
    return `${partes[3]}-${partes[2].padStart(2, "0")}-${partes[1].padStart(2, "0")}`;
  }

  return "";
}

function dataBrasileira(valor) {
  if (!valor) return "";

  const partes = String(valor).split("-");
  if (partes.length !== 3) return valor;

  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function dataLocalISO() {
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = String(hoje.getMonth() + 1).padStart(2, "0");
  const dia = String(hoje.getDate()).padStart(2, "0");
  return `${ano}-${mes}-${dia}`;
}

// ---------------------------------------------
// CATÁLOGO DE PRODUTOS
// ---------------------------------------------

function adicionarProdutoCatalogo(nome, preco) {
  if (!nome) return;
  catalogo.set(String(nome).trim(), numero(preco));
}

function atualizarListaProdutos() {
  const campoProduto = $("produto");
  if (!campoProduto) return;

  const produtoSelecionado = campoProduto.value;

  campoProduto.innerHTML =
    '<option value="">Selecione um produto</option>';

  [...catalogo.keys()].sort().forEach((nome) => {
    const opcao = document.createElement("option");
    opcao.value = nome;
    opcao.textContent = nome;
    campoProduto.appendChild(opcao);
  });

  if (catalogo.has(produtoSelecionado)) {
    campoProduto.value = produtoSelecionado;
  }
}

// Reconstrói o catálogo a partir da planilha e dos pedidos
function reconstruirCatalogo(linhas, registros) {
  catalogo.clear();

  linhas.forEach((linha) => {
    adicionarProdutoCatalogo(linha["Produtos.1"], linha["Valor"]);
  });

  registros.forEach((pedido) => {
    adicionarProdutoCatalogo(pedido.produto, pedido.valorUnitario);
  });

  atualizarListaProdutos();
}

// ---------------------------------------------
// INDICADORES E TABELA
// ---------------------------------------------

function atualizarIndicadores() {
  const faturamento = pedidos.reduce(
    (soma, pedido) => soma + numero(pedido.total),
    0
  );

  const quantidade = pedidos.reduce(
    (soma, pedido) => soma + numero(pedido.quantidade),
    0
  );

  const ticketMedio = pedidos.length
    ? faturamento / pedidos.length
    : 0;

  if ($("kpiRevenue")) $("kpiRevenue").textContent = moeda(faturamento);

  if ($("kpiOrders")) {
    $("kpiOrders").textContent = pedidos.length.toLocaleString("pt-BR");
  }

  if ($("kpiUnits")) {
    $("kpiUnits").textContent = quantidade.toLocaleString("pt-BR");
  }

  if ($("kpiTicket")) $("kpiTicket").textContent = moeda(ticketMedio);

  if ($("recordCount")) {
    $("recordCount").textContent = `${pedidos.length} registro(s)`;
  }

  if ($("dataStatus") && !arquivoExcelHandle) {
    $("dataStatus").textContent = pedidos.length
      ? `${pedidos.length} registros carregados`
      : "Aguardando importação da planilha";
  }
}

function atualizarTabela(filtro = "") {
  const corpo = $("ordersBody");
  if (!corpo) return;

  corpo.innerHTML = "";
  const termo = String(filtro).toLocaleLowerCase("pt-BR");

  const encontrados = pedidos.filter((pedido) => {
    const texto = [
      pedido.pedido,
      pedido.regiao,
      pedido.estado,
      pedido.produto
    ].join(" ").toLocaleLowerCase("pt-BR");

    return texto.includes(termo);
  });

  if (encontrados.length === 0) {
    const mensagem = pedidos.length
      ? "Nenhum registro encontrado."
      : "Importe sua planilha Excel para visualizar os registros.";

    corpo.innerHTML =
      `<tr><td colspan="9" class="empty">${mensagem}</td></tr>`;
    return;
  }

  encontrados
    .slice()
    .sort((a, b) => numero(b.pedido) - numero(a.pedido))
    .forEach((pedido) => {
      const linha = document.createElement("tr");

      const valores = [
        pedido.pedido,
        dataBrasileira(pedido.data),
        pedido.regiao,
        pedido.estado,
        pedido.produto,
        pedido.quantidade,
        moeda(pedido.valorUnitario),
        moeda(pedido.frete),
        moeda(pedido.total)
      ];

      valores.forEach((valor) => {
        const celula = document.createElement("td");
        celula.textContent = valor ?? "";
        linha.appendChild(celula);
      });

      corpo.appendChild(linha);
    });
}

// ---------------------------------------------
// GRÁFICOS
// ---------------------------------------------

function atualizarGraficos() {
  if (!window.Chart) {
    console.warn("Chart.js não foi carregado.");
    return;
  }

  const valoresPorMes = new Map();
  const valoresPorRegiao = new Map();

  pedidos.forEach((pedido) => {
    const mes = pedido.data ? pedido.data.slice(0, 7) : "Sem data";
    const regiao = pedido.regiao || "Não informada";

    valoresPorMes.set(
      mes,
      (valoresPorMes.get(mes) || 0) + numero(pedido.total)
    );

    valoresPorRegiao.set(
      regiao,
      (valoresPorRegiao.get(regiao) || 0) + numero(pedido.total)
    );
  });

  const meses = [...valoresPorMes.keys()].sort();

  const nomesMeses = meses.map((mes) => {
    if (mes === "Sem data") return mes;

    const [ano, numeroMes] = mes.split("-");

    return new Date(Number(ano), Number(numeroMes) - 1, 1)
      .toLocaleDateString("pt-BR", {
        month: "short",
        year: "2-digit"
      });
  });

  const elementoMensal = $("monthlyChart");

  if (elementoMensal) {
    if (graficoMensal) graficoMensal.destroy();

    graficoMensal = new Chart(elementoMensal, {
      type: "bar",
      data: {
        labels: nomesMeses,
        datasets: [{
          label: "Faturamento",
          data: meses.map((mes) => valoresPorMes.get(mes)),
          backgroundColor: "#3985b8",
          borderRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              callback: (valor) =>
                "R$ " + Number(valor).toLocaleString("pt-BR")
            }
          },
          x: {
            grid: { display: false }
          }
        }
      }
    });
  }

  const elementoRegiao = $("regionChart");

  if (elementoRegiao) {
    if (graficoRegiao) graficoRegiao.destroy();

    graficoRegiao = new Chart(elementoRegiao, {
      type: "doughnut",
      data: {
        labels: [...valoresPorRegiao.keys()],
        datasets: [{
          data: [...valoresPorRegiao.values()],
          backgroundColor: [
            "#2878ad",
            "#69a6c9",
            "#9bbdce",
            "#e2a85b",
            "#52718b",
            "#9aabb9"
          ]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              boxWidth: 10,
              font: { size: 10 }
            }
          }
        }
      }
    });
  }
}

function atualizarDashboard() {
  atualizarIndicadores();

  atualizarTabela(
    $("searchInput") ? $("searchInput").value : ""
  );

  atualizarGraficos();
  atualizarGraficosExtras();
}

// ---------------------------------------------
// IMPORTAÇÃO E LEITURA DO EXCEL
// ---------------------------------------------

// Converte uma linha da planilha em um pedido
function converterLinhaPedido(linha) {
  const quantidade = numero(linha["Quantidade"]);
  const valorUnitario = numero(linha["Valor Unitário"]);
  const frete = numero(linha["Frete"]);

  return {
    pedido: numero(linha["Pedido"]),
    regiao: linha["Região"] || "",
    estado: String(linha["Estado"] || "").toUpperCase(),
    produto: linha["Produtos"] || "",
    data: converterData(linha["Data"]),
    quantidade,
    valorUnitario,
    frete,

    // O total é recalculado para acompanhar mudanças na quantidade,
    // no preço unitário ou no frete.
    total: quantidade * valorUnitario + frete
  };
}

// Lê os dados de um arquivo Excel e atualiza o dashboard
async function processarArquivoExcel(arquivo, automatico = false, bufferExistente = null) {
  if (!window.XLSX) {
    throw new Error("A biblioteca Excel não foi carregada.");
  }

  const buffer = bufferExistente || await arquivo.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const nomeAba = workbook.SheetNames[0];

  if (!nomeAba) {
    throw new Error("A planilha não possui nenhuma aba.");
  }

  const primeiraAba = workbook.Sheets[nomeAba];

  const linhas = XLSX.utils.sheet_to_json(
    primeiraAba,
    { defval: "" }
  );

  const pedidosImportados = linhas
    .filter((linha) =>
      linha["Pedido"] !== "" &&
      linha["Pedido"] !== null &&
      linha["Pedido"] !== undefined
    )
    .map(converterLinhaPedido)
    .filter((pedido) => pedido.pedido);

  // Permite que a planilha fique sem registros, desde que tenha sido lida.
  pedidos = pedidosImportados;

  reconstruirCatalogo(linhas, pedidos);
  salvarDados();
  atualizarDashboard();

  if ($("dataStatus")) {
    $("dataStatus").textContent = automatico
      ? `Excel atualizado às ${new Date().toLocaleTimeString("pt-BR")}`
      : `${pedidos.length} registros carregados`;
  }

  return pedidos.length;
}

// Importação manual pelo input de arquivo existente no HTML
async function importarExcel(evento) {
  const arquivo = evento.target.files[0];
  if (!arquivo) return;

  try {
    const quantidade = await processarArquivoExcel(arquivo, false);

    alert(`Importação concluída: ${quantidade} pedidos carregados.`);
  } catch (erro) {
    console.error("Erro ao importar Excel:", erro);

    alert(
      "Não foi possível importar a planilha. Confira o arquivo e os nomes das colunas."
    );
  } finally {
    evento.target.value = "";
  }
}

// ---------------------------------------------
// MONITORAMENTO AUTOMÁTICO DO ARQUIVO LOCAL
// ---------------------------------------------

// Gera uma assinatura do conteúdo. Assim, também detecta alterações
// quando o tamanho do arquivo permanece igual.
function gerarAssinaturaArquivo(buffer, arquivo) {
  const bytes = new Uint8Array(buffer);
  let hash = 2166136261;

  for (let i = 0; i < bytes.length; i++) {
    hash ^= bytes[i];
    hash = Math.imul(hash, 16777619);
  }

  return `${arquivo.lastModified}-${arquivo.size}-${hash >>> 0}`;
}

function iniciarMonitoramentoExcel() {
  if (monitoramentoExcel) {
    clearInterval(monitoramentoExcel);
  }

  monitoramentoExcel = setInterval(
    verificarAlteracaoExcel,
    2000
  );
}

async function verificarAlteracaoExcel() {
  if (!arquivoExcelHandle || verificandoExcel) return;

  verificandoExcel = true;

  try {
    const permissao = await arquivoExcelHandle.queryPermission({
      mode: "read"
    });

    if (permissao !== "granted") {
      if ($("dataStatus")) {
        $("dataStatus").textContent =
          "Acesso ao Excel pausado. Reconecte o arquivo.";
      }

      clearInterval(monitoramentoExcel);
      monitoramentoExcel = null;
      return;
    }

    const arquivo = await arquivoExcelHandle.getFile();
    const buffer = await arquivo.arrayBuffer();
    const assinatura = gerarAssinaturaArquivo(buffer, arquivo);

    if (assinatura !== ultimaModificacaoExcel) {
      await processarArquivoExcel(arquivo, true, buffer);
      ultimaModificacaoExcel = assinatura;

      console.log("Dashboard atualizado com a versão mais recente do Excel.");
    }
  } catch (erro) {
    // Se o Excel estiver salvando ou temporariamente bloqueado,
    // a próxima verificação tentará novamente.
    console.warn("Não foi possível verificar o Excel nesta tentativa:", erro);
  } finally {
    verificandoExcel = false;
  }
}

// Cria o botão de conexão com o Excel
function criarBotaoConexaoExcel() {
  const campoExcel = $("excelInput");

  if (!campoExcel || $("connectExcelBtn")) return;

  const botao = document.createElement("button");

  botao.id = "connectExcelBtn";
  botao.type = "button";
  botao.textContent = "Conectar arquivo Excel";
  botao.className = "btn btn-primary";
  botao.style.marginLeft = "10px";

  botao.addEventListener("click", conectarArquivoExcel);

  campoExcel.insertAdjacentElement("afterend", botao);
}

// Abre o seletor de arquivos e conecta a planilha local
async function conectarArquivoExcel() {
  if (!window.showOpenFilePicker) {
    alert(
      "Seu navegador não oferece suporte à conexão automática. " +
      "Use uma versão atualizada do Google Chrome ou Microsoft Edge " +
      "e abra o projeto pelo Live Server (localhost)."
    );
    return;
  }

  try {
    const arquivos = await window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: "Planilhas Excel",
        accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel": [".xls"]
        }
      }]
    });

    if (!arquivos || !arquivos[0]) return;

    const novoHandle = arquivos[0];

    const permissao = await novoHandle.requestPermission({
      mode: "read"
    });

    if (permissao !== "granted") {
      alert("Permissão de leitura do arquivo não concedida.");
      return;
    }

    arquivoExcelHandle = novoHandle;

    const arquivo = await arquivoExcelHandle.getFile();
    const buffer = await arquivo.arrayBuffer();

    const quantidade = await processarArquivoExcel(
      arquivo,
      false,
      buffer
    );

    ultimaModificacaoExcel =
      gerarAssinaturaArquivo(buffer, arquivo);

    iniciarMonitoramentoExcel();

    if ($("dataStatus")) {
      $("dataStatus").textContent =
        `Excel conectado: ${quantidade} registros monitorados`;
    }

    alert(
      "Planilha conectada com sucesso!\n\n" +
      "Salve as alterações no Excel para que o dashboard seja atualizado."
    );
  } catch (erro) {
    if (erro.name !== "AbortError") {
      console.error("Erro ao conectar Excel:", erro);
      alert("Não foi possível conectar a planilha. Verifique o arquivo e as permissões.");
    }
  }
}

// ---------------------------------------------
// FORMULÁRIO DE CADASTRO
// ---------------------------------------------

function atualizarTotalPreview() {
  const quantidade = numero($("quantidade")?.value);
  const valorUnitario = numero($("valorUnitario")?.value);
  const frete = numero($("frete")?.value);

  const total = quantidade * valorUnitario + frete;

  if ($("totalPreview")) {
    $("totalPreview").textContent = moeda(total);
  }
}

function cadastrarPedido(evento) {
  evento.preventDefault();

  const novoPedido = numero($("pedido")?.value);

  if (!novoPedido) {
    if ($("formMessage")) {
      $("formMessage").textContent = "Informe um número de pedido válido.";
    }
    return;
  }

  const pedidoExiste = pedidos.some(
    (pedido) => numero(pedido.pedido) === novoPedido
  );

  if (pedidoExiste) {
    if ($("formMessage")) {
      $("formMessage").textContent =
        "Esse número de pedido já existe. Informe outro número.";
    }
    return;
  }

  const quantidade = numero($("quantidade")?.value);
  const valorUnitario = numero($("valorUnitario")?.value);
  const frete = numero($("frete")?.value);

  const registro = {
    pedido: novoPedido,
    data: $("data")?.value || dataLocalISO(),
    regiao: $("regiao")?.value || "",
    estado: ($("estado")?.value || "").trim().toUpperCase(),
    produto: $("produto")?.value || "",
    quantidade,
    valorUnitario,
    frete,
    total: quantidade * valorUnitario + frete
  };

  pedidos.push(registro);

  adicionarProdutoCatalogo(
    registro.produto,
    registro.valorUnitario
  );

  salvarDados();
  atualizarListaProdutos();
  atualizarDashboard();

  if ($("formMessage")) {
    $("formMessage").textContent =
      `Pedido ${novoPedido} cadastrado com sucesso!`;
  }

  $("orderForm")?.reset();

  if ($("pedido")) {
    $("pedido").value = Math.max(
      1000,
      ...pedidos.map((pedido) => numero(pedido.pedido))
    ) + 1;
  }

  if ($("data")) $("data").value = dataLocalISO();

  atualizarTotalPreview();
}

// ---------------------------------------------
// EXPORTAÇÃO PARA EXCEL
// ---------------------------------------------

function exportarExcel() {
  if (!window.XLSX) {
    alert("A biblioteca Excel não foi carregada.");
    return;
  }

  if (pedidos.length === 0) {
    alert("Não existem pedidos para exportar.");
    return;
  }

  const dados = pedidos.map((pedido) => ({
    "Pedido": pedido.pedido,
    "Região": pedido.regiao,
    "Estado": pedido.estado,
    "Produtos": pedido.produto,
    "Data": pedido.data,
    "Quantidade": pedido.quantidade,
    "Valor Unitário": pedido.valorUnitario,
    "Frete": pedido.frete,
    "Total": pedido.total
  }));

  const planilha = XLSX.utils.json_to_sheet(dados);
  const arquivo = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    arquivo,
    planilha,
    "Vendas"
  );

  XLSX.writeFile(
    arquivo,
    "base_vendas_atualizada.xlsx"
  );
}

// ---------------------------------------------
// INICIALIZAÇÃO DO DASHBOARD
// ---------------------------------------------

function iniciarDashboard() {
  criarBotaoConexaoExcel();

  if ($("excelInput")) {
    $("excelInput").addEventListener("change", importarExcel);
  }

  if ($("orderForm")) {
    $("orderForm").addEventListener("submit", cadastrarPedido);

    $("orderForm").addEventListener("reset", () => {
      setTimeout(() => {
        if ($("formMessage")) {
          $("formMessage").textContent = "";
        }

        atualizarTotalPreview();
      }, 0);
    });
  }

  ["quantidade", "valorUnitario", "frete"].forEach((id) => {
    if ($(id)) {
      $(id).addEventListener("input", atualizarTotalPreview);
    }
  });

  if ($("produto")) {
    $("produto").addEventListener("change", () => {
      const preco = catalogo.get($("produto").value);

      if (preco !== undefined && $("valorUnitario")) {
        $("valorUnitario").value = preco;
      }

      atualizarTotalPreview();
    });
  }

  if ($("searchInput")) {
    $("searchInput").addEventListener("input", (evento) => {
      atualizarTabela(evento.target.value);
    });
  }

  if ($("exportBtn")) {
    $("exportBtn").addEventListener("click", exportarExcel);
  }

  if ($("data") && !$("data").value) {
    $("data").value = dataLocalISO();
  }

  if ($("pedido")) {
    $("pedido").value = Math.max(
      1000,
      ...pedidos.map((pedido) => numero(pedido.pedido))
    ) + 1;
  }

  pedidos.forEach((pedido) => {
    adicionarProdutoCatalogo(
      pedido.produto,
      pedido.valorUnitario
    );
  });

  atualizarListaProdutos();
  atualizarDashboard();
  atualizarTotalPreview();


// ==========================================
// GRÁFICOS COMPLEMENTARES
// ==========================================

const graficosExtras = {};

function campoPedido(pedido, nomes) {
    for (const nome of nomes) {
        if (pedido[nome] !== undefined && pedido[nome] !== null) {
            return pedido[nome];
        }
    }

    return "";
}

function numeroDashboard(valor) {
    if (typeof valor === "number") {
        return Number.isFinite(valor) ? valor : 0;
    }

    if (valor === null || valor === undefined || valor === "") {
        return 0;
    }

    let texto = String(valor).trim();

    if (texto.includes(",")) {
        texto = texto.replace(/\./g, "").replace(",", ".");
    }

    const numero = Number(texto);

    return Number.isFinite(numero) ? numero : 0;
}

function moedaDashboard(valor) {
    return numeroDashboard(valor).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL"
    });
}

function obterDadosDashboard() {
    return (Array.isArray(pedidos) ? pedidos : []).map(pedido => {
        const produto = campoPedido(pedido, [
            "produto", "Produtos", "Produtos.1"
        ]) || "Não identificado";

        const regiao = campoPedido(pedido, [
            "regiao", "Região"
        ]) || "Não identificada";

        const estado = campoPedido(pedido, [
            "estado", "Estado"
        ]) || "Não identificado";

        const quantidade = numeroDashboard(
            campoPedido(pedido, ["quantidade", "Quantidade"])
        );

        const valorUnitario = numeroDashboard(
            campoPedido(pedido, [
                "valorUnitario", "Valor Unitário"
            ])
        );

        const frete = numeroDashboard(
            campoPedido(pedido, ["frete", "Frete"])
        );

        const totalInformado = campoPedido(pedido, [
            "total", "Total"
        ]);

        const total = totalInformado !== ""
            ? numeroDashboard(totalInformado)
            : quantidade * valorUnitario + frete;

        const pedidoId = campoPedido(pedido, [
            "pedido", "Pedido"
        ]);

        return {
            produto,
            regiao,
            estado,
            quantidade,
            valorUnitario,
            frete,
            total,
            faturamentoProduto: quantidade * valorUnitario,
            pedidoId
        };
    });
}

// ==========================================
// FUNÇÃO PARA CRIAR OU ATUALIZAR GRÁFICOS
// ==========================================

function criarGraficoExtra(id, tipo, labels, dados, configuracao = {}) {
    const elemento = document.getElementById(id);

    if (!elemento || typeof Chart === "undefined") {
        return;
    }

    if (graficosExtras[id]) {
        graficosExtras[id].destroy();
    }

    const cores = [
        "#668bb5",
        "#70a990",
        "#a18bbd",
        "#b49a70",
        "#6f9da5",
        "#9b8290",
        "#7d91a8",
        "#8b9c76"
    ];

    graficosExtras[id] = new Chart(elemento, {
        type: tipo,

        data: {
            labels: labels,

            datasets: [{
                label: configuracao.label || "Vendas",
                data: dados,
                backgroundColor: tipo === "doughnut"
                    ? labels.map((_, i) => cores[i % cores.length])
                    : configuracao.cor || "#668bb5",

                borderColor: tipo === "doughnut"
                    ? "#181e25"
                    : configuracao.borda || "#7b9fc7",

                borderWidth: tipo === "doughnut" ? 3 : 1,
                borderRadius: tipo === "bar" ? 5 : 0,
                hoverOffset: tipo === "doughnut" ? 8 : 0
            }]
        },

        options: {
            responsive: true,
            maintainAspectRatio: false,

            indexAxis: configuracao.horizontal ? "y" : "x",

            plugins: {
                legend: {
                    display: tipo === "doughnut",
                    position: "bottom",

                    labels: {
                        color: "#aab6c4",
                        padding: 18,
                        usePointStyle: true
                    }
                },

                tooltip: {
                    backgroundColor: "#202832",
                    titleColor: "#ffffff",
                    bodyColor: "#dce5ef",
                    borderColor: "#344252",
                    borderWidth: 1,

                    callbacks: {
                        label: function(context) {
                            const valor = context.parsed.y ??
                                          context.parsed.x ??
                                          context.parsed;

                            if (configuracao.moeda) {
                                return moedaDashboard(valor);
                            }

                            return Number(valor).toLocaleString("pt-BR");
                        }
                    }
                }
            },

            scales: tipo === "doughnut" ? {} : {
                x: {
                    beginAtZero: true,
                    grid: {
                        color: "rgba(255,255,255,0.045)"
                    },
                    ticks: {
                        color: "#8f9eae"
                    }
                },

                y: {
                    beginAtZero: true,
                    grid: {
                        color: "rgba(255,255,255,0.045)"
                    },
                    ticks: {
                        color: "#8f9eae"
                    }
                }
            }
        }
    });
}

// ==========================================
// ATUALIZAR TODOS OS GRÁFICOS COMPLEMENTARES
// ==========================================

function atualizarGraficosExtras() {
    const dados = obterDadosDashboard();

    if (!dados.length) {
        return;
    }

    // 1. FATURAMENTO POR PRODUTO
    const faturamentoProdutos = {};

    dados.forEach(item => {
        faturamentoProdutos[item.produto] =
            (faturamentoProdutos[item.produto] || 0) +
            item.faturamentoProduto;
    });

    const produtosFaturamento = Object.entries(faturamentoProdutos)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

    criarGraficoExtra(
        "productChart",
        "bar",
        produtosFaturamento.map(item => item[0]),
        produtosFaturamento.map(item => item[1]),
        {
            horizontal: true,
            moeda: true,
            label: "Faturamento",
            cor: "#668bb5"
        }
    );

    // 2. PEDIDOS POR ESTADO
    const pedidosEstado = {};

    dados.forEach(item => {
        pedidosEstado[item.estado] =
            (pedidosEstado[item.estado] || 0) + 1;
    });

    const estados = Object.entries(pedidosEstado)
        .sort((a, b) => b[1] - a[1]);

    criarGraficoExtra(
        "stateChart",
        "doughnut",
        estados.map(item => item[0]),
        estados.map(item => item[1]),
        {
            label: "Pedidos"
        }
    );

    // 3. FRETE POR REGIÃO
    const freteRegiao = {};

    dados.forEach(item => {
        freteRegiao[item.regiao] =
            (freteRegiao[item.regiao] || 0) + item.frete;
    });

    const regioesFrete = Object.entries(freteRegiao)
        .sort((a, b) => b[1] - a[1]);

    criarGraficoExtra(
        "shippingChart",
        "bar",
        regioesFrete.map(item => item[0]),
        regioesFrete.map(item => item[1]),
        {
            moeda: true,
            label: "Frete",
            cor: "#70a990"
        }
    );

    // 4. QUANTIDADE VENDIDA POR PRODUTO
    const quantidadeProdutos = {};

    dados.forEach(item => {
        quantidadeProdutos[item.produto] =
            (quantidadeProdutos[item.produto] || 0) +
            item.quantidade;
    });

    const rankingProdutos = Object.entries(quantidadeProdutos)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

    criarGraficoExtra(
        "quantityChart",
        "bar",
        rankingProdutos.map(item => item[0]),
        rankingProdutos.map(item => item[1]),
        {
            horizontal: true,
            label: "Unidades vendidas",
            cor: "#a18bbd"
        }
    );

    // ATUALIZAR PAINEL DE INSIGHTS
    atualizarInsights(dados, quantidadeProdutos, pedidosEstado,
        faturamentoProdutos);

    // ATUALIZAR ÚLTIMOS PEDIDOS
    atualizarUltimosPedidos(dados);
}

// ==========================================
// PAINEL DE INSIGHTS
// ==========================================

function atualizarInsights(
    dados,
    quantidadeProdutos,
    pedidosEstado,
    faturamentoProdutos
) {
    const produtoMaisVendido = Object.entries(quantidadeProdutos)
        .sort((a, b) => b[1] - a[1])[0];

    const estadoMaisVendido = Object.entries(pedidosEstado)
        .sort((a, b) => b[1] - a[1])[0];

    const produtoMaiorFaturamento = Object.entries(faturamentoProdutos)
        .sort((a, b) => b[1] - a[1])[0];

    const freteTotal = dados.reduce(
        (soma, item) => soma + item.frete, 0
    );

    document.getElementById("topProduct").textContent =
        produtoMaisVendido ? produtoMaisVendido[0] : "--";

    document.getElementById("topState").textContent =
        estadoMaisVendido ? estadoMaisVendido[0] : "--";

    document.getElementById("totalShipping").textContent =
        moedaDashboard(freteTotal);

    document.getElementById("topRevenueProduct").textContent =
        produtoMaiorFaturamento ? produtoMaiorFaturamento[0] : "--";
}

// ==========================================
// ÚLTIMOS PEDIDOS
// ==========================================

function atualizarUltimosPedidos(dados) {
    const tabela = document.getElementById("recentOrdersBody");

    if (!tabela) {
        return;
    }

    const ultimos = dados.slice(-5).reverse();

    tabela.innerHTML = "";

    if (!ultimos.length) {
        tabela.innerHTML = `
            <tr>
                <td colspan="5">Nenhum pedido encontrado.</td>
            </tr>
        `;
        return;
    }

    ultimos.forEach(item => {
        const linha = document.createElement("tr");

        const valores = [
            item.pedidoId || "--",
            item.produto,
            item.estado,
            item.quantidade.toLocaleString("pt-BR"),
            moedaDashboard(item.total)
        ];

        valores.forEach(valor => {
            const celula = document.createElement("td");
            celula.textContent = valor;
            linha.appendChild(celula);
        });

        tabela.appendChild(linha);
    });
}

}

document.addEventListener("DOMContentLoaded", iniciarDashboard);