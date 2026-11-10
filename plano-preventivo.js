const CONFIG = {
    arquivo: window.PLANO_PREVENTIVO_CONFIG?.arquivo || "/static/dados.xlsx",
    aba: "Acomp_Modernização"
};

let registros = [];
let graficoEvolucao = null;
let dataAgendaSelecionada = "";

function normalizarTexto(valor) {
    return String(valor ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

function escaparHtml(valor) {
    return String(valor ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function dataValida(data) {
    return data instanceof Date && !Number.isNaN(data.getTime());
}

function criarDataSegura(ano, mes, dia) {
    const data = new Date(ano, mes - 1, dia);
    return data.getFullYear() === ano && data.getMonth() === mes - 1 && data.getDate() === dia ? data : null;
}

function dataExcel(valor) {
    if (dataValida(valor)) return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());

    if (typeof valor === "number" && Number.isFinite(valor)) {
        const partes = XLSX.SSF.parse_date_code(valor);
        return partes ? criarDataSegura(partes.y, partes.m, partes.d) : null;
    }

    const texto = String(valor ?? "").trim();
    if (!texto) return null;

    let match = texto.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s.*)?$/);
    if (match) return criarDataSegura(Number(match[3]), Number(match[2]), Number(match[1]));

    match = texto.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
    if (match) return criarDataSegura(Number(match[1]), Number(match[2]), Number(match[3]));

    return null;
}

function dataIso(data) {
    if (!dataValida(data)) return "";
    return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, "0")}-${String(data.getDate()).padStart(2, "0")}`;
}

function dataBr(data) {
    return dataValida(data) ? data.toLocaleDateString("pt-BR") : "--";
}

function mesTitulo(ano, mes) {
    return new Date(ano, mes - 1, 1).toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
}

function obterCampo(linha, nomes) {
    const mapa = new Map(Object.entries(linha).map(([chave, valor]) => [normalizarTexto(chave), valor]));
    for (const nome of nomes) {
        const chave = normalizarTexto(nome);
        if (mapa.has(chave)) return mapa.get(chave);
    }
    return "";
}

function lerRegistros(aba) {
    const linhas = XLSX.utils.sheet_to_json(aba, { defval: "", raw: true });

    return linhas.map(linha => {
        const valorFibras = obterCampo(linha, ["FIBRAS", "FIBRA", "QTDE FIBRAS", "QUANTIDADE DE FIBRAS"]);
        let fibras = null;
        if (valorFibras !== "" && valorFibras != null && String(valorFibras).trim() !== "-") {
            const numero = Number(String(valorFibras).replace(/\./g, "").replace(",", "."));
            fibras = Number.isFinite(numero) ? numero : null;
        }

        return {
            at: String(obterCampo(linha, ["AT", "AREA TECNICA", "ÁREA TÉCNICA"])).trim(),
            dataCadastro: dataExcel(obterCampo(linha, ["DATA CADASTRO", "DATA DE CADASTRO"])),
            dataInicio: dataExcel(obterCampo(linha, ["DATA INICIO", "DATA INÍCIO", "DATA DE INICIO", "DATA DE INÍCIO"])),
            dataFinal: dataExcel(obterCampo(linha, ["DATA FINAL", "DATA FIM", "DATA DE FINALIZACAO", "DATA DE FINALIZAÇÃO", "DATA CONCLUSAO", "DATA CONCLUSÃO"])),
            endereco: String(obterCampo(linha, ["ENDERECO", "ENDEREÇO", "LOGRADOURO"])).trim(),
            numero: String(obterCampo(linha, ["NUMERO", "NÚMERO", "Nº", "N°"])).trim(),
            tipo: String(obterCampo(linha, ["TIPO", "TIPO DE SERVICO", "TIPO DE SERVIÇO"])).trim(),
            status: String(obterCampo(linha, ["STATUS", "SITUACAO", "SITUAÇÃO"])).trim(),
            fibras
        };
    }).filter(item => item.at || item.dataCadastro || item.dataInicio || item.dataFinal || item.status);
}

function statusTipo(item) {
    const status = normalizarTexto(item.status);
    if (["AGENDADO", "AGENDADA"].includes(status)) return "agendado";
    if (["EXECUCAO", "EM EXECUCAO"].includes(status)) return "execucao";
    if (["SERVICO OK", "CONCLUIDO", "CONCLUIDA", "FINALIZADO", "FINALIZADA"].includes(status)) return "ok";
    if (["CANCELADO", "CANCELADA"].includes(status)) return "cancelado";
    return "";
}

function registrosPorAt(lista = registros) {
    const at = document.getElementById("filtroAt").value;
    return at === "TODAS" ? lista : lista.filter(item => normalizarTexto(item.at) === normalizarTexto(at));
}

function somenteConcluidos(lista = registros) {
    return lista.filter(item => statusTipo(item) === "ok" && item.dataFinal);
}

function somenteAgendados(lista = registros) {
    return lista.filter(item => statusTipo(item) === "agendado" && (item.dataInicio || item.dataFinal));
}

function preencherFiltros() {
    const datas = registros.flatMap(item => [item.dataFinal, item.dataInicio]).filter(dataValida);
    const anos = [...new Set(datas.map(data => data.getFullYear()))].sort((a, b) => b - a);
    const filtroAno = document.getElementById("filtroAno");
    filtroAno.innerHTML = "";
    if (!anos.length) anos.push(new Date().getFullYear());
    anos.forEach(ano => filtroAno.add(new Option(String(ano), String(ano))));
    const anoAtual = new Date().getFullYear();
    filtroAno.value = String(anos.includes(anoAtual) ? anoAtual : anos[0]);

    const filtroAt = document.getElementById("filtroAt");
    filtroAt.innerHTML = '<option value="TODAS">Todas</option>';
    [...new Set(registros.map(item => item.at).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "pt-BR"))
        .forEach(at => filtroAt.add(new Option(at, at)));

    const referencias = registros
        .flatMap(item => [item.dataInicio, item.dataFinal])
        .filter(dataValida)
        .sort((a, b) => b - a);
    const referencia = referencias[0] || new Date();
    document.getElementById("filtroMesAgenda").value = `${referencia.getFullYear()}-${String(referencia.getMonth() + 1).padStart(2, "0")}`;
}

function atualizarResumo() {
    const ano = Number(document.getElementById("filtroAno").value);
    const filtrados = registrosPorAt();
    const concluidos = somenteConcluidos(filtrados);
    const doAno = concluidos.filter(item => item.dataFinal.getFullYear() === ano);
    const porMes = Array.from({ length: 12 }, (_, mes) => doAno.filter(item => item.dataFinal.getMonth() === mes).length);
    const ultimaConclusaoNoAno = doAno.map(item => item.dataFinal).sort((a, b) => b - a)[0];
    const referenciaMes = ultimaConclusaoNoAno || new Date(ano, new Date().getMonth(), 1);
    const totalMesAtual = doAno.filter(item => item.dataFinal.getMonth() === referenciaMes.getMonth()).length;
    const agendaConsolidada = filtrados.filter(item => ["agendado", "execucao"].includes(statusTipo(item)) && (item.dataInicio || item.dataFinal));

    document.getElementById("totalAno").textContent = String(doAno.length);
    document.getElementById("totalMesAtual").textContent = String(totalMesAtual);
    document.getElementById("legendaMesAtual").textContent = referenciaMes.toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
    document.getElementById("totalAgendasConsolidadas").textContent = String(agendaConsolidada.length);
    montarGrafico(ano, porMes);
}

function montarGrafico(ano, valores) {
    graficoEvolucao?.destroy();
    const rotulosValores = {
        id: "rotulosValores",
        afterDatasetsDraw(chart) {
            const { ctx } = chart;
            const meta = chart.getDatasetMeta(0);
            ctx.save();
            ctx.fillStyle = "#4b126f";
            ctx.font = '800 12px "Segoe UI", Arial, sans-serif';
            ctx.textAlign = "center";
            ctx.textBaseline = "bottom";
            meta.data.forEach((ponto, indice) => {
                const valor = Number(chart.data.datasets[0].data[indice]) || 0;
                ctx.fillText(String(valor), ponto.x, Math.max(chart.chartArea.top + 15, ponto.y - 11));
            });
            ctx.restore();
        }
    };

    graficoEvolucao = new Chart(document.getElementById("graficoEvolucao"), {
        type: "line",
        plugins: [rotulosValores],
        data: {
            labels: ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"],
            datasets: [{ label: `Concluídos em ${ano}`, data: valores, borderColor: "#5b1fa6", backgroundColor: "#5b1fa622", borderWidth: 3, pointRadius: 5, pointHoverRadius: 7, pointBackgroundColor: "#7a3fe0", pointBorderColor: "#5b1fa6", pointBorderWidth: 2, tension: .3, fill: true }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { intersect: false, mode: "index" },
            layout: { padding: { top: 25, right: 16, bottom: 8, left: 16 } },
            plugins: { legend: { display: false }, tooltip: { displayColors: false, callbacks: { label: contexto => `${contexto.parsed.y} concluído${contexto.parsed.y === 1 ? "" : "s"}` } } },
            scales: { x: { offset: true, grid: { drawBorder: false }, ticks: { padding: 10 } }, y: { beginAtZero: true, suggestedMax: Math.max(4, ...valores) + 3, ticks: { precision: 0, padding: 8 }, grid: { drawBorder: false } } }
        }
    });
}

function obterMesAgenda() {
    const valor = document.getElementById("filtroMesAgenda").value;
    if (!/^\d{4}-\d{2}$/.test(valor)) return { ano: 0, mes: 0 };
    const [ano, mes] = valor.split("-").map(Number);
    return { ano, mes };
}

function statusHabilitado(tipo) {
    const mapa = { agendado: "exibirAgendado", execucao: "exibirExecucao", ok: "exibirServicoOk" };
    return Boolean(mapa[tipo] && document.getElementById(mapa[tipo]).checked);
}

function intervaloItem(item) {
    const fim = item.dataFinal || item.dataInicio;
    const inicio = item.dataInicio || fim;
    return { inicio, fim };
}

function itensDoCalendario(ano, mes) {
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);
    return registrosPorAt().filter(item => {
        const tipo = statusTipo(item);
        if (!tipo || !statusHabilitado(tipo)) return false;
        const { inicio, fim } = intervaloItem(item);
        if (!inicio || !fim) return false;
        if (tipo === "ok") return fim >= inicioMes && fim <= fimMes;
        return inicio <= fimMes && fim >= inicioMes;
    });
}

function atualizarAgenda() {
    const { ano, mes } = obterMesAgenda();
    if (!ano || !mes) return;

    const filtrados = registrosPorAt();
    const agendados = somenteAgendados(filtrados);
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);

    const agendadosMes = agendados.filter(item => {
        const { inicio, fim } = intervaloItem(item);
        return inicio && fim && inicio <= fimMes && fim >= inicioMes;
    });

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const agendaHoje = filtrados.filter(item => {
        const tipo = statusTipo(item);
        if (!["agendado", "execucao"].includes(tipo)) return false;
        const { inicio, fim } = intervaloItem(item);
        return inicio && fim && inicio <= hoje && fim >= hoje;
    });

    const totalExecucao = filtrados.filter(
        item => statusTipo(item) === "execucao"
    ).length;

    const canceladosHoje = filtrados.filter(item =>
        statusTipo(item) === "cancelado" &&
        item.dataFinal &&
        dataIso(item.dataFinal) === dataIso(hoje)
    );

    document.getElementById("totalAgendaMes").textContent = String(agendadosMes.length);
    document.getElementById("legendaAgendaMes").textContent = mesTitulo(ano, mes);
    document.getElementById("totalAgendaHoje").textContent = String(agendaHoje.length);
    document.getElementById("legendaAgendaHoje").textContent = dataBr(hoje);
    document.getElementById("totalExecucao").textContent = String(totalExecucao);
    document.getElementById("totalCanceladosHoje").textContent = String(canceladosHoje.length);
    document.getElementById("legendaCanceladosHoje").textContent = dataBr(hoje);

    montarCalendario(ano, mes, itensDoCalendario(ano, mes));
}

function adicionarAoMapa(mapa, chave, item) {
    if (!mapa.has(chave)) mapa.set(chave, []);
    mapa.get(chave).push(item);
}

function calcularPascoa(ano) {
    const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100;
    const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const mes = Math.floor((h + l - 7 * m + 114) / 31);
    const dia = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(ano, mes - 1, dia);
}

function obterFeriados(ano) {
    const pascoa = calcularPascoa(ano);
    const paixao = new Date(pascoa); paixao.setDate(pascoa.getDate() - 2);
    const lista = [
        [new Date(ano, 0, 1), "Confraternização Universal", "Nacional"], [paixao, "Paixão de Cristo", "Nacional"],
        [new Date(ano, 3, 21), "Tiradentes", "Nacional"], [new Date(ano, 4, 1), "Dia Mundial do Trabalho", "Nacional"],
        [new Date(ano, 6, 9), "Revolução Constitucionalista de 1932", "Estadual - São Paulo"],
        [new Date(ano, 8, 7), "Independência do Brasil", "Nacional"], [new Date(ano, 9, 12), "Nossa Senhora Aparecida", "Nacional"],
        [new Date(ano, 10, 2), "Finados", "Nacional"], [new Date(ano, 10, 15), "Proclamação da República", "Nacional"],
        [new Date(ano, 10, 20), "Dia Nacional de Zumbi e da Consciência Negra", "Nacional"], [new Date(ano, 11, 25), "Natal", "Nacional"]
    ];
    return new Map(lista.map(([data, nome, abrangencia]) => [dataIso(data), { data, nome, abrangencia }]));
}

function montarCalendario(ano, mes, itensMes) {
    document.getElementById("tituloCalendario").textContent = mesTitulo(ano, mes);
    const mapa = new Map();
    const feriados = obterFeriados(ano);
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);

    itensMes.forEach(item => {
        const tipo = statusTipo(item);
        const { inicio, fim } = intervaloItem(item);
        if (!inicio || !fim) return;
        if (tipo === "ok") {
            adicionarAoMapa(mapa, dataIso(fim), item);
            return;
        }
        const inicioVisivel = inicio < inicioMes ? new Date(inicioMes) : new Date(inicio);
        const fimVisivel = fim > fimMes ? new Date(fimMes) : new Date(fim);
        for (const cursor = new Date(inicioVisivel); cursor <= fimVisivel; cursor.setDate(cursor.getDate() + 1)) {
            adicionarAoMapa(mapa, dataIso(cursor), item);
        }
    });

    const primeiro = new Date(ano, mes - 1, 1);
    const inicioGrade = new Date(ano, mes - 1, 1 - primeiro.getDay());
    const hoje = dataIso(new Date());
    const celulas = [];

    for (let i = 0; i < 42; i++) {
        const data = new Date(inicioGrade);
        data.setDate(inicioGrade.getDate() + i);
        const chave = dataIso(data);
        const itens = mapa.get(chave) || [];
        const feriado = feriados.get(chave) || null;
        const fora = data.getMonth() !== mes - 1;
        const contagens = { agendado: 0, execucao: 0, ok: 0 };
        itens.forEach(item => { const tipo = statusTipo(item); if (tipo) contagens[tipo]++; });
        const ats = [...new Set(itens.map(item => item.at).filter(Boolean))].join(" · ");
        const badges = [
            feriado && !fora ? '<span class="badge-feriado">Feriado</span>' : "",
            contagens.agendado ? `<span class="badge-agenda badge-agendado">${contagens.agendado} agendado${contagens.agendado > 1 ? "s" : ""}</span>` : "",
            contagens.execucao ? `<span class="badge-agenda badge-execucao">${contagens.execucao} em execução</span>` : "",
            contagens.ok ? `<span class="badge-agenda badge-ok">${contagens.ok} serviço${contagens.ok > 1 ? "s" : ""} OK</span>` : ""
        ].join("");
        const interativo = itens.length > 0 || Boolean(feriado && !fora);
        celulas.push(`<button type="button" class="dia-calendario ${fora ? "fora-mes" : ""} ${itens.length ? "com-agenda" : ""} ${feriado && !fora ? "dia-feriado" : ""} ${chave === dataAgendaSelecionada ? "selecionado" : ""} ${chave === hoje ? "hoje" : ""}" data-data="${chave}" data-interativo="${interativo}" ${interativo ? "" : "disabled"}><span class="numero-dia">${data.getDate()}</span>${badges}${itens.length ? `<span class="ats-dia">${escaparHtml(ats)}</span>` : ""}</button>`);
    }

    document.getElementById("calendario").innerHTML = celulas.join("");
    document.querySelectorAll('.dia-calendario[data-interativo="true"]').forEach(botao => {
        botao.addEventListener("click", () => selecionarDia(botao.dataset.data, mapa.get(botao.dataset.data) || [], feriados.get(botao.dataset.data) || null));
    });

    const feriadoSelecionado = feriados.get(dataAgendaSelecionada) || null;
    const haFeriadoNoMes = [...feriados.values()].some(item => item.data.getFullYear() === ano && item.data.getMonth() === mes - 1);
    if (!itensMes.length && !haFeriadoNoMes) {
        document.getElementById("tituloDetalheAgenda").textContent = "Nenhum caso";
        document.getElementById("listaAgenda").innerHTML = `<p class="estado-vazio">Nenhum caso dos status selecionados em ${escaparHtml(mesTitulo(ano, mes))}.</p>`;
    } else if (!mapa.has(dataAgendaSelecionada) && !feriadoSelecionado) {
        document.getElementById("tituloDetalheAgenda").textContent = "Selecione um dia";
        document.getElementById("listaAgenda").innerHTML = '<p class="estado-vazio">Clique em uma data com ocorrência ou feriado para visualizar os detalhes.</p>';
    }
}

function selecionarDia(chave, itens, feriado = null) {
    dataAgendaSelecionada = chave;

    const rotulosStatus = {
        agendado: "AGENDADO",
        execucao: "EM EXECUÇÃO",
        ok: "SERVIÇO OK",
        cancelado: "CANCELADO"
    };

    const statusPresentes = [...new Set(
        itens
            .map(item => statusTipo(item))
            .filter(tipo => rotulosStatus[tipo])
    )];

    let tituloStatus = "DETALHAMENTO";
    if (statusPresentes.length === 1) {
        tituloStatus = rotulosStatus[statusPresentes[0]];
    } else if (statusPresentes.length > 1) {
        tituloStatus = "MÚLTIPLOS STATUS";
    } else if (feriado) {
        tituloStatus = "FERIADO";
    }

    document.getElementById("tituloDetalheAgenda").textContent = tituloStatus;

    const ordem = { execucao: 1, agendado: 2, ok: 3, cancelado: 4 };
    const detalheFeriado = feriado
        ? `<article class="item-feriado"><span class="item-feriado-rotulo">Feriado</span><strong>${escaparHtml(feriado.nome)}</strong><span>${escaparHtml(feriado.abrangencia)}</span></article>`
        : "";

    const detalhes = [...itens]
        .sort((a, b) =>
            (ordem[statusTipo(a)] || 9) - (ordem[statusTipo(b)] || 9) ||
            a.at.localeCompare(b.at, "pt-BR")
        )
        .map(item => {
            const tipoStatus = statusTipo(item);
            const rotulo = rotulosStatus[tipoStatus] || normalizarTexto(item.status) || "NÃO INFORMADO";
            return `<article class="item-agenda item-${tipoStatus}"><strong>${escaparHtml(item.at || "AT não informada")} · ${escaparHtml(item.tipo || "Tipo não informado")}</strong><span>Início: ${dataBr(item.dataInicio)} | Final: ${dataBr(item.dataFinal)}</span><span class="item-endereco"><strong>Endereço:</strong> ${escaparHtml([item.endereco, item.numero].filter(Boolean).join(", ") || "Não informado")}</span><span class="status-detalhe">Status: ${rotulo}</span><span class="item-fibras">Fibras: <strong>${item.fibras ?? "Não informado"}</strong></span></article>`;
        }).join("");

    document.getElementById("listaAgenda").innerHTML =
        detalheFeriado +
        (detalhes || '<p class="estado-vazio">Nenhuma ocorrência nesta data.</p>');

    atualizarAgenda();
}

function mudarMesAgenda(deslocamento) {
    const { ano, mes } = obterMesAgenda();
    if (!ano || !mes) return;
    const nova = new Date(ano, mes - 1 + deslocamento, 1);
    document.getElementById("filtroMesAgenda").value = `${nova.getFullYear()}-${String(nova.getMonth() + 1).padStart(2, "0")}`;
    dataAgendaSelecionada = "";
    atualizarAgenda();
}

function atualizarTudo() {
    atualizarResumo();
    dataAgendaSelecionada = "";
    atualizarAgenda();
}

function mostrarErro(mensagem) {
    const elemento = document.getElementById("mensagemErro");
    elemento.hidden = false;
    elemento.textContent = mensagem;
}

async function iniciar() {
    try {
        if (typeof XLSX === "undefined") throw new Error("A biblioteca XLSX não foi carregada.");
        if (typeof Chart === "undefined") throw new Error("A biblioteca Chart.js não foi carregada.");
        const resposta = await fetch(CONFIG.arquivo, { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status} ao carregar ${CONFIG.arquivo}.`);
        const ultimaAtualizacao = resposta.headers.get("Last-Modified");
        document.getElementById("dataAtualizacao").textContent = ultimaAtualizacao ? new Date(ultimaAtualizacao).toLocaleString("pt-BR") : new Date().toLocaleString("pt-BR");
        const workbook = XLSX.read(await resposta.arrayBuffer(), { type: "array", cellDates: true });
        const aba = workbook.Sheets[CONFIG.aba];
        if (!aba) throw new Error(`A aba "${CONFIG.aba}" não foi encontrada. Abas disponíveis: ${workbook.SheetNames.join(", ")}.`);
        registros = lerRegistros(aba);
        if (!registros.length) throw new Error(`A aba "${CONFIG.aba}" foi encontrada, mas nenhum registro válido foi lido. Verifique os cabeçalhos e a linha inicial da tabela.`);
        console.info("Plano Preventivo", { registros: registros.length, status: [...new Set(registros.map(item => item.status))], ats: [...new Set(registros.map(item => item.at))] });
        preencherFiltros();
        atualizarTudo();
    } catch (erro) {
        console.error(erro);
        mostrarErro(`Não foi possível carregar o painel: ${erro.message}`);
        document.getElementById("dataAtualizacao").textContent = "Falha no carregamento";
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("filtroAno").addEventListener("change", atualizarResumo);
    document.getElementById("filtroAt").addEventListener("change", atualizarTudo);
    document.getElementById("filtroMesAgenda").addEventListener("change", () => { dataAgendaSelecionada = ""; atualizarAgenda(); });
    document.getElementById("mesAnterior").addEventListener("click", () => mudarMesAgenda(-1));
    document.getElementById("mesSeguinte").addEventListener("click", () => mudarMesAgenda(1));
    ["exibirAgendado", "exibirExecucao", "exibirServicoOk"].forEach(id => document.getElementById(id).addEventListener("change", () => { dataAgendaSelecionada = ""; atualizarAgenda(); }));
    document.getElementById("btnTelaCheia").addEventListener("click", async () => {
        try {
            if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
            else await document.exitFullscreen();
        } catch (erro) {
            mostrarErro(`Não foi possível alternar a tela cheia: ${erro.message}`);
        }
    });
    iniciar();
});
