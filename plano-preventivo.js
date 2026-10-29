const CONFIG = {
    arquivo: "dados.xlsx",
    aba: "Acomp_Modernização"
};

let registros = [];
let graficoEvolucao = null;
let dataAgendaSelecionada = "";

function normalizarTexto(valor) {
    return String(valor ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().replace(/\s+/g, " ").toUpperCase();
}

function escaparHtml(valor) {
    return String(valor ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}

function dataExcel(valor) {
    if (valor instanceof Date && !Number.isNaN(valor.getTime())) return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
    if (typeof valor === "number" && Number.isFinite(valor)) {
        const partes = XLSX.SSF.parse_date_code(valor);
        return partes ? new Date(partes.y, partes.m - 1, partes.d) : null;
    }
    const texto = String(valor ?? "").trim();
    let match = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    match = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return null;
}

function dataIso(data) {
    if (!data) return "";
    return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, "0")}-${String(data.getDate()).padStart(2, "0")}`;
}

function dataBr(data) {
    return data ? data.toLocaleDateString("pt-BR") : "--";
}

function mesTitulo(ano, mes) {
    return new Date(ano, mes - 1, 1).toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
}

function lerRegistros(aba) {
    return XLSX.utils.sheet_to_json(aba, { defval: "", raw: true }).map(linha => ({
        at: String(linha.AT ?? "").trim(),
        dataCadastro: dataExcel(linha["DATA CADASTRO"]),
        dataInicio: dataExcel(linha["DATA INICIO"]),
        dataFinal: dataExcel(linha["DATA FINAL"]),
        tipo: String(linha.TIPO ?? "").trim(),
        status: String(linha.STATUS ?? "").trim(),
        fibras: (() => {
            const valor = linha.Fibras ?? linha.FIBRAS ?? linha.fibras ?? "";
            if (valor === "" || valor === null || valor === undefined || String(valor).trim() === "-") return null;
            const numero = Number(String(valor).replace(",", "."));
            return Number.isFinite(numero) ? numero : null;
        })()
    })).filter(item => item.at || item.dataFinal || item.status);
}

function registrosPorAt(lista = registros) {
    const at = document.getElementById("filtroAt").value;
    return at === "TODAS" ? lista : lista.filter(item => normalizarTexto(item.at) === normalizarTexto(at));
}

function somenteConcluidos(lista = registros) {
    return lista.filter(item => normalizarTexto(item.status) === "SERVICO OK" && item.dataFinal);
}

function somenteAgendados(lista = registros) {
    return lista.filter(item => normalizarTexto(item.status) === "AGENDADO" && item.dataFinal);
}

function preencherFiltros() {
    const datas = registros.map(item => item.dataFinal).filter(Boolean);
    const anos = [...new Set(datas.map(data => data.getFullYear()))].sort((a, b) => b - a);
    const filtroAno = document.getElementById("filtroAno");
    anos.forEach(ano => filtroAno.add(new Option(ano, ano)));
    const anoAtual = new Date().getFullYear();
    filtroAno.value = anos.includes(anoAtual) ? anoAtual : anos[0] ?? anoAtual;

    const ats = [...new Set(registros.map(item => item.at).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
    const filtroAt = document.getElementById("filtroAt");
    ats.forEach(at => filtroAt.add(new Option(at, at)));

    const ultimaData = datas.sort((a, b) => b - a)[0] || new Date();
    document.getElementById("filtroData").value = dataIso(ultimaData);
    const hoje = new Date();
    document.getElementById("filtroMesAgenda").value = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, "0")}`;
}

function atualizarResumo() {
    const ano = Number(document.getElementById("filtroAno").value);
    const dataEscolhida = dataExcel(document.getElementById("filtroData").value);
    const concluidos = somenteConcluidos(registrosPorAt());
    const agendados = somenteAgendados(registrosPorAt());
    const doAno = concluidos.filter(item => item.dataFinal.getFullYear() === ano);
    const porMes = Array.from({ length: 12 }, (_, mes) =>
        doAno.filter(item => item.dataFinal.getMonth() === mes).length
    );

    const referenciaMes = dataEscolhida || new Date();
    const totalMesAtual = concluidos.filter(item =>
        item.dataFinal.getFullYear() === referenciaMes.getFullYear() &&
        item.dataFinal.getMonth() === referenciaMes.getMonth()
    ).length;

    document.getElementById("totalAno").textContent = doAno.length;
    document.getElementById("legendaAno").textContent = ano;
    document.getElementById("totalMesAtual").textContent = totalMesAtual;
    document.getElementById("legendaMesAtual").textContent = referenciaMes.toLocaleDateString("pt-BR", {
        month: "long",
        year: "numeric"
    });
    document.getElementById("totalAgendasConsolidadas").textContent = agendados.length;
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
                const y = Math.max(chart.chartArea.top + 15, ponto.y - 11);
                ctx.fillText(String(valor), ponto.x, y);
            });
            ctx.restore();
        }
    };

    graficoEvolucao = new Chart(document.getElementById("graficoEvolucao"), {
        type: "line",
        plugins: [rotulosValores],
        data: {
            labels: ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"],
            datasets: [{
                label: `Concluídos em ${ano}`,
                data: valores,
                borderColor: "#5b1fa6",
                backgroundColor: "#5b1fa622",
                borderWidth: 3,
                pointRadius: 5,
                pointHoverRadius: 7,
                pointBackgroundColor: "#7a3fe0",
                pointBorderColor: "#5b1fa6",
                pointBorderWidth: 2,
                tension: .3,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { intersect: false, mode: "index" },
            layout: {
                padding: { top: 25, right: 16, bottom: 8, left: 16 }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    displayColors: false,
                    callbacks: {
                        label: contexto => `${contexto.parsed.y} concluído${contexto.parsed.y === 1 ? "" : "s"}`
                    }
                }
            },
            scales: {
                x: {
                    offset: true,
                    grid: { drawBorder: false },
                    ticks: { padding: 10 }
                },
                y: {
                    beginAtZero: true,
                    suggestedMax: Math.max(4, ...valores) + 3,
                    ticks: { precision: 0, padding: 8 },
                    grid: { drawBorder: false }
                }
            }
        }
    });
}

function obterMesAgenda() {
    const [ano, mes] = document.getElementById("filtroMesAgenda").value.split("-").map(Number);
    return { ano, mes };
}

function statusTipo(item) {
    const status = normalizarTexto(item.status);
    if (status === "AGENDADO") return "agendado";
    if (status === "EXECUCAO") return "execucao";
    if (status === "SERVICO OK") return "ok";
    return "";
}

function statusHabilitado(tipo) {
    const mapa = {
        agendado: "exibirAgendado",
        execucao: "exibirExecucao",
        ok: "exibirServicoOk"
    };
    return mapa[tipo] && document.getElementById(mapa[tipo]).checked;
}

function itensDoCalendario(ano, mes) {
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);
    return registrosPorAt().filter(item => {
        const tipo = statusTipo(item);
        if (!tipo || !statusHabilitado(tipo) || !item.dataFinal) return false;
        if (tipo === "ok") {
            return item.dataFinal >= inicioMes && item.dataFinal <= fimMes;
        }
        const inicio = item.dataInicio || item.dataFinal;
        return inicio <= fimMes && item.dataFinal >= inicioMes;
    });
}

function atualizarAgenda() {
    const { ano, mes } = obterMesAgenda();
    if (!ano || !mes) return;

    const agendados = somenteAgendados(registrosPorAt());
    const inicioSeguinte = new Date(ano, mes, 1);
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);
    const agendadosMes = agendados.filter(item => {
        const inicio = item.dataInicio || item.dataFinal;
        return inicio <= fimMes && item.dataFinal >= inicioMes;
    });
    const futuros = agendados.filter(item => item.dataFinal >= inicioSeguinte);

    document.getElementById("totalAgendaMes").textContent = agendadosMes.length;
    document.getElementById("legendaAgendaMes").textContent = mesTitulo(ano, mes);
    document.getElementById("totalAgendaFutura").textContent = futuros.length;
    document.getElementById("legendaAgendaFutura").textContent = `Após ${dataBr(fimMes)}`;
    montarCalendario(ano, mes, itensDoCalendario(ano, mes));
}

function adicionarAoMapa(mapa, chave, item) {
    if (!mapa.has(chave)) mapa.set(chave, []);
    mapa.get(chave).push(item);
}

function montarCalendario(ano, mes, itensMes) {
    document.getElementById("tituloCalendario").textContent = mesTitulo(ano, mes);
    const mapa = new Map();
    const inicioMes = new Date(ano, mes - 1, 1);
    const fimMes = new Date(ano, mes, 0);

    itensMes.forEach(item => {
        const tipo = statusTipo(item);
        if (tipo === "ok") {
            adicionarAoMapa(mapa, dataIso(item.dataFinal), item);
            return;
        }
        const inicioItem = item.dataInicio || item.dataFinal;
        const inicioVisivel = inicioItem < inicioMes ? new Date(inicioMes) : new Date(inicioItem);
        const fimVisivel = item.dataFinal > fimMes ? new Date(fimMes) : new Date(item.dataFinal);
        const cursor = new Date(inicioVisivel);
        while (cursor <= fimVisivel) {
            adicionarAoMapa(mapa, dataIso(cursor), item);
            cursor.setDate(cursor.getDate() + 1);
        }
    });

    const primeiro = new Date(ano, mes - 1, 1);
    const inicio = new Date(ano, mes - 1, 1 - primeiro.getDay());
    const hoje = dataIso(new Date());
    const celulas = [];

    for (let i = 0; i < 42; i++) {
        const data = new Date(inicio);
        data.setDate(inicio.getDate() + i);
        const chave = dataIso(data);
        const itens = mapa.get(chave) || [];
        const fora = data.getMonth() !== mes - 1;
        const contagens = { agendado: 0, execucao: 0, ok: 0 };
        itens.forEach(item => contagens[statusTipo(item)]++);
        const ats = [...new Set(itens.map(item => item.at))].join(" · ");
        const badges = [
            contagens.agendado ? `<span class="badge-agenda badge-agendado">${contagens.agendado} agendado${contagens.agendado > 1 ? "s" : ""}</span>` : "",
            contagens.execucao ? `<span class="badge-agenda badge-execucao">${contagens.execucao} em execução</span>` : "",
            contagens.ok ? `<span class="badge-agenda badge-ok">${contagens.ok} serviço${contagens.ok > 1 ? "s" : ""} OK</span>` : ""
        ].join("");

        celulas.push(`<button type="button" class="dia-calendario ${fora ? "fora-mes" : ""} ${itens.length ? "com-agenda" : ""} ${chave === dataAgendaSelecionada ? "selecionado" : ""} ${chave === hoje ? "hoje" : ""}" data-data="${chave}" ${itens.length ? "" : "disabled"}>
            <span class="numero-dia">${data.getDate()}</span>${badges}
            ${itens.length ? `<span class="ats-dia">${escaparHtml(ats)}</span>` : ""}
        </button>`);
    }

    document.getElementById("calendario").innerHTML = celulas.join("");
    document.querySelectorAll(".dia-calendario.com-agenda").forEach(botao =>
        botao.addEventListener("click", () => selecionarDia(botao.dataset.data, mapa.get(botao.dataset.data) || []))
    );

    if (!itensMes.length) {
        document.getElementById("tituloDetalheAgenda").textContent = "Nenhum caso";
        document.getElementById("listaAgenda").innerHTML = `<p class="estado-vazio">Nenhum caso dos status selecionados em ${escaparHtml(mesTitulo(ano, mes))}.</p>`;
    } else if (!mapa.has(dataAgendaSelecionada)) {
        document.getElementById("tituloDetalheAgenda").textContent = "Selecione um dia";
        document.getElementById("listaAgenda").innerHTML = '<p class="estado-vazio">Clique em uma data com ocorrências para visualizar os casos.</p>';
    }
}

function selecionarDia(chave, itens) {
    dataAgendaSelecionada = chave;
    document.getElementById("tituloDetalheAgenda").textContent = `Ocorrências de ${dataBr(dataExcel(chave))}`;
    const ordem = { execucao: 1, agendado: 2, ok: 3 };
    document.getElementById("listaAgenda").innerHTML = itens
        .sort((a, b) => ordem[statusTipo(a)] - ordem[statusTipo(b)] || a.at.localeCompare(b.at))
        .map(item => {
            const tipoStatus = statusTipo(item);
            const rotulo = tipoStatus === "execucao" ? "EXECUÇÃO" : tipoStatus === "ok" ? "SERVIÇO OK" : "AGENDADO";
            return `<article class="item-agenda item-${tipoStatus}">
                <strong>${escaparHtml(item.at || "AT não informada")} · ${escaparHtml(item.tipo || "Tipo não informado")}</strong>
                <span>Início: ${dataBr(item.dataInicio)} | Final: ${dataBr(item.dataFinal)}</span>
                <span class="status-detalhe">Status: ${rotulo}</span>
                <span class="item-fibras">Fibras: <strong>${item.fibras ?? "Não informado"}</strong></span>
            </article>`;
        }).join("");
    atualizarAgenda();
}

function mudarMesAgenda(deslocamento) {
    const { ano, mes } = obterMesAgenda();
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
        const resposta = await fetch(CONFIG.arquivo, { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status} ao carregar ${CONFIG.arquivo}.`);
        const ultimaAtualizacao = resposta.headers.get("Last-Modified");
        document.getElementById("dataAtualizacao").textContent = ultimaAtualizacao ? new Date(ultimaAtualizacao).toLocaleString("pt-BR") : new Date().toLocaleString("pt-BR");
        const workbook = XLSX.read(await resposta.arrayBuffer(), { type: "array", cellDates: true });
        const aba = workbook.Sheets[CONFIG.aba];
        if (!aba) throw new Error(`A aba ${CONFIG.aba} não foi encontrada.`);
        registros = lerRegistros(aba);
        preencherFiltros();
        atualizarTudo();
    } catch (erro) {
        console.error(erro);
        mostrarErro(`Não foi possível carregar o painel. Verifique a aba ${CONFIG.aba} no arquivo ${CONFIG.arquivo}.`);
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("filtroAno").addEventListener("change", atualizarResumo);
    document.getElementById("filtroAt").addEventListener("change", atualizarTudo);
    document.getElementById("filtroData").addEventListener("change", atualizarResumo);
    document.getElementById("filtroMesAgenda").addEventListener("change", () => { dataAgendaSelecionada = ""; atualizarAgenda(); });
    document.getElementById("mesAnterior").addEventListener("click", () => mudarMesAgenda(-1));
    document.getElementById("mesSeguinte").addEventListener("click", () => mudarMesAgenda(1));
    ["exibirAgendado", "exibirExecucao", "exibirServicoOk"].forEach(id =>
        document.getElementById(id).addEventListener("change", () => {
            dataAgendaSelecionada = "";
            atualizarAgenda();
        })
    );
    document.getElementById("btnTelaCheia").addEventListener("click", async () => {
        if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
        else await document.exitFullscreen();
    });
    iniciar();
});
