let notasAtuais = [];
let modal90;
let modalDet;
let chartEmpresa = null;
let chartStatus = null; 
let modalUser;

window.onload = async () => {
    iniciarCanvasBackground();
    modal90 = new bootstrap.Modal(document.getElementById('modal90Dias'));
    modalDet = new bootstrap.Modal(document.getElementById('modalDetalhes'));
    modalUser = new bootstrap.Modal(document.getElementById('modalUsuario'));

    let perfil = await eel.obter_perfil_usuario()();
    if (perfil.is_admin) {
        // CORREÇÃO: Trocamos "block" por "flex" e garantimos o important!
        document.getElementById('area-admin').style.setProperty("display", "flex", "important"); 
    }
    
    // Carrega contas de e-mail
    let contas = await eel.obter_contas_outlook()();
    let combo = document.getElementById('combo-email');
    contas.forEach(c => {
        let opt = document.createElement('option');
        opt.value = c; opt.innerText = c;
        combo.appendChild(opt);
    });

    // Puxa e filtra dados
    await eel.carregar_dados_reais()();
    carregarDadosEFiltrar();

    // Eventos de Filtro
    document.getElementById('filtro-busca').addEventListener('keyup', carregarDadosEFiltrar);
    document.getElementById('filtro-empresa').addEventListener('change', carregarDadosEFiltrar);

    navegarPara('dashboard');
};

// ==========================================
// SISTEMA DE NAVEGAÇÃO DE TELAS (SIDEBAR)
// ==========================================
function navegarPara(tela) {
    const telas = ['tela-dashboard', 'tela-cobrancas', 'tela-admin', 'tela-templates'];
    
    // 1. Esconde tudo
    telas.forEach(t => {
        let el = document.getElementById(t);
        if(el) el.style.setProperty("display", "none", "important");
    });

    // 2. Reseta os botões
    document.querySelectorAll('.nav-main-link').forEach(btn => {
        btn.className = "btn btn-outline-light fw-bold text-white text-start nav-main-link";
    });

    // 3. Mostra a tela certa
    if (tela === 'cobrancas') {
        document.getElementById('tela-cobrancas').style.setProperty("display", "flex", "important");
        document.getElementById('nav-cob').className = "btn btn-light fw-bold text-dark text-start nav-main-link";
    } 
    else if (tela === 'dashboard') {
        document.getElementById('tela-dashboard').style.setProperty("display", "flex", "important");
        document.getElementById('nav-dash').className = "btn btn-info fw-bold text-dark text-start nav-main-link";
        atualizarGraficos(); 
    } 
    else if (tela === 'admin') {
        document.getElementById('tela-admin').style.setProperty("display", "flex", "important");
        document.getElementById('nav-admin').className = "btn btn-warning fw-bold text-dark text-start nav-main-link";
        carregarUsuarios(); 
    }
    else if (tela === 'templates') {
        document.getElementById('tela-templates').style.setProperty("display", "flex", "important");
        document.getElementById('nav-temp').className = "btn btn-success fw-bold text-dark text-start nav-main-link";
        carregarTemplates(); // Função que vamos criar agora
    }

    let sidebarEl = document.getElementById('sidebarMenu');
    let bsOffcanvas = bootstrap.Offcanvas.getInstance(sidebarEl);
    if (bsOffcanvas) bsOffcanvas.hide();
}

async function carregarDadosEFiltrar() {
    let termo = document.getElementById('filtro-busca').value;
    let empresa = document.getElementById('filtro-empresa').value;
    
    let res = await eel.filtrar_dados(termo, empresa)();
    notasAtuais = res.dados;
    
    // 1. Atualiza os textos pequenos na Tela de Cobranças (Tabela)
    document.getElementById('lbl-total-geral').innerText = "Total Geral: " + res.total_geral;
    document.getElementById('lbl-total-atraso').innerText = "Total Atraso (30+): " + res.total_atraso;
    
    // 2. ATUALIZA OS CARDS GIGANTES NO DASHBOARD
    document.getElementById('dash-total-geral').innerText = res.total_geral;
    document.getElementById('dash-total-atraso').innerText = res.total_atraso;
    
    renderizarTabela();
    atualizarGraficos(); 
}

let modoVisaoGrafico = 'valor'; 

function limparFiltros() {
    // Reseta os campos de busca e o combo de empresa
    document.getElementById('filtro-busca').value = "";
    document.getElementById('filtro-empresa').value = "TODAS";
    
    // Recarrega tudo e atualiza os gráficos
    carregarDadosEFiltrar();
}

function alternarModoGrafico() {
    let btn = document.getElementById('btn-toggle-graficos');
    if (modoVisaoGrafico === 'valor') {
        modoVisaoGrafico = 'quantidade';
        btn.innerHTML = "💰 Ver por Valores (R$)";
        btn.className = "btn btn-outline-success fw-bold btn-sm px-4";
        btn.style.boxShadow = "0 0 10px rgba(0, 230, 118, 0.3)";
    } else {
        modoVisaoGrafico = 'valor';
        btn.innerHTML = "📄 Ver por Qtd. de Notas";
        btn.className = "btn btn-outline-info fw-bold btn-sm px-4";
        btn.style.boxShadow = "0 0 10px rgba(0, 229, 255, 0.3)";
    }
    atualizarGraficos(); // Redesenha os gráficos com a nova visão!
}

function atualizarGraficos() {
    let statsEmpresa = {};
    let statsStatus = {};

    notasAtuais.forEach(nota => {
        // Por Empresa
        if(!statsEmpresa[nota.empresa]) statsEmpresa[nota.empresa] = { valor: 0, qtd: 0 };
        statsEmpresa[nota.empresa].valor += nota.valor_num;
        statsEmpresa[nota.empresa].qtd += 1;

        // Por Status
        let statusCurto = nota.parecer.split('(')[0].trim();
        if(!statsStatus[statusCurto]) statsStatus[statusCurto] = { valor: 0, qtd: 0 };
        statsStatus[statusCurto].valor += nota.valor_num;
        statsStatus[statusCurto].qtd += 1;
    });

    let labelsEmpresa = Object.keys(statsEmpresa);
    let dataEmpresa = labelsEmpresa.map(e => modoVisaoGrafico === 'valor' ? statsEmpresa[e].valor : statsEmpresa[e].qtd);
    
    // 🎨 MÁGICA 1: Dicionário Fixo de Cores por Empresa
    const paletaEmpresas = {
        'CANAA': '#00E5FF', // Azul Ciano
        'HAGG': '#FF6600',  // Laranja
        'SN': '#FF1744'     // Vermelho
    };
    // Se aparecer uma empresa nova no futuro, ele pega uma destas como reserva
    let coresReserva = ['#00E676', '#FFD600', '#D500F9']; 
    let bgCoresEmpresa = labelsEmpresa.map((emp, i) => paletaEmpresas[emp] || coresReserva[i % coresReserva.length]);

    // 📊 MÁGICA 2: Ordenando as Barras de Status Logicamente
    const ordemCorreta = [
        "NO PRAZO", 
        "AVISO AMIGÁVEL", 
        "FAIXA VERDE", 
        "FAIXA AMARELA", 
        "FAIXA LARANJA", 
        "FAIXA VERMELHA", 
        "FAIXA ROXA", 
        "FAIXA PRETA"
    ];
    let labelsStatus = Object.keys(statsStatus).sort((a, b) => {
        let posA = ordemCorreta.indexOf(a);
        let posB = ordemCorreta.indexOf(b);
        if (posA === -1) posA = 99; // Joga pro final se não achar
        if (posB === -1) posB = 99;
        return posA - posB;
    });
    let dataStatus = labelsStatus.map(s => modoVisaoGrafico === 'valor' ? statsStatus[s].valor : statsStatus[s].qtd);

    Chart.defaults.color = '#FFFFFF';
    Chart.defaults.font.family = 'Calibri, Arial, sans-serif';
    let formatarBR = (v) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(v);

    // ==========================================
    // GRÁFICO 1: EMPRESA (ROSCA)
    // ==========================================
    let ctxEmpresa = document.getElementById('graficoEmpresa').getContext('2d');
    if(chartEmpresa) chartEmpresa.destroy(); 
    
    chartEmpresa = new Chart(ctxEmpresa, {
        type: 'doughnut',
        data: {
            labels: labelsEmpresa,
            datasets: [{
                data: dataEmpresa,
                backgroundColor: bgCoresEmpresa, // 👈 Aplicando as cores fixas aqui!
                borderColor: '#1e1e2f',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            onClick: (e, elements) => {
                if (elements.length > 0) {
                    let empresaClicada = labelsEmpresa[elements[0].index];
                    document.getElementById('filtro-empresa').value = empresaClicada;
                    carregarDadosEFiltrar(); 
                }
            },
            plugins: {
                legend: { position: 'right' },
                title: { display: true, text: modoVisaoGrafico === 'valor' ? 'Volume Financeiro por Empresa' : 'Quantidade de Notas por Empresa', font: { size: 15 } },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` ${formatarBR(statsEmpresa[ctx.label].valor)} (${statsEmpresa[ctx.label].qtd} notas)`
                    }
                }
            }
        }
    });

    // ==========================================
    // GRÁFICO 2: STATUS (BARRAS)
    // ==========================================
    let ctxStatus = document.getElementById('graficoStatus').getContext('2d');
    if(chartStatus) chartStatus.destroy();

    chartStatus = new Chart(ctxStatus, {
        type: 'bar',
        data: {
            labels: labelsStatus,
            datasets: [{
                label: modoVisaoGrafico === 'valor' ? 'Valor (R$)' : 'Qtd. de Notas',
                data: dataStatus,
                backgroundColor: '#FF6600',
                borderRadius: 4
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            onClick: (e, elements) => {
                if (elements.length > 0) {
                    let statusClicado = labelsStatus[elements[0].index];
                    document.getElementById('filtro-busca').value = statusClicado;
                    carregarDadosEFiltrar();
                }
            },
            plugins: {
                legend: { display: false },
                title: { display: true, text: modoVisaoGrafico === 'valor' ? 'Valores Presos por Faixa' : 'Volume de Notas por Faixa', font: { size: 15 } },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` Total: ${formatarBR(statsStatus[ctx.label].valor)} | Volume: ${statsStatus[ctx.label].qtd} notas`
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.1)' } },
                x: { grid: { display: false } }
            }
        }
    });
}

function getCssClass(parecer) {
    let p = parecer.toUpperCase();
    
    if (p.includes("AVISO AMIGÁVEL")) return "faixa-aviso"; 
    if (p.includes("NO PRAZO")) return "faixa-prazo";
    // Agora o sistema procura apenas pelo nome da Faixa, sem conflito de números!
    if (p.includes("FAIXA VERDE")) {
        return "faixa-verde"; 
    }
    if (p.includes("FAIXA AMARELA")) {
        return "faixa-amarela"; 
    }
    if (p.includes("FAIXA LARANJA")) {
        return "faixa-laranja"; 
    }
    if (p.includes("FAIXA VERMELHA")) {
        return "faixa-vermelha"; 
    }
    if (p.includes("FAIXA ROXA")) {
        return "faixa-roxa"; 
    }
    if (p.includes("FAIXA PRETA")) {
        return "faixa-preta"; 
    }
    return "";
}

function renderizarTabela() {
    let tbody = document.getElementById('corpo-tabela');
    tbody.innerHTML = "";

    notasAtuais.forEach((nota, index) => {
        let tr = document.createElement('tr');
        tr.className = getCssClass(nota.parecer);
        
        // 🎩 MÁGICA VISUAL: Corta a palavra "FAIXA X" e deixa só os dias
        // Ex: "FAIXA ROXA (91 A 120 DIAS)" vira "91 A 120 DIAS"
        let statusExibicao = nota.parecer;
        if (statusExibicao.includes("(")) {
            statusExibicao = statusExibicao.split("(")[1].replace(")", "");
        }
        
        // Coluna de Checkbox customizada
        let tdCheck = document.createElement('td');
        let check = document.createElement('input');
        check.type = "checkbox";
        check.className = "form-check-input check-nota";
        check.dataset.index = index;
        check.onclick = (e) => e.stopPropagation(); // Impede duplo clique ao selecionar
        tdCheck.appendChild(check);

        tr.innerHTML += `
            <td>${nota.emissao}</td>
            <td>${nota.nota}</td>
            <td class="fw-bold">${nota.cliente}</td>
            <td>${nota.empresa}</td>
            <td>${nota.valor_str}</td>
            <td style="color: #00E5FF; font-weight: bold;">${nota.total_contrato_str}</td>
            <td>${nota.dias}</td>
            <td class="fw-bold">${statusExibicao}</td>
            <td>${nota.ultimo_envio}</td>
        `;
        
        tr.prepend(tdCheck); // Bota o check no começo
        
        // Duplo clique abre os detalhes
        tr.ondblclick = () => abrirDetalhes(nota);
        
        // Clique simples seleciona/deseleciona
        tr.onclick = () => { check.checked = !check.checked; };
        
        tbody.appendChild(tr);
    });
}

function obterSelecionadas() {
    let checks = document.querySelectorAll('.check-nota:checked');
    return Array.from(checks).map(c => notasAtuais[c.dataset.index]);
}

async function exportarExcel() {
    let res = await eel.exportar_relatorio()();
    if (res.status === "erro") alert("Erro: " + res.msg);
}

function iniciarEnvio() {
    let sel = obterSelecionadas();
    if (sel.length === 0) return alert("Selecione pelo menos uma nota!");
    
    let niveis = {
        "AVISO AMIGÁVEL (PRÉ-VENCIMENTO)": 0,
        "FAIXA VERDE (ATÉ 15 DIAS)": 1, 
        "FAIXA AMARELA (16 A 30 DIAS)": 2, 
        "FAIXA LARANJA (31 A 60 DIAS)": 3, 
        "FAIXA VERMELHA (61 A 90 DIAS)": 4, 
        "FAIXA ROXA (91 A 120 DIAS)": 5,
        "FAIXA PRETA (+120 DIAS)": 6
    };
    
    let maior = -1;
    let nomeMaiorFaixa = "";
    
    for (let n of sel) {
        if (n.cliente !== sel[0].cliente) return alert("Selecione notas do MESMO CLIENTE.");
        
        // Corrige o nome da faixa para o formato oficial
        let nomeOficial = Object.keys(niveis).find(k => k.includes(n.parecer.split('(')[0].trim())) || n.parecer;
        
        let nv = niveis[nomeOficial] || 0;
        if (nv > maior) {
            maior = nv;
            nomeMaiorFaixa = nomeOficial; // Guarda o nome da maior faixa selecionada
        }
    }

    // Chama a nossa nova função de verificação antes de mandar bala!
    verificarEContinuar(sel, nomeMaiorFaixa, maior);
}

// ==========================================
// 🛡️ NOVO ESCUDO ANTI-DUPLICIDADE
// ==========================================
async function verificarEContinuar(sel, nomeFaixa, maior) {
    let btn = document.querySelector('.btn-action-custom');
    let textoAntigo = btn.innerText;
    
    btn.innerText = "⏳ CHECANDO..."; 
    btn.disabled = true;

    // Vai no Python perguntar se essa faixa já foi disparada pra essa nota
    let res = await eel.verificar_cobranca_duplicada_python(sel, nomeFaixa)();
    
    btn.innerText = textoAntigo; 
    btn.disabled = false;

    // Se o Python avisar que já tem e-mail disparado, joga na cara do usuário!
    // (Final da função verificarEContinuar atual...)
    if (res.duplicado) {
        let confirma = confirm(`⚠️ ALERTA DE DUPLICIDADE!\n\n${res.msg}\n\nTem certeza que deseja gerar e enviar esta cobrança repetida?`);
        if (!confirma) return; 
    }

    // 🚀 LÓGICA NOVA: Se for Laranja (3), Vermelha (4) ou Roxa (5), abre o seletor!
    if (maior >= 3 && maior <= 5) {
        prepararModalEscolha(maior);
        modal90.show();
    } else {
        executarEnvioBackend(""); 
    }
}

// ========================================================
// 🧠 CONSTRUTOR DINÂMICO DE MODAL (INJETA HTML NA HORA)
// ========================================================
function prepararModalEscolha(nivel) {
    let corpoModal = document.querySelector('#modal90Dias .modal-content');
    if (!corpoModal) return;

    let titulo = ""; let btn1Text = ""; let btn2Text = "";

    if (nivel === 3) {
        titulo = "Faixa Laranja (31 a 60 Dias)";
        btn1Text = "📄 Anexo I (Cobrança Administrativa - Financeiro)";
        btn2Text = "⚖️ Anexo II (Notificação Jurídico - Juridico)";
    } else if (nivel === 4) {
        titulo = "Faixa Vermelha (61 a 90 Dias)";
        btn1Text = "📄 Anexo III (Notificação Formal - Financeiro)";
        btn2Text = "⚖️ Anexo IV (Autoridade Sup. II - Jurídico)";
    } else if (nivel === 5) {
        titulo = "Faixa Roxa (91 a 120 Dias)";
        btn1Text = "⚖️ Minuta para Jurídico";
        btn2Text = "🛡️ Parecer Interno Preventivo";
    }

    // Sobrescreve o modal inteiro com os botões certos!
    corpoModal.innerHTML = `
        <div class="modal-header border-0 pb-0 justify-content-center mt-3">
            <h4 class="modal-title fw-bold text-warning">🚦 Escolha o Tipo de Documento</h4>
        </div>
        <div class="modal-body px-4 pb-4 text-center">
            <p class="text-white-50 mb-4">Você está disparando um alerta da <b>${titulo}</b>. Qual modelo deseja utilizar?</p>
            <div class="d-flex flex-column gap-3">
                <button class="btn btn-primary fw-bold py-2" onclick="confirmarEnvio90Dias('1')">${btn1Text}</button>
                <button class="btn btn-danger fw-bold py-2" onclick="confirmarEnvio90Dias('2')">${btn2Text}</button>
            </div>
            <button class="btn btn-outline-secondary text-white fw-bold w-100 rounded-pill py-2 mt-4" data-bs-dismiss="modal">Cancelar</button>
        </div>
    `;
}

async function executarEnvioBackend(escolha90) {
    modal90.hide();
    let sel = obterSelecionadas();
    let email = document.getElementById('combo-email').value;
    
    let btn = document.querySelector('.btn-action-custom');
    btn.innerText = "⏳ ENVIANDO..."; btn.disabled = true;

    let res = await eel.enviar_email_backend(sel, email, escolha90)();
    
    btn.innerText = "📧 GERAR E-MAIL"; btn.disabled = false;
    
    if (res.status === "erro") alert("Erro: " + res.msg);
    else carregarDadosEFiltrar(); // Atualiza a tela pós envio
}

function abrirDetalhes(nota) {
    let c = document.getElementById('corpo-detalhes');
    
    let caminhoSeguro = "";
    if (nota.caminho_evidencia) {
        caminhoSeguro = nota.caminho_evidencia.replace(/\\/g, '\\\\');
    }

    c.innerHTML = `
        <div class="row mb-1"><div class="col-3 fw-bold">Cliente:</div><div class="col-9">${nota.cliente}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Empresa:</div><div class="col-9">${nota.empresa}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Nº Nota:</div><div class="col-9">${nota.nota}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Competência:</div><div class="col-9">${nota.competencia}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Emissão:</div><div class="col-9">${nota.emissao}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Atraso:</div><div class="col-9">${nota.dias} dias</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Status:</div><div class="col-9">${nota.parecer}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Valor:</div><div class="col-9 text-success fw-bold">${nota.valor_str}</div></div>
        <div class="row mb-1"><div class="col-3 fw-bold">Últ. Envio:</div><div class="col-9">${nota.ultimo_envio} (por ${nota.usuario_envio})</div></div>
        <hr class="border-secondary my-2">
        <div class="row mb-3"><div class="col-3 fw-bold">Origem:</div><div class="col-9" style="font-size: 0.9em;">${nota.arquivo} (Linha ${nota.linha_excel})</div></div>
        
        <div class="mt-3 mb-4 p-3 rounded" style="background-color: rgba(255, 255, 255, 0.05); border: 1px dashed #aaa;">
            <h6 class="fw-bold text-info mb-3">📁 Documentação da Nota</h6>
            <div class="d-flex flex-wrap gap-2">
                <button class="btn btn-sm btn-outline-info fw-bold px-3" onclick="abrirPdf('${caminhoSeguro}')">
                    📄 ABRIR ARQUIVO
                </button>
                <button class="btn btn-sm btn-outline-warning fw-bold px-3" onclick="substituirEvidencia('${nota.nota}', '${nota.cliente}', '${nota.empresa}')">
                    📎 ANEXAR MANUAL
                </button>
                <button class="btn btn-sm btn-outline-secondary fw-bold px-3 text-white" onclick="abrirPastaCliente('${nota.cliente}', '${nota.empresa}')">
                    📂 ABRIR PASTA
                </button>
            </div>
        </div>
        
        <button class="btn btn-outline-info w-100 fw-bold" onclick="abrirTimeline('${nota.cliente}')">📜 VER HISTÓRICO COMPLETO DESTE CONTRATO</button>
    `;
    modalDet.show();
}

let modalTimeline; // Crie esta variável global

async function abrirTimeline(cliente) {
    if (!modalTimeline) modalTimeline = new bootstrap.Modal(document.getElementById('modalTimeline'));
    
    document.getElementById('timeline-cliente').innerText = "Cliente: " + cliente;
    
    // 👇 Garante que o botão do ZIP (que criamos antes) funcione para este cliente 👇
    let btnDossie = document.getElementById('btn-dossie');
    if (btnDossie) btnDossie.setAttribute('onclick', `gerarDossie('${cliente}')`);

    let corpo = document.getElementById('corpo-timeline');
    corpo.innerHTML = "<p class='text-center'>Buscando registros no banco de dados...</p>";
    
    modalTimeline.show();

    let logs = await eel.buscar_historico_cliente(cliente)();
    
    if (logs.length === 0) {
        corpo.innerHTML = "<p class='text-center text-secondary'>Nenhuma ação foi registrada para este contrato ainda.</p>";
        return;
    }

    let html = '<ul class="list-group list-group-flush">';
    logs.forEach(log => {
        // Pinta de verde se for whats, azul se for e-mail
        let corAcao = log.acao.includes('WhatsApp') ? 'color: #00E676;' : 'color: #00E5FF;';
        
        // ==========================================
        // 🧠 LÓGICA INTELIGENTE DOS BOTÕES DE ANEXO
        // ==========================================
        let botoesAnexo = "";
        if (log.anexo && log.anexo.trim() !== "") {
            // Se tem anexo, mostra o botão de VER e a LIXEIRA de remover
            let caminhoSeguro = log.anexo.replace(/\\/g, '\\\\');
            botoesAnexo = `
                <div class="d-flex gap-1">
                    <button class="btn btn-sm btn-outline-info" onclick="abrirPdf('${caminhoSeguro}')">📄 Ver</button>
                    <button class="btn btn-sm btn-outline-danger" title="Remover Anexo" onclick="removerAnexoTimeline('${log.id}', '${cliente}')">🗑️</button>
                </div>`;
        } else {
            // Se não tem, mostra o botão de ANEXAR
            botoesAnexo = `<button class="btn btn-sm btn-outline-warning" onclick="anexarNaTimeline('${log.id}', '${cliente}', '${log.nota}')">📎 Anexar Doc</button>`;
        }
        // ==========================================

        html += `
            <li class="list-group-item bg-transparent text-light border-secondary">
                <div class="d-flex w-100 justify-content-between">
                    <h6 class="mb-1 fw-bold" style="${corAcao}">${log.acao}</h6>
                    <small class="text-white-50">${log.data}</small>
                </div>
                <p class="mb-1">Nota Fiscal: <b>${log.nota}</b></p>
                <div class="d-flex justify-content-between align-items-center mt-2">
                    <small>Realizado por: <b>${log.usuario}</b></small>
                    ${botoesAnexo}
                </div>
            </li>
        `;
    });
    html += '</ul>';
    corpo.innerHTML = html;
}

//// zap

async function iniciarWhatsApp() {
    let sel = obterSelecionadas();
    
    if (sel.length === 0) return alert("Selecione pelo menos uma nota para gerar o alerta no WhatsApp!");

    let numeroDestino = document.getElementById('combo-operacao').value;
    if (!numeroDestino) return alert("Por favor, selecione o contato na caixinha ao lado do botão!");
    
    let clienteRef = sel[0].cliente;
    for (let n of sel) {
        if (n.cliente !== clienteRef) return alert("Selecione notas de apenas UM contrato para mandar o alerta.");
    }

    // 🚀 MÁGICA: Vai no banco de dados e busca o texto fresquinho que a chefia editou!
    let templateBanco = await eel.obter_texto_whatsapp()();
    
    // Troca a tag {cliente} pelo nome real do contrato
    let mensagem = templateBanco.replace(/{cliente}/g, clienteRef);

    let textoCodificado = encodeURIComponent(mensagem);
    
    for (let n of sel) {
        eel.registrar_log_atividade(n.cliente, n.nota, "📱 Alerta WhatsApp")();
    }

    let url = `https://wa.me/${numeroDestino}?text=${textoCodificado}`;
    window.open(url, 'aba_whatsapp_cobranca'); 
}

//// --- ANIMAÇÃO DO CANVAS NO BACKGROUND ---
function iniciarCanvasBackground() {
    const canvas = document.getElementById("bgCanvas");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;

    let particlesArray = [];
    class Particle {
        constructor() {
            this.x = Math.random() * canvas.width;
            this.y = Math.random() * canvas.height;
            this.size = Math.random() * 2 + 0.5;
            this.speedX = Math.random() * 0.8 - 0.4;
            this.speedY = Math.random() * 0.8 - 0.4;
            // Cores baseadas no estilo neon de vocês
            this.color = Math.random() > 0.5 ? '#FF6600' : '#00E5FF'; 
        }
        update() {
            this.x += this.speedX;
            this.y += this.speedY;
            if (this.x < 0 || this.x > canvas.width) this.speedX = -this.speedX;
            if (this.y < 0 || this.y > canvas.height) this.speedY = -this.speedY;
        }
        draw() {
            ctx.fillStyle = this.color;
            ctx.beginPath();
            ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
            ctx.fill();
        }
    }
    
    for (let i = 0; i < 80; i++) {
        particlesArray.push(new Particle());
    }

    function animate() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        for (let i = 0; i < particlesArray.length; i++) {
            particlesArray[i].update();
            particlesArray[i].draw();
        }
        requestAnimationFrame(animate);
    }
    animate();

    window.addEventListener('resize', () => {
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
    });
}

// ==========================================
// FUNÇÕES DO NOVO MODAL DE WHATSAPP (SNIPER)
// ==========================================

let textoWhatsAtual = ""; 

async function prepararAlertaWhats() {
    let sel = obterSelecionadas();
    if(sel.length === 0) return alert("Selecione pelo menos uma nota para cobrar!");

    let clienteRef = sel[0].cliente;
    for (let n of sel) {
        if (n.cliente !== clienteRef) return alert("Selecione apenas notas do MESMO contrato para o alerta!");
    }

    let modalEl = document.getElementById('modalWhatsPreview');
    let modal = new bootstrap.Modal(modalEl);
    
    document.getElementById('whats-texto-preview').innerText = "⏳ Gerando texto...";
    document.getElementById('area-botoes-whats').innerHTML = '<div class="text-center text-white-50"><small>Buscando contatos na base...</small></div>';
    
    modal.show();

    // 1. Gera o Texto
    let resTexto = await eel.gerar_preview_whats_python(sel)();
    if (resTexto.status === "sucesso") {
        textoWhatsAtual = resTexto.texto;
        document.getElementById('whats-texto-preview').innerHTML = resTexto.texto.replace(/\n/g, '<br>');
    } else {
        document.getElementById('whats-texto-preview').innerText = "❌ Erro ao gerar texto.";
    }

    // 2. Busca o Daison e o Supervisor no Python
    let resContatos = await eel.obter_contatos_whats_contrato(clienteRef)();
    if (resContatos.status === "sucesso") {
        let htmlBotoes = "";

        // Botão do Daison
        if (resContatos.daison_tel) {
            htmlBotoes += `
                <button class="btn btn-success fw-bold text-start px-3 py-2" style="box-shadow: 0 0 10px rgba(0, 230, 118, 0.4);" onclick="enviarWhatsDireto('${resContatos.daison_tel}')">
                    ⭐ Enviar para Daison (Gerente)
                </button>`;
        } else {
            htmlBotoes += `<div class="alert alert-danger p-2 small mb-0">⚠️ Daison sem número cadastrado!</div>`;
        }

        // Botão do Supervisor
        if (resContatos.sup_tel) {
            htmlBotoes += `
                <button class="btn btn-outline-info fw-bold text-start px-3 py-2 text-white" onclick="enviarWhatsDireto('${resContatos.sup_tel}')">
                    👤 Enviar para ${resContatos.sup_nome} (Supervisor)
                </button>`;
        } else {
            htmlBotoes += `<div class="alert alert-warning p-2 small mb-0">⚠️ Supervisor não encontrado na base.</div>`;
        }

        document.getElementById('area-botoes-whats').innerHTML = htmlBotoes;
    }
}

// Quando clicar no botão do modal, abre a aba do zap
async function enviarWhatsDireto(numero) {
    if(!numero || numero.length < 5) return alert("Número inválido no banco de dados!");
    if(!textoWhatsAtual) return alert("A mensagem não foi gerada corretamente.");

    let textoEncode = encodeURIComponent(textoWhatsAtual);
    let link = `https://api.whatsapp.com/send?phone=${numero}&text=${textoEncode}`;
    
    // Abre a aba do WhatsApp
    window.open(link, '_blank');

    // Registra no Histórico
    let sel = obterSelecionadas();
    let cliente = sel[0].cliente;
    let notaFiscal = sel[0].nota;
    await eel.registrar_log_atividade(cliente, notaFiscal, "📱 Alerta WhatsApp")();
}

async function confirmarEnvioWhats() {
    let numero = document.getElementById('whats-destinatario').value;
    
    if(!numero) return alert("Por favor, selecione um destinatário!");
    if(!textoWhatsAtual) return alert("A mensagem não foi gerada corretamente.");

    let textoEncode = encodeURIComponent(textoWhatsAtual);
    let link = `https://api.whatsapp.com/send?phone=${numero}&text=${textoEncode}`;
    window.open(link, '_blank');

    let sel = obterSelecionadas();
    let notaFiscal = sel[0].nota;
    let cliente = sel[0].cliente;
    await eel.registrar_log_python(cliente, "📱 Alerta WhatsApp", notaFiscal, "Ruan")();
    
    let modalEl = document.getElementById('modalWhatsPreview');
    bootstrap.Modal.getInstance(modalEl).hide();
}

// ==========================================
// FUNÇÕES DO PAINEL DE ADMINISTRAÇÃO
// ==========================================

async function carregarUsuarios() {
    let lista = await eel.listar_usuarios()();
    let tbody = document.getElementById('corpo-tabela-usuarios');
    tbody.innerHTML = "";
    
    lista.forEach(u => {
        tbody.innerHTML += `
            <tr>
                <td class="fw-bold text-white">${u.usuario}</td>
                <td>${u.funcao}</td>
                <td class="text-success">${u.telefone}</td>
                <td>
                    <button class="btn btn-sm btn-outline-info me-2 fw-bold" onclick="abrirModalUsuario(${u.id}, '${u.usuario}', '${u.funcao}', '${u.telefone}')">✏️ Editar</button>
                    <button class="btn btn-sm btn-outline-danger fw-bold" onclick="excluirUser(${u.id}, '${u.usuario}')">🗑️ Excluir</button>
                </td>
            </tr>
        `;
    });
}

function abrirModalUsuario(id='', nome='', funcao='', tel='') {
    document.getElementById('edit-user-id').value = id;
    document.getElementById('edit-user-nome').value = nome;
    document.getElementById('edit-user-senha').value = ""; // Senha sempre em branco por segurança
    document.getElementById('edit-user-funcao').value = funcao;
    document.getElementById('edit-user-tel').value = tel;
    
    if(id) {
        document.getElementById('tituloModalUsuario').innerText = "✏️ Editar Usuário";
        document.getElementById('hint-senha').innerText = "(Deixe em branco para manter a atual)";
    } else {
        document.getElementById('tituloModalUsuario').innerText = "👤 Novo Usuário";
        document.getElementById('hint-senha').innerText = "(Obrigatória)";
    }
    
    modalUser.show();
}

async function salvarUsuario() {
    let id = document.getElementById('edit-user-id').value;
    let nome = document.getElementById('edit-user-nome').value;
    let senha = document.getElementById('edit-user-senha').value;
    let funcao = document.getElementById('edit-user-funcao').value;
    let tel = document.getElementById('edit-user-tel').value;
    
    if(!nome) return alert("O nome do usuário é obrigatório!");
    if(!id && !senha) return alert("A senha é obrigatória para criar um novo usuário!");
    
    let res = await eel.salvar_usuario(id, nome, senha, funcao, tel)();
    if(res.status === "sucesso") {
        modalUser.hide();
        carregarUsuarios(); // Atualiza a tela de fundo
    } else {
        alert(res.msg);
    }
}

async function excluirUser(id, nome) {
    if(confirm(`🚨 ATENÇÃO!\nTem certeza que deseja apagar o acesso do usuário: ${nome}?`)) {
        let res = await eel.excluir_usuario(id)();
        if(res.status === "sucesso") carregarUsuarios();
        else alert("Erro ao excluir: " + res.msg);
    }
}

// ==========================================
// UTILITÁRIOS: SENHA E LOGOUT
// ==========================================

function fazerLogout() {
    if(confirm("Deseja realmente sair do sistema?")) {
        window.location.href = "login.html"; // Volta para a tela de login
    }
}

function abrirModalSenha() {
    document.getElementById("nova-senha-input").value = "";
    new bootstrap.Modal(document.getElementById('modalMudarSenha')).show();
}

async function salvarNovaSenha() {
    let nova = document.getElementById("nova-senha-input").value;
    if(nova.length < 3) return alert("A senha deve ter pelo menos 3 caracteres!");
    
    let res = await eel.alterar_senha_python(nova)();
    if(res.status === "sucesso") {
        alert("✅ Senha alterada com sucesso!");
        bootstrap.Modal.getInstance(document.getElementById('modalMudarSenha')).hide();
    } else {
        alert("❌ Erro: " + res.msg);
    }
}

// ==========================================
// GESTÃO DE TEMPLATES DE E-MAIL
// ==========================================

async function carregarTemplates() {
    let lista = await eel.listar_templates_email()();
    let tbody = document.getElementById('corpo-tabela-templates');
    tbody.innerHTML = "";
    
    lista.forEach(t => {
        tbody.innerHTML += `
            <tr>
                <td class="text-start ps-4 fw-bold text-white">${t.nome}</td>
                <td class="text-white-50">${t.assunto}</td>
                <td>
                    <button class="btn btn-sm btn-outline-success fw-bold" 
                        onclick="abrirModalTemplate('${t.id}', '${t.nome}', '${t.assunto}', \`${t.corpo}\`, '${t.anexo}', '${t.responsavel}')">✏️ Editar</button>
                </td>
            </tr>
        `;
    });
}

let modalTemp;
function abrirModalTemplate(id='', nome='', assunto='', corpo='', anexo='', responsavel='') {
    if(!modalTemp) modalTemp = new bootstrap.Modal(document.getElementById('modalTemplate'));
    
    document.getElementById('edit-temp-id').value = id;
    document.getElementById('edit-temp-nome').value = nome;
    document.getElementById('edit-temp-assunto').value = assunto;
    document.getElementById('edit-temp-corpo').value = corpo;
    
    // Alimenta os campos novos
    document.getElementById('edit-temp-anexo').value = anexo;
    document.getElementById('edit-temp-responsavel').value = responsavel;
    
    document.getElementById('titulo-modal-temp').innerText = id ? "📝 Editar Modelo" : "➕ Novo Modelo";
    modalTemp.show();
}


async function salvarTemplate() {
    let id = document.getElementById('edit-temp-id').value;
    let nome = document.getElementById('edit-temp-nome').value.toUpperCase();
    let assunto = document.getElementById('edit-temp-assunto').value;
    let corpo = document.getElementById('edit-temp-corpo').value;
    
    // Pega os valores novos
    let anexo = document.getElementById('edit-temp-anexo').value;
    let responsavel = document.getElementById('edit-temp-responsavel').value;
    
    if(!nome) return alert("O Nome do Modelo é obrigatório!");

    // Passa tudo pro Python!
    let res = await eel.salvar_template_email(id, nome, assunto, corpo, anexo, responsavel)();
    if(res.status === "sucesso") {
        alert("✅ Modelo salvo com sucesso!");
        modalTemp.hide();
        carregarTemplates();
    } else {
        alert("❌ Erro ao salvar: " + res.msg);
    }
}

function inserirTag(tag) {
    let area = document.getElementById('edit-temp-corpo');
    area.value += tag;
    area.focus();
}

// ==========================================
// FUNÇÃO DO MODAL DA FAIXA ROXA (MINUTA / PARECER)
// ==========================================
async function confirmarEnvio90Dias(escolha) {
    // 1. Esconde o modal para a tela ficar limpa e não travar
    let modalEl = document.getElementById('modal90Dias');
    let modalInst = bootstrap.Modal.getInstance(modalEl);
    if(modalInst) modalInst.hide();

    // 2. Pega as notas selecionadas na tabela principal
    let sel = obterSelecionadas();
    if(sel.length === 0) return alert("Nenhuma nota selecionada!");

    // 3. Pega a conta de e-mail (se existir o select de contas na sua tela)
    let conta = "Conta Padrão";
    let combo = document.getElementById('combo-operacao');
    if (combo) conta = combo.value;

    // 4. MANDA PRO PYTHON! O 'escolha' vai ser "1" ou "2" dependendo do botão que vc clicou
    let res = await eel.enviar_email_backend(sel, conta, escolha)();

    // 5. Trata a resposta
    if(res.status === "sucesso") {
        console.log("✅ Documento gerado no Outlook!");
        // carregarDadosReais(); // Se quiser que a tela atualize sozinha depois de enviar
    } else {
        alert("❌ Erro ao gerar documento: " + res.msg);
    }
}

// Exemplo de como ficaria a função no JS
async function abrirPdf(caminho) {
    if (caminho === "Não disponível" || caminho === "Salvo em versão anterior") {
        return alert("Não há evidência física salva para esta nota.");
    }
    
    let res = await eel.abrir_arquivo_evidencia(caminho)();
    if (res.status === "erro") {
        alert("❌ Erro: " + res.msg);
    }
}

// ==========================================
// FUNÇÕES DE GESTÃO DE ARQUIVOS E DOSSIÊS
// ==========================================

async function substituirEvidencia(nota, cliente, empresa) {
    let res = await eel.substituir_evidencia_python(nota, cliente, empresa)();
    
    if (res.status === "sucesso") {
        alert("✅ Nova evidência anexada e salva no sistema com sucesso!");
        modalDet.hide(); // Fecha o modal de detalhes
        carregarDadosEFiltrar(); // Atualiza a tabela por trás
    } else if (res.status !== "cancelado") {
        alert("❌ Erro ao substituir: " + res.msg);
    }
}

async function gerarDossie(cliente) {
    let btn = document.getElementById('btn-dossie');
    let textoAntigo = btn.innerHTML;
    
    // Efeito visual de carregamento
    btn.innerHTML = "⏳ EMPACOTANDO ARQUIVOS...";
    btn.disabled = true;

    let res = await eel.gerar_dossie_zip_python(cliente)();
    
    // Devolve o botão ao normal
    btn.innerHTML = textoAntigo;
    btn.disabled = false;

    if (res.status === "erro") {
        alert("❌ Erro: " + res.msg);
    }
}

async function anexarNaTimeline(idLog, cliente, notaFiscal) {
    let res = await eel.anexar_doc_timeline(idLog, cliente, notaFiscal)();
    if(res.status === "sucesso") {
        alert("✅ Arquivo anexado à linha do tempo com sucesso!");
        abrirTimeline(cliente); // Recarrega a timeline para o botão novo aparecer
    } else if(res.status !== "cancelado") {
        alert("❌ Erro: " + res.msg);
    }
}

async function removerAnexoTimeline(idLog, cliente) {
    if (confirm("🚨 ATENÇÃO!\nTem certeza que deseja remover este anexo? A ação não pode ser desfeita.")) {
        let res = await eel.remover_anexo_timeline_python(idLog)();
        if (res.status === "sucesso") {
            // Recarrega a timeline silenciosamente para o botão sumir e virar "Anexar" de novo
            abrirTimeline(cliente); 
        } else {
            alert("❌ Erro ao remover anexo: " + res.msg);
        }
    }
}

async function abrirPastaCliente(cliente, empresa) {
    let res = await eel.abrir_pasta_cliente_python(cliente, empresa)();
    if (res.status === "erro") {
        alert("❌ Erro: " + res.msg);
    }
}