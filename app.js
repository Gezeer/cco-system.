<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />

  <title>Intranet Executiva CCO</title>

  <link rel="icon" type="image/png" href="logo.png">
  <link rel="stylesheet" href="style.css" />

  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>

<body>
  <div class="layout">

    <aside class="sidebar">
      <div class="brand">
        <div class="brand-icon">CCO</div>

        <div>
          <h2>ÍNDICE DE EFICIÊNCIA OPERACIONAL</h2>
          <p>Painel Executivo</p>
        </div>
      </div>

      <div class="sidebar-text">
        Gestão executiva dos indicadores operacionais, contratuais e KPI’s dos serviços P1 a P12.
      </div>

      <nav class="menu">
        <button class="active" onclick="mostrarTela('executivo', this)">
          📊 Painel Executivo
        </button>

        <button onclick="mostrarTela('contrato', this)">
          📈 Execução P1 a P12
        </button>

        <button id="menuKpi" onclick="mostrarTela('kpi', this)">
          ⚡ KPI's
        </button>

        <button onclick="mostrarTela('dados', this)">
          🗂️ Base Importada
        </button>

        <button onclick="mostrarTela('historico', this); carregarHistorico();">
          📜 Histórico
        </button>

        <button class="logout" onclick="sair()">
          ➜ Sair
        </button>
      </nav>
    </aside>

    <main class="content">

      <header class="topbar">
        <div>
          <span class="eyebrow">Centro de Controle Operacional 👥</span>

          <h1>Intranet Executiva de KPI’s</h1>

          <p>
            Consolidação dos indicadores de limpeza urbana, execução contratual,
            produtividade, volume operacional e respostas executivas por serviço.
          </p>
        </div>

        <div class="top-actions">
          <div class="date-box" id="dataAtual"></div>

          <label class="file-label" id="btnImportar">
            📥 Importar planilhas
            <input type="file" id="arquivoExcel" accept=".xlsx,.xls,.csv" multiple />
          </label>

          <button id="btnExportar" class="btn secondary" onclick="window.print()">
            📄 Exportar PDF
          </button>

          <button id="btnLimpar" class="btn danger" onclick="limparBanco()">
            🚮 Limpar dados
          </button>

          <div class="file-status" id="nomeArquivo">
            Nenhuma planilha importada
          </div>
        </div>
      </header>

      <div id="loadingOverlay" class="loading-overlay">
        <div class="loading-box">
          <div class="spinner"></div>
          <strong>Processando planilha...</strong>
          <span>Aguarde enquanto os dados são analisados.</span>
        </div>
      </div>

      <!-- PAINEL EXECUTIVO -->
      <section id="tela-executivo" class="tela ativa">

        <section class="filter-panel">
          <div class="filter-header">
            <div>
              <span>FILTRO OPERACIONAL</span>
              <h3>Painel diário e mensal 🗓️</h3>
            </div>

            <button class="btn-limpar-filtro" onclick="limparFiltroPeriodo()">
              Limpar filtros
            </button>
          </div>

          <div class="filter-grid">

            <div class="filter-group">
              <label>Filtrar por dia</label>
              <input type="date" id="filtroDia" onchange="aplicarFiltroDiario()" />
            </div>

            <div class="filter-group">
              <label>Mês</label>

              <select id="filtroMes" onchange="aplicarFiltroMensal()">
                <option value="">Selecionar</option>
                <option value="01">Janeiro</option>
                <option value="02">Fevereiro</option>
                <option value="03">Março</option>
                <option value="04">Abril</option>
                <option value="05">Maio</option>
                <option value="06">Junho</option>
                <option value="07">Julho</option>
                <option value="08">Agosto</option>
                <option value="09">Setembro</option>
                <option value="10">Outubro</option>
                <option value="11">Novembro</option>
                <option value="12">Dezembro</option>
              </select>
            </div>

            <div class="filter-group">
              <label>Ano</label>

              <select id="filtroAno" onchange="aplicarFiltroMensal()">
                <option value="">Selecionar</option>
                <option value="2026">2026</option>
                <option value="2027">2027</option>
                <option value="2028">2028</option>
              </select>
            </div>

          </div>
        </section>

        <section class="cards">
          <div class="card">
            <span>Serviços com dados</span>
            <strong id="kpiServicosDados">0</strong>
            <small>P1 a P12</small>
          </div>

          <div class="card">
            <span>Média de execução</span>
            <strong id="kpiExecucaoMedia">0%</strong>
            <small>máximo 100%</small>
          </div>

          <div class="card">
            <span>Abas importadas</span>
            <strong id="kpiAbas">0</strong>
            <small>planilhas processadas</small>
          </div>
        </section>

        <section class="section">
          <div class="section-title">
            <h2>🤖 Resumo Automático Executivo</h2>
          </div>

          <div id="resumoAutomaticoDiretoria" class="resumo-ia">
            Importe uma planilha para gerar o resumo automático.
          </div>
        </section>

        <section class="grid-2">
          <div class="section chart-card">
            <div class="section-title">
              <span>Indicador</span>
              <h2>Execução por Serviço</h2>
            </div>

            <canvas id="graficoExecucao"></canvas>
          </div>

          <div class="section chart-card">
            <div class="section-title">
              <span>Produção</span>
              <h2>Acumulado por Serviço</h2>
            </div>

            <canvas id="graficoAcumulado"></canvas>
          </div>
        </section>

        <section class="grid-2">
          <div class="section">
            <div class="section-title">
              <span>Resumo</span>
              <h2>📋 Resumo Executivo</h2>
            </div>

            <div class="metric-list" id="resumoExecutivo"></div>
          </div>

          <div class="section">
            <div class="section-title">
              <span>Performance</span>
              <h2>🏆 Ranking Operacional</h2>
            </div>

            <div id="rankingOperacional" class="ranking-list">
              Nenhum dado disponível.
            </div>
          </div>
        </section>

        <section class="section">
          <div class="section-title">
            <span>Monitoramento</span>
            <h2>🚨 Alertas Executivos</h2>
          </div>

          <div id="alertasExecutivos" class="metric-list"></div>
        </section>

        <section class="section">
          <div class="section-title">
            <span>Tabela oficial</span>
            <h2>Painel Executivo</h2>
          </div>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Serviço</th>
                  <th>Nome Serviço</th>
                  <th>Acumulado no Mês</th>
                  <th>Medição</th>
                  <th>Previsto Mês</th>
                  <th>% Execução</th>
                  <th>Dias acumulados</th>
                  <th>Total dias mês</th>
                  <th>Valor</th>
                  <th>Status</th>
                </tr>
              </thead>

              <tbody id="tabelaPainelExecutivo"></tbody>
            </table>
          </div>
        </section>

      </section>

      <!-- EXECUÇÃO P1 A P12 -->
      <section id="tela-contrato" class="tela">
        <div class="section">
          <div class="section-title">
            <span>Execução contratual</span>
            <h2>✅ Serviços P1 a P12</h2>
          </div>

          <p class="section-description">
            Selecione um serviço para visualizar os dados detalhados da respectiva aba da planilha.
          </p>

          <div class="tabs-servicos">
            <button class="servico-btn active" onclick="mostrarServico('geral', this)">📚 Geral<br><small>Consolidado</small></button>
            <button class="servico-btn" onclick="mostrarServico('P1', this)">P1<br><small>Coleta Orgânica</small></button>
            <button class="servico-btn" onclick="mostrarServico('P2.1', this)">P2.1<br><small>Coleta Seletiva</small></button>
            <button class="servico-btn" onclick="mostrarServico('P2.2', this)">P2.2<br><small>Rejeito Seletivo</small></button>
            <button class="servico-btn" onclick="mostrarServico('P3', this)">P3<br><small>Remoção Manual</small></button>
            <button class="servico-btn" onclick="mostrarServico('P4', this)">P4<br><small>Remoção Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarServico('P5', this)">P5<br><small>Varrição Manual</small></button>
            <button class="servico-btn" onclick="mostrarServico('P6', this)">P6<br><small>Varrição Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarServico('P7', this)">P7<br><small>Lavagem de Vias</small></button>
            <button class="servico-btn" onclick="mostrarServico('P8', this)">P8<br><small>Limpeza de Bens</small></button>
            <button class="servico-btn" onclick="mostrarServico('P9', this)">P9<br><small>Catação</small></button>
            <button class="servico-btn" onclick="mostrarServico('P10', this)">P10<br><small>Pintura Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarServico('P11', this)">P11<br><small>Pós-eventos/Gordura</small></button>
            <button class="servico-btn" onclick="mostrarServico('P12', this)">P12<br><small>Transbordo</small></button>
          </div>

          <div id="servico-geral" class="servico-painel ativa">
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Programa</th>
                    <th>KPI</th>
                    <th>Previsto</th>
                    <th>Executado</th>
                    <th>Índice</th>
                    <th>Status</th>
                  </tr>
                </thead>

                <tbody id="tabelaContratual"></tbody>
              </table>
            </div>
          </div>

          <div id="servico-detalhe" class="servico-painel">
            <div id="detalheServico"></div>
          </div>
        </div>
      </section>

      <!-- KPI -->
      <section id="tela-kpi" class="tela">
        <div class="section">
          <div class="section-title">
            <span>Indicadores operacionais</span>
            <h2>⚡ KPI's por Serviço</h2>
          </div>

          <p class="section-description">
            Selecione um serviço para visualizar as perguntas, respostas e fórmulas dos KPI’s.
          </p>

          <div class="tabs-servicos">
            <button class="servico-btn active" onclick="mostrarKPI('geral', this)">📚 Geral<br><small>Indicadores</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P1', this)">P1<br><small>Coleta Orgânica</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P2.1', this)">P2.1<br><small>Coleta Seletiva</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P2.2', this)">P2.2<br><small>Rejeito Seletivo</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P3', this)">P3<br><small>Remoção Manual</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P4', this)">P4<br><small>Remoção Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P5', this)">P5<br><small>Varrição Manual</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P6', this)">P6<br><small>Varrição Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P7', this)">P7<br><small>Lavagem de Vias</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P8', this)">P8<br><small>Limpeza de Bens</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P9', this)">P9<br><small>Catação</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P10', this)">P10<br><small>Pintura Mecanizada</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P11', this)">P11<br><small>Pós-eventos/Gordura</small></button>
            <button class="servico-btn" onclick="mostrarKPI('P12', this)">P12<br><small>Transbordo</small></button>
          </div>

          <div id="conteudoKPI"></div>
        </div>
      </section>

      <!-- BASE IMPORTADA -->
      <section id="tela-dados" class="tela">
        <div class="section">
          <div class="section-title">
            <span>Conferência</span>
            <h2>🗂️ Base Importada</h2>
          </div>

          <div class="filtros">
            <input id="busca" placeholder="Buscar na base..." oninput="renderTabelaDados()" />

            <input type="date" id="filtroDataBase" onchange="renderTabelaDados()" />

            <select id="filtroPrograma" onchange="renderTabelaDados()">
              <option>Todos</option>
            </select>

            <select id="filtroStatus" onchange="renderTabelaDados()">
              <option>Todos</option>
            </select>

            <button class="btn-limpar-filtro" onclick="limparFiltroBase()">
              Limpar
            </button>
          </div>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Serviço</th>
                  <th>Origem</th>
                  <th>Data</th>
                  <th>Turno</th>
                  <th>RA</th>
                  <th>Peso</th>
                  <th>Viagens</th>
                  <th>KM</th>
                  <th>Equipe</th>
                </tr>
              </thead>

              <tbody id="tabelaDados"></tbody>
            </table>
          </div>
        </div>
      </section>

      <!-- HISTÓRICO -->
      <section id="tela-historico" class="tela">
        <div class="section">
          <div class="section-title">
            <span>Controle de versões</span>
            <h2>📜 Histórico de Importações</h2>
          </div>

          <p class="section-description">
            Consulte as importações realizadas, veja qual base está ativa e restaure versões anteriores.
          </p>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Usuário</th>
                  <th>Perfil</th>
                  <th>Arquivo</th>
                  <th>Abas</th>
                  <th>Status</th>
                  <th>Ações</th>
                </tr>
              </thead>

              <tbody id="tabelaHistorico">
                <tr>
                  <td colspan="7">Carregando histórico...</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </section>

      <footer>
        Intranet KPI CCO • Serviço de Limpeza Urbana • Valor Ambiental
      </footer>
    </main>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
  <script src="app.js"></script>
</body>
</html>
