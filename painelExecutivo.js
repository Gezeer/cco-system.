const metas = {
  "P1":  { nome: "Coleta Domiciliar", medicao: "T/mês", previsto: 21223, valor: 296 },
  "P2.1":{ nome: "Coleta Seletiva", medicao: "Vg/mês", previsto: 780, valor: 1027.42 },
  "P2.2":{ nome: "Rejeito Seletivo Das IRR", medicao: "Vg/mês", previsto: 260, valor: 1027.42 },
  "P3":  { nome: "Remoção Manual", medicao: "Equipe", previsto: 12, valor: 41992.93 },
  "P4":  { nome: "Remoção Mecanizada", medicao: "T/mês", previsto: 15779, valor: 68.8 },
  "P5":  { nome: "Varrição Manual", medicao: "Km/mês", previsto: 38541.8, valor: 160.94, aguardando: true },
  "P6":  { nome: "Varrição Mecanizada", medicao: "Km/mês", previsto: 9040, valor: 76.24, aguardando: true },
  "P7":  { nome: "Lavagem De Vias E Logradouros", medicao: "Equipe", previsto: 2, valor: 49811.72 },
  "P8":  { nome: "Limpeza De Equipamentos E Bens", medicao: "Equipe", previsto: 2, valor: 81001.04 },
  "P9":  { nome: "Catação Em Área Verde", medicao: "Equipe", previsto: 11, valor: 122039.23 },
  "P10": { nome: "Pintura Mecanizada", medicao: "Equipe", previsto: 3, valor: 346660.01 },
  "P11": { nome: "Limpeza Pós-Eventos e Coleta De Gordura", medicao: "Equipe", previsto: 1, valor: 272459.08 },
  "P12": { nome: "Transbordo", medicao: "T x Km/Vg", previsto: 1698432, valor: 0.83 }
};

function parseDataBR(dataTexto) {
  if (!dataTexto) return null;

  const partes = String(dataTexto).slice(0, 10).split("/");

  if (partes.length === 3) {
    return new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  }

  const data = new Date(dataTexto);
  return isNaN(data) ? null : data;
}

function semanaDoMes(dataTexto) {
  const data = parseDataBR(dataTexto);

  if (!data) return 0;

  const dia = data.getDate();

  if (dia <= 8) return 1;
  if (dia <= 14) return 2;
  if (dia <= 20) return 3;
  return 4;
}

function codigoExecutivo(item) {
  const texto = `${item.programa || ""} ${item.programa_nome || ""} ${item.servico || ""}`.toLowerCase();

  if (texto.includes("p1") || texto.includes("domiciliar") || texto.includes("orgânica") || texto.includes("organica")) return "P1";

  if (texto.includes("rejeito")) return "P2.2";

  if (texto.includes("p2") || texto.includes("seletiva")) return "P2.1";

  if (texto.includes("p3") || texto.includes("remoção manual") || texto.includes("remocao manual")) return "P3";

  if (texto.includes("p4") || texto.includes("remoção mecanizada") || texto.includes("remocao mecanizada")) return "P4";

  if (texto.includes("p5") || texto.includes("varrição manual") || texto.includes("varricao manual")) return "P5";

  if (texto.includes("p6") || texto.includes("varrição mecanizada") || texto.includes("varricao mecanizada")) return "P6";

  if (texto.includes("p7") || texto.includes("lavagem")) return "P7";

  if (texto.includes("p8") || texto.includes("equipamentos") || texto.includes("bens")) return "P8";

  if (texto.includes("p9") || texto.includes("catação") || texto.includes("catacao")) return "P9";

  if (texto.includes("p10") || texto.includes("pintura")) return "P10";

  if (texto.includes("p11") || texto.includes("evento") || texto.includes("gordura")) return "P11";

  if (texto.includes("p12") || texto.includes("transbordo")) return "P12";

  return null;
}

function valorExecutado(codigo, item) {
  if (["P1", "P4"].includes(codigo)) {
    return Number(item.peso || 0) / 1000;
  }

  if (["P2.1", "P2.2"].includes(codigo)) {
    return Number(item.viagens || 0) || 1;
  }

  if (["P5", "P6"].includes(codigo)) {
    return Number(item.km || 0);
  }

  if (codigo === "P12") {
    return Number(item.peso || 0) * Number(item.km || 0);
  }

  return item.veiculo ? 1 : 0;
}

function montarPainelExecutivo(operacoes) {
  const painel = {};

  Object.keys(metas).forEach(codigo => {
    painel[codigo] = {
      servico: codigo,
      nome_servico: metas[codigo].nome,
      acumulado_mes: 0,
      semana_1: 0,
      semana_2: 0,
      semana_3: 0,
      semana_4: 0,
      medicao: metas[codigo].medicao,
      previsto_mes: metas[codigo].previsto,
      porcentagem_execucao: 0,
      valor: metas[codigo].valor,
      status: metas[codigo].aguardando ? "Aguardando planilha" : "Sem dados"
    };
  });

  operacoes.forEach(item => {
    const codigo = codigoExecutivo(item);

    if (!codigo || !painel[codigo]) return;

    const valor = valorExecutado(codigo, item);
    const semana = semanaDoMes(item.data);

    painel[codigo].acumulado_mes += valor;

    if (semana === 1) painel[codigo].semana_1 += valor;
    if (semana === 2) painel[codigo].semana_2 += valor;
    if (semana === 3) painel[codigo].semana_3 += valor;
    if (semana === 4) painel[codigo].semana_4 += valor;

    painel[codigo].status = "Com dados";
  });

  Object.values(painel).forEach(linha => {
    linha.acumulado_mes = Number(linha.acumulado_mes.toFixed(2));
    linha.semana_1 = Number(linha.semana_1.toFixed(2));
    linha.semana_2 = Number(linha.semana_2.toFixed(2));
    linha.semana_3 = Number(linha.semana_3.toFixed(2));
    linha.semana_4 = Number(linha.semana_4.toFixed(2));

    linha.porcentagem_execucao = linha.previsto_mes
      ? Number(((linha.acumulado_mes / linha.previsto_mes) * 100).toFixed(2))
      : 0;
  });

  return Object.values(painel);
}

module.exports = { montarPainelExecutivo };