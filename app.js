dayjs.extend(dayjs_plugin_customParseFormat);

const DATE_FORMATS = ["DD-MM-YYYY", "DD/MM/YYYY", "YYYY-MM-DD", "YYYY/MM/DD"];
const GOODS_AND_SERVICES = new Set([
  "aquisição de bens móveis",
  "aquisição de serviços",
  "locação de bens móveis",
]);
const PUBLIC_WORKS = new Set(["empreitadas de obras públicas"]);
const FUNDAMENTACAO_ART_20_D = normalizeText(
  "Artigo 20.º, n.º 1, alínea d) do Código dos Contratos Públicos"
);
const FUNDAMENTACAO_ART_19_D = normalizeText(
  "Artigo 19.º, alínea d) do Código dos Contratos Públicos"
);
const FUNDAMENTACAO_ART_20_C = normalizeText(
  "Artigo 20.º, n.º 1, alínea c) do Código dos Contratos Públicos"
);
const FUNDAMENTACAO_ART_19_C = normalizeText(
  "Artigo 19.º, alínea c) do Código dos Contratos Públicos"
);
const PROCEDIMENTOS_EXCECAO_D = new Set(
  [
    "Artigo 24.º, n.º 1, alínea a) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea b) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea c) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea d) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea e), subalínea i) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea e), subalínea ii) do Código dos Contratos Públicos",
    "Artigo 24.º, n.º 1, alínea e), subalínea iii) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea a) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea b) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea c) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea d) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea e) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea g) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea h) do Código dos Contratos Públicos",
    "Artigo 27.º, n.º 1, alínea i) do Código dos Contratos Públicos",
  ].map(normalizeText)
);
const PROCEDIMENTOS_URGENTES_G = new Set(
  [
    "Artigo 155.º do Código dos Contratos Públicos",
    "Artigo 155.º, alínea a) do Código dos Contratos Públicos",
    "Artigo 155.º, alínea b) do Código dos Contratos Públicos",
  ].map(normalizeText)
);
const PROCEDIMENTOS_ACORDO_H = new Set(
  [
    "Artigo 258.º do Código dos Contratos Públicos",
    "Artigo 259.º do Código dos Contratos Públicos",
    "Artigo 252.º, n.º 1, alínea b) do Código dos Contratos Públicos",
  ].map(normalizeText)
);
const PROCEDIMENTO_ART_6A = normalizeText(
  "artigo 6.º-A n.º 1 do Código dos Contratos Públicos"
);

const statusBox = document.getElementById("status");
const generateBtn = document.getElementById("generate-btn");
const previewBtn = document.getElementById("preview-btn");
const analysisBtn = document.getElementById("analysis-btn");
const auditorInput = document.getElementById("auditor-name");
const reportPreview = document.getElementById("report-preview");
const previewContent = document.getElementById("preview-content");
const closePreview = document.getElementById("close-preview");
const cancelPreviewBtn = document.getElementById("cancel-preview-btn");
const confirmGenerateBtn = document.getElementById("confirm-generate-btn");
const dataAnalysis = document.getElementById("data-analysis");
const closeAnalysis = document.getElementById("close-analysis");
const auditHistory = document.getElementById("audit-history");
const historyBtn = document.getElementById("history-btn");
const closeHistory = document.getElementById("close-history");
const historyList = document.getElementById("history-list");

let previewData = null; // Guardar dados para gerar PDF depois
let tableData = []; // Dados da tabela
let filteredData = []; // Dados filtrados
let currentPage = 1;
let pageSize = 25;
let sortColumn = null;
let sortDirection = "asc";

const { jsPDF } = window.jspdf;

function capitalizeWords(text) {
  if (!text) return "";
  return text.replace(/\w\S*/g, (word) => word.charAt(0).toUpperCase() + word.slice(1));
}

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.style.color = isError ? "#c0392b" : "#163a7d";
}

function normalizeText(value) {
  if (!value) return "";
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function parsePrice(value) {
  if (!value) return 0;
  const cleaned = value
    .toString()
    .replace(/[^0-9,.-]/g, "")
    .replace(/\.(?=\d{3}(?:\D|$))/g, "")
    .replace(/,(\d{1,2})$/, ".$1");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  if (!value) return null;
  const trimmed = value.trim();
  for (const fmt of DATE_FORMATS) {
    const parsed = dayjs(trimmed, fmt, true);
    if (parsed.isValid()) return parsed;
  }
  const fallback = dayjs(trimmed);
  return fallback.isValid() ? fallback : null;
}

function parseYear(value) {
  const parsed = parseDate(value);
  return parsed ? parsed.year() : null;
}

// Feriados nacionais portugueses (formato MM-DD)
const FERIADOS_NACIONAIS = [
  "01-01", // Ano Novo
  "04-25", // Dia da Liberdade
  "05-01", // Dia do Trabalhador
  "06-10", // Dia de Portugal
  "08-15", // Assunção de Nossa Senhora
  "10-05", // Implantação da República
  "11-01", // Todos os Santos
  "12-01", // Restauração da Independência
  "12-08", // Imaculada Conceição
  "12-25", // Natal
];

// Calcular dias úteis entre duas datas (não conta o dia inicial, conta sábados/domingos e feriados)
function calculateWorkingDays(startDate, endDate) {
  if (!startDate || !endDate) return null;
  
  const start = dayjs(startDate);
  const end = dayjs(endDate);
  
  if (!start.isValid() || !end.isValid()) return null;
  if (end.isBefore(start)) return null;
  
  // Começar a contar no dia seguinte à celebração (não conta o dia do evento)
  let current = start.add(1, "day");
  let workingDays = 0;
  
  while (current.isBefore(end) || current.isSame(end, "day")) {
    const dayOfWeek = current.day(); // 0 = domingo, 6 = sábado
    const monthDay = current.format("MM-DD");
    
    // Não contar sábados (6) e domingos (0)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      // Não contar feriados nacionais
      if (!FERIADOS_NACIONAIS.includes(monthDay)) {
        workingDays++;
      }
    }
    
    current = current.add(1, "day");
  }
  
  return workingDays;
}

function extractContractTypes(contract) {
  return (contract.tiposContrato || "")
    .split(/<br\s*\/?>|\||;|\n/g)
    .map((value) =>
      normalizeText(
        value
          .replace(/&nbsp;/gi, " ")
          .replace(/&amp;/gi, "&")
      )
    )
    .filter(Boolean);
}

function matchesAnyContractType(contract, allowedSet) {
  const values = extractContractTypes(contract);
  return values.some((value) => allowedSet.has(value));
}

function matchesFundamentacao(contract, fundamentacao) {
  return normalizeText(contract.fundamentacao) === fundamentacao;
}

function matchesProcedimento(contract, targetSet) {
  return targetSet.has(normalizeText(contract.tipoProcedimento));
}

function hasNumericValue(value) {
  if (!value) return false;
  return /^\d+$/.test(value.trim());
}

function formatCurrency(value) {
  return new Intl.NumberFormat("pt-PT", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function numberToPortugueseWords(value) {
  if (value === 0) return "zero";
  const units = [
    "zero",
    "um",
    "dois",
    "três",
    "quatro",
    "cinco",
    "seis",
    "sete",
    "oito",
    "nove",
  ];
  const teens = [
    "dez",
    "onze",
    "doze",
    "treze",
    "catorze",
    "quinze",
    "dezasseis",
    "dezassete",
    "dezasseito",
    "dezanove",
  ];
  const tens = [
    "",
    "dez",
    "vinte",
    "trinta",
    "quarenta",
    "cinquenta",
    "sessenta",
    "setenta",
    "oitenta",
    "noventa",
  ];
  const hundreds = [
    "",
    "cento",
    "duzentos",
    "trezentos",
    "quatrocentos",
    "quinhentos",
    "seiscentos",
    "setecentos",
    "oitocentos",
    "novecentos",
  ];

  function belowHundred(n) {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    const ten = Math.floor(n / 10);
    const unit = n % 10;
    if (unit === 0) return tens[ten];
    return `${tens[ten]} e ${units[unit]}`;
  }

  function belowThousand(n) {
    if (n === 0) return "";
    if (n === 100) return "cem";
    const hundred = Math.floor(n / 100);
    const remainder = n % 100;
    const parts = [];
    if (hundred > 0) parts.push(hundreds[hundred]);
    if (remainder > 0) parts.push(belowHundred(remainder));
    return parts.join(" e ");
  }

  const thousands = Math.floor(value / 1000);
  const remainder = value % 1000;
  const parts = [];

  if (thousands > 0) {
    if (thousands === 1) parts.push("mil");
    else parts.push(`${belowThousand(thousands)} mil`);
  }

  if (remainder > 0) parts.push(belowThousand(remainder));
  return parts.join(" e ").replace(/\s+/g, " ").trim();
}

function formatYearWithWords(year) {
  return `${year} (${numberToPortugueseWords(year)})`;
}

function ensureYearOrder(contracts) {
  return Array.from(
    new Set(contracts.map((contract) => contract.year).filter((y) => y > 0))
  ).sort((a, b) => a - b);
}

function describeContract(contract) {
  return [
    `Objeto: ${contract.objeto || "(sem descrição)"}`,
    `Adjudicatária: ${contract.adjudicataria || "(desconhecida)"}`,
    `Preço contratual: ${formatCurrency(contract.precoContratual)}`,
    `Data de celebração: ${contract.dataCelebracao || "(não indicada)"}`,
    `Tipo de procedimento: ${contract.tipoProcedimento}`,
    `Fundamentação: ${contract.fundamentacao}`,
  ].join("\n");
}

function prepareContext(contracts) {
  const years = ensureYearOrder(contracts);
  const yearN = years[years.length - 1];
  const yearNMinus1 = years.includes(yearN - 1) ? yearN - 1 : null;
  const yearNMinus2 = years.includes(yearN - 2) ? yearN - 2 : null;
  return { yearN, yearNMinus1, yearNMinus2 };
}

function mapRowToContract(row, source) {
  const precoContratual = parsePrice(row["Preço Contratual"]);
  const dataCelebracao = (row["Data de Celebração do Contrato"] || "").trim();
  const year =
    parseYear(dataCelebracao) ||
    parseYear(row["Data de Publicação"]) ||
    0;

  return {
    id: `${row["Objeto do Contrato"] || ""}::${row["Entidade(s) Adjudicatária(s)"] || ""}::${dataCelebracao}::${row["Preço Contratual"] || ""}`,
    year,
    source,
    objeto: row["Objeto do Contrato"] || "",
    tipoProcedimento: row["Tipo de Procedimento"] || "",
    tiposContrato: row["Tipo(s) de Contrato"] || "",
    fundamentacao: row["Fundamentação"] || "",
    cpv: row["CPV"] || "",
    adjudicante: row["Entidade(s) Adjudicante(s)"] || "",
    adjudicataria: row["Entidade(s) Adjudicatária(s)"] || "",
    precoContratual,
    dataPublicacao: (row["Data de Publicação"] || "").trim(),
    dataCelebracao,
    prazoExecucao: (row["Prazo de Execução"] || "").trim(),
    localExecucao: (row["Local de Execução"] || "").trim(),
    numeroAcordoQuadro: (row["N.º registo do Acordo Quadro"] || "").trim(),
    estado: (row["Estado"] || "").trim(),
  };
}

function registerFinding(section, rule, contract, details, flagged) {
  section.findings.push({ contract, details, rule });
  flagged.add(contract.id);
}

function createLimitRule(context, contracts, section, flagged, config) {
  const { ruleCode, limit, fundamentacao, contractTypes, description } = config;

  const filtered = contracts.filter(
    (contract) =>
      matchesFundamentacao(contract, fundamentacao) &&
      matchesAnyContractType(contract, contractTypes)
  );

  filtered
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => contract.precoContratual > limit)
    .forEach((contract) =>
      registerFinding(
        section,
        `${ruleCode}1`,
        contract,
        `${description}: valor individual (${formatCurrency(
          contract.precoContratual
        )}) acima do limite de ${formatCurrency(limit)}.`,
        flagged
      )
    );

  const sameYear = filtered.filter((contract) => contract.year === context.yearN);
  const grouped = new Map();
  sameYear.forEach((contract) => {
    const key = normalizeText(contract.adjudicataria);
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(contract);
  });

  // Calcular soma dos anos anteriores (n-2 + n-1) por entidade
  const pastYears = filtered.filter(
    (contract) =>
      contract.year === context.yearNMinus1 ||
      contract.year === context.yearNMinus2
  );
  const pastYearSums = new Map();
  pastYears.forEach((contract) => {
    const key = normalizeText(contract.adjudicataria);
    pastYearSums.set(key, (pastYearSums.get(key) || 0) + contract.precoContratual);
  });

  grouped.forEach((contractsByEntity) => {
    contractsByEntity.sort((a, b) => {
      const dateA = parseDate(a.dataCelebracao || "")?.valueOf() || 0;
      const dateB = parseDate(b.dataCelebracao || "")?.valueOf() || 0;
      return dateA - dateB;
    });

    // Começar com a soma dos anos anteriores
    const entityKey = normalizeText(contractsByEntity[0]?.adjudicataria || "");
    let cumulative = pastYearSums.get(entityKey) || 0;
    
    contractsByEntity.forEach((contract) => {
      if (cumulative >= limit) {
        if (!flagged.has(contract.id)) {
          registerFinding(
            section,
            `${ruleCode}2`,
            contract,
            `${description}: soma acumulada (anos anteriores + ano n) atingiu ${formatCurrency(
              cumulative
            )} antes deste contrato, excedendo ${formatCurrency(limit)}.`,
            flagged
          );
        }
      }
      cumulative += contract.precoContratual;
    });
  });

  // Regra A3: Se a soma dos anos anteriores já ultrapassou o limite, sinalizar todos os contratos do ano n
  filtered
    .filter((contract) => contract.year === context.yearN)
    .forEach((contract) => {
      const key = normalizeText(contract.adjudicataria);
      const accumulated = pastYearSums.get(key) || 0;
      if (accumulated > limit && !flagged.has(contract.id)) {
        registerFinding(
          section,
          `${ruleCode}3`,
          contract,
          `${description}: soma nos anos n-2 e n-1 (${formatCurrency(
            accumulated
          )}) ultrapassou o limite antes do contrato em ${context.yearN}.`,
          flagged
        );
      }
    });
}

function auditContracts(contracts, context) {
  const sections = [
    {
      id: "A",
      title: "A - Contratos que excederam o limite do procedimento",
      findings: [],
    },
    {
      id: "B",
      title: "B - Contratos publicitados após o prazo de 20 dias úteis",
      findings: [],
    },
    {
      id: "C",
      title: "C - Contratos por \"Acordo Quadro\" que não mencionam o número do contrato",
      findings: [],
    },
    {
      id: "D",
      title: "D - Contração especializada excecionada",
      findings: [],
    },
    {
      id: "E",
      title: "E - Contração excluída",
      findings: [],
    },
    {
      id: "F",
      title: "F - Concursos públicos urgentes",
      findings: [],
    },
    {
      id: "G",
      title: "G - Contratos a solicitar a fiscalização prévia do Tribunal de Contas",
      findings: [],
    },
    {
      id: "H",
      title: "H - Contratação nos sectores da água, da energia, dos transportes e dos serviços postais",
      findings: [],
    },
    {
      id: "I",
      title: "I – Outras Contratações não enquadradas acima",
      findings: [],
    },
  ];
  const sectionMap = new Map(sections.map((section) => [section.id, section]));
  const flagged = new Set();

  const limitSection = sectionMap.get("A");

  createLimitRule(context, contracts, limitSection, flagged, {
    ruleCode: "A",
    limit: 20000,
    fundamentacao: FUNDAMENTACAO_ART_20_D,
    contractTypes: GOODS_AND_SERVICES,
    description: "Limite de 20.000 € (Artigo 20.º, n.º 1, alínea d))",
  });

  createLimitRule(context, contracts, limitSection, flagged, {
    ruleCode: "B",
    limit: 30000,
    fundamentacao: FUNDAMENTACAO_ART_19_D,
    contractTypes: PUBLIC_WORKS,
    description: "Limite de 30.000 € (Artigo 19.º, alínea d))",
  });

  createLimitRule(context, contracts, limitSection, flagged, {
    ruleCode: "C",
    limit: 75000,
    fundamentacao: FUNDAMENTACAO_ART_20_C,
    contractTypes: GOODS_AND_SERVICES,
    description: "Limite de 75.000 € (Artigo 20.º, n.º 1, alínea c))",
  });

  createLimitRule(context, contracts, limitSection, flagged, {
    ruleCode: "D",
    limit: 150000,
    fundamentacao: FUNDAMENTACAO_ART_19_C,
    contractTypes: PUBLIC_WORKS,
    description: "Limite de 150.000 € (Artigo 19.º, alínea c))",
  });

  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => normalizeText(contract.tipoProcedimento) === PROCEDIMENTO_ART_6A)
    .filter((contract) => contract.precoContratual > 750000)
    .forEach((contract) =>
      registerFinding(
        limitSection,
        "E",
        contract,
        "Contratação excluída pelo Artigo 6.º-A acima de 750.000 €.",
        flagged
      )
    );

  // Regra B: Contratos publicitados após o prazo de 20 dias úteis
  const sectionB = sectionMap.get("B");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .forEach((contract) => {
      const dataCelebracao = parseDate(contract.dataCelebracao);
      const dataPublicacao = parseDate(contract.dataPublicacao);
      
      if (dataCelebracao && dataPublicacao) {
        const workingDays = calculateWorkingDays(dataCelebracao, dataPublicacao);
        
        if (workingDays !== null && workingDays > 20) {
          registerFinding(
            sectionB,
            "B",
            contract,
            `Contrato publicitado após ${workingDays} dias úteis (limite: 20 dias úteis). Data de celebração: ${contract.dataCelebracao}, Data de publicação: ${contract.dataPublicacao}.`,
            flagged
          );
        }
      }
    });

  // Regra C: Contratos Acordo Quadro sem número
  const sectionC = sectionMap.get("C");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => matchesProcedimento(contract, PROCEDIMENTOS_ACORDO_H))
    .filter((contract) => !hasNumericValue(contract.numeroAcordoQuadro))
    .forEach((contract) =>
      registerFinding(
        sectionC,
        "C",
        contract,
        "Contrato de acordo-quadro sem número de registo válido.",
        flagged
      )
    );

  // Regra D: Contração especializada excecionada
  const sectionD = sectionMap.get("D");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => matchesProcedimento(contract, PROCEDIMENTOS_EXCECAO_D))
    .forEach((contract) =>
      registerFinding(
        sectionD,
        "D",
        contract,
        "Contratação especializada excecionada (Artigo 24.º ou 27.º do CCP).",
        flagged
      )
    );

  // Regra E: Contratação excluída (Artigo 5.º)
  const sectionE = sectionMap.get("E");
  const FUNDAMENTACAO_ART_5 = normalizeText("Artigo 5.º");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => {
      const fundamentacao = normalizeText(contract.fundamentacao);
      return fundamentacao.includes("artigo 5") || fundamentacao.includes("artigo 5.");
    })
    .forEach((contract) =>
      registerFinding(
        sectionE,
        "E",
        contract,
        "Contratação excluída ao abrigo do Artigo 5.º do CCP.",
        flagged
      )
    );

  // Regra F: Concursos públicos urgentes
  const sectionF = sectionMap.get("F");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => matchesProcedimento(contract, PROCEDIMENTOS_URGENTES_G))
    .forEach((contract) =>
      registerFinding(
        sectionF,
        "F",
        contract,
        "Concurso público urgente ao abrigo do Artigo 155.º.",
        flagged
      )
    );

  // Regra G: Tribunal de Contas
  const sectionG = sectionMap.get("G");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .forEach((contract) => {
      const types = extractContractTypes(contract);
      const typeSet = new Set(types);
      const isPublicWorks = typeSet.has(normalizeText("Empreitadas de obras públicas"));
      const isGoodsServices = [
        normalizeText("Aquisição de bens móveis"),
        normalizeText("Aquisição de serviços"),
        normalizeText("Locação de bens móveis"),
      ].some((value) => typeSet.has(value));
      const isConcession = [
        normalizeText("Concessão de obras públicas"),
        normalizeText("Concessão de serviços públicos"),
      ].some((value) => typeSet.has(value));

      if (
        (isPublicWorks && contract.precoContratual > 750000) ||
        (isGoodsServices && contract.precoContratual > 750000) ||
        isConcession
      ) {
        registerFinding(
          sectionG,
          "G",
          contract,
          "Contrato sujeito a fiscalização prévia do Tribunal de Contas (valor superior a 750.000€ ou concessão).",
          flagged
        );
      }
    });

  // Regra H: Sectores da água, energia, transportes e serviços postais
  const sectionH = sectionMap.get("H");
  const FUNDAMENTACAO_ART_11 = normalizeText("Artigo 11.º do Código dos Contratos Públicos");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => {
      const fundamentacao = normalizeText(contract.fundamentacao);
      // Verificar correspondência exata ou muito próxima do Artigo 11.º
      return fundamentacao === FUNDAMENTACAO_ART_11 || 
             fundamentacao === normalizeText("Artigo 11º do Código dos Contratos Públicos") ||
             fundamentacao === normalizeText("artigo 11.º do código dos contratos públicos");
    })
    .forEach((contract) =>
      registerFinding(
        sectionH,
        "H",
        contract,
        "Contratação nos sectores da água, da energia, dos transportes e dos serviços postais (Artigo 11.º do CCP).",
        flagged
      )
    );

  const knownFundamentacoes = new Set([
    FUNDAMENTACAO_ART_20_D,
    FUNDAMENTACAO_ART_19_D,
    FUNDAMENTACAO_ART_20_C,
    FUNDAMENTACAO_ART_19_C,
    ...PROCEDIMENTOS_EXCECAO_D,
    ...PROCEDIMENTOS_URGENTES_G,
    ...PROCEDIMENTOS_ACORDO_H,
    PROCEDIMENTO_ART_6A,
    FUNDAMENTACAO_ART_11,
  ]);
  
  // Adicionar artigos 5.º às fundamentações conhecidas para não aparecerem na secção I
  contracts
    .filter((contract) => contract.year === context.yearN)
    .forEach((contract) => {
      const fundamentacao = normalizeText(contract.fundamentacao);
      // Adicionar apenas se for Artigo 5.º (não outros artigos que contenham "5")
      if (fundamentacao.includes("artigo 5.º") || fundamentacao.includes("artigo 5º") || 
          (fundamentacao.includes("artigo 5") && !fundamentacao.includes("artigo 15") && !fundamentacao.includes("artigo 25") && !fundamentacao.includes("artigo 35") && !fundamentacao.includes("artigo 45") && !fundamentacao.includes("artigo 55") && !fundamentacao.includes("artigo 65") && !fundamentacao.includes("artigo 75") && !fundamentacao.includes("artigo 85") && !fundamentacao.includes("artigo 95"))) {
        knownFundamentacoes.add(fundamentacao);
      }
      // Adicionar Artigo 11.º se corresponder exatamente
      if (fundamentacao === FUNDAMENTACAO_ART_11 || 
          fundamentacao === normalizeText("Artigo 11º do Código dos Contratos Públicos") ||
          fundamentacao === normalizeText("artigo 11.º do código dos contratos públicos")) {
        knownFundamentacoes.add(fundamentacao);
      }
    });

  // Regra I: Outras contratações não enquadradas
  const sectionI = sectionMap.get("I");
  contracts
    .filter((contract) => contract.year === context.yearN)
    .filter((contract) => !flagged.has(contract.id))
    .filter((contract) => {
      const normalized = normalizeText(contract.fundamentacao);
      return normalized.length > 0 && !knownFundamentacoes.has(normalized);
    })
    .forEach((contract) =>
      registerFinding(
        sectionI,
        "I",
        contract,
        "Fundamentação não enquadrada nas regras A-H (outras contratações).",
        flagged
      )
    );

  sections.forEach((section) =>
    section.findings.sort((a, b) =>
      a.contract.objeto.localeCompare(b.contract.objeto, "pt")
    )
  );

  return sections;
}

async function generatePdf(sections, context, metadata) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const margin = 72; // Aumentado de 60 para 72 para mais espaço nas margens
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const { auditorName, adjudicantes, reportDate, totalContracts } = metadata;
  const adjudicanteMap = new Map();
  
  // Normalizar e deduplicar entidades (remover NIF, pontos, espaços extras, etc.)
  adjudicantes.forEach((value) => {
    // Remover NIF entre parênteses, normalizar espaços, pontos, vírgulas
    let normalized = value
      .toLowerCase()
      .replace(/\s*\([^)]*\)\s*/g, "") // Remove NIF e outros parênteses
      .replace(/\./g, "") // Remove pontos
      .replace(/,/g, "") // Remove vírgulas
      .replace(/\s+/g, " ") // Múltiplos espaços em um só
      .trim();
    
    // Se já existe uma entidade normalizada similar, usar a primeira
    let found = false;
    for (const [key, existing] of adjudicanteMap.entries()) {
      // Comparar sem considerar pequenas diferenças
      if (normalized === key || 
          (normalized.length > 10 && key.length > 10 && 
           (normalized.includes(key.substring(0, 15)) || key.includes(normalized.substring(0, 15))))) {
        found = true;
        break;
      }
    }
    
    if (!found) {
      adjudicanteMap.set(normalized, value);
    }
  });
  
  // Usar apenas a primeira entidade encontrada (ou todas se forem muito diferentes)
  const entidadesText = adjudicanteMap.size
    ? capitalizeWords(Array.from(adjudicanteMap.values())[0])
    : "Não identificado";

  doc.setFillColor("#eff4fb");
  doc.rect(0, 0, pageWidth, pageHeight, "F");

  doc.setFontSize(26);
  doc.setTextColor("#1a4d8f");
  doc.text(
    "Relatório de Auditoria à Contratação Pública",
    pageWidth / 2,
    margin + 40,
    { align: "center" }
  );

  // Adicionar nome da empresa após o título
  doc.setFontSize(14);
  doc.setTextColor("#1a4d8f");
  doc.text(
    "MRT Auditores",
    pageWidth / 2,
    margin + 75,
    { align: "center" }
  );

  doc.setFontSize(12);
  doc.setTextColor("#1f2a44");
  const infoStart = margin + 110;
  const infoLines = doc.splitTextToSize(
    `Entidade auditada: ${entidadesText}`,
    pageWidth - margin * 2
  );
  doc.text(infoLines, margin, infoStart);

  const auditorY = infoStart + infoLines.length * 16;
  doc.text(`Auditor responsável: ${capitalizeWords(auditorName)}`, margin, auditorY);

  // Informações dos anos e total de contratos após "Auditor responsável"
  let yearInfoY = auditorY + 24;
  doc.text(`Ano n: ${formatYearWithWords(context.yearN)}`, margin, yearInfoY);
  yearInfoY += 18;
  if (context.yearNMinus1) {
    doc.text(`Ano n-1: ${formatYearWithWords(context.yearNMinus1)}`, margin, yearInfoY);
    yearInfoY += 18;
  }
  if (context.yearNMinus2) {
    doc.text(`Ano n-2: ${formatYearWithWords(context.yearNMinus2)}`, margin, yearInfoY);
    yearInfoY += 18;
  }
  doc.text(`Total de contratos analisados: ${totalContracts}`, margin, yearInfoY);
  yearInfoY += 24;

  // Texto adicional após o número de contratos
  doc.setFontSize(9);
  doc.setTextColor("#1f2a44");
  
  // Primeiro parágrafo
  const para1 = "Este relatório procura dar resposta ao previsto na GAT 18 (Revisto) aplicável a \"Entidades que aplicam o SNC-AP\", em que, o ROC deve analisar, na contratação que entender relevante, os aspetos financeiros e orçamentais associados aos processos de contratação pública, nomeadamente:";
  const para1Lines = doc.splitTextToSize(para1, pageWidth - margin * 2);
  para1Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Item 1
  const item1 = "o Verificar se os limites dos procedimentos pré-contratuais para a realização da despesa, foram cumpridos pela entidade;";
  const item1Lines = doc.splitTextToSize(item1, pageWidth - margin * 2);
  item1Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Item 2
  const item2 = "o Verificar se a aquisição de bens e serviços ou empreitadas, executadas ou em curso, foram:";
  const item2Lines = doc.splitTextToSize(item2, pageWidth - margin * 2);
  item2Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Sub-item 1 com bullet
  const subitem1 = "• Objeto de fiscalização prévia do Tribunal de Contas, quando legalmente exigido;";
  const subitem1Lines = doc.splitTextToSize(subitem1, pageWidth - margin * 2 - 20);
  subitem1Lines.forEach((line, idx) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin + 20, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Sub-item 2 com bullet
  const subitem2 = "• Cumpridas as regras de publicitação dos contratos.";
  const subitem2Lines = doc.splitTextToSize(subitem2, pageWidth - margin * 2 - 20);
  subitem2Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin + 20, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 8;
  
  // Parágrafo de solicitação
  const para2 = "Sendo assim, solicitamos os esclarecimentos que tiverem por convenientes, quando aplicável, às situações identificadas nos títulos:";
  const para2Lines = doc.splitTextToSize(para2, pageWidth - margin * 2);
  para2Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Títulos A e B
  doc.setFont("helvetica", "bold");
  doc.text("A - Contratos que excederam o limite do procedimento", margin, yearInfoY);
  yearInfoY += 12;
  doc.text("B - Contratos publicitados após o prazo de 20 dias úteis", margin, yearInfoY);
  yearInfoY += 12;
  doc.setFont("helvetica", "normal");
  yearInfoY += 4;
  
  // Parágrafo adicional
  const para3 = "Solicitamos ainda, caso seja aplicável, cópia do visto do tribunal de Contas para os contratos identificados no título:";
  const para3Lines = doc.splitTextToSize(para3, pageWidth - margin * 2);
  para3Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });
  yearInfoY += 4;
  
  // Título G
  doc.setFont("helvetica", "bold");
  doc.text("G - Contratos a solicitar a fiscalização prévia do Tribunal de Contas", margin, yearInfoY);
  yearInfoY += 12;
  doc.setFont("helvetica", "normal");
  yearInfoY += 4;
  
  // Último parágrafo
  const para4 = "A resposta deverá ser enviada para o mail do auditor que remeteu o presente relatório.";
  const para4Lines = doc.splitTextToSize(para4, pageWidth - margin * 2);
  para4Lines.forEach((line) => {
    if (yearInfoY > pageHeight - margin - 50) return;
    doc.text(line, margin, yearInfoY);
    yearInfoY += 12;
  });

  doc.setFontSize(12);
  doc.text(`Data de emissão: ${reportDate}`, margin, yearInfoY + 12);

  doc.setFontSize(10);
  doc.setTextColor("#62729d");
  doc.text("Confidencial", margin, pageHeight - margin);

  doc.addPage();
  let cursorY = margin;
  doc.setFontSize(18);
  doc.setTextColor("#1a4d8f");
  doc.text("Resumo Executivo", margin, cursorY);
  cursorY += 28;

  // Fontes de informação e cálculos
  doc.setFontSize(14);
  doc.setTextColor("#1a4d8f");
  doc.text("Fontes de informação e cálculos", margin, cursorY);
  cursorY += 20;

  doc.setFontSize(10);
  doc.setTextColor("#1f2a44");
  const sourcesText = `Esta análise foi realizada com base no Código dos Contratos Públicos (CCP) e pela exportação dos contratos publicados no "Portal BASE" que centraliza a informação sobre os contratos públicos celebrados em Portugal continental e regiões autónomas.

Os cálculos foram efetuados por Entidade Adjudicatária individualmente, e considerou uma janela temporal de 3 anos (n-2, n-1, n).`;

  const sourcesLines = doc.splitTextToSize(sourcesText, pageWidth - margin * 2);
  sourcesLines.forEach((line) => {
    if (cursorY > pageHeight - margin) {
      doc.addPage();
      cursorY = margin;
      doc.setFontSize(10);
    }
    doc.text(line, margin, cursorY);
    cursorY += 14;
  });
  cursorY += 10;

  // Regras e procedimentos
  doc.setFontSize(14);
  doc.setTextColor("#1a4d8f");
  doc.text("Regras e procedimentos:", margin, cursorY);
  cursorY += 20;

  doc.setFontSize(10);
  doc.setTextColor("#1f2a44");
  const rulesText = `A - Contratos que excederam o limite do procedimento

• Artigo 20.º, n.º 1, alínea d) do CCP: Limite de 20.000€ para bens móveis, serviços ou locação

• Artigo 19.º, alínea d) do CCP: Limite de 30.000€ para empreitadas de obras públicas

• Artigo 20.º, n.º 1, alínea c) do CCP: Limite de 75.000€ para bens móveis, serviços ou locação

• Artigo 19.º, alínea c) do CCP: Limite de 150.000€ para empreitadas de obras públicas

• Artigo 6.º-A n.º 1 do CCP: Limite de 750.000€ para contratação excluída

B - Contratos publicitados após o prazo de 20 dias úteis

Conforme previsto no Artigo 8.º, alínea j) da Portaria n.º 318-B/2023, de 25 de outubro a entidade adjudicante tem até 20 dias úteis para submeter no Portal BASE o Relatório de formação do contrato (RFC) após a celebração do contrato escrito.

Caso o mesmo não tenha sido outorgado por escrito, 20 dias úteis após o início da sua execução, que pode ser entendido como a formalização, por parte do contraente público, de uma evidência de celebração do contrato, nomeadamente através de uma nota de encomenda, uma requisição, entre outras.

Na contagem dos 20 dias úteis não se conta o dia do evento (assinatura ou início de execução), e logo, também não se conta os sábados e domingos, nem também os feriados nacionais e locais aplicáveis à Entidade Adjudicatária.

C - Contratos por "Acordo Quadro" que não mencionam o número do contrato

Neste pretende-se identificar os contratos que foram enquadrados em "Contratos Acordo Quadro" e que não mencionam o respetivo número na coluna "N.º registo do Acordo Quadro". Posteriormente será analisado se houve uma simples omissão de preenchimento.

• Artigo 258.º, 259.º ou 252.º, n.º 1, alínea b) do CCP sem número de registo válido

D - Contração especializada excecionada (Escolha do ajuste directo para a formação de quaisquer contratos)

Neste ponto só se pretende identificar os contratos que foram enquadrados em "contratação especializada excecionada" e que permite a escolha do ajuste direto, somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.

• Artigo 24.º, n.º 1, alínea a) do CCP

• Artigo 24.º, n.º 1, alínea b) do CCP

• Artigo 24.º, n.º 1, alínea c) do CCP

• Artigo 24.º, n.º 1, alínea d) do CCP

• Artigo 24.º, n.º 1, alínea e), subalínea i), subalínea ii) ou subalínea iii) do CCP

• Artigo 27.º, n.º 1, alínea a) do CCP

• Artigo 27.º, n.º 1, alínea b) do CCP

• Artigo 27.º, n.º 1, alínea c) do CCP

• Artigo 27.º, n.º 1, alínea d) do CCP

• Artigo 27.º, n.º 1, alínea e) do CCP

• Artigo 27.º, n.º 1, alínea g) do CCP

• Artigo 27.º, n.º 1, alínea h) do CCP

• Artigo 27.º, n.º 1, alínea i) do CCP

E - Contração excluída

Neste ponto só se pretende identificar os contratos que foram enquadrados em "contratação excluída" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.

• Artigo 5.º, n.º 1 do CCP

• Artigo 5.º, n.º 4, alínea a) do CCP

• Artigo 5.º, n.º 4, alínea c) do CCP

• Artigo 5.º, n.º 4, alínea d) do CCP

• Artigo 5.º, n.º 4, alínea e) do CCP

• Artigo 5.º, n.º 4, alínea f) do CCP

• Artigo 5.º, n.º 4, alínea g) do CCP

• Artigo 5.º, n.º 4, alínea h) do CCP

• Artigo 5.º, n.º 4, alínea i) do CCP

• Artigo 5.º, n.º 4, alínea j) do CCP

F - Concursos públicos urgentes

Neste ponto só se pretende identificar os contratos que foram enquadrados em "concursos públicos urgentes" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.

• Artigo 155.º do CCP (com ou sem alíneas a) ou b))

G - Contratos a solicitar a fiscalização prévia do Tribunal de Contas

Nos termos da Lei de Organização e Processo do Tribunal de Contas (LOPTC) e no seu Artigo 48.º ficam sujeitos a fiscalização prévia os contratos de valor superior a 750.000 Euros, com exclusão do montante do imposto sobre o valor acrescentado que for devido.

O limite referido, quanto ao valor global dos atos e contratos que estejam ou aparentem estar relacionados entre si, é de 950.000 Euros.

Neste ponto só se pretende identificar os contratos que foram enquadrados acima do limite de 750.000 Euros, não sendo, por si só, uma identificação de qualquer inconformidade.

H - Contratação nos sectores da água, da energia, dos transportes e dos serviços postais

Neste ponto só se pretende identificar os contratos que foram enquadrados nos "sectores da água, da energia, dos transportes e dos serviços postais" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.

• Artigo 11.º do CCP

I – Outras Contratações não enquadradas acima

• Fundamentações, ou seja, os artigos do CCP não abrangidos pelas regras A-H.`;

  const rulesLines = doc.splitTextToSize(rulesText, pageWidth - margin * 2);
  rulesLines.forEach((line) => {
    if (cursorY > pageHeight - margin) {
      doc.addPage();
      cursorY = margin;
      doc.setFontSize(10);
    }
    doc.text(line, margin, cursorY);
    cursorY += 14;
  });

  sections.forEach((section) => {
    doc.addPage();
    let sectionCursorY = margin;
    doc.setFontSize(16);
    doc.setTextColor("#1a4d8f");
    // Quebrar título em múltiplas linhas se necessário
    const titleLines = doc.splitTextToSize(section.title, pageWidth - margin * 2);
    titleLines.forEach((line) => {
      doc.text(line, margin, sectionCursorY);
      sectionCursorY += 20;
    });
    sectionCursorY += 4;

    doc.setFontSize(10);
    doc.setTextColor("#1f2a44");
    if (!section.findings.length) {
      doc.text(
        "Não foram detetadas ocorrências para esta regra.",
        margin,
        sectionCursorY
      );
      return;
    }

    section.findings.forEach((finding, index) => {
      doc.setFontSize(12);
      const titleText = `${index + 1}. [${finding.rule}] ${finding.contract.objeto || "Sem descrição"}`;
      const titleLines = doc.splitTextToSize(titleText, pageWidth - margin * 2);
      titleLines.forEach((textLine) => {
        if (sectionCursorY > pageHeight - margin) {
          doc.addPage();
          sectionCursorY = margin;
          doc.setFontSize(12);
        }
        doc.text(textLine, margin, sectionCursorY);
        sectionCursorY += 16;
      });
      sectionCursorY += 4; // Espaçamento extra após o título

      doc.setFontSize(10);
      const details = describeContract(finding.contract).split("\n");
      details.forEach((line) => {
        const wrapped = doc.splitTextToSize(line, pageWidth - margin * 2);
        wrapped.forEach((textLine) => {
          if (sectionCursorY > pageHeight - margin) {
            doc.addPage();
            sectionCursorY = margin;
            doc.setFontSize(10);
          }
          doc.text(textLine, margin, sectionCursorY);
          sectionCursorY += 14;
        });
      });

      const noteLines = doc.splitTextToSize(
        `Nota: ${finding.details}`,
        pageWidth - margin * 2
      );
      noteLines.forEach((textLine) => {
        if (sectionCursorY > pageHeight - margin) {
          doc.addPage();
          sectionCursorY = margin;
          doc.setFontSize(10);
        }
        doc.text(textLine, margin, sectionCursorY);
        sectionCursorY += 14;
      });

      sectionCursorY += 10;
    });
  });

  // Secção de Cruzamento de Dados
  if (metadata.crossReference && metadata.crossReference.length > 0) {
    doc.addPage();
    let crossCursorY = margin;
    doc.setFontSize(16);
    doc.setTextColor("#1a4d8f");
    doc.text("Cruzamento de Dados", margin, crossCursorY);
    crossCursorY += 20;
    doc.setFontSize(12);
    doc.text("(E-faturas vs Extractos Contabilísticos vs Base.gov)", margin, crossCursorY);
    crossCursorY += 24;

    doc.setFontSize(10);
    doc.setTextColor("#1f2a44");
    doc.text(`Total de fornecedores analisados: ${metadata.crossReference.length}`, margin, crossCursorY);
    crossCursorY += 20;

    // Tabela de cruzamento
    metadata.crossReference.forEach((item, idx) => {
      if (crossCursorY > pageHeight - margin - 100) {
        doc.addPage();
        crossCursorY = margin;
      }

      doc.setFontSize(11);
      doc.setTextColor("#1a4d8f");
      doc.text(`${idx + 1}. ${item.nome || "Desconhecido"} (NIF: ${item.nif})`, margin, crossCursorY);
      crossCursorY += 16;

      doc.setFontSize(10);
      doc.setTextColor("#1f2a44");
      const details = [
        `E-faturas: ${formatCurrency(item.totalEfaturas)}`,
        `Base.gov (contratos): ${formatCurrency(item.totalContratos)}`,
        `Movimentos a crédito: ${formatCurrency(item.movimentosCredito)}`,
        `Movimentos a débito: ${formatCurrency(item.movimentosDebito)}`,
        `Saldo extractos: ${formatCurrency(item.saldoExtractos)}`,
        `Diferença: ${formatCurrency(item.diferenca)}`,
        `Notas de crédito: ${formatCurrency(item.notasCredito)}`,
      ];

      details.forEach((line) => {
        if (crossCursorY > pageHeight - margin) {
          doc.addPage();
          crossCursorY = margin;
        }
        doc.text(line, margin + 20, crossCursorY);
        crossCursorY += 14;
      });

      crossCursorY += 10;
    });
  }

  const filename = `relatorio-auditoria-${metadata.adjudicantes[0]?.replace(/[^a-zA-Z0-9]/g, "-") || "entidade"}-${metadata.auditorName.replace(/[^a-zA-Z0-9]/g, "-")}-${dayjs().format("YYYY-MM-DD")}.pdf`;
  doc.save(filename);
  
  // Guardar no histórico
    await saveAuditToHistory(sections, context, metadata);
}

async function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = reader.result.toString().replace(/\ufeff/g, "");
        const result = Papa.parse(text, {
          header: true,
          delimiter: ";",
          skipEmptyLines: true,
        });
        if (result.errors?.length) {
          return reject(new Error(`Erro ao ler ${file.name}: ${result.errors[0].message}`));
        }
        const contracts = result.data.map((row) => mapRowToContract(row, file.name));
        resolve(contracts);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error(`Erro ao ler o ficheiro ${file.name}.`));
    reader.readAsText(file, "utf-8");
  });
}

// Extrair NIF de uma string (pode estar entre parênteses ou separado)
function extractNIF(text) {
  if (!text) return null;
  // Procurar padrões como (123456789) ou NIF: 123456789
  const nifMatch = text.match(/(\d{9})/);
  return nifMatch ? nifMatch[1] : null;
}

// Normalizar NIF (remover espaços, pontos, etc.)
function normalizeNIF(nif) {
  if (!nif) return null;
  return String(nif).replace(/\D/g, "").padStart(9, "0");
}

// Processar e-faturas
async function parseEfaturas(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = reader.result.toString().replace(/\ufeff/g, "");
        const result = Papa.parse(text, {
          header: true,
          delimiter: ";",
          skipEmptyLines: true,
        });
        
        const efaturas = [];
        result.data.forEach((row) => {
          // Tentar encontrar NIF em várias colunas possíveis
          const nifCols = ["NIF", "NIF Fornecedor", "Contribuinte", "NIF Contribuinte", "NIF Emitente"];
          const valorCols = ["Valor", "Valor Total", "Montante", "Importe", "Valor da Fatura"];
          const dataCols = ["Data", "Data Emissão", "Data Fatura", "Data Documento"];
          const tipoCols = ["Tipo", "Tipo Documento", "Natureza"];
          
          let nif = null;
          for (const col of nifCols) {
            if (row[col]) {
              nif = normalizeNIF(extractNIF(row[col]) || row[col]);
              break;
            }
          }
          
          if (!nif) {
            // Tentar extrair de qualquer coluna
            for (const key in row) {
              const extracted = extractNIF(row[key]);
              if (extracted) {
                nif = normalizeNIF(extracted);
                break;
              }
            }
          }
          
          if (!nif) return; // Ignorar se não encontrar NIF
          
          let valor = 0;
          for (const col of valorCols) {
            if (row[col]) {
              valor = parsePrice(row[col]);
              break;
            }
          }
          
          const tipoCol = tipoCols.find((col) => row[col]);
          const tipo = tipoCol ? row[tipoCol] : "Fatura";
          const tipoLower = tipo ? tipo.toLowerCase() : "";
          const isCreditNote = tipoLower.includes("crédito") || tipoLower.includes("credito") || (tipoLower.includes("nota") && tipoLower.includes("crédito"));
          
          efaturas.push({
            nif,
            valor: isCreditNote ? -Math.abs(valor) : Math.abs(valor), // Notas de crédito negativas
            data: row[dataCols.find((col) => row[col])] || "",
            tipo,
            row,
          });
        });
        
        resolve(efaturas);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error(`Erro ao ler o ficheiro ${file.name}.`));
    reader.readAsText(file, "utf-8");
  });
}

// Processar extractos contabilísticos
async function parseExtractos(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = reader.result.toString().replace(/\ufeff/g, "");
        const result = Papa.parse(text, {
          header: true,
          delimiter: ";",
          skipEmptyLines: true,
        });
        
        const extractos = [];
        result.data.forEach((row) => {
          // Tentar encontrar NIF
          const nifCols = ["NIF", "Contribuinte", "NIF Fornecedor", "Entidade"];
          const valorCols = ["Valor", "Montante", "Importe", "Movimento"];
          const tipoCols = ["Tipo", "Natureza", "Movimento", "Débito/Crédito"];
          const dataCols = ["Data", "Data Movimento", "Data Operação"];
          
          let nif = null;
          for (const col of nifCols) {
            if (row[col]) {
              nif = normalizeNIF(extractNIF(row[col]) || row[col]);
              break;
            }
          }
          
          if (!nif) {
            for (const key in row) {
              const extracted = extractNIF(row[key]);
              if (extracted) {
                nif = normalizeNIF(extracted);
                break;
              }
            }
          }
          
          if (!nif) return;
          
          let valor = 0;
          for (const col of valorCols) {
            if (row[col]) {
              valor = parsePrice(row[col]);
              break;
            }
          }
          
          // Determinar se é crédito ou débito
          let isCredito = true;
          const tipoStr = tipoCols.find((col) => row[col]) ? row[tipoCols.find((col) => row[col])] : "";
          if (tipoStr) {
            const tipoLower = tipoStr.toLowerCase();
            isCredito = tipoLower.includes("crédito") || tipoLower.includes("credito") || tipoLower.includes("fatura");
          } else {
            // Se não houver coluna de tipo, assumir que valores positivos são crédito
            isCredito = valor >= 0;
          }
          
          extractos.push({
            nif,
            valor: Math.abs(valor),
            isCredito,
            data: row[dataCols.find((col) => row[col])] || "",
            row,
          });
        });
        
        resolve(extractos);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error(`Erro ao ler o ficheiro ${file.name}.`));
    reader.readAsText(file, "utf-8");
  });
}

// Processar cruzamento de dados
async function processCrossReference(contracts, fileEfaturas, fileExtractos) {
  const results = [];
  const nifMap = new Map(); // NIF -> { nome, contratos, efaturas, extractos }
  
  // Agrupar contratos por NIF do fornecedor
  contracts.forEach((contract) => {
    const nif = normalizeNIF(extractNIF(contract.adjudicataria));
    if (!nif) return;
    
    if (!nifMap.has(nif)) {
      nifMap.set(nif, {
        nif,
        nome: contract.adjudicataria || "Desconhecido",
        contratos: [],
        efaturas: [],
        extractos: [],
      });
    }
    
    nifMap.get(nif).contratos.push(contract);
  });
  
  // Processar e-faturas
  if (fileEfaturas) {
    try {
      const efaturas = await parseEfaturas(fileEfaturas);
      efaturas.forEach((efatura) => {
        if (!nifMap.has(efatura.nif)) {
          nifMap.set(efatura.nif, {
            nif: efatura.nif,
            nome: "Não encontrado no Base.gov",
            contratos: [],
            efaturas: [],
            extractos: [],
          });
        }
        nifMap.get(efatura.nif).efaturas.push(efatura);
      });
    } catch (error) {
      console.error("Erro ao processar e-faturas:", error);
    }
  }
  
  // Processar extractos
  if (fileExtractos) {
    try {
      const extractos = await parseExtractos(fileExtractos);
      extractos.forEach((extracto) => {
        if (!nifMap.has(extracto.nif)) {
          nifMap.set(extracto.nif, {
            nif: extracto.nif,
            nome: "Não encontrado no Base.gov",
            contratos: [],
            efaturas: [],
            extractos: [],
          });
        }
        nifMap.get(extracto.nif).extractos.push(extracto);
      });
    } catch (error) {
      console.error("Erro ao processar extractos:", error);
    }
  }
  
  // Calcular totais e diferenças
  nifMap.forEach((data) => {
    const totalEfaturas = data.efaturas.reduce((sum, e) => sum + (e.valor || 0), 0);
    const totalContratos = data.contratos.reduce((sum, c) => sum + (c.precoContratual || 0), 0);
    
    const movimentosCredito = data.extractos.filter((e) => e.isCredito).reduce((sum, e) => sum + (e.valor || 0), 0);
    const movimentosDebito = data.extractos.filter((e) => !e.isCredito).reduce((sum, e) => sum + (e.valor || 0), 0);
    const saldoExtractos = movimentosCredito - movimentosDebito;
    
    const diferenca = totalEfaturas - saldoExtractos;
    
    // Notas de crédito
    const notasCredito = data.efaturas.filter((e) => e.valor < 0).reduce((sum, e) => sum + Math.abs(e.valor), 0);
    
    if (data.contratos.length > 0 || data.efaturas.length > 0 || data.extractos.length > 0) {
      results.push({
        nif: data.nif,
        nome: data.nome,
        totalEfaturas,
        totalContratos,
        movimentosCredito,
        movimentosDebito,
        saldoExtractos,
        diferenca,
        notasCredito,
        numContratos: data.contratos.length,
        numEfaturas: data.efaturas.length,
        numExtractos: data.extractos.length,
      });
    }
  });
  
  // Ordenar por diferença absoluta (maiores diferenças primeiro)
  results.sort((a, b) => Math.abs(b.diferenca) - Math.abs(a.diferenca));
  
  return results;
}

async function handleGenerate() {
  try {
    setStatus("A processar ficheiros...");
    generateBtn.disabled = true;

    const auditorName = auditorInput.value.trim();
    if (!auditorName) {
      setStatus("Indique o nome do auditor responsável.", true);
      return;
    }

    const fileN = document.getElementById("file-n").files[0];
    const fileN1 = document.getElementById("file-n1").files[0];
    const fileN2 = document.getElementById("file-n2").files[0];

    if (!fileN || !fileN1 || !fileN2) {
      setStatus("Por favor selecione os três ficheiros CSV.", true);
      return;
    }

    const allContracts = [fileN, fileN1, fileN2];
    const contractsArrays = await Promise.all(allContracts.map(parseCsvFile));
    const contracts = contractsArrays.flat();

    if (!contracts.length) {
      setStatus("Não foi possível extrair contratos dos CSV fornecidos.", true);
      return;
    }

    const context = prepareContext(contracts);
    const sections = auditContracts(contracts, context);
    // Filtrar apenas entidades do ano n para a capa
    const adjudicantes = Array.from(
      new Set(
        contracts
          .filter((contract) => contract.year === context.yearN)
          .map((contract) => contract.adjudicante)
          .filter(Boolean)
      )
    );
    
    // Processar cruzamento de dados (opcional)
    let crossReferenceData = null;
    const fileEfaturas = document.getElementById("file-efaturas").files[0];
    const fileExtractos = document.getElementById("file-extractos").files[0];
    
    if (fileEfaturas || fileExtractos) {
      setStatus("A processar cruzamento de dados...");
      crossReferenceData = await processCrossReference(
        contracts,
        fileEfaturas,
        fileExtractos
      );
    }
    
    const metadata = {
      auditorName,
      adjudicantes,
      reportDate: dayjs().format("DD/MM/YYYY"),
      totalContracts: contracts.length,
      crossReference: crossReferenceData,
    };
    
    setStatus("Relatório gerado com sucesso. A descarregar...");
    await generatePdf(sections, context, metadata);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Ocorreu um erro ao gerar o relatório.", true);
  } finally {
    generateBtn.disabled = false;
  }
}

// Inicializar SharePoint quando a página carregar
let sharePointInitialized = false;
async function initializeSharePoint() {
  if (sharePointInitialized) return;
  
  try {
    setStatus("A conectar ao SharePoint...");
    const success = await window.SharePointIntegration?.initialize();
    if (success) {
      sharePointInitialized = true;
      setStatus("Conectado ao SharePoint", false);
      // Carregar histórico se já houver
      if (historyBtn) {
        loadAuditHistory();
      }
    } else {
      console.warn("SharePoint não inicializado. A usar localStorage como fallback.");
      setStatus("", false);
    }
  } catch (error) {
    console.error("Erro ao inicializar SharePoint:", error);
    setStatus("Erro ao conectar ao SharePoint. A usar modo offline.", true);
  }
}

// Inicializar quando o DOM estiver pronto
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initializeSharePoint);
} else {
  initializeSharePoint();
}

previewBtn.addEventListener("click", handlePreview);
generateBtn.addEventListener("click", handleGenerate);
analysisBtn.addEventListener("click", () => {
  dataAnalysis.classList.remove("hidden");
  dataAnalysis.scrollIntoView({ behavior: "smooth" });
});
closeAnalysis.addEventListener("click", () => {
  dataAnalysis.classList.add("hidden");
});
historyBtn.addEventListener("click", async () => {
  await loadAuditHistory();
  auditHistory.classList.remove("hidden");
  auditHistory.scrollIntoView({ behavior: "smooth" });
});
closeHistory.addEventListener("click", () => {
  auditHistory.classList.add("hidden");
});
closePreview.addEventListener("click", () => {
  reportPreview.classList.add("hidden");
  previewData = null;
});
cancelPreviewBtn.addEventListener("click", () => {
  reportPreview.classList.add("hidden");
  previewData = null;
});
confirmGenerateBtn.addEventListener("click", async () => {
  if (previewData) {
    await generatePdf(previewData.sections, previewData.context, previewData.metadata);
    reportPreview.classList.add("hidden");
    previewData = null;
  }
});

// ========== TABELAS INTERATIVAS ==========

function prepareTableData(sections) {
  tableData = [];
  const entitySet = new Set();
  
  sections.forEach((section) => {
    section.findings.forEach((finding) => {
      const contract = finding.contract;
      tableData.push({
        rule: finding.rule,
        sectionId: section.id,
        objeto: contract.objeto || "Sem descrição",
        adjudicataria: contract.adjudicataria || "Desconhecida",
        precoContratual: contract.precoContratual || 0,
        dataCelebracao: contract.dataCelebracao || "",
        tipoProcedimento: contract.tipoProcedimento || "",
        fundamentacao: contract.fundamentacao || "",
        details: finding.details || "",
        contract: contract,
      });
      
      if (contract.adjudicataria) {
        entitySet.add(contract.adjudicataria);
      }
    });
  });
  
  // Popular dropdown de entidades
  const entitySelect = document.getElementById("filter-entity");
  entitySelect.innerHTML = '<option value="">Todas</option>';
  Array.from(entitySet).sort().forEach((entity) => {
    const option = document.createElement("option");
    option.value = entity;
    option.textContent = entity;
    entitySelect.appendChild(option);
  });
  
  filteredData = [...tableData];
  renderTable();
}

function applyFilters() {
  const searchTerm = document.getElementById("search-input").value.toLowerCase();
  const entityFilter = document.getElementById("filter-entity").value;
  const sectionFilter = document.getElementById("filter-section").value;
  const minValue = parseFloat(document.getElementById("filter-min-value").value) || 0;
  
  filteredData = tableData.filter((row) => {
    // Pesquisa geral
    if (searchTerm) {
      const searchable = [
        row.objeto,
        row.adjudicataria,
        row.tipoProcedimento,
        row.fundamentacao,
        row.details,
      ].join(" ").toLowerCase();
      if (!searchable.includes(searchTerm)) return false;
    }
    
    // Filtro por entidade
    if (entityFilter && row.adjudicataria !== entityFilter) return false;
    
    // Filtro por secção
    if (sectionFilter && row.sectionId !== sectionFilter) return false;
    
    // Filtro por valor mínimo
    if (row.precoContratual < minValue) return false;
    
    return true;
  });
  
  currentPage = 1;
  applySorting();
  renderTable();
}

function applySorting() {
  if (!sortColumn) return;
  
  filteredData.sort((a, b) => {
    let aVal = a[sortColumn];
    let bVal = b[sortColumn];
    
    // Tratamento especial para diferentes tipos
    if (sortColumn === "precoContratual") {
      aVal = aVal || 0;
      bVal = bVal || 0;
    } else if (sortColumn === "dataCelebracao") {
      aVal = aVal || "";
      bVal = bVal || "";
    } else {
      aVal = (aVal || "").toString().toLowerCase();
      bVal = (bVal || "").toString().toLowerCase();
    }
    
    if (aVal < bVal) return sortDirection === "asc" ? -1 : 1;
    if (aVal > bVal) return sortDirection === "asc" ? 1 : -1;
    return 0;
  });
}

function sortTable(column) {
  if (sortColumn === column) {
    sortDirection = sortDirection === "asc" ? "desc" : "asc";
  } else {
    sortColumn = column;
    sortDirection = "asc";
  }
  
  // Atualizar indicadores visuais
  document.querySelectorAll("#data-table th").forEach((th) => {
    th.classList.remove("sorted-asc", "sorted-desc");
  });
  const header = document.querySelector(`#data-table th[data-sort="${column}"]`);
  if (header) {
    header.classList.add(`sorted-${sortDirection}`);
  }
  
  applySorting();
  renderTable();
}

function renderTable() {
  const tbody = document.getElementById("table-body");
  tbody.innerHTML = "";
  
  // Paginação
  const start = pageSize > 0 ? (currentPage - 1) * pageSize : 0;
  const end = pageSize > 0 ? start + pageSize : filteredData.length;
  const pageData = filteredData.slice(start, end);
  
  pageData.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>[${row.rule}]</strong></td>
      <td>${escapeHtml(row.objeto)}</td>
      <td>${escapeHtml(row.adjudicataria)}</td>
      <td>${formatPrice(row.precoContratual)}</td>
      <td>${row.dataCelebracao || "-"}</td>
      <td>${escapeHtml(row.tipoProcedimento)}</td>
      <td>${escapeHtml(row.details)}</td>
    `;
    tbody.appendChild(tr);
  });
  
  // Atualizar informações
  const total = filteredData.length;
  const showing = pageSize > 0 ? `${start + 1}-${Math.min(end, total)}` : total;
  document.getElementById("table-info").textContent = `A mostrar ${showing} de ${total} contrato(s)`;
  
  // Atualizar paginação
  const totalPages = pageSize > 0 ? Math.ceil(total / pageSize) : 1;
  document.getElementById("page-info").textContent = `Página ${currentPage} de ${totalPages}`;
  document.getElementById("prev-page").disabled = currentPage === 1;
  document.getElementById("next-page").disabled = currentPage >= totalPages || pageSize === 0;
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function exportToExcel() {
  if (filteredData.length === 0) {
    alert("Não há dados para exportar.");
    return;
  }
  
  // Criar CSV
  const headers = ["Regra", "Objeto", "Adjudicatária", "Preço Contratual", "Data Celebração", "Tipo Procedimento", "Fundamentação", "Detalhes"];
  const rows = filteredData.map((row) => [
    row.rule,
    row.objeto,
    row.adjudicataria,
    row.precoContratual || 0,
    row.dataCelebracao || "",
    row.tipoProcedimento || "",
    row.fundamentacao || "",
    row.details || "",
  ]);
  
  const csvContent = [
    headers.join(";"),
    ...rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(";")),
  ].join("\n");
  
  const blob = new Blob(["\ufeff" + csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `analise-contratos-${new Date().toISOString().split("T")[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// Event listeners para tabelas
document.getElementById("search-input").addEventListener("input", applyFilters);
document.getElementById("filter-entity").addEventListener("change", applyFilters);
document.getElementById("filter-section").addEventListener("change", applyFilters);
document.getElementById("filter-min-value").addEventListener("input", applyFilters);
document.getElementById("reset-filters-btn").addEventListener("click", () => {
  document.getElementById("search-input").value = "";
  document.getElementById("filter-entity").value = "";
  document.getElementById("filter-section").value = "";
  document.getElementById("filter-min-value").value = "";
  applyFilters();
});
document.getElementById("export-btn").addEventListener("click", exportToExcel);
document.getElementById("page-size").addEventListener("change", (e) => {
  pageSize = parseInt(e.target.value) || 0;
  currentPage = 1;
  renderTable();
});
document.getElementById("prev-page").addEventListener("click", () => {
  if (currentPage > 1) {
    currentPage--;
    renderTable();
  }
});
document.getElementById("next-page").addEventListener("click", () => {
  const totalPages = pageSize > 0 ? Math.ceil(filteredData.length / pageSize) : 1;
  if (currentPage < totalPages) {
    currentPage++;
    renderTable();
  }
});

// Ordenação por clique nos cabeçalhos
document.querySelectorAll("#data-table th[data-sort]").forEach((th) => {
  th.addEventListener("click", () => {
    sortTable(th.dataset.sort);
  });
});

// ========== HISTÓRICO DE AUDITORIAS ==========

async function saveAuditToHistory(sections, context, metadata) {
  try {
    // Tentar guardar no SharePoint primeiro
    if (sharePointInitialized && window.SharePointIntegration) {
      try {
        await window.SharePointIntegration.saveAudit(sections, context, metadata);
        return; // Sucesso, não precisa de fallback
      } catch (spError) {
        console.warn("Erro ao guardar no SharePoint, a usar localStorage:", spError);
        // Continuar para fallback localStorage
      }
    }
    
    // Fallback para localStorage
    const historyKey = `audit_history_${metadata.adjudicantes[0]?.replace(/[^a-zA-Z0-9]/g, "_") || "entidade"}_${metadata.auditorName.replace(/[^a-zA-Z0-9]/g, "_")}`;
    const history = JSON.parse(localStorage.getItem(historyKey) || "[]");
    
    const auditRecord = {
      id: Date.now().toString(),
      entity: metadata.adjudicantes[0] || "Não especificado",
      auditor: metadata.auditorName,
      date: metadata.reportDate,
      timestamp: Date.now(),
      totalContracts: metadata.totalContracts,
      totalFindings: Object.values(sections).reduce((sum, section) => sum + (section.findings?.length || 0), 0),
      years: {
        n: context.yearN,
        n1: context.yearNMinus1,
        n2: context.yearNMinus2,
      },
      sections: sections.map((section) => ({
        id: section.id,
        title: section.title,
        findingsCount: section.findings.length,
      })),
      data: {
        sections,
        context,
        metadata,
      },
    };
    
    history.unshift(auditRecord); // Adicionar no início
    // Manter apenas os últimos 50 registos
    if (history.length > 50) {
      history.splice(50);
    }
    
    localStorage.setItem(historyKey, JSON.stringify(history));
  } catch (error) {
    console.error("Erro ao guardar histórico:", error);
  }
}

async function loadAuditHistory() {
  try {
    let allHistory = [];
    
    // Tentar carregar do SharePoint primeiro
    if (sharePointInitialized && window.SharePointIntegration) {
      try {
        allHistory = await window.SharePointIntegration.loadHistory();
      } catch (spError) {
        console.warn("Erro ao carregar do SharePoint, a usar localStorage:", spError);
        // Continuar para fallback localStorage
      }
    }
    
    // Fallback para localStorage se SharePoint não funcionar ou não estiver inicializado
    if (allHistory.length === 0) {
      const allKeys = Object.keys(localStorage).filter((key) => key.startsWith("audit_history_"));
      allKeys.forEach((key) => {
        const history = JSON.parse(localStorage.getItem(key) || "[]");
        allHistory.push(...history);
      });
    }
    
    // Ordenar por data (mais recente primeiro)
    allHistory.sort((a, b) => b.timestamp - a.timestamp);
    
    if (allHistory.length === 0) {
      document.getElementById("history-list").innerHTML = "";
      document.getElementById("history-empty").style.display = "block";
      return;
    }
    
    document.getElementById("history-empty").style.display = "none";
    const listHtml = allHistory.map((audit) => {
      const yearsText = [
        audit.years.n2 ? audit.years.n2 : null,
        audit.years.n1 ? audit.years.n1 : null,
        audit.years.n,
      ]
        .filter(Boolean)
        .join(", ");
      
      return `
        <div class="history-item">
          <div class="history-item-header">
            <h3 class="history-item-title">${escapeHtml(audit.entity)}</h3>
            <span class="history-item-date">${audit.date}</span>
          </div>
          <div class="history-item-info">
            <div><strong>Auditor:</strong> ${escapeHtml(audit.auditor)}</div>
            <div><strong>Contratos analisados:</strong> ${audit.totalContracts}</div>
            <div><strong>Total de achados:</strong> ${audit.totalFindings}</div>
            <div><strong>Anos:</strong> ${yearsText}</div>
          </div>
          <div class="history-item-actions">
            <button class="history-btn-view" onclick="viewHistoryAudit('${audit.id}')">Ver Detalhes</button>
            <button class="history-btn-download" onclick="downloadHistoryAudit('${audit.id}')">Descarregar PDF</button>
            <button class="history-btn-delete" onclick="deleteHistoryAudit('${audit.id}')">Eliminar</button>
          </div>
        </div>
      `;
    }).join("");
    
    document.getElementById("history-list").innerHTML = listHtml;
  } catch (error) {
    console.error("Erro ao carregar histórico:", error);
    document.getElementById("history-list").innerHTML = "<p style='color: red;'>Erro ao carregar histórico.</p>";
  }
}

async function findAuditInHistory(id) {
  // Tentar encontrar no SharePoint primeiro
  if (sharePointInitialized && window.SharePointIntegration) {
    try {
      const audit = await window.SharePointIntegration.getAudit(id);
      if (audit) {
        return { audit, source: "sharepoint" };
      }
    } catch (spError) {
      console.warn("Erro ao buscar no SharePoint, a usar localStorage:", spError);
    }
  }
  
  // Fallback para localStorage
  const allKeys = Object.keys(localStorage).filter((key) => key.startsWith("audit_history_"));
  for (const key of allKeys) {
    const history = JSON.parse(localStorage.getItem(key) || "[]");
    const audit = history.find((a) => a.id === id);
    if (audit) return { audit, key, source: "localStorage" };
  }
  
  return null;
}

window.viewHistoryAudit = async function (id) {
  const result = await findAuditInHistory(id);
  if (!result) {
    alert("Auditoria não encontrada.");
    return;
  }
  
  const { sections, context, metadata } = result.audit.data;
  previewData = { sections, context, metadata };
  generatePreview(sections, context, metadata);
  reportPreview.classList.remove("hidden");
  auditHistory.classList.add("hidden");
  reportPreview.scrollIntoView({ behavior: "smooth" });
};

window.downloadHistoryAudit = async function (id) {
  const result = await findAuditInHistory(id);
  if (!result) {
    alert("Auditoria não encontrada.");
    return;
  }
  
  const { sections, context, metadata } = result.audit.data;
  await generatePdf(sections, context, metadata);
};

window.deleteHistoryAudit = async function (id) {
  if (!confirm("Tem a certeza que deseja eliminar esta auditoria do histórico?")) {
    return;
  }
  
  try {
    // Tentar eliminar do SharePoint primeiro
    if (sharePointInitialized && window.SharePointIntegration) {
      try {
        await window.SharePointIntegration.deleteAudit(id);
        await loadAuditHistory();
        return;
      } catch (spError) {
        console.warn("Erro ao eliminar do SharePoint, a usar localStorage:", spError);
      }
    }
    
    // Fallback para localStorage
    const result = await findAuditInHistory(id);
    if (!result || result.source !== "localStorage") {
      alert("Auditoria não encontrada.");
      return;
    }
    
    const { audit, key } = result;
    const history = JSON.parse(localStorage.getItem(key) || "[]");
    const filtered = history.filter((a) => a.id !== id);
    
    if (filtered.length === 0) {
      localStorage.removeItem(key);
    } else {
      localStorage.setItem(key, JSON.stringify(filtered));
    }
    
    await loadAuditHistory();
  } catch (error) {
    console.error("Erro ao eliminar auditoria:", error);
    alert("Erro ao eliminar auditoria.");
  }
};

async function handlePreview() {
  try {
    setStatus("A processar ficheiros...");
    previewBtn.disabled = true;

    const auditorName = auditorInput.value.trim();
    if (!auditorName) {
      setStatus("Indique o nome do auditor responsável.", true);
      return;
    }

    const fileN = document.getElementById("file-n").files[0];
    const fileN1 = document.getElementById("file-n1").files[0];
    const fileN2 = document.getElementById("file-n2").files[0];

    if (!fileN || !fileN1 || !fileN2) {
      setStatus("Por favor selecione os três ficheiros CSV.", true);
      return;
    }

    const allContracts = [fileN, fileN1, fileN2];
    const contractsArrays = await Promise.all(allContracts.map(parseCsvFile));
    const contracts = contractsArrays.flat();

    if (!contracts.length) {
      setStatus("Não foi possível extrair contratos dos CSV fornecidos.", true);
      return;
    }

    const context = prepareContext(contracts);
    const sections = auditContracts(contracts, context);
    // Filtrar apenas entidades do ano n para a capa
    const adjudicantes = Array.from(
      new Set(
        contracts
          .filter((contract) => contract.year === context.yearN)
          .map((contract) => contract.adjudicante)
          .filter(Boolean)
      )
    );
    
    // Processar cruzamento de dados (opcional)
    let crossReferenceData = null;
    const fileEfaturas = document.getElementById("file-efaturas").files[0];
    const fileExtractos = document.getElementById("file-extractos").files[0];
    
    if (fileEfaturas || fileExtractos) {
      setStatus("A processar cruzamento de dados...");
      crossReferenceData = await processCrossReference(
        contracts,
        fileEfaturas,
        fileExtractos
      );
    }
    
    const metadata = {
      auditorName,
      adjudicantes,
      reportDate: dayjs().format("DD/MM/YYYY"),
      totalContracts: contracts.length,
      crossReference: crossReferenceData,
    };

    // Guardar dados para gerar PDF depois
    previewData = { sections, context, metadata };

    // Preparar dados para tabela
    prepareTableData(sections);
    
    // Mostrar botão de análise
    analysisBtn.style.display = "inline-block";

    // Gerar preview HTML
    generatePreview(sections, context, metadata);
    
    setStatus("Preview gerado com sucesso.");
    reportPreview.classList.remove("hidden");
    reportPreview.scrollIntoView({ behavior: "smooth" });
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Ocorreu um erro ao gerar o preview.", true);
  } finally {
    previewBtn.disabled = false;
  }
}

function generatePreview(sections, context, metadata) {
  let html = "";

  // Capa - usar apenas a primeira entidade (normalizada e deduplicada)
  const adjudicanteMap = new Map();
  metadata.adjudicantes.forEach((value) => {
    // Remover NIF entre parênteses, normalizar espaços, pontos, vírgulas
    let normalized = value
      .toLowerCase()
      .replace(/\s*\([^)]*\)\s*/g, "") // Remove NIF e outros parênteses
      .replace(/\./g, "") // Remove pontos
      .replace(/,/g, "") // Remove vírgulas
      .replace(/\s+/g, " ") // Múltiplos espaços em um só
      .trim();
    
    // Se já existe uma entidade normalizada similar, usar a primeira
    let found = false;
    for (const [key] of adjudicanteMap.entries()) {
      // Comparar sem considerar pequenas diferenças
      if (normalized === key || 
          (normalized.length > 10 && key.length > 10 && 
           (normalized.includes(key.substring(0, 15)) || key.includes(normalized.substring(0, 15))))) {
        found = true;
        break;
      }
    }
    
    if (!found) {
      adjudicanteMap.set(normalized, value);
    }
  });
  
  // Usar apenas a primeira entidade encontrada
  const entidadesText = adjudicanteMap.size
    ? capitalizeWords(Array.from(adjudicanteMap.values())[0])
    : "Não identificado";
  
  html += `<div class="preview-cover">
    <h1>Relatório de Auditoria à Contratação Pública</h1>
    <p style="text-align: center; font-size: 1.1rem; color: #1a4d8f; margin-top: 0.5rem; margin-bottom: 1.5rem;"><strong>MRT Auditores</strong></p>
    <div class="preview-cover-info">
      <p><strong>Entidade auditada:</strong> ${entidadesText}</p>
      <p><strong>Auditor responsável:</strong> ${metadata.auditorName}</p>
      <p><strong>Ano n:</strong> ${formatYearWithWords(context.yearN)}</p>
      ${context.yearNMinus1 ? `<p><strong>Ano n-1:</strong> ${formatYearWithWords(context.yearNMinus1)}</p>` : ""}
      ${context.yearNMinus2 ? `<p><strong>Ano n-2:</strong> ${formatYearWithWords(context.yearNMinus2)}</p>` : ""}
      <p><strong>Total de contratos analisados:</strong> ${metadata.totalContracts}</p>
      <div style="margin-top: 1rem; font-size: 0.9rem; line-height: 1.6; color: #1f2a44;">
        <p>Este relatório procura dar resposta ao previsto na GAT 18 (Revisto) aplicável a "Entidades que aplicam o SNC-AP", em que, o ROC deve analisar, na contratação que entender relevante, os aspetos financeiros e orçamentais associados aos processos de contratação pública, nomeadamente:</p>
        <p>o Verificar se os limites dos procedimentos pré-contratuais para a realização da despesa, foram cumpridos pela entidade;</p>
        <p>o Verificar se a aquisição de bens e serviços ou empreitadas, executadas ou em curso, foram:</p>
        <ul style="margin-left: 1.5rem;">
          <li>Objeto de fiscalização prévia do Tribunal de Contas, quando legalmente exigido;</li>
          <li>Cumpridas as regras de publicitação dos contratos.</li>
        </ul>
        <p>Sendo assim, solicitamos os esclarecimentos que tiverem por convenientes, quando aplicável, às situações identificadas nos títulos:</p>
        <p><strong>A - Contratos que excederam o limite do procedimento</strong></p>
        <p><strong>B - Contratos publicitados após o prazo de 20 dias úteis</strong></p>
        <p>Solicitamos ainda, caso seja aplicável, cópia do visto do tribunal de Contas para os contratos identificados no título:</p>
        <p><strong>G - Contratos a solicitar a fiscalização prévia do Tribunal de Contas</strong></p>
        <p>A resposta deverá ser enviada para o mail do auditor que remeteu o presente relatório.</p>
      </div>
      <p style="margin-top: 1rem;"><strong>Data de emissão:</strong> ${metadata.reportDate}</p>
    </div>
  </div>`;

  // Resumo Executivo
  const totalFindings = sections.reduce((sum, section) => sum + (section.findings?.length || 0), 0);
  
  html += `<div class="preview-section">
    <h3>Resumo Executivo</h3>
    <p><strong>Total de achados:</strong> ${totalFindings}</p>
    
    <div style="margin-top: 1.5rem; padding: 1.5rem; background: #f9fbff; border-radius: 8px; border-left: 3px solid #667eea;">
      <h4 style="margin-top: 0; color: #1a4d8f; font-size: 1.1rem;">Fontes de informação e cálculos</h4>
      <div style="color: #555; line-height: 1.8; font-size: 0.95rem; margin-bottom: 1.5rem;">
        <p>Esta análise foi realizada com base no Código dos Contratos Públicos (CCP) e pela exportação dos contratos publicados no "Portal BASE" que centraliza a informação sobre os contratos públicos celebrados em Portugal continental e regiões autónomas.</p>
        <p>Os cálculos foram efetuados por Entidade Adjudicatária individualmente, e considerou uma janela temporal de 3 anos (n-2, n-1, n).</p>
      </div>
      
      <h4 style="margin-top: 1.5rem; color: #1a4d8f; font-size: 1.1rem;">Regras e procedimentos:</h4>
      <div style="color: #555; line-height: 1.8; font-size: 0.95rem;">
        <p><strong>A - Contratos que excederam o limite do procedimento</strong></p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 20.º, n.º 1, alínea d) do CCP: Limite de 20.000€ para bens móveis, serviços ou locação</li>
          <li>Artigo 19.º, alínea d) do CCP: Limite de 30.000€ para empreitadas de obras públicas</li>
          <li>Artigo 20.º, n.º 1, alínea c) do CCP: Limite de 75.000€ para bens móveis, serviços ou locação</li>
          <li>Artigo 19.º, alínea c) do CCP: Limite de 150.000€ para empreitadas de obras públicas</li>
          <li>Artigo 6.º-A n.º 1 do CCP: Limite de 750.000€ para contratação excluída</li>
        </ul>
        
        <p><strong>B - Contratos publicitados após o prazo de 20 dias úteis</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Conforme previsto no Artigo 8.º, alínea j) da Portaria n.º 318-B/2023, de 25 de outubro a entidade adjudicante tem até 20 dias úteis para submeter no Portal BASE o Relatório de formação do contrato (RFC) após a celebração do contrato escrito.</p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Caso o mesmo não tenha sido outorgado por escrito, 20 dias úteis após o início da sua execução, que pode ser entendido como a formalização, por parte do contraente público, de uma evidência de celebração do contrato, nomeadamente através de uma nota de encomenda, uma requisição, entre outras.</p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Na contagem dos 20 dias úteis não se conta o dia do evento (assinatura ou início de execução), e logo, também não se conta os sábados e domingos, nem também os feriados nacionais e locais aplicáveis à Entidade Adjudicatária.</p>
        
        <p><strong>C - Contratos por "Acordo Quadro" que não mencionam o número do contrato</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste pretende-se identificar os contratos que foram enquadrados em "Contratos Acordo Quadro" e que não mencionam o respetivo número na coluna "N.º registo do Acordo Quadro". Posteriormente será analisado se houve uma simples omissão de preenchimento.</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 258.º, 259.º ou 252.º, n.º 1, alínea b) do CCP sem número de registo válido</li>
        </ul>
        
        <p><strong>D - Contração especializada excecionada (Escolha do ajuste directo para a formação de quaisquer contratos)</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste ponto só se pretende identificar os contratos que foram enquadrados em "contratação especializada excecionada" e que permite a escolha do ajuste direto, somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 24.º, n.º 1, alínea a) do CCP</li>
          <li>Artigo 24.º, n.º 1, alínea b) do CCP</li>
          <li>Artigo 24.º, n.º 1, alínea c) do CCP</li>
          <li>Artigo 24.º, n.º 1, alínea d) do CCP</li>
          <li>Artigo 24.º, n.º 1, alínea e), subalínea i), subalínea ii) ou subalínea iii) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea a) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea b) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea c) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea d) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea e) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea g) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea h) do CCP</li>
          <li>Artigo 27.º, n.º 1, alínea i) do CCP</li>
        </ul>
        
        <p><strong>E - Contração excluída</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste ponto só se pretende identificar os contratos que foram enquadrados em "contratação excluída" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 5.º, n.º 1 do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea a) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea c) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea d) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea e) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea f) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea g) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea h) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea i) do CCP</li>
          <li>Artigo 5.º, n.º 4, alínea j) do CCP</li>
        </ul>
        
        <p><strong>F - Concursos públicos urgentes</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste ponto só se pretende identificar os contratos que foram enquadrados em "concursos públicos urgentes" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 155.º do CCP (com ou sem alíneas a) ou b))</li>
        </ul>
        
        <p><strong>G - Contratos a solicitar a fiscalização prévia do Tribunal de Contas</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Nos termos da Lei de Organização e Processo do Tribunal de Contas (LOPTC) e no seu Artigo 48.º ficam sujeitos a fiscalização prévia os contratos de valor superior a 750.000 Euros, com exclusão do montante do imposto sobre o valor acrescentado que for devido.</p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">O limite referido, quanto ao valor global dos atos e contratos que estejam ou aparentem estar relacionados entre si, é de 950.000 Euros.</p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste ponto só se pretende identificar os contratos que foram enquadrados acima do limite de 750.000 Euros, não sendo, por si só, uma identificação de qualquer inconformidade.</p>
        
        <p><strong>H - Contratação nos sectores da água, da energia, dos transportes e dos serviços postais</strong></p>
        <p style="margin-left: 1.5rem; margin-bottom: 1rem;">Neste ponto só se pretende identificar os contratos que foram enquadrados nos "sectores da água, da energia, dos transportes e dos serviços postais" somente para sua identificação, não sendo, por si só, uma identificação de qualquer inconformidade.</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Artigo 11.º do CCP</li>
        </ul>
        
        <p><strong>I – Outras Contratações não enquadradas acima</strong></p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
          <li>Fundamentações, ou seja, os artigos do CCP não abrangidos pelas regras A-H.</li>
        </ul>
      </div>
    </div>
  </div>`;

  // Secções do relatório (sections é um array)
  sections.forEach((section) => {
    const findings = section.findings || [];
    html += `<div class="preview-section ${findings.length === 0 ? "empty" : ""}">
      <h3>${section.title}</h3>`;
    
    if (findings.length === 0) {
      html += `<p>Não foram detetadas ocorrências para esta regra.</p>`;
    } else {
      html += `<p><strong>Total:</strong> ${findings.length} contrato(s)</p>`;
      findings.forEach((finding, idx) => {
        const contract = finding.contract;
        html += `<div class="preview-contract">
          <p><strong>Contrato ${idx + 1} [${finding.rule}]:</strong></p>
          <p><strong>Objeto:</strong> ${contract.objeto || "Sem descrição"}</p>
          <p><strong>Entidade Adjudicatária:</strong> ${contract.adjudicataria || "Desconhecida"}</p>
          <p><strong>Preço Contratual:</strong> ${formatPrice(contract.precoContratual)}</p>
          <p><strong>Data de Celebração:</strong> ${contract.dataCelebracao || "Não indicada"}</p>
          <p><strong>Tipo de Procedimento:</strong> ${contract.tipoProcedimento || "Não especificado"}</p>
          ${contract.fundamentacao ? `<p><strong>Fundamentação:</strong> ${contract.fundamentacao}</p>` : ""}
          ${finding.details ? `<p><strong>Detalhes:</strong> ${finding.details}</p>` : ""}
        </div>`;
      });
    }
    
    html += `</div>`;
  });

  previewContent.innerHTML = html;
}

function formatPrice(value) {
  if (!value) return "Não especificado";
  const num = typeof value === "string" ? parseFloat(value.replace(/[^\d,.-]/g, "").replace(",", ".")) : value;
  if (isNaN(num)) return value;
  return new Intl.NumberFormat("pt-PT", { style: "currency", currency: "EUR" }).format(num);
}

// ========== CHATBOT LOCAL ==========

// Base de conhecimento sobre auditoria à contratação pública
const KNOWLEDGE_BASE = {
  "limites": {
    keywords: ["limite", "valor máximo", "montante", "até quanto", "quanto pode", "valor permitido", "limites", "quais os limites"],
    content: `**Limites Legais para Contratação Pública em Portugal**

Os limites legais variam conforme o tipo de contrato e o procedimento utilizado:

**Ajuste Direto Regime Geral:**
- **Artigo 20.º, n.º 1, alínea d) do CCP:** 20.000€ para bens móveis, serviços ou locação
- **Artigo 19.º, alínea d) do CCP:** 30.000€ para empreitadas de obras públicas
- **Artigo 20.º, n.º 1, alínea c) do CCP:** 75.000€ para bens móveis, serviços ou locação
- **Artigo 19.º, alínea c) do CCP:** 150.000€ para empreitadas de obras públicas
- **Artigo 6.º-A n.º 1 do CCP:** 750.000€ (contratação excluída)

**Aprovação Tribunal de Contas:**
- Empreitadas de obras públicas: acima de 950.000€
- Bens móveis/serviços/locação: acima de 950.000€
- Concessões: qualquer valor

**Importante:** Estes limites aplicam-se por Entidade Adjudicatária individualmente, não a nível global.`
  },
  "artigo 20 d": {
    keywords: ["artigo 20", "alínea d", "20.000", "vinte mil"],
    content: `**Artigo 20.º, n.º 1, alínea d) do Código dos Contratos Públicos**

Aplica-se a: Aquisição de bens móveis, Aquisição de serviços, ou Locação de bens móveis

**Limite:** 20.000€ por entidade adjudicatária

**Regras de verificação:**
1. Nenhum contrato individual pode exceder 20.000€
2. A soma dos contratos no mesmo ano (n) não pode exceder 20.000€ antes de celebrar novo contrato
3. A soma dos contratos dos anos n-2 e n-1 não pode exceder 20.000€ se houver contrato no ano n

A soma é calculada por cada Entidade Adjudicatária individualmente.`
  },
  "artigo 19 d": {
    keywords: ["artigo 19", "empreitadas", "30.000", "trinta mil", "obras públicas"],
    content: `**Artigo 19.º, alínea d) do Código dos Contratos Públicos**

Aplica-se a: Empreitadas de obras públicas

**Limite:** 30.000€ por entidade adjudicatária

**Regras de verificação:**
1. Nenhum contrato individual pode exceder 30.000€
2. A soma dos contratos no mesmo ano (n) não pode exceder 30.000€ antes de celebrar novo contrato
3. A soma dos contratos dos anos n-2 e n-1 não pode exceder 30.000€ se houver contrato no ano n

A soma é calculada por cada Entidade Adjudicatária individualmente.`
  },
  "artigo 20 c": {
    keywords: ["artigo 20", "alínea c", "75.000", "setenta e cinco mil"],
    content: `**Artigo 20.º, n.º 1, alínea c) do Código dos Contratos Públicos**

Aplica-se a: Aquisição de bens móveis, Aquisição de serviços, ou Locação de bens móveis

**Limite:** 75.000€ por entidade adjudicatária

**Regras de verificação:**
1. Nenhum contrato individual pode exceder 75.000€
2. A soma dos contratos no mesmo ano (n) não pode exceder 75.000€ antes de celebrar novo contrato
3. A soma dos contratos dos anos n-2 e n-1 não pode exceder 75.000€ se houver contrato no ano n

A soma é calculada por cada Entidade Adjudicatária individualmente.`
  },
  "artigo 19 c": {
    keywords: ["artigo 19", "alínea c", "150.000", "cento e cinquenta mil"],
    content: `**Artigo 19.º, alínea c) do Código dos Contratos Públicos**

Aplica-se a: Empreitadas de obras públicas

**Limite:** 150.000€ por entidade adjudicatária

**Regras de verificação:**
1. Nenhum contrato individual pode exceder 150.000€
2. A soma dos contratos no mesmo ano (n) não pode exceder 150.000€ antes de celebrar novo contrato
3. A soma dos contratos dos anos n-2 e n-1 não pode exceder 150.000€ se houver contrato no ano n

A soma é calculada por cada Entidade Adjudicatária individualmente.`
  },
  "artigo 6a": {
    keywords: ["artigo 6", "6-A", "750.000", "setecentos e cinquenta mil", "contratação excluída"],
    content: `**Artigo 6.º-A n.º 1 do Código dos Contratos Públicos**

Aplica-se a: Contratação excluída

**Limite:** 750.000€

**Verificação:** Para o ano "n", quando o Tipo de Procedimento seja "artigo 6.º-A n.º 1 do Código dos Contratos Públicos", nenhum contrato pode ter Preço Contratual superior a 750.000€.`
  },
  "contratação especializada": {
    keywords: ["especializada", "excecionada", "artigo 24", "subalínea"],
    content: `**Contração especializada excecionada**

Identifica-se quando o Tipo de Procedimento contém:
- "Artigo 24.º, n.º 1, alínea e), subalínea i) do Código dos Contratos Públicos"
- "Artigo 24.º, n.º 1, alínea e), subalínea ii) do Código dos Contratos Públicos"

Todos os contratos que cumpram esta condição no ano "n" devem ser identificados no relatório sob o título "Contração especializada excecionada".`
  },
  "concursos urgentes": {
    keywords: ["urgente", "artigo 155", "concurso público urgente", "concursos urgentes", "concurso urgente", "155", "urgentes"],
    content: `**Concursos Públicos Urgentes**

Os concursos públicos urgentes são procedimentos de contratação pública que podem ser realizados de forma acelerada devido a circunstâncias excecionais.

**Identificação:**
Identifica-se quando o Tipo de Procedimento contém:
- "Artigo 155.º do Código dos Contratos Públicos"
- "Artigo 155.º, alínea a) do Código dos Contratos Públicos"
- "Artigo 155.º, alínea b) do Código dos Contratos Públicos"

**No relatório:**
Todos os contratos que cumpram esta condição no ano "n" devem ser identificados no relatório sob o título "Concursos Públicos Urgentes".

Estes contratos são objeto de análise especial na auditoria, pois a urgência pode justificar procedimentos simplificados, mas deve ser verificada a fundamentação adequada.`
  },
  "acordo quadro": {
    keywords: ["acordo quadro", "artigo 258", "artigo 259", "artigo 252", "número registo"],
    content: `**Contratos Acordo Quadro que não mencionam o número do contrato**

Identifica-se quando:
- Tipo de Procedimento contém: "Artigo 258.º", "Artigo 259.º" ou "Artigo 252.º, n.º 1, alínea b) do Código dos Contratos Públicos"
- E a coluna "N.º registo do Acordo Quadro" está vazia, nula ou contém texto não numérico

Todos os contratos que cumpram estas condições no ano "n" devem ser listados no relatório.`
  },
  "tribunal contas": {
    keywords: ["tribunal de contas", "aprovação", "950.000", "novecentos e cinquenta mil", "concessão"],
    content: `**Contratos a solicitar aprovação do Tribunal de Contas**

Aplica-se quando o Preço Contratual excede os limites ou quando se trata de:

1. **Empreitadas de obras públicas:** Preço Contratual > 950.000€
2. **Bens móveis/serviços/locação:** Preço Contratual > 950.000€
3. **Concessões:** Qualquer valor (empreitadas de obras públicas ou serviços públicos)

A verificação aplica-se apenas aos contratos do ano "n" (mais recente).`
  },
  "exceções revisão": {
    keywords: ["exceções", "revisão", "fundamentação", "não abrangida"],
    content: `**Contratação enquadrada em exceções que merecem revisão**

Identifica-se quando a coluna "Fundamentação" contém um valor que não está abrangido pelas regras A-H (Contratos que excederam limite, Contratação especializada, Concursos urgentes, Acordo Quadro, Tribunal de Contas).

Esta regra funciona como um filtro residual (catch-all) e exclui os contratos já classificados nas outras categorias para evitar duplicações.`
  },
  "normalização": {
    keywords: ["normalização", "comparação", "espaços", "maiúsculas", "minúsculas"],
    content: `**Normalização de dados**

Para evitar falhas causadas por variações mínimas de escrita nos ficheiros CSV, a verificação do valor da coluna "Tipo de Procedimento" deve ser feita após aplicar uma normalização:

1. **Trim de espaços:** Remover espaços em branco no início e no fim do texto
2. **Remoção de espaços duplos:** Substituir ocorrências de múltiplos espaços por apenas um espaço
3. **Comparação case-insensitive:** Considerar iguais maiúsculas e minúsculas (ex.: "Artigo" = "artigo")

Só depois desta normalização é que deve ser feita a comparação exata com os textos previstos nas instruções.`
  },
  "soma contratos": {
    keywords: ["soma", "acumular", "janela", "3 anos", "n-2", "n-1", "n"],
    content: `**Cálculo de somas de contratos**

A soma do Preço Contratual deve ser efetuada:
- Por cada Entidade Adjudicatária individualmente (não por adjudicante nem a nível global)
- Considerando apenas contratos do mesmo tipo (bens móveis, serviços, locação, ou empreitadas)
- Com o mesmo Tipo de Procedimento e Fundamentação

**Janela de 3 anos:**
- n-2: ano mais antigo
- n-1: ano intermédio  
- n: ano mais recente

A soma dos anos n-2 + n-1 não pode exceder o limite se houver contrato no ano n.`
  },
  "ajuste direto": {
    keywords: ["ajuste direto", "ajuste", "direto", "regime geral"],
    content: `**Ajuste Direto Regime Geral**

O ajuste direto é um procedimento de contratação pública que permite a celebração direta de contratos sem prévia publicitação, dentro dos limites legais estabelecidos.

**Limites principais:**
- Artigo 20.º, n.º 1, alínea d): 20.000€ (bens móveis, serviços, locação)
- Artigo 20.º, n.º 1, alínea c): 75.000€ (bens móveis, serviços, locação)
- Artigo 19.º, alínea d): 30.000€ (empreitadas de obras públicas)
- Artigo 19.º, alínea c): 150.000€ (empreitadas de obras públicas)

Estes limites são calculados por Entidade Adjudicatária e consideram uma janela de 3 anos (n-2, n-1, n).`
  },
  "código contratos públicos": {
    keywords: ["código", "contratos públicos", "ccp", "lei", "legislação"],
    content: `**Código dos Contratos Públicos (CCP)**

O Código dos Contratos Públicos é a legislação que regula a contratação pública em Portugal. Estabelece os procedimentos, limites e regras para a celebração de contratos públicos.

**Principais artigos relevantes para auditoria:**
- Artigo 19.º: Empreitadas de obras públicas (limites 30.000€ e 150.000€)
- Artigo 20.º: Bens móveis, serviços e locação (limites 20.000€ e 75.000€)
- Artigo 24.º: Contratação especializada excecionada
- Artigo 155.º: Concursos públicos urgentes
- Artigo 252.º, 258.º, 259.º: Acordos-quadro
- Artigo 6.º-A: Contratação excluída (limite 750.000€)

A auditoria verifica o cumprimento destas regras e identifica contratos que não estão em conformidade.`
  }
};

// Função para normalizar texto (igual à usada na auditoria)
function normalizeTextForSearch(text) {
  return text
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // Remove acentos
}

// Função para encontrar resposta na base de conhecimento
function findLocalAnswer(question) {
  const normalizedQuestion = normalizeTextForSearch(question);
  const questionWords = normalizedQuestion.split(/\s+/).filter(w => w.length > 2); // Palavras com mais de 2 letras
  let bestMatch = null;
  let maxScore = 0;

  for (const [key, entry] of Object.entries(KNOWLEDGE_BASE)) {
    let score = 0;
    
    // Verificar keywords completas
    for (const keyword of entry.keywords) {
      const normalizedKeyword = normalizeTextForSearch(keyword);
      if (normalizedQuestion.includes(normalizedKeyword)) {
        score += keyword.length * 2; // Keywords completas têm mais peso
      }
    }
    
    // Verificar palavras individuais das keywords
    for (const keyword of entry.keywords) {
      const keywordWords = normalizeTextForSearch(keyword).split(/\s+/);
      for (const kw of keywordWords) {
        if (kw.length > 2 && questionWords.some(qw => qw.includes(kw) || kw.includes(qw))) {
          score += kw.length;
        }
      }
    }
    
    // Verificar se palavras da pergunta aparecem no conteúdo
    const normalizedContent = normalizeTextForSearch(entry.content);
    for (const qw of questionWords) {
      if (normalizedContent.includes(qw)) {
        score += qw.length;
      }
    }
    
    if (score > maxScore) {
      maxScore = score;
      bestMatch = entry;
    }
  }

  // Se encontrou alguma correspondência (mesmo que baixa), retornar
  if (maxScore > 0 && bestMatch) {
    // Se o score for muito baixo, adicionar nota
    if (maxScore < 5) {
      return `${bestMatch.content}\n\n*Nota: Esta resposta pode não ser totalmente específica para a sua pergunta. Pode reformular ou ser mais específico.*`;
    }
    return bestMatch.content;
  }

  return null;
}

// Função para pesquisar na web (usando DuckDuckGo Instant Answer)
async function searchWeb(query) {
  try {
    // DuckDuckGo Instant Answer API (gratuita, sem API key)
    const searchQuery = encodeURIComponent(`Código dos Contratos Públicos Portugal ${query}`);
    const response = await fetch(
      `https://api.duckduckgo.com/?q=${searchQuery}&format=json&no_html=1&skip_disambig=1`,
      { mode: "cors" }
    );
    
    if (response.ok) {
      const data = await response.json();
      if (data.AbstractText) {
        return data.AbstractText;
      }
      if (data.Answer) {
        return data.Answer;
      }
      if (data.RelatedTopics && data.RelatedTopics.length > 0) {
        return data.RelatedTopics[0].Text || data.RelatedTopics[0].FirstURL;
      }
    }
  } catch (error) {
    console.log("Erro na pesquisa web:", error);
  }
  
  return null;
}

const chatbotToggle = document.getElementById("chatbot-toggle");
const chatbotWindow = document.getElementById("chatbot-window");
const chatbotClose = document.getElementById("chatbot-close");
const chatbotMessages = document.getElementById("chatbot-messages");
const chatbotInput = document.getElementById("chatbot-input");
const chatbotSend = document.getElementById("chatbot-send");
// Esconder campo de API key (não é mais necessário)
const apiKeyContainer = document.getElementById("api-key-container");
if (apiKeyContainer) {
  apiKeyContainer.style.display = "none";
}

// Abrir/fechar chat
let chatInitialized = false;
chatbotToggle.addEventListener("click", () => {
  chatbotWindow.classList.toggle("hidden");
  if (!chatbotWindow.classList.contains("hidden")) {
    chatbotInput.focus();
    const history = loadChatHistory();
    if (!chatInitialized && history.length === 0) {
      addMessage(
        "assistant",
        "Olá! Sou o assistente de auditoria. Como posso ajudar? Pode fazer perguntas sobre as regras de auditoria, limites legais, ou interpretação de resultados. Se não encontrar a resposta na minha base de conhecimento, posso pesquisar na internet."
      );
      chatInitialized = true;
    }
  }
});

chatbotClose.addEventListener("click", () => {
  chatbotWindow.classList.add("hidden");
});

// Enviar mensagem com Enter
chatbotInput.addEventListener("keypress", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    sendMessage();
  }
});

chatbotSend.addEventListener("click", sendMessage);

async function sendMessage() {
  const message = chatbotInput.value.trim();
  if (!message) return;

  // Adicionar mensagem do utilizador
  addMessage("user", message);
  chatbotInput.value = "";
  chatbotSend.disabled = true;

  // Adicionar indicador de carregamento
  const loadingId = addMessage("assistant", "A procurar resposta...", "loading");

  try {
    // Primeiro, tentar encontrar resposta na base de conhecimento local
    let answer = findLocalAnswer(message);

    // Se não encontrou, pesquisar na web
    if (!answer) {
      removeMessage(loadingId);
      const searchingId = addMessage("assistant", "Não encontrei na base de conhecimento. A pesquisar na internet...", "loading");
      
      const webResult = await searchWeb(message);
      
      if (webResult) {
        removeMessage(searchingId);
        answer = `Encontrei esta informação na internet:\n\n${webResult}\n\n*Nota: Esta informação foi obtida através de pesquisa online e pode precisar de verificação.*`;
      } else {
        removeMessage(searchingId);
        answer = `Não consegui encontrar uma resposta específica para a sua pergunta na minha base de conhecimento nem através de pesquisa online.\n\nPode reformular a pergunta ou tentar perguntar sobre:\n- Limites legais de contratação\n- Artigos específicos do Código dos Contratos Públicos\n- Regras de auditoria\n- Procedimentos de contratação`;
      }
    } else {
      removeMessage(loadingId);
    }

    // Adicionar resposta
    addMessage("assistant", answer);
    saveChatHistory();
  } catch (error) {
    removeMessage(loadingId);
    addMessage(
      "assistant",
      `Ocorreu um erro ao processar a sua pergunta: ${error.message}. Por favor, tente novamente.`
    );
  } finally {
    chatbotSend.disabled = false;
    chatbotInput.focus();
  }
}

function addMessage(role, content, className = null) {
  const messageDiv = document.createElement("div");
  messageDiv.className = `chatbot-message ${role} ${className || ""}`;
  messageDiv.textContent = content;
  messageDiv.dataset.messageId = Date.now().toString();
  chatbotMessages.appendChild(messageDiv);
  chatbotMessages.scrollTop = chatbotMessages.scrollHeight;
  return messageDiv.dataset.messageId;
}

function removeMessage(messageId) {
  const message = chatbotMessages.querySelector(`[data-message-id="${messageId}"]`);
  if (message) {
    message.remove();
  }
}

function getChatHistory() {
  const history = JSON.parse(localStorage.getItem("chatbot_history") || "[]");
  return history.slice(-10); // Últimas 10 mensagens para contexto
}

function saveChatHistory() {
  const messages = Array.from(chatbotMessages.querySelectorAll(".chatbot-message:not(.loading)"));
  const history = messages.map((msg) => ({
    role: msg.classList.contains("user") ? "user" : "assistant",
    content: msg.textContent,
  }));
  localStorage.setItem("chatbot_history", JSON.stringify(history));
}

function loadChatHistory() {
  const history = getChatHistory();
  chatbotMessages.innerHTML = "";
  history.forEach((msg) => {
    addMessage(msg.role, msg.content);
  });
  return history;
}
