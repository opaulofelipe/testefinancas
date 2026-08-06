(function () {
  "use strict";

  const MONTHS = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro"
  ];

  const MONTH_SHORT = [
    "Jan",
    "Fev",
    "Mar",
    "Abr",
    "Mai",
    "Jun",
    "Jul",
    "Ago",
    "Set",
    "Out",
    "Nov",
    "Dez"
  ];

  const DEFAULT_FILE_URLS = [
    "./Financas.xlsx",
    "./financas.xlsx",
    "./Finanças.xlsx",
    "./Financ\u0327as.xlsx",
    "./FINANCAS.xlsx",
    "./upload/Financ\u0327as.xlsx"
  ];

  const HEADER_ALIASES = {
    month: ["mes", "month", "periodo", "competencia"],
    year: ["ano", "year"],
    total: ["totalgasto", "gastototal", "totaldespesa", "despesastotais"],
    balance: ["balanca", "saldo", "balance", "resultado"],
    income: [
      "renda",
      "quintino",
      "receita",
      "receitas",
      "entrada",
      "entradas",
      "salario",
      "salarios"
    ]
  };

  const GROUPS = {
    essential: {
      label: "Base da casa",
      color: "#ffc575",
      aliases: [
        "aluguel",
        "condominio",
        "luz",
        "energia",
        "gas",
        "agua",
        "internet",
        "telefone",
        "moradia",
        "financiamento",
        "prestacao",
        "escola",
        "saude",
        "plano"
      ]
    },
    flexible: {
      label: "Flexível",
      color: "#67dbe5",
      aliases: [
        "cartao",
        "comida",
        "lazer",
        "restaurante",
        "delivery",
        "mercado",
        "compras",
        "viagem",
        "entretenimento"
      ]
    },
    other: {
      label: "Outros",
      color: "#9b8cff",
      aliases: []
    }
  };

  const CATEGORY_COLORS = [
    "#67dbe5",
    "#9b8cff",
    "#ffc575",
    "#ff86ad",
    "#62dba8",
    "#7da8ff",
    "#d58cff",
    "#ff9b6c"
  ];

  const money = new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });

  const percentage = new Intl.NumberFormat("pt-BR", {
    style: "percent",
    minimumFractionDigits: 0,
    maximumFractionDigits: 1
  });

  const state = {
    rows: [],
    categories: [],
    categoryGroups: {},
    incomeColumns: [],
    filteredRows: [],
    sourceName: "",
    charts: {},
    filters: {
      years: new Set(),
      months: new Set(),
      status: "all",
      view: "all"
    }
  };

  const elements = {};
  const byId = (id) => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheElements();
    bindEvents();
    configureCharts();
    renderEmptyDashboard();
    loadDefaultWorkbook();
  }

  function cacheElements() {
    [
      "statusDot",
      "sourceName",
      "sourceMeta",
      "openFileButton",
      "fileInput",
      "periodLabel",
      "updatedAt",
      "heroTitle",
      "heroDescription",
      "heroNumberLabel",
      "heroNumber",
      "heroNumberFoot",
      "resultCount",
      "clearFilters",
      "yearChips",
      "monthChips",
      "statusChips",
      "viewChips",
      "incomeAverage",
      "incomeFoot",
      "spendAverage",
      "spendFoot",
      "balanceAverage",
      "balanceFoot",
      "positiveRate",
      "positiveFoot",
      "cashflowNote",
      "spendCommitment",
      "essentialCommitment",
      "flexibleCommitment",
      "spendProgress",
      "essentialProgress",
      "flexibleProgress",
      "healthSummary",
      "categoryLegend",
      "insightsGrid",
      "exportButton",
      "categoryTableBody",
      "tableCount",
      "monthlyTableBody",
      "uploadZone",
      "methodNote",
      "emptyState",
      "emptyUploadButton",
      "emptyError",
      "toast"
    ].forEach((id) => {
      elements[id] = byId(id);
    });
  }

  function bindEvents() {
    elements.openFileButton.addEventListener("click", openFilePicker);
    elements.emptyUploadButton.addEventListener("click", openFilePicker);
    elements.uploadZone.addEventListener("click", openFilePicker);

    elements.uploadZone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openFilePicker();
      }
    });

    elements.fileInput.addEventListener("change", (event) => {
      const file = event.target.files && event.target.files[0];

      if (file) {
        loadLocalFile(file);
      }

      event.target.value = "";
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      elements.uploadZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.uploadZone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      elements.uploadZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.uploadZone.classList.remove("is-dragging");
      });
    });

    elements.uploadZone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files && event.dataTransfer.files[0];

      if (file) {
        loadLocalFile(file);
      }
    });

    elements.yearChips.addEventListener("click", handleFilterClick);
    elements.monthChips.addEventListener("click", handleFilterClick);
    elements.statusChips.addEventListener("click", handleFilterClick);
    elements.viewChips.addEventListener("click", handleFilterClick);

    elements.clearFilters.addEventListener("click", () => {
      state.filters.years.clear();
      state.filters.months.clear();
      state.filters.status = "all";
      state.filters.view = "all";
      renderFilters();
      renderDashboard();
      showToast("Filtros limpos.");
    });

    elements.exportButton.addEventListener("click", exportCurrentView);
  }

  function openFilePicker() {
    elements.fileInput.click();
  }

  async function loadDefaultWorkbook() {
    if (!window.XLSX) {
      showLoadError(
        "A biblioteca de leitura não carregou. Verifique sua conexão com a internet."
      );
      return;
    }

    setSourceStatus(
      "loading",
      "Procurando planilha…",
      "Buscando Financas.xlsx na raiz do projeto."
    );

    const failures = [];

    for (const url of DEFAULT_FILE_URLS) {
      try {
        const response = await fetch(url, { cache: "no-store" });

        if (!response.ok) {
          failures.push(`${url}: ${response.status}`);
          continue;
        }

        const buffer = await response.arrayBuffer();
        const fileName = decodeURIComponent(url.split("/").pop());
        parseWorkbook(buffer, fileName);
        return;
      } catch (error) {
        failures.push(`${url}: ${error.message}`);
      }
    }

    console.info("Planilha padrão não encontrada:", failures);
    showLoadError("");
  }

  async function loadLocalFile(file) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      showToast("Escolha um arquivo XLSX, XLS ou CSV.", true);
      return;
    }

    if (!window.XLSX) {
      showToast("A biblioteca XLSX não está disponível.", true);
      return;
    }

    try {
      setSourceStatus("loading", file.name, "Lendo a primeira aba…");
      const buffer = await file.arrayBuffer();
      parseWorkbook(buffer, file.name);
      showToast("Planilha carregada com sucesso.");
    } catch (error) {
      console.error(error);
      setSourceStatus("error", "Erro ao ler a tabela", error.message);
      elements.emptyState.hidden = false;
      elements.emptyError.textContent = error.message;
      showToast(error.message, true);
    }
  }

  function parseWorkbook(buffer, sourceName) {
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("A planilha não possui nenhuma aba.");
    }

    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    const rawRows = XLSX.utils.sheet_to_json(firstSheet, {
      defval: null,
      raw: true
    });

    const normalized = normalizeWorkbookRows(rawRows);

    state.rows = normalized.rows;
    state.categories = normalized.categories;
    state.categoryGroups = normalized.categoryGroups;
    state.incomeColumns = normalized.incomeColumns;
    state.sourceName = sourceName;

    state.filters.years.clear();
    state.filters.months.clear();
    state.filters.status = "all";
    state.filters.view = "all";

    elements.emptyState.hidden = true;
    elements.emptyError.textContent = "";

    setSourceStatus(
      "ready",
      sourceName,
      `${state.rows.length} meses · ${state.categories.length} categorias · aba ${workbook.SheetNames[0]}`
    );

    elements.updatedAt.textContent = new Intl.DateTimeFormat("pt-BR", {
      hour: "2-digit",
      minute: "2-digit"
    }).format(new Date());

    elements.methodNote.textContent =
      `Entradas = ${state.incomeColumns.map(labelize).join(" + ")} · ` +
      "Gastos = soma das categorias · Saldo = entradas − gastos.";

    renderFilters();
    renderDashboard();
  }

  function normalizeWorkbookRows(rawRows) {
    if (!Array.isArray(rawRows) || !rawRows.length) {
      throw new Error("A primeira aba está vazia.");
    }

    const headers = [
      ...new Set(
        rawRows.flatMap((row) => Object.keys(row || {}))
      )
    ];

    const headerInfo = headers.map((header) => ({
      original: header,
      normalized: normalizeKey(header)
    }));

    const monthColumn = findColumn(headerInfo, HEADER_ALIASES.month);
    const yearColumn = findColumn(headerInfo, HEADER_ALIASES.year);
    const totalColumn = findColumn(headerInfo, HEADER_ALIASES.total);
    const balanceColumn = findColumn(headerInfo, HEADER_ALIASES.balance);

    if (!monthColumn || !yearColumn) {
      throw new Error("Não encontrei as colunas Mês e Ano.");
    }

    const incomeColumns = headerInfo
      .filter(({ normalized }) => {
        return HEADER_ALIASES.income.some((alias) => {
          const normalizedAlias = normalizeKey(alias);
          return normalized === normalizedAlias || normalized.includes(normalizedAlias);
        });
      })
      .map(({ original }) => original);

    if (!incomeColumns.length) {
      throw new Error("Não encontrei colunas de entrada, como Renda ou Quintino.");
    }

    const excludedColumns = new Set([
      monthColumn,
      yearColumn,
      totalColumn,
      balanceColumn,
      ...incomeColumns
    ].filter(Boolean));

    const categories = headerInfo
      .filter(({ original }) => {
        if (excludedColumns.has(original)) {
          return false;
        }

        return rawRows.some((row) => isNumeric(row[original]));
      })
      .map(({ original }) => original);

    if (!categories.length) {
      throw new Error("Não encontrei colunas numéricas de despesas.");
    }

    const categoryGroups = Object.fromEntries(
      categories.map((category) => [category, inferCategoryGroup(category)])
    );

    const rows = rawRows
      .map((rawRow, rowIndex) => {
        const monthIndex = parseMonth(rawRow[monthColumn]);
        const year = Math.round(toNumber(rawRow[yearColumn]));

        if (monthIndex < 0 || year < 1900) {
          return null;
        }

        const categoryValues = {};

        categories.forEach((category) => {
          categoryValues[category] = Math.max(0, toNumber(rawRow[category]));
        });

        const calculatedSpend = categories.reduce((total, category) => {
          return total + categoryValues[category];
        }, 0);

        const informedSpend = totalColumn && isNumeric(rawRow[totalColumn])
          ? toNumber(rawRow[totalColumn])
          : calculatedSpend;

        const spend = informedSpend >= 0 ? informedSpend : calculatedSpend;

        const income = incomeColumns.reduce((total, column) => {
          return total + toNumber(rawRow[column]);
        }, 0);

        const informedBalance = balanceColumn && isNumeric(rawRow[balanceColumn])
          ? toNumber(rawRow[balanceColumn])
          : income - spend;

        const groups = {
          essential: 0,
          flexible: 0,
          other: 0
        };

        categories.forEach((category) => {
          const group = categoryGroups[category];
          groups[group] += categoryValues[category];
        });

        return {
          id: `${year}-${String(monthIndex + 1).padStart(2, "0")}-${rowIndex}`,
          year,
          monthIndex,
          month: MONTHS[monthIndex],
          date: new Date(year, monthIndex, 1),
          income,
          spend,
          balance: informedBalance,
          categories: categoryValues,
          groups
        };
      })
      .filter(Boolean)
      .sort((first, second) => first.date - second.date);

    if (!rows.length) {
      throw new Error("Nenhuma linha válida com Mês e Ano foi encontrada.");
    }

    return {
      rows,
      categories,
      categoryGroups,
      incomeColumns
    };
  }

  function renderFilters() {
    const availableYears = uniqueSorted(state.rows.map((row) => row.year));
    const availableMonths = uniqueSorted(state.rows.map((row) => row.monthIndex));

    elements.yearChips.innerHTML = [
      filterButton("Todos", "years", "all", state.filters.years.size === 0),
      ...availableYears.map((year) => {
        return filterButton(
          String(year),
          "years",
          String(year),
          state.filters.years.has(year)
        );
      })
    ].join("");

    elements.monthChips.innerHTML = [
      filterButton("Todos", "months", "all", state.filters.months.size === 0),
      ...availableMonths.map((monthIndex) => {
        return filterButton(
          MONTH_SHORT[monthIndex],
          "months",
          String(monthIndex),
          state.filters.months.has(monthIndex)
        );
      })
    ].join("");

    elements.statusChips.innerHTML = [
      ["all", "Todos"],
      ["positive", "Com sobra"],
      ["negative", "No vermelho"]
    ].map(([value, label]) => {
      return segmentButton(
        label,
        "status",
        value,
        state.filters.status === value
      );
    }).join("");

    elements.viewChips.innerHTML = [
      ["all", "Tudo"],
      ["essential", "Base da casa"],
      ["flexible", "Flexível"]
    ].map(([value, label]) => {
      return segmentButton(
        label,
        "view",
        value,
        state.filters.view === value
      );
    }).join("");
  }

  function filterButton(label, group, value, isActive) {
    return `
      <button
        class="filter-chip${isActive ? " is-active" : ""}"
        type="button"
        data-filter-group="${group}"
        data-filter-value="${value}"
        aria-pressed="${isActive}"
      >${escapeHtml(label)}</button>
    `;
  }

  function segmentButton(label, group, value, isActive) {
    return `
      <button
        class="segment-chip${isActive ? " is-active" : ""}"
        type="button"
        data-filter-group="${group}"
        data-filter-value="${value}"
        aria-pressed="${isActive}"
      >${escapeHtml(label)}</button>
    `;
  }

  function handleFilterClick(event) {
    const button = event.target.closest("button[data-filter-group]");

    if (!button) {
      return;
    }

    const group = button.dataset.filterGroup;
    const value = button.dataset.filterValue;

    if (group === "years" || group === "months") {
      const targetSet = state.filters[group];

      if (value === "all") {
        targetSet.clear();
      } else {
        const numericValue = Number(value);

        if (targetSet.has(numericValue)) {
          targetSet.delete(numericValue);
        } else {
          targetSet.add(numericValue);
        }
      }
    }

    if (group === "status") {
      state.filters.status = value;
    }

    if (group === "view") {
      state.filters.view = value;
    }

    renderFilters();
    renderDashboard();
  }

  function getFilteredRows() {
    return state.rows.filter((row) => {
      const yearMatches =
        state.filters.years.size === 0 ||
        state.filters.years.has(row.year);

      const monthMatches =
        state.filters.months.size === 0 ||
        state.filters.months.has(row.monthIndex);

      const statusMatches =
        state.filters.status === "all" ||
        (state.filters.status === "positive"
          ? row.balance >= 0
          : row.balance < 0);

      return yearMatches && monthMatches && statusMatches;
    });
  }

  function calculateMetrics(rows) {
    const count = rows.length;
    const totalIncome = sum(rows, (row) => row.income);
    const totalSpend = sum(rows, (row) => row.spend);
    const totalBalance = sum(rows, (row) => row.balance);
    const averageIncome = count ? totalIncome / count : 0;
    const averageSpend = count ? totalSpend / count : 0;
    const averageBalance = count ? totalBalance / count : 0;
    const medianIncome = median(rows.map((row) => row.income));

    const groupTotals = {
      essential: sum(rows, (row) => row.groups.essential),
      flexible: sum(rows, (row) => row.groups.flexible),
      other: sum(rows, (row) => row.groups.other)
    };

    const groupAverages = {
      essential: count ? groupTotals.essential / count : 0,
      flexible: count ? groupTotals.flexible / count : 0,
      other: count ? groupTotals.other / count : 0
    };

    const positiveMonths = rows.filter((row) => row.balance >= 0).length;
    const incomeStandardDeviation = standardDeviation(rows.map((row) => row.income));
    const incomeVolatility = averageIncome
      ? incomeStandardDeviation / averageIncome
      : 0;

    const categoryStats = state.categories.map((category) => {
      const total = sum(rows, (row) => row.categories[category] || 0);
      const recentRows = rows.slice(-Math.min(3, count));
      const previousRows = rows.slice(
        -Math.min(6, count),
        -Math.min(3, count)
      );
      const recentAverage = average(recentRows, (row) => row.categories[category] || 0);
      const previousAverage = average(previousRows, (row) => row.categories[category] || 0);
      const trend = previousAverage
        ? (recentAverage - previousAverage) / previousAverage
        : 0;

      return {
        name: category,
        group: state.categoryGroups[category],
        total,
        average: count ? total / count : 0,
        share: totalSpend ? total / totalSpend : 0,
        trend
      };
    }).sort((first, second) => second.total - first.total);

    const recentWindow = rows.slice(-Math.min(3, count));
    const previousWindow = rows.slice(
      -Math.min(6, count),
      -Math.min(3, count)
    );
    const recentSpendAverage = average(recentWindow, (row) => row.spend);
    const previousSpendAverage = average(previousWindow, (row) => row.spend);
    const recentSpendChange = previousSpendAverage
      ? (recentSpendAverage - previousSpendAverage) / previousSpendAverage
      : 0;

    const bestMonth = count
      ? rows.reduce((best, row) => row.balance > best.balance ? row : best, rows[0])
      : null;

    const worstMonth = count
      ? rows.reduce((worst, row) => row.balance < worst.balance ? row : worst, rows[0])
      : null;

    const annual = {};

    rows.forEach((row) => {
      if (!annual[row.year]) {
        annual[row.year] = {
          count: 0,
          income: 0,
          spend: 0,
          balance: 0
        };
      }

      annual[row.year].count += 1;
      annual[row.year].income += row.income;
      annual[row.year].spend += row.spend;
      annual[row.year].balance += row.balance;
    });

    Object.values(annual).forEach((yearData) => {
      yearData.income /= yearData.count;
      yearData.spend /= yearData.count;
      yearData.balance /= yearData.count;
    });

    const extraordinaryIncomeMonths = rows.filter((row) => {
      return medianIncome > 0 && row.income > medianIncome * 1.35;
    }).length;

    return {
      count,
      totalIncome,
      totalSpend,
      totalBalance,
      averageIncome,
      averageSpend,
      averageBalance,
      medianIncome,
      recurringCapacity: medianIncome - averageSpend,
      groupTotals,
      groupAverages,
      positiveMonths,
      positiveRate: count ? positiveMonths / count : 0,
      spendCommitment: averageIncome ? averageSpend / averageIncome : 0,
      essentialCommitment: averageIncome ? groupAverages.essential / averageIncome : 0,
      flexibleCommitment: averageIncome ? groupAverages.flexible / averageIncome : 0,
      incomeVolatility,
      extraordinaryIncomeMonths,
      categoryStats,
      recentSpendAverage,
      previousSpendAverage,
      recentSpendChange,
      bestMonth,
      worstMonth,
      annual
    };
  }

  function renderDashboard() {
    const rows = getFilteredRows();
    const metrics = calculateMetrics(rows);

    state.filteredRows = rows;
    elements.resultCount.textContent = `${rows.length} ${rows.length === 1 ? "mês" : "meses"}`;
    elements.periodLabel.textContent = getPeriodLabel(rows);

    if (!rows.length) {
      renderNoMatches();
      return;
    }

    renderHero(metrics);
    renderKpis(metrics);
    renderHealth(metrics);
    renderCharts(rows, metrics);
    renderInsights(metrics);
    renderCategoryTable(metrics);
    renderMonthlyTable(rows);
  }

  function renderHero(metrics) {
    const hasPositiveAverage = metrics.averageBalance >= 0;
    const recurringIsPositive = metrics.recurringCapacity >= 0;

    if (hasPositiveAverage && !recurringIsPositive && metrics.extraordinaryIncomeMonths) {
      elements.heroTitle.textContent = "A sobra existe, mas depende dos meses de renda mais alta.";
      elements.heroDescription.textContent =
        `A renda mediana é ${money.format(metrics.medianIncome)}, enquanto o gasto médio é ` +
        `${money.format(metrics.averageSpend)}. ${metrics.extraordinaryIncomeMonths} mês(es) de entrada extraordinária elevam a média.`;
    } else if (hasPositiveAverage) {
      elements.heroTitle.textContent = "O período fecha no azul — agora o foco é consistência.";
      elements.heroDescription.textContent =
        `A casa preservou ${percentage.format(metrics.positiveRate)} dos meses com saldo positivo. ` +
        `A base recorrente consome ${percentage.format(metrics.essentialCommitment)} da entrada média.`;
    } else {
      elements.heroTitle.textContent = "A rotina da casa ainda gasta mais do que recebe.";
      elements.heroDescription.textContent =
        `Faltaram, em média, ${money.format(Math.abs(metrics.averageBalance))} por mês. ` +
        "Separe o peso fixo da casa dos gastos flexíveis antes de definir qualquer corte.";
    }

    elements.heroNumberLabel.textContent = "Saldo médio mensal";
    elements.heroNumber.textContent = money.format(metrics.averageBalance);
    elements.heroNumber.classList.toggle("is-positive", hasPositiveAverage);
    elements.heroNumber.classList.toggle("is-negative", !hasPositiveAverage);
    elements.heroNumberFoot.textContent =
      `${money.format(metrics.totalBalance)} acumulados · ` +
      `capacidade recorrente estimada em ${money.format(metrics.recurringCapacity)}/mês`;
  }

  function renderKpis(metrics) {
    elements.incomeAverage.textContent = money.format(metrics.averageIncome);
    elements.incomeFoot.textContent = `Mediana mensal: ${money.format(metrics.medianIncome)}`;

    elements.spendAverage.textContent = money.format(metrics.averageSpend);
    elements.spendFoot.textContent = `${percentage.format(metrics.spendCommitment)} da entrada média`;

    elements.balanceAverage.textContent = money.format(metrics.averageBalance);
    elements.balanceFoot.textContent = `${money.format(metrics.totalBalance)} no período`;
    elements.balanceAverage.classList.toggle("is-positive", metrics.averageBalance >= 0);
    elements.balanceAverage.classList.toggle("is-negative", metrics.averageBalance < 0);

    elements.positiveRate.textContent = percentage.format(metrics.positiveRate);
    elements.positiveFoot.textContent = `${metrics.positiveMonths} de ${metrics.count} meses`;
  }

  function renderHealth(metrics) {
    elements.spendCommitment.textContent = percentage.format(metrics.spendCommitment);
    elements.essentialCommitment.textContent = percentage.format(metrics.essentialCommitment);
    elements.flexibleCommitment.textContent = percentage.format(metrics.flexibleCommitment);

    setProgress(elements.spendProgress, metrics.spendCommitment);
    setProgress(elements.essentialProgress, metrics.essentialCommitment);
    setProgress(elements.flexibleProgress, metrics.flexibleCommitment);

    if (metrics.spendCommitment > 1) {
      elements.healthSummary.textContent =
        `Os gastos médios superam a entrada média em ${percentage.format(metrics.spendCommitment - 1)}. ` +
        `A base da casa custa ${money.format(metrics.groupAverages.essential)} e os flexíveis ${money.format(metrics.groupAverages.flexible)} por mês.`;
    } else {
      const available = Math.max(0, metrics.averageIncome - metrics.averageSpend);
      elements.healthSummary.textContent =
        `Depois dos gastos médios, restam ${money.format(available)} por mês. ` +
        `Uma redução simulada de 10% nos flexíveis liberaria ${money.format(metrics.groupAverages.flexible * 0.1)} mensais.`;
    }
  }

  function renderCharts(rows, metrics) {
    if (!window.Chart) {
      elements.cashflowNote.textContent = "Gráficos indisponíveis: Chart.js não carregou.";
      return;
    }

    const labels = rows.map((row) => `${MONTH_SHORT[row.monthIndex]} ${String(row.year).slice(-2)}`);

    replaceChart("cashflow", byId("cashflowChart"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Entradas",
            data: rows.map((row) => row.income),
            backgroundColor: "rgba(103, 219, 229, 0.62)",
            borderColor: "#67dbe5",
            borderWidth: 1,
            borderRadius: 5,
            maxBarThickness: 18
          },
          {
            label: "Gastos",
            data: rows.map((row) => row.spend),
            backgroundColor: "rgba(255, 134, 173, 0.55)",
            borderColor: "#ff86ad",
            borderWidth: 1,
            borderRadius: 5,
            maxBarThickness: 18
          },
          {
            type: "line",
            label: "Saldo",
            data: rows.map((row) => row.balance),
            borderColor: "#f6f7fb",
            backgroundColor: "rgba(246, 247, 251, 0.1)",
            pointBackgroundColor: rows.map((row) => row.balance >= 0 ? "#62dba8" : "#ff7d91"),
            pointBorderWidth: 0,
            pointRadius: rows.length > 30 ? 0 : 2.5,
            pointHoverRadius: 5,
            borderWidth: 2,
            tension: 0.3,
            yAxisID: "balance"
          }
        ]
      },
      options: cashflowChartOptions()
    });

    const visibleCategories = metrics.categoryStats
      .filter((category) => {
        return state.filters.view === "all" || category.group === state.filters.view;
      })
      .filter((category) => category.total > 0);

    replaceChart("category", byId("categoryChart"), {
      type: "doughnut",
      data: {
        labels: visibleCategories.map((category) => labelize(category.name)),
        datasets: [{
          data: visibleCategories.map((category) => category.total),
          backgroundColor: visibleCategories.map((category, index) => CATEGORY_COLORS[index % CATEGORY_COLORS.length]),
          borderColor: "rgba(8, 11, 22, 0.75)",
          borderWidth: 3,
          hoverOffset: 6
        }]
      },
      options: doughnutChartOptions()
    });

    elements.categoryLegend.innerHTML = visibleCategories.slice(0, 7).map((category, index) => {
      return `
        <div class="legend-item">
          <i style="background:${CATEGORY_COLORS[index % CATEGORY_COLORS.length]}"></i>
          <span>${escapeHtml(labelize(category.name))}</span>
          <strong>${percentage.format(category.share)}</strong>
        </div>
      `;
    }).join("");

    const annualEntries = Object.entries(metrics.annual)
      .sort(([firstYear], [secondYear]) => Number(firstYear) - Number(secondYear));

    replaceChart("year", byId("yearChart"), {
      type: "bar",
      data: {
        labels: annualEntries.map(([year]) => year),
        datasets: [
          {
            label: "Entrada média",
            data: annualEntries.map(([, values]) => values.income),
            backgroundColor: "rgba(103, 219, 229, 0.72)",
            borderRadius: 5,
            maxBarThickness: 28
          },
          {
            label: "Gasto médio",
            data: annualEntries.map(([, values]) => values.spend),
            backgroundColor: "rgba(255, 134, 173, 0.62)",
            borderRadius: 5,
            maxBarThickness: 28
          },
          {
            label: "Saldo médio",
            data: annualEntries.map(([, values]) => values.balance),
            backgroundColor: annualEntries.map(([, values]) => values.balance >= 0
              ? "rgba(98, 219, 168, 0.72)"
              : "rgba(255, 125, 145, 0.72)"),
            borderRadius: 5,
            maxBarThickness: 28
          }
        ]
      },
      options: yearChartOptions()
    });

    elements.cashflowNote.textContent = `${rows.length} meses em ordem cronológica`;
  }

  function renderInsights(metrics) {
    const topCategory = metrics.categoryStats[0];
    const worst = metrics.worstMonth;
    const best = metrics.bestMonth;
    const recentTrendType = metrics.recentSpendChange > 0.05
      ? "negative"
      : metrics.recentSpendChange < -0.05
        ? "positive"
        : "warning";

    const cards = [
      {
        type: metrics.recurringCapacity >= 0 ? "positive" : "negative",
        icon: metrics.recurringCapacity >= 0 ? "✓" : "!",
        label: "Orçamento recorrente",
        title: `${money.format(metrics.recurringCapacity)} por mês`,
        text: `Diferença entre a renda mediana (${money.format(metrics.medianIncome)}) e o gasto médio (${money.format(metrics.averageSpend)}).`
      },
      {
        type: metrics.essentialCommitment <= 0.6 ? "positive" : "warning",
        icon: "⌂",
        label: "Peso fixo",
        title: `${percentage.format(metrics.essentialCommitment)} da entrada`,
        text: `Aluguel, condomínio, luz, gás e internet representam ${money.format(metrics.groupAverages.essential)} por mês.`
      },
      {
        type: "warning",
        icon: "◉",
        label: "Maior categoria",
        title: topCategory ? labelize(topCategory.name) : "—",
        text: topCategory
          ? `${money.format(topCategory.average)} por mês e ${percentage.format(topCategory.share)} de todos os gastos.`
          : "Nenhuma categoria disponível."
      },
      {
        type: worst && worst.balance < 0 ? "negative" : "warning",
        icon: "↓",
        label: "Mês mais apertado",
        title: worst ? `${capitalize(worst.month)} ${worst.year}` : "—",
        text: worst
          ? `Saldo de ${money.format(worst.balance)} com ${money.format(worst.income)} de entradas e ${money.format(worst.spend)} de gastos.`
          : "Sem dados."
      },
      {
        type: recentTrendType,
        icon: metrics.recentSpendChange > 0 ? "↗" : "↘",
        label: "Ritmo recente",
        title: `${metrics.recentSpendChange >= 0 ? "+" : ""}${percentage.format(metrics.recentSpendChange)} nos gastos`,
        text: `Os últimos três meses ficaram em ${money.format(metrics.recentSpendAverage)}/mês contra ${money.format(metrics.previousSpendAverage)}/mês no trimestre anterior.`
      },
      {
        type: metrics.incomeVolatility > 0.3 ? "warning" : "positive",
        icon: "∿",
        label: "Variação da renda",
        title: incomeVolatilityLabel(metrics.incomeVolatility),
        text: `${percentage.format(metrics.incomeVolatility)} de variação relativa. Melhor mês do recorte: ${best ? `${capitalize(best.month)} ${best.year}, ${money.format(best.balance)}` : "—"}.`
      }
    ];

    elements.insightsGrid.innerHTML = cards.map((card) => {
      return `
        <article class="insight ${card.type}">
          <div class="insight-top">
            <span class="insight-icon">${card.icon}</span>
            <span>${escapeHtml(card.label)}</span>
          </div>
          <strong>${escapeHtml(card.title)}</strong>
          <p>${escapeHtml(card.text)}</p>
        </article>
      `;
    }).join("");
  }

  function renderCategoryTable(metrics) {
    const visibleCategories = metrics.categoryStats.filter((category) => {
      return state.filters.view === "all" || category.group === state.filters.view;
    });

    if (!visibleCategories.length) {
      elements.categoryTableBody.innerHTML =
        `<tr><td class="empty-row" colspan="5">Nenhuma categoria nesse recorte.</td></tr>`;
      return;
    }

    elements.categoryTableBody.innerHTML = visibleCategories.map((category) => {
      const trendClass = category.trend > 0.05
        ? "up"
        : category.trend < -0.05
          ? "down"
          : "stable";

      const trendLabel = category.trend > 0.05
        ? `↑ ${percentage.format(category.trend)}`
        : category.trend < -0.05
          ? `↓ ${percentage.format(Math.abs(category.trend))}`
          : "Estável";

      return `
        <tr>
          <td>${escapeHtml(labelize(category.name))}</td>
          <td><span class="type-badge ${category.group}">${GROUPS[category.group].label}</span></td>
          <td>${money.format(category.average)}</td>
          <td>${percentage.format(category.share)}</td>
          <td><span class="trend-badge ${trendClass}">${trendLabel}</span></td>
        </tr>
      `;
    }).join("");
  }

  function renderMonthlyTable(rows) {
    const sortedRows = [...rows].sort((first, second) => second.date - first.date);
    elements.tableCount.textContent = `${rows.length} registros`;

    elements.monthlyTableBody.innerHTML = sortedRows.map((row) => {
      const positive = row.balance >= 0;

      return `
        <tr>
          <td>${capitalize(row.month)} ${row.year}</td>
          <td>${money.format(row.income)}</td>
          <td>${money.format(row.groups.essential)}</td>
          <td>${money.format(row.groups.flexible)}</td>
          <td>${money.format(row.spend)}</td>
          <td class="${positive ? "positive-value" : "negative-value"}">${money.format(row.balance)}</td>
          <td>
            <span class="status-badge ${positive ? "positive" : "negative"}">
              ${positive ? "Com sobra" : "No vermelho"}
            </span>
          </td>
        </tr>
      `;
    }).join("");
  }

  function renderNoMatches() {
    elements.heroTitle.textContent = "Nenhum mês corresponde aos filtros.";
    elements.heroDescription.textContent = "Selecione mais anos, meses ou volte para a situação ‘Todos’.";
    elements.heroNumber.textContent = "—";
    elements.heroNumberFoot.textContent = "Sem dados no recorte";
    elements.heroNumber.classList.remove("is-positive", "is-negative");

    [
      "incomeAverage",
      "spendAverage",
      "balanceAverage",
      "positiveRate"
    ].forEach((id) => {
      elements[id].textContent = "—";
      elements[id].classList.remove("is-positive", "is-negative");
    });

    elements.healthSummary.textContent = "Amplie o recorte para recalcular os indicadores.";
    [elements.spendProgress, elements.essentialProgress, elements.flexibleProgress].forEach((bar) => {
      bar.style.width = "0%";
    });

    elements.spendCommitment.textContent = "—";
    elements.essentialCommitment.textContent = "—";
    elements.flexibleCommitment.textContent = "—";
    elements.insightsGrid.innerHTML = `<div class="insight"><strong>Sem dados</strong><p>Altere os filtros para continuar.</p></div>`;
    elements.categoryTableBody.innerHTML = `<tr><td class="empty-row" colspan="5">Sem dados no recorte.</td></tr>`;
    elements.monthlyTableBody.innerHTML = `<tr><td class="empty-row" colspan="7">Sem dados no recorte.</td></tr>`;
    elements.categoryLegend.innerHTML = "";
    elements.tableCount.textContent = "0 registros";
    destroyCharts();
  }

  function renderEmptyDashboard() {
    renderFilters();
    elements.heroTitle.textContent = "Carregue a planilha para começar.";
    elements.heroDescription.textContent = "O painel usará as colunas reais da tabela para montar os indicadores.";
    elements.heroNumber.textContent = "—";
    elements.heroNumberFoot.textContent = "Aguardando dados";

    [
      "incomeAverage",
      "spendAverage",
      "balanceAverage",
      "positiveRate"
    ].forEach((id) => {
      elements[id].textContent = "—";
    });

    elements.categoryTableBody.innerHTML = `<tr><td class="empty-row" colspan="5">Aguardando planilha.</td></tr>`;
    elements.monthlyTableBody.innerHTML = `<tr><td class="empty-row" colspan="7">Aguardando planilha.</td></tr>`;
    elements.insightsGrid.innerHTML = `<div class="insight"><strong>Aguardando dados</strong><p>Carregue Financas.xlsx para gerar os insights.</p></div>`;
  }

  function exportCurrentView() {
    if (!state.filteredRows.length) {
      showToast("Não há dados para exportar.", true);
      return;
    }

    const headers = [
      "Mês",
      "Ano",
      "Entradas",
      "Base da casa",
      "Flexíveis",
      "Outros",
      "Total gasto",
      "Saldo",
      ...state.categories.map(labelize)
    ];

    const body = state.filteredRows.map((row) => [
      capitalize(row.month),
      row.year,
      row.income,
      row.groups.essential,
      row.groups.flexible,
      row.groups.other,
      row.spend,
      row.balance,
      ...state.categories.map((category) => row.categories[category])
    ]);

    const csv = [headers, ...body]
      .map((row) => row.map(csvValue).join(";"))
      .join("\n");

    const blob = new Blob(["\uFEFF" + csv], {
      type: "text/csv;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "financas-filtradas.csv";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    showToast("Recorte exportado.");
  }

  function configureCharts() {
    if (!window.Chart) {
      return;
    }

    Chart.defaults.color = "#97a1ba";
    Chart.defaults.font.family = "\"DM Sans\", sans-serif";
    Chart.defaults.font.size = 10;
    Chart.defaults.animation.duration = 450;
    Chart.defaults.plugins.legend.labels = Chart.defaults.plugins.legend.labels || {};
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
    Chart.defaults.plugins.legend.labels.pointStyle = "circle";
    Chart.defaults.plugins.legend.labels.boxWidth = 7;
    Chart.defaults.plugins.legend.labels.boxHeight = 7;
    Chart.defaults.plugins.tooltip.backgroundColor = "rgba(17, 23, 43, 0.97)";
    Chart.defaults.plugins.tooltip.borderColor = "rgba(255,255,255,0.12)";
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.padding = 11;
  }

  function replaceChart(name, canvas, configuration) {
    if (state.charts[name]) {
      state.charts[name].destroy();
    }

    state.charts[name] = new Chart(canvas, configuration);
  }

  function destroyCharts() {
    Object.values(state.charts).forEach((chart) => chart.destroy());
    state.charts = {};
  }

  function cashflowChartOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: "#7e89a3",
            maxRotation: 0,
            autoSkip: true,
            maxTicksLimit: window.innerWidth < 640 ? 6 : 14
          }
        },
        y: {
          beginAtZero: true,
          grid: { color: "rgba(255,255,255,0.06)" },
          ticks: {
            color: "#7e89a3",
            maxTicksLimit: 6,
            callback: compactMoney
          }
        },
        balance: {
          position: "right",
          grid: { drawOnChartArea: false },
          ticks: {
            color: "#aab2c8",
            maxTicksLimit: 6,
            callback: compactMoney
          }
        }
      },
      plugins: {
        legend: {
          display: true,
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label(context) {
              return ` ${context.dataset.label}: ${money.format(context.raw || 0)}`;
            }
          }
        }
      }
    };
  }

  function doughnutChartOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "68%",
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label(context) {
              const values = context.dataset.data;
              const total = values.reduce((sumValue, value) => sumValue + value, 0);
              const share = total ? context.raw / total : 0;
              return ` ${context.label}: ${money.format(context.raw)} · ${percentage.format(share)}`;
            }
          }
        }
      }
    };
  }

  function yearChartOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: "#7e89a3" }
        },
        y: {
          grid: { color: "rgba(255,255,255,0.06)" },
          ticks: {
            color: "#7e89a3",
            maxTicksLimit: 6,
            callback: compactMoney
          }
        }
      },
      plugins: {
        legend: {
          display: true,
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label(context) {
              return ` ${context.dataset.label}: ${money.format(context.raw || 0)}`;
            }
          }
        }
      }
    };
  }

  function setSourceStatus(status, name, meta) {
    elements.sourceName.textContent = name;
    elements.sourceMeta.textContent = meta;
    elements.statusDot.classList.toggle("is-ready", status === "ready");
    elements.statusDot.classList.toggle("is-error", status === "error");
  }

  function showLoadError(detail) {
    setSourceStatus(
      "error",
      "Planilha não encontrada",
      "Use Financas.xlsx na raiz ou carregue outra tabela."
    );

    elements.emptyError.textContent = detail;
    elements.emptyState.hidden = false;
  }

  function showToast(message, isError = false) {
    window.clearTimeout(showToast.timer);
    elements.toast.textContent = message;
    elements.toast.classList.toggle("is-error", isError);
    elements.toast.classList.add("is-visible");

    showToast.timer = window.setTimeout(() => {
      elements.toast.classList.remove("is-visible");
    }, 3200);
  }

  function setProgress(element, value) {
    const safeValue = Math.max(0, Math.min(value, 1.2));
    element.style.width = `${Math.min(100, safeValue * 100)}%`;
  }

  function inferCategoryGroup(category) {
    const normalized = normalizeKey(category);

    if (GROUPS.essential.aliases.some((alias) => normalized.includes(normalizeKey(alias)))) {
      return "essential";
    }

    if (GROUPS.flexible.aliases.some((alias) => normalized.includes(normalizeKey(alias)))) {
      return "flexible";
    }

    return "other";
  }

  function findColumn(headerInfo, aliases) {
    const normalizedAliases = aliases.map(normalizeKey);
    const exact = headerInfo.find(({ normalized }) => normalizedAliases.includes(normalized));

    if (exact) {
      return exact.original;
    }

    const partial = headerInfo.find(({ normalized }) => {
      return normalizedAliases.some((alias) => normalized.includes(alias));
    });

    return partial ? partial.original : null;
  }

  function parseMonth(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value.getMonth();
    }

    const normalized = normalizeText(value);
    const textIndex = MONTHS.map(normalizeText).indexOf(normalized);

    if (textIndex >= 0) {
      return textIndex;
    }

    const abbreviationIndex = MONTH_SHORT
      .map(normalizeText)
      .indexOf(normalized.slice(0, 3));

    if (abbreviationIndex >= 0) {
      return abbreviationIndex;
    }

    const numericValue = Number(value);
    return numericValue >= 1 && numericValue <= 12 ? numericValue - 1 : -1;
  }

  function isNumeric(value) {
    if (typeof value === "number") {
      return Number.isFinite(value);
    }

    if (value === null || value === undefined || String(value).trim() === "") {
      return false;
    }

    const text = String(value).trim();

    return /^[-+]?\(?\s*R?\$?\s*[\d.]+(?:,\d+)?\s*\)?$/.test(text) ||
      /^[-+]?\(?\s*\d+(?:\.\d+)?\s*\)?$/.test(text);
  }

  function toNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : 0;
    }

    let text = String(value ?? "").trim();

    if (!text) {
      return 0;
    }

    const isNegativeByParentheses = /^\(.*\)$/.test(text);
    text = text.replace(/[()\sR$]/gi, "");

    if (text.includes(",") && text.includes(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else if (text.includes(",")) {
      text = text.replace(",", ".");
    }

    const parsed = Number(text.replace(/[^0-9.-]/g, ""));

    if (!Number.isFinite(parsed)) {
      return 0;
    }

    return isNegativeByParentheses ? -Math.abs(parsed) : parsed;
  }

  function normalizeText(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  function normalizeKey(value) {
    return normalizeText(value).replace(/[^a-z0-9]/g, "");
  }

  function labelize(value) {
    return String(value ?? "")
      .trim()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ")
      .replace(/\b\w/g, (letter) => letter.toUpperCase());
  }

  function capitalize(value) {
    const text = String(value ?? "");
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  function getPeriodLabel(rows) {
    if (!rows.length) {
      return "Sem dados no recorte";
    }

    const first = rows[0];
    const last = rows[rows.length - 1];

    if (first.id === last.id) {
      return `${capitalize(first.month)} ${first.year}`;
    }

    return `${capitalize(first.month)} ${first.year} — ${capitalize(last.month)} ${last.year}`;
  }

  function incomeVolatilityLabel(value) {
    if (value <= 0.15) {
      return "Renda previsível";
    }

    if (value <= 0.3) {
      return "Renda variável";
    }

    return "Renda irregular";
  }

  function compactMoney(value) {
    const numericValue = Number(value) || 0;
    const absolute = Math.abs(numericValue);
    const sign = numericValue < 0 ? "−" : "";

    if (absolute >= 1000000) {
      return `${sign}R$ ${(absolute / 1000000).toFixed(1).replace(".", ",")} mi`;
    }

    if (absolute >= 1000) {
      return `${sign}R$ ${(absolute / 1000).toFixed(1).replace(".", ",")} mil`;
    }

    return `${sign}R$ ${Math.round(absolute)}`;
  }

  function csvValue(value) {
    const text = String(value ?? "").replace(/"/g, '""');
    return /[;"\n]/.test(text) ? `"${text}"` : text;
  }

  function sum(items, getter) {
    return items.reduce((total, item) => total + getter(item), 0);
  }

  function average(items, getter) {
    return items.length ? sum(items, getter) / items.length : 0;
  }

  function median(values) {
    if (!values.length) {
      return 0;
    }

    const sorted = [...values].sort((first, second) => first - second);
    const middle = Math.floor(sorted.length / 2);

    return sorted.length % 2
      ? sorted[middle]
      : (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function standardDeviation(values) {
    if (!values.length) {
      return 0;
    }

    const mean = values.reduce((total, value) => total + value, 0) / values.length;
    const variance = values.reduce((total, value) => {
      return total + Math.pow(value - mean, 2);
    }, 0) / values.length;

    return Math.sqrt(variance);
  }

  function uniqueSorted(values) {
    return [...new Set(values)].sort((first, second) => first - second);
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (character) => {
      const entities = {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      };

      return entities[character];
    });
  }
})();
