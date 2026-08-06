/*
 * Finanças da casa
 *
 * O painel foi pensado para orçamento doméstico:
 * - separa contas essenciais, gastos flexíveis e outros;
 * - compara vários anos e meses sem menus suspensos;
 * - mostra médias mensais, pressão sobre a renda e movimentos recentes;
 * - permite importar outra planilha sem alterar o arquivo original.
 */
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

  const MONTH_LABELS = [
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

  const DEFAULT_FILES = [
    "./Finanças.xlsx",
    "./Finanças.xlsx",
    "./Financas.xlsx",
    "./financas.xlsx",
    "./upload/Finanças.xlsx"
  ];

  const INCOME_ALIASES = [
    "renda",
    "quintino",
    "receita",
    "receitas",
    "entrada",
    "entradas",
    "salario",
    "salário",
    "income"
  ];

  const SUMMARY_ALIASES = [
    "totalgasto",
    "totaldespesa",
    "gastototal",
    "despesastotal",
    "balanca",
    "saldo",
    "balance",
    "poupanca",
    "taxa de poupanca"
  ];

  const GROUP_RULES = {
    essential: {
      label: "Essenciais",
      shortLabel: "Essenciais",
      color: "#ffc178",
      aliases: [
        "aluguel",
        "condominio",
        "condomínio",
        "luz",
        "energia",
        "gas",
        "gás",
        "agua",
        "água",
        "internet",
        "telefone",
        "escola",
        "saude",
        "saúde",
        "plano",
        "financiamento",
        "prestacao",
        "prestação"
      ]
    },
    flexible: {
      label: "Flexíveis",
      shortLabel: "Flexíveis",
      color: "#78d9e9",
      aliases: [
        "cartao",
        "cartão",
        "comida",
        "lazer",
        "entretenimento",
        "restaurante",
        "delivery",
        "compras",
        "viagem",
        "roupa"
      ]
    },
    other: {
      label: "Outros",
      shortLabel: "Outros",
      color: "#9c8eff",
      aliases: []
    }
  };

  const GROUP_ORDER = ["essential", "flexible", "other"];

  const palette = [
    "#78d9e9",
    "#9c8eff",
    "#ff83b4",
    "#ffc178",
    "#6edbaa",
    "#7aa9ff",
    "#d19bff",
    "#6dd3c4"
  ];

  const currency = new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });

  const number = new Intl.NumberFormat("pt-BR", {
    maximumFractionDigits: 0
  });

  const percent = new Intl.NumberFormat("pt-BR", {
    style: "percent",
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  });

  const state = {
    rows: [],
    sourceName: "",
    categories: [],
    categoryGroups: {},
    incomeSources: [],
    filtered: [],
    charts: {},
    ignoredRows: 0
  };

  const filterState = {
    years: new Set(),
    months: new Set(),
    status: "all",
    lens: "all"
  };

  const els = {};
  const $ = (id) => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    [
      "dataStatusDot",
      "sourceName",
      "sourceMeta",
      "fileInput",
      "dropzone",
      "lastUpdated",
      "periodLabel",
      "yearHint",
      "monthHint",
      "yearChips",
      "monthChips",
      "statusChips",
      "lensChips",
      "activeFilters",
      "filterResult",
      "resetFilters",
      "exportCsv",
      "heroTitle",
      "heroText",
      "heroLastBalance",
      "heroLastLabel",
      "heroAverageBalance",
      "heroAverageLabel",
      "heroEssentialShare",
      "heroEssentialLabel",
      "rowCountTag",
      "categoryCountTag",
      "kpiBalance",
      "kpiBalanceFoot",
      "kpiIncomeAvg",
      "kpiIncomeAvgFoot",
      "kpiSpendAvg",
      "kpiSpendAvgFoot",
      "kpiEssential",
      "kpiEssentialFoot",
      "kpiFlexible",
      "kpiFlexibleFoot",
      "kpiPositive",
      "kpiPositiveFoot",
      "snapshotStatus",
      "snapshotBalance",
      "snapshotLabel",
      "snapshotIncome",
      "snapshotEssential",
      "snapshotFlexible",
      "snapshotSpend",
      "snapshotNote",
      "allocationCaption",
      "allocationBar",
      "allocationLegend",
      "allocationNote",
      "cashflowNote",
      "categoryChartTitle",
      "insightsList",
      "actionList",
      "categoryTableBody",
      "tableCaption",
      "dataTableBody",
      "quickRead",
      "essentialLegend",
      "flexibleLegend",
      "otherLegend",
      "classificationNote",
      "footerMethod",
      "toast"
    ].forEach((id) => {
      els[id] = $(id);
    });

    bindInteractions();
    setupChartDefaults();
    await tryDefaultFiles();
  }

  function bindInteractions() {
    [
      [els.yearChips, "years"],
      [els.monthChips, "months"],
      [els.statusChips, "status"],
      [els.lensChips, "lens"]
    ].forEach(([container, group]) => {
      container.addEventListener("click", (event) => {
        const button = event.target.closest("button[data-filter-value]");

        if (!button) {
          return;
        }

        updateFilter(group, button.dataset.filterValue);
      });
    });

    els.resetFilters.addEventListener("click", () => {
      clearFilters();
      showToast("Filtros limpos");
    });

    els.exportCsv.addEventListener("click", exportFilteredCsv);

    els.dropzone.addEventListener("click", () => {
      els.fileInput.click();
    });

    els.dropzone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        els.fileInput.click();
      }
    });

    els.fileInput.addEventListener("change", (event) => {
      const file = event.target.files && event.target.files[0];

      if (file) {
        loadFile(file);
      }

      event.target.value = "";
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        els.dropzone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        els.dropzone.classList.remove("is-dragging");
      });
    });

    els.dropzone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files && event.dataTransfer.files[0];

      if (file) {
        loadFile(file);
      }
    });
  }

  function setupChartDefaults() {
    if (!window.Chart) {
      return;
    }

    Chart.defaults.color = "#9aa7c4";
    Chart.defaults.font.family = '"DM Sans", sans-serif';
    Chart.defaults.font.size = 10;
    Chart.defaults.animation.duration = 500;
    Chart.defaults.plugins.legend.display = false;
    Chart.defaults.plugins.tooltip.backgroundColor = "#101a36";
    Chart.defaults.plugins.tooltip.borderColor = "rgba(255,255,255,.15)";
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.titleFont = {
      family: "Space Grotesk",
      weight: "600"
    };
  }

  function updateFilter(group, value) {
    if (group === "years" || group === "months") {
      const target = filterState[group];

      if (value === "all") {
        target.clear();
      } else {
        const numericValue = Number(value);
        const parsedValue = Number.isNaN(numericValue)
          ? value
          : numericValue;

        if (target.has(parsedValue)) {
          target.delete(parsedValue);
        } else {
          target.add(parsedValue);
        }
      }
    } else if (group === "status") {
      filterState.status = value;
    } else if (group === "lens") {
      filterState.lens = value;
    }

    renderFilterControls();
    renderAll();
  }

  function clearFilters() {
    filterState.years.clear();
    filterState.months.clear();
    filterState.status = "all";
    filterState.lens = "all";

    renderFilterControls();
    renderAll();
  }

  async function tryDefaultFiles() {
    setSourceState(
      "loading",
      "Carregando arquivo da raiz…",
      "Tentando encontrar sua planilha"
    );

    for (const path of DEFAULT_FILES) {
      try {
        const response = await fetch(path, { cache: "no-store" });

        if (!response.ok) {
          continue;
        }

        const data = await response.arrayBuffer();

        await loadArrayBuffer(
          data,
          path.split("/").pop() || "Finanças.xlsx"
        );

        return;
      } catch (error) {
        /* Fallback para upload manual quando o arquivo não está acessível. */
      }
    }

    setSourceState(
      "waiting",
      "Nenhum arquivo carregado",
      "Escolha uma planilha para começar"
    );

    renderEmpty();
  }

  async function loadFile(file) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      showToast("Escolha um arquivo .xlsx, .xls ou .csv", true);
      return;
    }

    try {
      setSourceState(
        "loading",
        file.name,
        "Lendo e organizando os dados da casa…"
      );

      const data = await file.arrayBuffer();

      await loadArrayBuffer(data, file.name);
      showToast("Tabela carregada com sucesso");
    } catch (error) {
      console.error(error);

      setSourceState(
        "waiting",
        "Não foi possível ler o arquivo",
        "Verifique a aba e os cabeçalhos"
      );

      showToast(
        "Não consegui ler essa planilha. Confira se ela tem mês e ano.",
        true
      );
    }
  }

  async function loadArrayBuffer(data, sourceName) {
    if (!window.XLSX) {
      throw new Error("Biblioteca XLSX indisponível");
    }

    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("Planilha sem abas");
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      defval: null,
      raw: true
    });

    const normalized = normalizeRows(rawRows);

    if (!normalized.rows.length) {
      throw new Error("Nenhuma linha válida encontrada");
    }

    state.rows = normalized.rows;
    state.categories = normalized.categories;
    state.categoryGroups = normalized.categoryGroups;
    state.incomeSources = normalized.incomeSources;
    state.sourceName = sourceName;
    state.ignoredRows = normalized.ignoredRows;

    filterState.years.clear();
    filterState.months.clear();
    filterState.status = "all";
    filterState.lens = "all";

    updateSidebarLegend();
    renderFilterControls();

    setSourceState(
      "ready",
      sourceName,
      `${state.rows.length} meses · ${state.categories.length} categorias`
    );

    els.lastUpdated.textContent = `Atualizado às ${new Intl.DateTimeFormat(
      "pt-BR",
      {
        hour: "2-digit",
        minute: "2-digit"
      }
    ).format(new Date())}`;

    renderAll();
  }

  function normalizeRows(rawRows) {
    if (!Array.isArray(rawRows) || !rawRows.length) {
      return {
        rows: [],
        categories: [],
        categoryGroups: {},
        incomeSources: [],
        ignoredRows: 0
      };
    }

    const keys = Object.keys(rawRows[0] || {});

    const keyInfo = keys.map((key) => ({
      key,
      id: normalizeKey(key)
    }));

    const monthKey = findKey(keyInfo, [
      "mes",
      "month",
      "periodo",
      "period"
    ]);

    const yearKey = findKey(keyInfo, [
      "ano",
      "year"
    ]);

    const totalKey = findKey(keyInfo, [
      "totalgasto",
      "totaldespesa",
      "gastototal",
      "despesastotal"
    ]);

    const balanceKey = findKey(keyInfo, [
      "balanca",
      "saldo",
      "balance"
    ]);

    if (!monthKey || !yearKey) {
      throw new Error("A planilha precisa ter colunas de mês e ano");
    }

    const incomeSources = keyInfo
      .filter(({ id }) => {
        return INCOME_ALIASES.some((alias) => {
          const normalizedAlias = normalizeKey(alias);

          return (
            id === normalizedAlias ||
            id.includes(normalizedAlias)
          );
        });
      })
      .map(({ key }) => key);

    const excluded = new Set([
      monthKey,
      yearKey,
      totalKey,
      balanceKey,
      ...incomeSources
    ]);

    const categories = keyInfo
      .filter(({ key, id }) => {
        const isSummary = SUMMARY_ALIASES.some((alias) => {
          return id === normalizeKey(alias);
        });

        const hasNumericValues = rawRows.some((row) => {
          return isNumeric(row[key]);
        });

        return !excluded.has(key) && !isSummary && hasNumericValues;
      })
      .map(({ key }) => key);

    const uniqueCategories = [...new Set(categories)];

    const categoryGroups = Object.fromEntries(
      uniqueCategories.map((category) => {
        return [category, inferGroup(category)];
      })
    );

    const rows = [];
    let ignoredRows = 0;

    rawRows.forEach((raw, index) => {
      const monthText = raw[monthKey];
      const year = Math.round(toNumber(raw[yearKey]));
      const monthIndex = monthIndexOf(monthText);

      if (!monthText || !year || monthIndex < 0) {
        ignoredRows += 1;
        return;
      }

      const values = {};

      uniqueCategories.forEach((category) => {
        values[category] = toNumber(raw[category]);
      });

      const spend = totalKey && isNumeric(raw[totalKey])
        ? toNumber(raw[totalKey])
        : uniqueCategories.reduce((sumValue, category) => {
            return sumValue + values[category];
          }, 0);

      const income = incomeSources.reduce((sumValue, source) => {
        return sumValue + toNumber(raw[source]);
      }, 0);

      const balance = balanceKey && isNumeric(raw[balanceKey])
        ? toNumber(raw[balanceKey])
        : income - spend;

      const groups = summarizeGroups(values, categoryGroups);
      const date = new Date(year, monthIndex, 1);

      rows.push({
        id: `${year}-${String(monthIndex + 1).padStart(2, "0")}-${index}`,
        month: MONTHS[monthIndex],
        monthLabel: capitalize(MONTHS[monthIndex]),
        monthIndex,
        year,
        date,
        values,
        groups,
        spend,
        income,
        balance,
        savingsRate: income ? balance / income : 0,
        leader: findLeader(values)
      });
    });

    rows.sort((first, second) => first.date - second.date);

    return {
      rows,
      categories: uniqueCategories,
      categoryGroups,
      incomeSources,
      ignoredRows
    };
  }

  function renderFilterControls() {
    const years = uniqueSorted(state.rows.map((row) => row.year));

    const availableMonths = uniqueSorted(
      state.rows.map((row) => row.monthIndex)
    );

    els.yearChips.innerHTML = [
      makeChip("Todos", "all", !filterState.years.size, "years"),
      ...years.map((year) => {
        return makeChip(
          String(year),
          String(year),
          filterState.years.has(year),
          "years"
        );
      })
    ].join("");

    els.monthChips.innerHTML = [
      makeChip("Todos", "all", !filterState.months.size, "months"),
      ...availableMonths.map((monthIndex) => {
        return makeChip(
          capitalize(MONTH_LABELS[monthIndex]),
          String(monthIndex),
          filterState.months.has(monthIndex),
          "months"
        );
      })
    ].join("");

    const statusOptions = [
      ["all", "Todos"],
      ["positive", "No azul"],
      ["negative", "No vermelho"]
    ];

    els.statusChips.innerHTML = statusOptions
      .map(([value, label]) => {
        return makeChip(
          label,
          value,
          filterState.status === value,
          "status",
          "segment-chip"
        );
      })
      .join("");

    const lensOptions = [
      ["all", "Tudo"],
      ["essential", "Essenciais"],
      ["flexible", "Flexíveis"],
      ["other", "Outros"]
    ];

    els.lensChips.innerHTML = lensOptions
      .map(([value, label]) => {
        return makeChip(
          label,
          value,
          filterState.lens === value,
          "lens",
          "segment-chip"
        );
      })
      .join("");

    els.yearHint.textContent = filterState.years.size
      ? `${filterState.years.size} selecionado(s)`
      : "Todos selecionados";

    els.monthHint.textContent = filterState.months.size
      ? `${filterState.months.size} selecionado(s)`
      : "Todos selecionados";

    renderActiveFilters();
  }

  function makeChip(
    label,
    value,
    active,
    group,
    extraClass = "choice-chip"
  ) {
    return `
      <button
        type="button"
        class="${extraClass}${active ? " is-active" : ""}"
        data-filter-group="${group}"
        data-filter-value="${escapeHtml(value)}"
        aria-pressed="${active}"
      >${escapeHtml(label)}</button>
    `;
  }

  function renderActiveFilters() {
    const tags = [];

    if (filterState.years.size) {
      tags.push(
        `Anos: ${[...filterState.years].sort().join(", ")}`
      );
    }

    if (filterState.months.size) {
      tags.push(
        `Meses: ${[...filterState.months]
          .sort((a, b) => a - b)
          .map((month) => capitalize(MONTH_LABELS[month]))
          .join(", ")}`
      );
    }

    if (filterState.status !== "all") {
      tags.push(
        filterState.status === "positive"
          ? "Somente meses no azul"
          : "Somente meses no vermelho"
      );
    }

    if (filterState.lens !== "all") {
      tags.push(
        `Lente: ${GROUP_RULES[filterState.lens].label}`
      );
    }

    els.activeFilters.innerHTML = tags.length
      ? tags
          .map(
            (tag) =>
              `<span class="active-filter">${escapeHtml(tag)}</span>`
          )
          .join("")
      : `<span class="active-filter is-empty">Nenhum filtro aplicado · mostrando todos os dados</span>`;
  }

  function getFilteredRows() {
    return state.rows.filter((row) => {
      const yearMatches =
        !filterState.years.size ||
        filterState.years.has(row.year);

      const monthMatches =
        !filterState.months.size ||
        filterState.months.has(row.monthIndex);

      const statusMatches =
        filterState.status === "all" ||
        (filterState.status === "positive"
          ? row.balance > 0
          : row.balance < 0);

      return yearMatches && monthMatches && statusMatches;
    });
  }

  function aggregate(rows) {
    const income = sum(rows, (row) => row.income);
    const spend = sum(rows, (row) => row.spend);
    const balance = sum(rows, (row) => row.balance);

    const averageIncome = rows.length ? income / rows.length : 0;
    const averageSpend = rows.length ? spend / rows.length : 0;
    const averageBalance = rows.length ? balance / rows.length : 0;

    const categoryTotals = Object.fromEntries(
      state.categories.map((category) => {
        return [
          category,
          sum(rows, (row) => row.values[category] || 0)
        ];
      })
    );

    const groupTotals = Object.fromEntries(
      GROUP_ORDER.map((group) => {
        return [
          group,
          sum(rows, (row) => row.groups[group] || 0)
        ];
      })
    );

    const averageGroups = Object.fromEntries(
      GROUP_ORDER.map((group) => {
        return [
          group,
          rows.length ? groupTotals[group] / rows.length : 0
        ];
      })
    );

    const positive = rows.filter((row) => row.balance > 0).length;
    const negative = rows.filter((row) => row.balance < 0).length;

    const savingsRate = income ? balance / income : 0;
    const essentialIncomeShare = income
      ? groupTotals.essential / income
      : 0;

    const flexibleIncomeShare = income
      ? groupTotals.flexible / income
      : 0;

    const essentialSpendShare = spend
      ? groupTotals.essential / spend
      : 0;

    const flexibleSpendShare = spend
      ? groupTotals.flexible / spend
      : 0;

    const incomeValues = rows.map((row) => row.income);
    const incomeMedian = median(incomeValues);
    const incomeStd = standardDeviation(incomeValues);

    const incomeVolatility = averageIncome
      ? incomeStd / averageIncome
      : 0;

    const firstWindowSize = Math.max(
      1,
      Math.floor(rows.length / 3)
    );

    const firstWindow = rows.slice(0, firstWindowSize);
    const recentWindow = rows.slice(-firstWindowSize);

    const firstWindowSpend = average(
      firstWindow,
      (row) => row.spend
    );

    const recentWindowSpend = average(
      recentWindow,
      (row) => row.spend
    );

    const firstWindowBalance = average(
      firstWindow,
      (row) => row.balance
    );

    const recentWindowBalance = average(
      recentWindow,
      (row) => row.balance
    );

    const categoryStats = state.categories
      .map((category) => {
        const total = categoryTotals[category] || 0;

        const firstAverage = average(
          firstWindow,
          (row) => row.values[category] || 0
        );

        const recentAverage = average(
          recentWindow,
          (row) => row.values[category] || 0
        );

        const trend = firstAverage
          ? (recentAverage - firstAverage) / firstAverage
          : recentAverage > 0
            ? 1
            : 0;

        return {
          category,
          group: state.categoryGroups[category] || "other",
          total,
          average: rows.length ? total / rows.length : 0,
          share: spend ? total / spend : 0,
          trend,
          firstAverage,
          recentAverage
        };
      })
      .sort((first, second) => second.total - first.total);

    const yearAverages = groupByYearAverage(rows);
    const monthAverages = groupByMonthAverage(rows);

    return {
      income,
      spend,
      balance,
      averageIncome,
      averageSpend,
      averageBalance,
      categoryTotals,
      categoryStats,
      groupTotals,
      averageGroups,
      positive,
      negative,
      savingsRate,
      positiveRate: rows.length ? positive / rows.length : 0,
      essentialIncomeShare,
      flexibleIncomeShare,
      essentialSpendShare,
      flexibleSpendShare,
      incomeMedian,
      incomeStd,
      incomeVolatility,
      minIncome: incomeValues.length
        ? Math.min(...incomeValues)
        : 0,
      maxIncome: incomeValues.length
        ? Math.max(...incomeValues)
        : 0,
      extraordinaryIncomeCount: incomeValues.filter((value) => {
        return incomeMedian > 0 && value > incomeMedian * 1.35;
      }).length,
      firstWindowSpend,
      recentWindowSpend,
      firstWindowBalance,
      recentWindowBalance,
      recentSpendChange: firstWindowSpend
        ? (recentWindowSpend - firstWindowSpend) / firstWindowSpend
        : 0,
      recentBalanceChange: recentWindowBalance - firstWindowBalance,
      yearAverages,
      monthAverages,
      focusCategories: getFocusCategories(categoryStats)
    };
  }

  function renderAll() {
    if (!state.rows.length) {
      renderEmpty();
      return;
    }

    const rows = getFilteredRows();
    const metrics = aggregate(rows);

    state.filtered = rows;

    els.filterResult.textContent = `${rows.length} ${
      rows.length === 1 ? "mês" : "meses"
    }`;

    els.periodLabel.textContent = getPeriodLabel(rows);
    els.footerMethod.textContent = buildFooterMethod();

    renderHero(rows, metrics);
    renderKpis(rows, metrics);
    renderSnapshot(rows, metrics);
    renderAllocation(metrics);
    renderCharts(rows, metrics);
    renderInsights(rows, metrics);
    renderActions(rows, metrics);
    renderCategoryTable(metrics);
    renderTable(rows);
    updateQuickRead(rows, metrics);
  }

  function renderHero(rows, metrics) {
    if (!rows.length) {
      els.heroTitle.textContent = "Nenhum mês nesse recorte.";

      els.heroText.textContent =
        "Amplie os anos, meses ou a situação do mês para voltar a enxergar a rotina da casa.";

      els.heroLastBalance.textContent = "—";
      els.heroLastLabel.textContent = "—";
      els.heroAverageBalance.textContent = "—";
      els.heroAverageLabel.textContent = "—";
      els.heroEssentialShare.textContent = "—";
      els.heroEssentialLabel.textContent = "da renda média";

      return;
    }

    const recent = rows[rows.length - 1];
    const averageTone = metrics.averageBalance >= 0;

    els.heroTitle.textContent = averageTone
      ? "A média do recorte deixa uma sobra."
      : "A média do recorte pede espaço.";

    const lensText = filterState.lens === "all"
      ? "contas essenciais e gastos flexíveis"
      : `a lente de ${GROUP_RULES[filterState.lens].label.toLowerCase()}`;

    els.heroText.textContent =
      `Em ${rows.length} ${
        rows.length === 1 ? "mês" : "meses"
      }, ` +
      `a leitura compara ${lensText} com todas as entradas registradas. ` +
      `Use os chips para cruzar anos sem perder a visão mensal.`;

    els.heroLastBalance.textContent = currency.format(
      recent.balance
    );

    els.heroLastLabel.textContent =
      `${capitalize(recent.month)} ${recent.year} · ${
        recent.balance >= 0 ? "no azul" : "no vermelho"
      }`;

    els.heroAverageBalance.textContent = currency.format(
      metrics.averageBalance
    );

    els.heroAverageLabel.textContent =
      `${currency.format(metrics.balance)} acumulados`;

    els.heroEssentialShare.textContent = percent.format(
      metrics.essentialIncomeShare
    );

    els.heroEssentialLabel.textContent = "da renda média";
  }

  function renderKpis(rows, metrics) {
    if (!rows.length) {
      [
        "kpiBalance",
        "kpiIncomeAvg",
        "kpiSpendAvg",
        "kpiEssential",
        "kpiFlexible",
        "kpiPositive"
      ].forEach((id) => {
        els[id].textContent = "—";
      });

      els.kpiBalanceFoot.textContent = "Nenhum mês no recorte";
      els.kpiIncomeAvgFoot.textContent = "Nenhum mês no recorte";
      els.kpiSpendAvgFoot.textContent = "Nenhum mês no recorte";
      els.kpiEssentialFoot.textContent = "Nenhum mês no recorte";
      els.kpiFlexibleFoot.textContent = "Nenhum mês no recorte";
      els.kpiPositiveFoot.textContent = "Nenhum mês no recorte";

      return;
    }

    els.kpiBalance.textContent = currency.format(
      metrics.averageBalance
    );

    els.kpiBalanceFoot.textContent =
      `Saldo acumulado ${currency.format(metrics.balance)}`;

    els.kpiIncomeAvg.textContent = currency.format(
      metrics.averageIncome
    );

    els.kpiIncomeAvgFoot.textContent =
      `${incomeVolatilityLabel(metrics.incomeVolatility)} · ${
        state.incomeSources.length || 1
      } fonte(s)`;

    els.kpiSpendAvg.textContent = currency.format(
      metrics.averageSpend
    );

    els.kpiSpendAvgFoot.textContent =
      `${percent.format(
        metrics.averageIncome
          ? metrics.averageSpend / metrics.averageIncome
          : 0
      )} da entrada média`;

    els.kpiEssential.textContent = currency.format(
      metrics.averageGroups.essential
    );

    els.kpiEssentialFoot.textContent =
      `${percent.format(metrics.essentialIncomeShare)} da renda · ${
        percent.format(metrics.essentialSpendShare)
      } dos gastos`;

    els.kpiFlexible.textContent = currency.format(
      metrics.averageGroups.flexible
    );

    els.kpiFlexibleFoot.textContent =
      `Cenário -10%: libera ${currency.format(
        metrics.averageGroups.flexible * 0.1
      )} / mês`;

    els.kpiPositive.textContent =
      `${metrics.positive} / ${rows.length}`;

    els.kpiPositiveFoot.textContent =
      `${percent.format(metrics.positiveRate)} dos meses selecionados`;

    els.kpiBalance.classList.toggle(
      "is-positive",
      metrics.averageBalance >= 0
    );
  }

  function renderSnapshot(rows, metrics) {
    if (!rows.length) {
      els.snapshotStatus.textContent = "sem dados";
      els.snapshotStatus.classList.remove("is-negative");
      els.snapshotBalance.textContent = "—";
      els.snapshotLabel.textContent = "—";

      [
        "snapshotIncome",
        "snapshotEssential",
        "snapshotFlexible",
        "snapshotSpend"
      ].forEach((id) => {
        els[id].textContent = "—";
      });

      els.snapshotNote.textContent =
        "Amplie o recorte para obter um retrato mensal.";

      return;
    }

    const recent = rows[rows.length - 1];
    const difference = recent.balance - metrics.averageBalance;
    const statusPositive = recent.balance >= 0;

    els.snapshotStatus.textContent = statusPositive
      ? "no azul"
      : "no vermelho";

    els.snapshotStatus.classList.toggle(
      "is-negative",
      !statusPositive
    );

    els.snapshotBalance.textContent = currency.format(
      recent.balance
    );

    els.snapshotLabel.textContent =
      `${capitalize(recent.month)} ${recent.year}`;

    els.snapshotIncome.textContent = currency.format(
      recent.income
    );

    els.snapshotEssential.textContent = currency.format(
      recent.groups.essential
    );

    els.snapshotFlexible.textContent = currency.format(
      recent.groups.flexible
    );

    els.snapshotSpend.textContent = currency.format(
      recent.spend
    );

    els.snapshotNote.textContent = difference >= 0
      ? `${currency.format(
          Math.abs(difference)
        )} acima da sobra média do recorte.`
      : `${currency.format(
          Math.abs(difference)
        )} abaixo da sobra média do recorte.`;
  }

  function renderAllocation(metrics) {
    const total = metrics.averageSpend;

    if (!total) {
      els.allocationBar.innerHTML = "";

      els.allocationLegend.innerHTML =
        `<span class="active-filter is-empty">Sem gastos no recorte.</span>`;

      els.allocationCaption.textContent = "sem dados";

      return;
    }

    els.allocationCaption.textContent =
      currency.format(total) + " / mês";

    els.allocationBar.innerHTML = GROUP_ORDER.map((group) => {
      const width =
        clamp(metrics.averageGroups[group] / total, 0, 1) * 100;

      return `
        <span
          class="allocation-${group}"
          style="width:${width}%"
          title="${GROUP_RULES[group].label}: ${currency.format(
            metrics.averageGroups[group]
          )}"
        ></span>
      `;
    }).join("");

    els.allocationLegend.innerHTML = GROUP_ORDER.map((group) => {
      const share = total
        ? metrics.averageGroups[group] / total
        : 0;

      return `
        <div class="allocation-item">
          <div>
            <i class="${group}-color"></i>
            <span>${GROUP_RULES[group].shortLabel}</span>
          </div>
          <strong>
            ${currency.format(metrics.averageGroups[group])} ·
            ${percent.format(share)}
          </strong>
        </div>
      `;
    }).join("");

    els.allocationNote.textContent =
      filterState.lens === "all"
        ? "Essenciais são contas recorrentes; flexíveis são categorias que variam mais mês a mês."
        : `A lente selecionada destaca ${
            GROUP_RULES[filterState.lens].label.toLowerCase()
          } no gráfico de categorias.`;
  }

  function renderCharts(rows, metrics) {
    if (!window.Chart) {
      return;
    }

    if (!rows.length) {
      destroyCharts();
      return;
    }

    const labels = rows.map((row) => {
      return `${MONTH_LABELS[row.monthIndex]} ${String(
        row.year
      ).slice(2)}`;
    });

    updateChart("cashflow", $("cashflowChart"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          barDataset(
            "Essenciais",
            rows.map((row) => row.groups.essential),
            "rgba(255,193,120,.75)"
          ),
          barDataset(
            "Flexíveis",
            rows.map((row) => row.groups.flexible),
            "rgba(120,217,233,.72)"
          ),
          barDataset(
            "Outros",
            rows.map((row) => row.groups.other),
            "rgba(156,142,255,.68)"
          ),
          {
            type: "line",
            label: "Entradas",
            data: rows.map((row) => row.income),
            borderColor: "#f4f7ff",
            backgroundColor: "rgba(244,247,255,.12)",
            borderWidth: 2.2,
            tension: 0.35,
            pointRadius: 0,
            pointHoverRadius: 4,
            pointHoverBackgroundColor: "#f4f7ff",
            yAxisID: "income"
          }
        ]
      },
      options: cashflowOptions()
    });

    const focusEntries = metrics.focusCategories.slice(0, 8);

    els.categoryChartTitle.textContent =
      filterState.lens === "all"
        ? "Categorias que mais pesam"
        : `Categorias · ${GROUP_RULES[filterState.lens].label}`;

    updateChart("category", $("categoryChart"), {
      type: "bar",
      data: {
        labels: focusEntries.map((entry) =>
          labelize(entry.category)
        ),
        datasets: [
          {
            label: "Total gasto",
            data: focusEntries.map((entry) => entry.total),
            backgroundColor: focusEntries.map((entry, index) => {
              return filterState.lens === "all"
                ? palette[index % palette.length]
                : GROUP_RULES[filterState.lens].color;
            }),
            borderRadius: 5,
            maxBarThickness: 22
          }
        ]
      },
      options: categoryOptions()
    });

    const years = Object.keys(metrics.yearAverages);
    const yearGroups = Object.values(metrics.yearAverages);

    updateChart("year", $("yearChart"), {
      type: "bar",
      data: {
        labels: years,
        datasets: [
          barDataset(
            "Entrada média",
            yearGroups.map((group) => group.income),
            "rgba(120,217,233,.78)"
          ),
          barDataset(
            "Gasto médio",
            yearGroups.map((group) => group.spend),
            "rgba(255,131,180,.72)"
          ),
          barDataset(
            "Sobra média",
            yearGroups.map((group) => group.balance),
            "rgba(161,141,255,.82)"
          )
        ]
      },
      options: groupedBarOptions()
    });

    const monthGroups = metrics.monthAverages;

    const monthData = MONTHS.map((month, index) => {
      return monthGroups[index]
        ? monthGroups[index].balance
        : null;
    });

    updateChart("seasonality", $("seasonalityChart"), {
      type: "bar",
      data: {
        labels: MONTH_LABELS,
        datasets: [
          {
            label: "Saldo médio",
            data: monthData,
            backgroundColor: monthData.map((value) => {
              if (value === null) {
                return "rgba(255,255,255,.08)";
              }

              return value >= 0
                ? "rgba(95,225,174,.72)"
                : "rgba(255,140,157,.72)";
            }),
            borderRadius: 5,
            maxBarThickness: 22
          }
        ]
      },
      options: seasonalityOptions()
    });
  }

  function renderInsights(rows, metrics) {
    if (!rows.length) {
      els.insightsList.innerHTML = `
        <div class="insight-card">
          <strong>Sem dados nesse recorte</strong>
          <p>Amplie os anos, meses ou a situação do mês para continuar.</p>
        </div>
      `;

      return;
    }

    const best = rows.reduce((current, row) => {
      return row.balance > current.balance ? row : current;
    }, rows[0]);

    const worst = rows.reduce((current, row) => {
      return row.balance < current.balance ? row : current;
    }, rows[0]);

    const topCategory = metrics.categoryStats[0];
    const bestSeason = seasonEntry(metrics.monthAverages, "max");
    const worstSeason = seasonEntry(metrics.monthAverages, "min");

    const volatilityTone = metrics.incomeVolatility <= 0.15
      ? "tone-good"
      : metrics.incomeVolatility <= 0.3
        ? "tone-warn"
        : "tone-danger";

    const recentTone = metrics.recentSpendChange <= 0
      ? "tone-good"
      : metrics.recentSpendChange <= 0.1
        ? "tone-warn"
        : "tone-danger";

    const cards = [
      {
        tone: metrics.averageBalance >= 0
          ? "tone-good"
          : "tone-danger",
        icon: metrics.averageBalance >= 0 ? "↗" : "↘",
        title: "Sobra média",
        value: `${currency.format(metrics.averageBalance)} / mês`,
        text: metrics.averageBalance >= 0
          ? `A média deixa ${currency.format(
              metrics.averageBalance
            )} depois das despesas. No recorte inteiro, o saldo acumulado foi ${currency.format(
              metrics.balance
            )}.`
          : `A média termina ${currency.format(
              Math.abs(metrics.averageBalance)
            )} abaixo de zero. O primeiro ajuste deve olhar para recorrentes e flexíveis separadamente.`
      },
      {
        tone:
          metrics.essentialIncomeShare <= 0.6
            ? "tone-good"
            : metrics.essentialIncomeShare <= 0.75
              ? "tone-warn"
              : "tone-danger",
        icon: "⌂",
        title: "Contas da casa",
        value: `${percent.format(
          metrics.essentialIncomeShare
        )} da renda`,
        text: `${currency.format(
          metrics.averageGroups.essential
        )} por mês vão para essenciais. Isso equivale a ${percent.format(
          metrics.essentialSpendShare
        )} de todos os gastos.`
      },
      {
        tone: "tone-warn",
        icon: "≈",
        title: "Gastos flexíveis",
        value: `${currency.format(
          metrics.averageGroups.flexible
        )} / mês`,
        text: `Uma simulação de 10% menos nessa cesta liberaria ${currency.format(
          metrics.averageGroups.flexible * 0.1
        )} por mês — uma referência de cenário, não uma meta automática.`
      },
      {
        tone: "tone-danger",
        icon: "!",
        title: "Mês mais apertado",
        value: `${capitalize(worst.month)} ${
          worst.year
        } · ${currency.format(worst.balance)}`,
        text: `Foi o menor saldo do recorte. Nesse mês, entraram ${currency.format(
          worst.income
        )} e saíram ${currency.format(worst.spend)}.`
      },
      {
        tone: volatilityTone,
        icon: "∿",
        title: "Previsibilidade da renda",
        value: `${incomeVolatilityLabel(
          metrics.incomeVolatility
        )} · ${percent.format(metrics.incomeVolatility)}`,
        text: `A menor entrada foi ${currency.format(
          metrics.minIncome
        )} e a maior ${currency.format(
          metrics.maxIncome
        )}. ${
          metrics.extraordinaryIncomeCount
            ? `${metrics.extraordinaryIncomeCount} mês(es) ficaram bem acima da mediana.`
            : "Não há picos muito acima da mediana."
        }`
      },
      {
        tone: recentTone,
        icon: metrics.recentSpendChange <= 0 ? "↓" : "↑",
        title: "Movimento recente",
        value: `${
          metrics.recentSpendChange >= 0 ? "+" : ""
        }${percent.format(metrics.recentSpendChange)} em gastos`,
        text: `Comparando o terço final com o inicial, o gasto médio passou de ${currency.format(
          metrics.firstWindowSpend
        )} para ${currency.format(
          metrics.recentWindowSpend
        )} por mês. Melhor mês médio: ${
          bestSeason
            ? capitalize(MONTHS[bestSeason.index])
            : "—"
        }.`
      }
    ];

    if (topCategory && rows.length > 1) {
      cards[2].text += ` A categoria que mais pesa no total é ${labelize(
        topCategory.category
      )} (${percent.format(topCategory.share)}).`;
    }

    if (
      bestSeason &&
      worstSeason &&
      bestSeason.index !== worstSeason.index
    ) {
      cards[5].text += ` O calendário mais apertado foi ${capitalize(
        MONTHS[worstSeason.index]
      )}, com saldo médio de ${currency.format(
        worstSeason.value
      )}.`;
    }

    els.insightsList.innerHTML = cards
      .map((card) => {
        return `
          <article class="insight-card ${card.tone}">
            <div class="insight-top">
              <span class="insight-icon">${card.icon}</span>
              <span class="insight-title">${card.title}</span>
            </div>
            <strong>${escapeHtml(card.value)}</strong>
            <p>${escapeHtml(card.text)}</p>
          </article>
        `;
      })
      .join("");
  }

  function renderActions(rows, metrics) {
    if (!rows.length) {
      els.actionList.innerHTML =
        `<p class="empty-copy">Aguardando dados para montar as leituras práticas.</p>`;

      return;
    }

    const weakestIncomeRow = rows.reduce((current, row) => {
      return row.income < current.income ? row : current;
    }, rows[0]);

    const essentialBuffer =
      metrics.averageIncome - metrics.averageGroups.essential;

    const actions = [
      {
        title: "Conheça o piso da casa",
        text: `As contas essenciais ficam em torno de ${currency.format(
          metrics.averageGroups.essential
        )} por mês. Esse é o valor recorrente que precisa caber antes dos gastos flexíveis.`
      },
      {
        title: "Teste um cenário flexível",
        text: `Sem alterar a planilha, um corte simulado de 10% em flexíveis representa ${currency.format(
          metrics.averageGroups.flexible * 0.1
        )} por mês e ${currency.format(
          metrics.averageGroups.flexible * 0.1 * 12
        )} em um ano cheio.`
      },
      {
        title: "Proteja o mês de menor entrada",
        text: `A menor entrada foi ${currency.format(
          weakestIncomeRow.income
        )} em ${capitalize(weakestIncomeRow.month)} ${
          weakestIncomeRow.year
        }; sobrariam ${currency.format(
          weakestIncomeRow.income -
            metrics.averageGroups.essential
        )} depois dos essenciais médios.`
      }
    ];

    if (essentialBuffer < 0) {
      actions[0].title = "Atenção ao piso recorrente";

      actions[0].text = `Os essenciais médios (${currency.format(
        metrics.averageGroups.essential
      )}) já superam a entrada média em ${currency.format(
        Math.abs(essentialBuffer)
      )}. Vale separar esse sinal dos gastos flexíveis para entender o peso estrutural.`;
    }

    els.actionList.innerHTML = actions
      .map((action, index) => {
        return `
          <div class="action-item">
            <span class="action-number">${index + 1}</span>
            <div>
              <strong>${escapeHtml(action.title)}</strong>
              <p>${escapeHtml(action.text)}</p>
            </div>
          </div>
        `;
      })
      .join("");
  }

  function renderCategoryTable(metrics) {
    if (
      !metrics.categoryStats.length ||
      metrics.categoryStats.every((entry) => entry.total === 0)
    ) {
      els.categoryTableBody.innerHTML =
        `<tr><td colspan="5">Nenhuma categoria encontrada.</td></tr>`;

      return;
    }

    els.categoryTableBody.innerHTML = metrics.categoryStats
      .slice(0, 10)
      .map((entry) => {
        const trendClass = entry.trend > 0.05
          ? "trend-up"
          : entry.trend < -0.05
            ? "trend-down"
            : "trend-flat";

        const trendLabel = entry.trend > 0.05
          ? `↑ ${percent.format(entry.trend)}`
          : entry.trend < -0.05
            ? `↓ ${percent.format(Math.abs(entry.trend))}`
            : "estável";

        return `
          <tr>
            <td>${escapeHtml(labelize(entry.category))}</td>
            <td>
              <span class="group-badge ${entry.group}">
                ${GROUP_RULES[entry.group].shortLabel}
              </span>
            </td>
            <td>${currency.format(entry.average)}</td>
            <td>${percent.format(entry.share)}</td>
            <td class="${trendClass}">${trendLabel}</td>
          </tr>
        `;
      })
      .join("");
  }

  function renderTable(rows) {
    const sorted = [...rows].sort(
      (first, second) => second.date - first.date
    );

    els.tableCaption.textContent = `${rows.length} ${
      rows.length === 1 ? "linha" : "linhas"
    }`;

    if (!sorted.length) {
      els.dataTableBody.innerHTML =
        `<tr><td colspan="8">Nenhum mês corresponde aos filtros.</td></tr>`;

      return;
    }

    els.dataTableBody.innerHTML = sorted
      .map((row) => {
        const balanceClass =
          row.balance >= 0 ? "positive" : "negative";

        const statusLabel =
          row.balance >= 0 ? "No azul" : "No vermelho";

        return `
          <tr>
            <td>${capitalize(row.month)} ${row.year}</td>
            <td class="amount">${currency.format(row.income)}</td>
            <td class="amount">${currency.format(
              row.groups.essential
            )}</td>
            <td class="amount">${currency.format(
              row.groups.flexible
            )}</td>
            <td class="amount">${currency.format(
              row.groups.other
            )}</td>
            <td class="amount">${currency.format(row.spend)}</td>
            <td class="${balanceClass}">${currency.format(
              row.balance
            )}</td>
            <td class="${balanceClass}">
              <span class="status-pill">${statusLabel}</span>
            </td>
          </tr>
        `;
      })
      .join("");
  }

  function renderEmpty() {
    els.filterResult.textContent = "0 meses";
    els.periodLabel.textContent = "Aguardando uma planilha";

    els.heroTitle.textContent =
      "Carregue uma tabela para começar.";

    els.heroText.textContent =
      "O painel vai separar contas essenciais de gastos flexíveis e permitir comparar vários anos de uma só vez.";

    els.heroLastBalance.textContent = "—";
    els.heroLastLabel.textContent = "—";
    els.heroAverageBalance.textContent = "—";
    els.heroAverageLabel.textContent = "—";
    els.heroEssentialShare.textContent = "—";
    els.heroEssentialLabel.textContent = "da renda média";

    [
      "kpiBalance",
      "kpiIncomeAvg",
      "kpiSpendAvg",
      "kpiEssential",
      "kpiFlexible",
      "kpiPositive",
      "snapshotBalance",
      "snapshotIncome",
      "snapshotEssential",
      "snapshotFlexible",
      "snapshotSpend"
    ].forEach((id) => {
      els[id].textContent = "—";
    });

    els.insightsList.innerHTML = `
      <article class="insight-card">
        <div class="insight-top">
          <span class="insight-icon">↥</span>
          <span class="insight-title">Comece por aqui</span>
        </div>
        <strong>Carregue uma tabela</strong>
        <p>Use a área “Trocar tabela” para importar os dados da casa.</p>
      </article>
    `;

    els.actionList.innerHTML =
      `<p class="empty-copy">As leituras práticas aparecem depois da importação.</p>`;

    els.categoryTableBody.innerHTML =
      `<tr><td colspan="5">Aguardando dados…</td></tr>`;

    els.dataTableBody.innerHTML =
      `<tr><td colspan="8">Aguardando dados…</td></tr>`;

    els.allocationBar.innerHTML = "";
    els.allocationLegend.innerHTML = "";

    els.snapshotStatus.textContent = "sem dados";
    els.snapshotStatus.classList.remove("is-negative");
    els.snapshotLabel.textContent = "—";

    els.snapshotNote.textContent =
      "Importe uma planilha para obter um retrato mensal.";

    renderFilterControls();
    destroyCharts();
  }

  function updateSidebarLegend() {
    const groups = Object.fromEntries(
      GROUP_ORDER.map((group) => {
        return [
          group,
          state.categories
            .filter(
              (category) =>
                state.categoryGroups[category] === group
            )
            .map(labelize)
        ];
      })
    );

    els.essentialLegend.textContent = groups.essential.length
      ? groups.essential.join(", ")
      : "Nenhuma coluna identificada";

    els.flexibleLegend.textContent = groups.flexible.length
      ? groups.flexible.join(", ")
      : "Nenhuma coluna identificada";

    els.otherLegend.textContent = groups.other.length
      ? groups.other.join(", ")
      : "Nenhuma coluna identificada";

    els.classificationNote.textContent = state.ignoredRows
      ? `${state.ignoredRows} linha(s) sem mês/ano foram ignoradas durante a leitura.`
      : "A classificação é inferida pelo nome das colunas e usada para orientar a leitura.";
  }

  function updateQuickRead(rows, metrics) {
    if (!rows.length) {
      els.quickRead.textContent =
        "Amplie o recorte para continuar a leitura.";

      return;
    }

    if (metrics.essentialIncomeShare > 0.75) {
      els.quickRead.textContent =
        `Os essenciais ocupam ${percent.format(
          metrics.essentialIncomeShare
        )} da renda média; esse é o primeiro número para acompanhar.`;

      return;
    }

    if (metrics.recentSpendChange > 0.1) {
      els.quickRead.textContent =
        `O gasto médio recente subiu ${percent.format(
          metrics.recentSpendChange
        )} contra o início do recorte.`;

      return;
    }

    els.quickRead.textContent = metrics.averageBalance >= 0
      ? "A média termina no azul; veja se a sobra aparece também nos meses de menor entrada."
      : "A média termina no vermelho; compare primeiro os meses de menor entrada e os gastos flexíveis.";
  }

  function updateChart(name, canvas, config) {
    if (state.charts[name]) {
      state.charts[name].destroy();
    }

    state.charts[name] = new Chart(canvas, config);
  }

  function destroyCharts() {
    Object.values(state.charts).forEach((chart) => {
      chart.destroy();
    });

    state.charts = {};
  }

  function cashflowOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      scales: {
        x: {
          stacked: true,
          grid: {
            display: false
          },
          ticks: {
            maxRotation: 0,
            autoSkip: true,
            maxTicksLimit: 10,
            color: "#76829e"
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          grid: {
            color: "rgba(255,255,255,.07)",
            drawBorder: false
          },
          ticks: {
            color: "#76829e",
            maxTicksLimit: 5,
            callback: compactCurrency
          }
        },
        income: {
          position: "right",
          beginAtZero: true,
          grid: {
            drawOnChartArea: false
          },
          ticks: {
            color: "#c5cce2",
            maxTicksLimit: 5,
            callback: compactCurrency
          }
        }
      },
      plugins: {
        tooltip: {
          callbacks: {
            label(context) {
              return ` ${context.dataset.label}: ${currency.format(
                context.raw || 0
              )}`;
            }
          }
        }
      }
    };
  }

  function categoryOptions() {
    return {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          beginAtZero: true,
          grid: {
            color: "rgba(255,255,255,.07)",
            drawBorder: false
          },
          ticks: {
            color: "#76829e",
            maxTicksLimit: 5,
            callback: compactCurrency
          }
        },
        y: {
          grid: {
            display: false
          },
          ticks: {
            color: "#a9b4cf",
            autoSkip: false
          }
        }
      },
      plugins: {
        tooltip: {
          callbacks: {
            label(context) {
              return ` ${currency.format(context.raw || 0)}`;
            }
          }
        }
      }
    };
  }

  function groupedBarOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          grid: {
            display: false
          },
          ticks: {
            color: "#76829e"
          }
        },
        y: {
          beginAtZero: true,
          grid: {
            color: "rgba(255,255,255,.07)",
            drawBorder: false
          },
          ticks: {
            color: "#76829e",
            maxTicksLimit: 5,
            callback: compactCurrency
          }
        }
      },
      plugins: {
        tooltip: {
          callbacks: {
            label(context) {
              return ` ${context.dataset.label}: ${currency.format(
                context.raw || 0
              )}`;
            }
          }
        }
      }
    };
  }

  function seasonalityOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          grid: {
            display: false
          },
          ticks: {
            color: "#76829e"
          }
        },
        y: {
          beginAtZero: false,
          grid: {
            color: "rgba(255,255,255,.07)",
            drawBorder: false
          },
          ticks: {
            color: "#76829e",
            maxTicksLimit: 5,
            callback: compactCurrency
          }
        }
      },
      plugins: {
        tooltip: {
          callbacks: {
            label(context) {
              return ` Saldo médio: ${currency.format(
                context.raw || 0
              )}`;
            }
          }
        }
      }
    };
  }

  function barDataset(label, data, backgroundColor) {
    return {
      label,
      data,
      backgroundColor,
      borderRadius: 5,
      maxBarThickness: 20,
      stack: "expenses"
    };
  }

  function getFocusCategories(categoryStats) {
    if (filterState.lens === "all") {
      return categoryStats;
    }

    return categoryStats.filter((entry) => {
      return entry.group === filterState.lens;
    });
  }

  function groupByYearAverage(rows) {
    const groups = {};

    rows.forEach((row) => {
      if (!groups[row.year]) {
        groups[row.year] = {
          months: 0,
          income: 0,
          spend: 0,
          balance: 0
        };
      }

      groups[row.year].months += 1;
      groups[row.year].income += row.income;
      groups[row.year].spend += row.spend;
      groups[row.year].balance += row.balance;
    });

    Object.values(groups).forEach((group) => {
      group.income /= group.months;
      group.spend /= group.months;
      group.balance /= group.months;
    });

    return Object.fromEntries(
      Object.entries(groups).sort(
        ([first], [second]) => Number(first) - Number(second)
      )
    );
  }

  function groupByMonthAverage(rows) {
    const groups = {};

    rows.forEach((row) => {
      if (!groups[row.monthIndex]) {
        groups[row.monthIndex] = {
          months: 0,
          balance: 0,
          spend: 0,
          income: 0
        };
      }

      groups[row.monthIndex].months += 1;
      groups[row.monthIndex].balance += row.balance;
      groups[row.monthIndex].spend += row.spend;
      groups[row.monthIndex].income += row.income;
    });

    Object.values(groups).forEach((group) => {
      group.balance /= group.months;
      group.spend /= group.months;
      group.income /= group.months;
    });

    return groups;
  }

  function seasonEntry(monthAverages, mode) {
    const entries = Object.entries(monthAverages).map(
      ([index, values]) => {
        return {
          index: Number(index),
          value: values.balance
        };
      }
    );

    if (!entries.length) {
      return null;
    }

    return entries.reduce((current, entry) => {
      if (mode === "max") {
        return entry.value > current.value
          ? entry
          : current;
      }

      return entry.value < current.value
        ? entry
        : current;
    }, entries[0]);
  }

  function summarizeGroups(values, categoryGroups) {
    const groups = {
      essential: 0,
      flexible: 0,
      other: 0
    };

    Object.entries(values).forEach(([category, value]) => {
      const group = categoryGroups[category] || "other";
      groups[group] += value;
    });

    return groups;
  }

  function findLeader(values) {
    const entries = Object.entries(values);

    if (!entries.length) {
      return "";
    }

    return entries.sort((first, second) => {
      return second[1] - first[1];
    })[0][0];
  }

  function inferGroup(category) {
    const normalized = normalizeKey(category);

    if (
      GROUP_RULES.essential.aliases.some((alias) =>
        normalized.includes(normalizeKey(alias))
      )
    ) {
      return "essential";
    }

    if (
      GROUP_RULES.flexible.aliases.some((alias) =>
        normalized.includes(normalizeKey(alias))
      )
    ) {
      return "flexible";
    }

    return "other";
  }

  function incomeVolatilityLabel(value) {
    if (value <= 0.15) {
      return "renda previsível";
    }

    if (value <= 0.3) {
      return "renda variável";
    }

    return "renda irregular";
  }

  function buildFooterMethod() {
    const sourceLabel = state.incomeSources.length
      ? state.incomeSources.map(labelize).join(" + ")
      : "fontes de renda identificadas";

    return `Entradas = ${sourceLabel} · gastos = categorias · sobra = entradas − gastos.`;
  }

  function setSourceState(status, name, meta) {
    els.sourceName.textContent = name;
    els.sourceMeta.textContent = meta;
    els.dataStatusDot.classList.toggle(
      "is-ready",
      status === "ready"
    );
  }

  function showToast(message, isError = false) {
    els.toast.textContent = message;
    els.toast.classList.toggle("error", isError);
    els.toast.classList.add("show");

    window.clearTimeout(showToast.timer);

    showToast.timer = window.setTimeout(() => {
      els.toast.classList.remove("show");
    }, 3200);
  }

  function exportFilteredCsv() {
    if (!state.filtered.length) {
      showToast(
        "Não há dados filtrados para exportar",
        true
      );

      return;
    }

    const headers = [
      "Mês",
      "Ano",
      "Entradas",
      "Essenciais",
      "Flexíveis",
      "Outros",
      "Total Gasto",
      "Balança",
      "Taxa de poupança",
      ...state.categories.map(labelize)
    ];

    const dataRows = state.filtered.map((row) => {
      return [
        capitalize(row.month),
        row.year,
        row.income,
        row.groups.essential,
        row.groups.flexible,
        row.groups.other,
        row.spend,
        row.balance,
        row.savingsRate,
        ...state.categories.map(
          (category) => row.values[category]
        )
      ];
    });

    const lines = [headers, ...dataRows].map((line) => {
      return line.map(csvEscape).join(";");
    });

    const blob = new Blob(
      ["\uFEFF" + lines.join("\n")],
      {
        type: "text/csv;charset=utf-8"
      }
    );

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");

    anchor.href = url;
    anchor.download =
      "financas-da-casa-visao-filtrada.csv";

    anchor.click();

    URL.revokeObjectURL(url);
    showToast("CSV exportado");
  }

  function csvEscape(value) {
    const text = String(
      value == null ? "" : value
    ).replace(/"/g, '""');

    return /[;"\n]/.test(text)
      ? `"${text}"`
      : text;
  }

  function monthIndexOf(value) {
    const normalized = normalizeText(value);

    if (!normalized) {
      return -1;
    }

    const direct = MONTHS
      .map(normalizeText)
      .indexOf(normalized);

    if (direct >= 0) {
      return direct;
    }

    const numeric = Number(value);

    return numeric >= 1 && numeric <= 12
      ? numeric - 1
      : -1;
  }

  function findKey(keyInfo, aliases) {
    const normalizedAliases = aliases.map(normalizeKey);

    const exact = keyInfo.find(({ id }) =>
      normalizedAliases.includes(id)
    );

    if (exact) {
      return exact.key;
    }

    const partial = keyInfo.find(({ id }) => {
      return normalizedAliases.some((alias) =>
        id.includes(alias)
      );
    });

    return partial ? partial.key : null;
  }

  function normalizeKey(value) {
    return normalizeText(value).replace(
      /[^a-z0-9]/g,
      ""
    );
  }

  function normalizeText(value) {
    return String(value == null ? "" : value)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  function labelize(value) {
    return String(value || "")
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/\b\w/g, (letter) =>
        letter.toUpperCase()
      );
  }

  function capitalize(value) {
    const text = String(value || "");

    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  function isNumeric(value) {
    if (
      value === null ||
      value === "" ||
      value === undefined
    ) {
      return false;
    }

    if (typeof value === "number") {
      return Number.isFinite(value);
    }

    const text = String(value).trim();

    if (!text) {
      return false;
    }

    return (
      /^[-+]?\(?\s*R?\$?\s*[\d.]+(?:,\d+)?\s*\)?$/.test(
        text
      ) ||
      /^[-+]?\(?\s*\d+(?:\.\d+)?\s*\)?$/.test(text)
    );
  }

  function toNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : 0;
    }

    let text = String(
      value == null ? "" : value
    ).trim();

    if (!text) {
      return 0;
    }

    const negative = /^\(.*\)$/.test(text);

    text = text.replace(/[()R$\s]/gi, "");

    if (text.includes(",") && text.includes(".")) {
      text = text
        .replace(/\./g, "")
        .replace(",", ".");
    } else {
      text = text.replace(",", ".");
    }

    const parsed = Number(
      text.replace(/[^\d.-]/g, "")
    );

    const safeNumber = Number.isFinite(parsed)
      ? parsed
      : 0;

    return (negative ? -1 : 1) * safeNumber;
  }

  function compactCurrency(value) {
    const abs = Math.abs(value);
    const sign = value < 0 ? "−" : "";

    if (abs >= 1000000) {
      return `${sign}R$ ${(abs / 1000000)
        .toFixed(1)
        .replace(".", ",")} mi`;
    }

    if (abs >= 1000) {
      return `${sign}R$ ${(abs / 1000)
        .toFixed(1)
        .replace(".", ",")} mil`;
    }

    return `${sign}R$ ${Math.round(abs)}`;
  }

  function hexToRgba(hex, alpha) {
    const clean = String(hex).replace("#", "");

    const value = clean.length === 3
      ? clean
          .split("")
          .map((part) => part + part)
          .join("")
      : clean;

    const red = parseInt(value.slice(0, 2), 16);
    const green = parseInt(value.slice(2, 4), 16);
    const blue = parseInt(value.slice(4, 6), 16);

    return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
  }

  function uniqueSorted(values) {
    return [...new Set(values)].sort(
      (first, second) => first - second
    );
  }

  function sum(items, getter) {
    return items.reduce(
      (total, item) => total + getter(item),
      0
    );
  }

  function average(items, getter) {
    return items.length
      ? sum(items, getter) / items.length
      : 0;
  }

  function median(values) {
    if (!values.length) {
      return 0;
    }

    const sorted = [...values].sort(
      (first, second) => first - second
    );

    const middle = Math.floor(sorted.length / 2);

    return sorted.length % 2
      ? sorted[middle]
      : (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function standardDeviation(values) {
    if (!values.length) {
      return 0;
    }

    const mean =
      values.reduce(
        (total, value) => total + value,
        0
      ) / values.length;

    const variance =
      values.reduce((total, value) => {
        return total + (value - mean) ** 2;
      }, 0) / values.length;

    return Math.sqrt(variance);
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

    return `${capitalize(first.month)} ${first.year} — ${capitalize(
      last.month
    )} ${last.year}`;
  }

  function escapeHtml(value) {
    return String(
      value == null ? "" : value
    ).replace(/[&<>'"]/g, (character) => {
      const entities = {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "'": "&#39;",
        '"': "&quot;"
      };

      return entities[character];
    });
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }
})();
