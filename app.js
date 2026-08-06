/*
 * Finanças em foco
 * - Carrega automaticamente uma planilha na raiz do site.
 * - Permite substituir a fonte por upload/drag-and-drop.
 * - Calcula métricas a partir das linhas importadas.
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
    "./Finanças.xlsx",
    "./Finanças.xlsx",
    "./Financas.xlsx",
    "./financas.xlsx",
    "./upload/Finanças.xlsx"
  ];

  const FIXED_ALIASES = [
    "aluguel",
    "condominio",
    "condomínio",
    "luz",
    "gas",
    "gás",
    "internet"
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
    "gasto total",
    "despesas totais",
    "balanca",
    "saldo",
    "balance",
    "poupanca",
    "taxa de poupanca"
  ];

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
    incomeSources: [],
    filtered: [],
    charts: {},
    ignoredRows: 0
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
      "rowCountTag",
      "categoryCountTag",
      "yearFilter",
      "monthFilter",
      "statusFilter",
      "leaderFilter",
      "filterResult",
      "resetFilters",
      "exportCsv",
      "kpiIncome",
      "kpiIncomeFoot",
      "kpiSpend",
      "kpiSpendFoot",
      "kpiBalance",
      "kpiBalanceFoot",
      "kpiSavings",
      "kpiSavingsFoot",
      "kpiFixed",
      "kpiFixedFoot",
      "kpiPositive",
      "kpiPositiveFoot",
      "cashflowNote",
      "compositionLegend",
      "insightsList",
      "healthBadge",
      "healthScore",
      "healthLabel",
      "healthDescription",
      "healthPositive",
      "healthPositiveBar",
      "healthSavings",
      "healthSavingsBar",
      "healthEssential",
      "healthEssentialBar",
      "dataTableBody",
      "tableCaption",
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
      els.yearFilter,
      els.monthFilter,
      els.statusFilter,
      els.leaderFilter
    ].forEach((element) => {
      element.addEventListener("change", renderAll);
    });

    els.resetFilters.addEventListener("click", () => {
      els.yearFilter.value = "all";
      els.monthFilter.value = "all";
      els.statusFilter.value = "all";
      els.leaderFilter.value = "all";

      renderAll();
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
      const file =
        event.dataTransfer.files && event.dataTransfer.files[0];

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
    Chart.defaults.animation.duration = 550;
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

  async function tryDefaultFiles() {
    setSourceState(
      "loading",
      "Carregando arquivo da raiz…",
      "Tentando encontrar sua planilha"
    );

    for (const path of DEFAULT_FILES) {
      try {
        const response = await fetch(path, {
          cache: "no-store"
        });

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
        /*
         * Arquivos abertos via file:// ou servidores sem a
         * planilha caem automaticamente no upload manual.
         */
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
      showToast(
        "Escolha um arquivo .xlsx, .xls ou .csv",
        true
      );
      return;
    }

    try {
      setSourceState(
        "loading",
        file.name,
        "Lendo e normalizando os dados…"
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
    state.incomeSources = normalized.incomeSources;
    state.sourceName = sourceName;
    state.ignoredRows = normalized.ignoredRows;

    populateFilters();

    setSourceState(
      "ready",
      sourceName,
      `${state.rows.length} meses · ${state.categories.length} categorias`
    );

    els.lastUpdated.textContent =
      `Atualizado às ${
        new Intl.DateTimeFormat("pt-BR", {
          hour: "2-digit",
          minute: "2-digit"
        }).format(new Date())
      }`;

    renderAll();
  }

  function normalizeRows(rawRows) {
    if (!Array.isArray(rawRows) || !rawRows.length) {
      return {
        rows: [],
        categories: [],
        incomeSources: [],
        ignoredRows: 0
      };
    }

    const sample = rawRows[0];
    const keys = Object.keys(sample || {});

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
      throw new Error(
        "A planilha precisa ter colunas de mês e ano"
      );
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
        const isExcluded = excluded.has(key);

        const isSummary = SUMMARY_ALIASES.some((alias) => {
          return id === normalizeKey(alias);
        });

        const hasNumericValues = rawRows.some((row) => {
          return isNumeric(row[key]);
        });

        return !isExcluded && !isSummary && hasNumericValues;
      })
      .map(({ key }) => key);

    const uniqueCategories = [...new Set(categories)];
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

      const spend =
        totalKey && isNumeric(raw[totalKey])
          ? toNumber(raw[totalKey])
          : uniqueCategories.reduce((sum, category) => {
              return sum + values[category];
            }, 0);

      const income = incomeSources.reduce((sum, source) => {
        return sum + toNumber(raw[source]);
      }, 0);

      const balance =
        balanceKey && isNumeric(raw[balanceKey])
          ? toNumber(raw[balanceKey])
          : income - spend;

      const date = new Date(year, monthIndex, 1);

      rows.push({
        id: `${year}-${String(monthIndex + 1).padStart(2, "0")}-${index}`,
        month: MONTHS[monthIndex],
        monthLabel: capitalize(MONTHS[monthIndex]),
        monthIndex,
        year,
        date,
        values,
        spend,
        income,
        balance,
        savingsRate: income ? balance / income : 0,
        leader: findLeader(values)
      });
    });

    rows.sort((first, second) => {
      return first.date - second.date;
    });

    return {
      rows,
      categories: uniqueCategories,
      incomeSources,
      ignoredRows
    };
  }

  function populateFilters() {
    const years = [
      ...new Set(
        state.rows.map((row) => row.year)
      )
    ].sort((first, second) => first - second);

    els.yearFilter.innerHTML = `
      <option value="all">Todos os anos</option>
      ${years
        .map((year) => {
          return `<option value="${year}">${year}</option>`;
        })
        .join("")}
    `;

    els.monthFilter.innerHTML = `
      <option value="all">Todos os meses</option>
      ${MONTHS
        .map((month, index) => {
          return `
            <option value="${index}">
              ${capitalize(month)}
            </option>
          `;
        })
        .join("")}
    `;

    els.leaderFilter.innerHTML = `
      <option value="all">Todas as categorias</option>
      ${state.categories
        .map((category) => {
          return `
            <option value="${escapeHtml(category)}">
              ${escapeHtml(labelize(category))}
            </option>
          `;
        })
        .join("")}
    `;

    els.yearFilter.value = "all";
    els.monthFilter.value = "all";
    els.statusFilter.value = "all";
    els.leaderFilter.value = "all";
  }

  function renderAll() {
    if (!state.rows.length) {
      renderEmpty();
      return;
    }

    const filtered = getFilteredRows();

    state.filtered = filtered;

    const metrics = aggregate(filtered);

    els.filterResult.textContent =
      `${filtered.length} de ${state.rows.length} meses`;

    els.periodLabel.textContent =
      getPeriodLabel(filtered);

    els.rowCountTag.textContent =
      number.format(filtered.length);

    els.categoryCountTag.textContent =
      number.format(state.categories.length);

    renderKpis(metrics, filtered);
    renderCharts(filtered, metrics);
    renderInsights(filtered, metrics);
    renderHealth(metrics, filtered);
    renderTable(filtered);
  }

  function getFilteredRows() {
    const year = els.yearFilter.value;
    const month = els.monthFilter.value;
    const status = els.statusFilter.value;
    const leader = els.leaderFilter.value;

    return state.rows.filter((row) => {
      const yearMatches =
        year === "all" ||
        row.year === Number(year);

      const monthMatches =
        month === "all" ||
        row.monthIndex === Number(month);

      const statusMatches =
        status === "all" ||
        (
          status === "positive"
            ? row.balance > 0
            : row.balance < 0
        );

      const leaderMatches =
        leader === "all" ||
        row.leader === leader;

      return (
        yearMatches &&
        monthMatches &&
        statusMatches &&
        leaderMatches
      );
    });
  }

  function aggregate(rows) {
    const income = rows.reduce((sum, row) => {
      return sum + row.income;
    }, 0);

    const spend = rows.reduce((sum, row) => {
      return sum + row.spend;
    }, 0);

    const balance = rows.reduce((sum, row) => {
      return sum + row.balance;
    }, 0);

    const categoryTotals = Object.fromEntries(
      state.categories.map((category) => {
        const total = rows.reduce((sum, row) => {
          return sum + (row.values[category] || 0);
        }, 0);

        return [category, total];
      })
    );

    const fixedCategories = state.categories.filter((category) => {
      return FIXED_ALIASES.some((alias) => {
        return normalizeKey(category).includes(
          normalizeKey(alias)
        );
      });
    });

    const fixed = fixedCategories.reduce((sum, category) => {
      return sum + categoryTotals[category];
    }, 0);

    const positive = rows.filter((row) => {
      return row.balance > 0;
    }).length;

    const negative = rows.filter((row) => {
      return row.balance < 0;
    }).length;

    const savingsRate =
      income ? balance / income : 0;

    const essentialRatio =
      spend ? fixed / spend : 0;

    const positiveRate =
      rows.length ? positive / rows.length : 0;

    return {
      income,
      spend,
      balance,
      categoryTotals,
      fixed,
      fixedCategories,
      positive,
      negative,
      savingsRate,
      essentialRatio,
      positiveRate
    };
  }

  function renderKpis(metrics, rows) {
    els.kpiIncome.textContent =
      currency.format(metrics.income);

    els.kpiSpend.textContent =
      currency.format(metrics.spend);

    els.kpiBalance.textContent =
      currency.format(metrics.balance);

    els.kpiBalance.classList.toggle(
      "is-positive",
      metrics.balance >= 0
    );

    els.kpiSavings.textContent =
      percent.format(metrics.savingsRate);

    els.kpiFixed.textContent =
      currency.format(metrics.fixed);

    els.kpiPositive.textContent =
      `${metrics.positive} / ${rows.length}`;

    els.kpiIncomeFoot.textContent =
      `${rows.length
        ? currency.format(metrics.income / rows.length)
        : "—"
      } por mês`;

    els.kpiSpendFoot.textContent =
      `${rows.length
        ? currency.format(metrics.spend / rows.length)
        : "—"
      } por mês`;

    els.kpiBalanceFoot.textContent =
      metrics.balance >= 0
        ? "Saldo acumulado positivo"
        : "Saldo acumulado negativo";

    els.kpiSavingsFoot.textContent =
      metrics.savingsRate >= 0
        ? "Margem disponível no período"
        : "Gastos acima das entradas";

    els.kpiFixedFoot.textContent =
      metrics.spend
        ? `${percent.format(metrics.essentialRatio)} dos gastos`
        : "Sem despesas no período";

    els.kpiPositiveFoot.textContent =
      rows.length
        ? `${percent.format(metrics.positiveRate)} dos meses`
        : "Sem meses selecionados";

    els.cashflowNote.textContent =
      rows.length
        ? `${rows.length} pontos no tempo`
        : "Sem dados";
  }

  function renderCharts(rows, metrics) {
    if (!window.Chart) {
      return;
    }

    const labels = rows.map((row) => {
      return `${MONTH_LABELS[row.monthIndex]} ${String(row.year).slice(2)}`;
    });

    updateChart(
      "cashflow",
      $("cashflowChart"),
      {
        type: "line",
        data: {
          labels,
          datasets: [
            lineDataset(
              "Entradas",
              rows.map((row) => row.income),
              "#78d9e9",
              true
            ),
            lineDataset(
              "Gastos",
              rows.map((row) => row.spend),
              "#ff83b4",
              true
            ),
            lineDataset(
              "Balança",
              rows.map((row) => row.balance),
              "#a18dff",
              false
            )
          ]
        },
        options: chartOptions("currency", false)
      }
    );

    const categoryEntries = Object
      .entries(metrics.categoryTotals)
      .sort((first, second) => {
        return second[1] - first[1];
      });

    updateChart(
      "composition",
      $("compositionChart"),
      {
        type: "doughnut",
        data: {
          labels: categoryEntries.map(([category]) => {
            return labelize(category);
          }),
          datasets: [
            {
              data: categoryEntries.map(([, value]) => value),
              backgroundColor: palette,
              borderColor: "#111b39",
              borderWidth: 3,
              hoverOffset: 5
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: "69%",
          plugins: {
            tooltip: {
              callbacks: {
                label(context) {
                  const share = metrics.spend
                    ? context.raw / metrics.spend
                    : 0;

                  return (
                    ` ${context.label}: ` +
                    `${currency.format(context.raw)} ` +
                    `(${percent.format(share)})`
                  );
                }
              }
            }
          }
        }
      }
    );

    els.compositionLegend.innerHTML = categoryEntries
      .slice(0, 6)
      .map(([category, value], index) => {
        const share =
          metrics.spend
            ? value / metrics.spend
            : 0;

        return `
          <div class="legend-item">
            <i
              class="legend-swatch"
              style="background:${palette[index % palette.length]}"
            ></i>
            <span class="legend-name">
              ${escapeHtml(labelize(category))}
            </span>
            <span class="legend-value">
              ${percent.format(share)}
            </span>
          </div>
        `;
      })
      .join("");

    const byYear = groupByYear(rows);
    const yearlyGroups = Object.values(byYear);

    updateChart(
      "year",
      $("yearChart"),
      {
        type: "bar",
        data: {
          labels: Object.keys(byYear),
          datasets: [
            {
              label: "Entradas",
              data: yearlyGroups.map((group) => group.income),
              backgroundColor: "rgba(120,217,233,.78)",
              borderRadius: 5,
              maxBarThickness: 18
            },
            {
              label: "Gastos",
              data: yearlyGroups.map((group) => group.spend),
              backgroundColor: "rgba(255,131,180,.72)",
              borderRadius: 5,
              maxBarThickness: 18
            },
            {
              label: "Balança",
              data: yearlyGroups.map((group) => group.balance),
              backgroundColor: "rgba(161,141,255,.82)",
              borderRadius: 5,
              maxBarThickness: 18
            }
          ]
        },
        options: chartOptions("currency", false)
      }
    );

    const byMonth = MONTHS.map((month, index) => {
      const sameMonth = rows.filter((row) => {
        return row.monthIndex === index;
      });

      if (!sameMonth.length) {
        return null;
      }

      return sameMonth.reduce((sum, row) => {
        return sum + row.balance;
      }, 0) / sameMonth.length;
    });

    updateChart(
      "seasonality",
      $("seasonalityChart"),
      {
        type: "bar",
        data: {
          labels: MONTH_LABELS,
          datasets: [
            {
              label: "Saldo médio",
              data: byMonth,
              backgroundColor: byMonth.map((value) => {
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
        options: chartOptions("currency", false)
      }
    );
  }

  function lineDataset(label, data, color, fill) {
    return {
      label,
      data,
      borderColor: color,
      backgroundColor: hexToRgba(color, 0.12),
      borderWidth: 2.2,
      tension: 0.36,
      fill,
      pointRadius: 0,
      pointHoverRadius: 4,
      pointHoverBackgroundColor: color,
      pointHoverBorderColor: "#fff",
      pointHoverBorderWidth: 2
    };
  }

  function chartOptions(valueType, stacked) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      scales: {
        x: {
          stacked,
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
          stacked,
          beginAtZero: false,
          grid: {
            color: "rgba(255,255,255,.07)",
            drawBorder: false
          },
          ticks: {
            color: "#76829e",
            maxTicksLimit: 5,
            callback(value) {
              return valueType === "currency"
                ? compactCurrency(value)
                : value;
            }
          }
        }
      },
      plugins: {
        tooltip: {
          callbacks: {
            label(context) {
              return (
                ` ${context.dataset.label}: ` +
                currency.format(context.raw || 0)
              );
            }
          }
        }
      }
    };
  }

  function updateChart(name, canvas, config) {
    if (state.charts[name]) {
      state.charts[name].destroy();
    }

    state.charts[name] = new Chart(canvas, config);
  }

  function renderInsights(rows, metrics) {
    if (!rows.length) {
      els.insightsList.innerHTML = `
        <div class="insight-card">
          <strong>Sem dados nesse recorte</strong>
          <p>
            Amplie os filtros para voltar a encontrar padrões.
          </p>
        </div>
      `;
      return;
    }

    const best = rows.reduce((current, row) => {
      return row.balance > current.balance
        ? row
        : current;
    }, rows[0]);

    const worst = rows.reduce((current, row) => {
      return row.balance < current.balance
        ? row
        : current;
    }, rows[0]);

    const topCategory = Object
      .entries(metrics.categoryTotals)
      .sort((first, second) => {
        return second[1] - first[1];
      })[0];

    const first = rows[0];
    const last = rows[rows.length - 1];

    const spendChange = first.spend
      ? (last.spend - first.spend) / first.spend
      : 0;

    const streak = longestPositiveStreak(rows);

    const cards = [
      {
        tone: "tone-good",
        icon: "↗",
        title: "Melhor mês",
        value:
          `${capitalize(best.month)} de ${best.year} · ` +
          currency.format(best.balance),
        text:
          `Entradas de ${currency.format(best.income)} ` +
          `contra ${currency.format(best.spend)} em gastos.`
      },
      {
        tone: "tone-danger",
        icon: "!",
        title: "Ponto de atenção",
        value:
          `${capitalize(worst.month)} de ${worst.year} · ` +
          currency.format(worst.balance),
        text:
          "Foi o maior déficit do recorte; vale investigar " +
          "o que puxou as despesas."
      },
      {
        tone: "tone-warn",
        icon: "◌",
        title: "Categoria dominante",
        value:
          `${labelize(topCategory[0])} · ` +
          percent.format(
            metrics.spend
              ? topCategory[1] / metrics.spend
              : 0
          ),
        text:
          `${currency.format(topCategory[1])} ` +
          "acumulados no período selecionado."
      },
      {
        tone:
          spendChange <= 0
            ? "tone-good"
            : "tone-warn",
        icon:
          spendChange <= 0
            ? "↓"
            : "↑",
        title: "Trajetória dos gastos",
        value:
          `${spendChange >= 0 ? "+" : ""}` +
          `${percent.format(spendChange)} ` +
          "desde o primeiro mês",
        text:
          `De ${currency.format(first.spend)} ` +
          `para ${currency.format(last.spend)} ` +
          "no recorte atual."
      },
      {
        tone:
          streak >= 3
            ? "tone-good"
            : "tone-warn",
        icon: "✦",
        title: "Sequência positiva",
        value:
          `${streak} ${
            streak === 1 ? "mês" : "meses"
          } consecutivos`,
        text:
          streak
            ? "A maior sequência de saldo acima de zero no período."
            : "Ainda não há uma sequência positiva neste recorte."
      },
      {
        tone:
          metrics.savingsRate >= 0.1
            ? "tone-good"
            : metrics.savingsRate >= 0
              ? "tone-warn"
              : "tone-danger",
        icon: "◎",
        title: "Margem acumulada",
        value: percent.format(metrics.savingsRate),
        text:
          metrics.savingsRate >= 0
            ? "Parte das entradas que permaneceu após os gastos."
            : "O recorte gastou mais do que recebeu."
      }
    ];

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

  function renderHealth(metrics, rows) {
    if (!rows.length) {
      els.healthScore.textContent = "—";
      els.healthBadge.textContent = "sem recorte";
      els.healthLabel.textContent = "Sem dados";

      els.healthDescription.textContent =
        "Amplie os filtros para calcular o índice orientativo do período.";

      els.healthPositive.textContent = "—";
      els.healthSavings.textContent = "—";
      els.healthEssential.textContent = "—";

      els.healthPositiveBar.style.width = "0%";
      els.healthSavingsBar.style.width = "0%";
      els.healthEssentialBar.style.width = "0%";

      return;
    }

    const savingsScore =
      clamp(metrics.savingsRate / 0.25, 0, 1) * 35;

    const positiveScore =
      metrics.positiveRate * 45;

    const essentialScore =
      (1 - clamp(metrics.essentialRatio, 0, 1)) * 20;

    const score = Math.round(
      clamp(
        savingsScore +
          positiveScore +
          essentialScore,
        0,
        100
      )
    );

    const label =
      score >= 72
        ? "Confortável"
        : score >= 48
          ? "Em construção"
          : "Pede atenção";

    els.healthScore.textContent = score;

    els.healthBadge.textContent =
      score >= 72
        ? "ritmo saudável"
        : score >= 48
          ? "acompanhar"
          : "olhar de perto";

    els.healthLabel.textContent = label;

    els.healthDescription.textContent =
      "Índice orientativo: combina meses positivos, " +
      "taxa de poupança e peso do custo essencial.";

    els.healthPositive.textContent =
      percent.format(metrics.positiveRate);

    els.healthSavings.textContent =
      percent.format(metrics.savingsRate);

    els.healthEssential.textContent =
      percent.format(metrics.essentialRatio);

    els.healthPositiveBar.style.width =
      `${clamp(metrics.positiveRate, 0, 1) * 100}%`;

    els.healthSavingsBar.style.width =
      `${clamp(metrics.savingsRate / 0.25, 0, 1) * 100}%`;

    els.healthEssentialBar.style.width =
      `${clamp(metrics.essentialRatio, 0, 1) * 100}%`;
  }

  function renderTable(rows) {
    const sorted = [...rows].sort((first, second) => {
      return second.date - first.date;
    });

    els.tableCaption.textContent =
      `${rows.length} ${
        rows.length === 1 ? "linha" : "linhas"
      }`;

    if (!sorted.length) {
      els.dataTableBody.innerHTML = `
        <tr>
          <td colspan="6">
            Nenhum mês corresponde aos filtros.
          </td>
        </tr>
      `;
      return;
    }

    els.dataTableBody.innerHTML = sorted
      .map((row) => {
        const balanceClass =
          row.balance >= 0
            ? "positive"
            : "negative";

        const savingsClass =
          row.savingsRate >= 0
            ? "positive"
            : "negative";

        return `
          <tr>
            <td>
              ${capitalize(row.month)} ${row.year}
            </td>
            <td class="amount">
              ${currency.format(row.income)}
            </td>
            <td class="amount">
              ${currency.format(row.spend)}
            </td>
            <td class="${balanceClass}">
              <span class="status-pill">
                ${currency.format(row.balance)}
              </span>
            </td>
            <td class="${savingsClass}">
              ${percent.format(row.savingsRate)}
            </td>
            <td class="leader-cell">
              ${escapeHtml(labelize(row.leader || "—"))}
            </td>
          </tr>
        `;
      })
      .join("");
  }

  function renderEmpty() {
    els.filterResult.textContent = "0 de 0 meses";
    els.periodLabel.textContent = "Aguardando uma planilha";
    els.rowCountTag.textContent = "0";
    els.categoryCountTag.textContent = "0";

    [
      "kpiIncome",
      "kpiSpend",
      "kpiBalance",
      "kpiSavings",
      "kpiFixed",
      "kpiPositive"
    ].forEach((id) => {
      els[id].textContent = "—";
    });

    els.insightsList.innerHTML = `
      <article class="insight-card">
        <div class="insight-top">
          <span class="insight-icon">↥</span>
          <span class="insight-title">
            Comece por aqui
          </span>
        </div>
        <strong>Carregue uma tabela</strong>
        <p>
          Use a área “Trocar tabela” para importar seus
          dados e ativar todos os insights.
        </p>
      </article>
    `;

    els.dataTableBody.innerHTML = `
      <tr>
        <td colspan="6">Aguardando dados…</td>
      </tr>
    `;

    els.compositionLegend.innerHTML = "";
  }

  function groupByYear(rows) {
    const groups = {};

    rows.forEach((row) => {
      if (!groups[row.year]) {
        groups[row.year] = {
          income: 0,
          spend: 0,
          balance: 0
        };
      }

      groups[row.year].income += row.income;
      groups[row.year].spend += row.spend;
      groups[row.year].balance += row.balance;
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

  function longestPositiveStreak(rows) {
    let best = 0;
    let current = 0;

    rows.forEach((row) => {
      current =
        row.balance > 0
          ? current + 1
          : 0;

      best = Math.max(best, current);
    });

    return best;
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

    return (
      `${capitalize(first.month)} ${first.year}` +
      ` — ${capitalize(last.month)} ${last.year}`
    );
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
        row.spend,
        row.balance,
        row.savingsRate,
        ...state.categories.map((category) => {
          return row.values[category];
        })
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
    anchor.download = "financas-visao-filtrada.csv";
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
    const normalizedAliases =
      aliases.map(normalizeKey);

    const exact = keyInfo.find(({ id }) => {
      return normalizedAliases.includes(id);
    });

    if (exact) {
      return exact.key;
    }

    const partial = keyInfo.find(({ id }) => {
      return normalizedAliases.some((alias) => {
        return id.includes(alias);
      });
    });

    return partial ? partial.key : null;
  }

  function normalizeKey(value) {
    return normalizeText(value)
      .replace(/[^a-z0-9]/g, "");
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
      .replace(/\b\w/g, (letter) => {
        return letter.toUpperCase();
      });
  }

  function capitalize(value) {
    const text = String(value || "");

    return text.charAt(0).toUpperCase() +
      text.slice(1);
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
      /^[-+]?\(?\s*R?\$?\s*[\d.]+(?:,\d+)?\s*\)?$/.test(text) ||
      /^[-+]?\(?\s*\d+(?:\.\d+)?\s*\)?$/.test(text)
    );
  }

  function toNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value)
        ? value
        : 0;
    }

    let text = String(
      value == null ? "" : value
    ).trim();

    if (!text) {
      return 0;
    }

    const negative = /^\(.*\)$/.test(text);

    text = text.replace(/[()R$\s]/gi, "");

    if (
      text.includes(",") &&
      text.includes(".")
    ) {
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
      return (
        `${sign}R$ ` +
        `${(abs / 1000000)
          .toFixed(1)
          .replace(".", ",")} mi`
      );
    }

    if (abs >= 1000) {
      return (
        `${sign}R$ ` +
        `${(abs / 1000)
          .toFixed(1)
          .replace(".", ",")} mil`
      );
    }

    return `${sign}R$ ${Math.round(abs)}`;
  }

  function hexToRgba(hex, alpha) {
    const clean = String(hex).replace("#", "");

    const value =
      clean.length === 3
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

  function clamp(value, min, max) {
    return Math.min(
      max,
      Math.max(min, value)
    );
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/[&<>'"]/g, (character) => {
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
})();