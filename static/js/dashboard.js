const fileInput = document.getElementById("dashboard-file");
const statusMessage = document.getElementById("dashboard-status");
const emptyState = document.getElementById("dashboard-empty");
const content = document.getElementById("dashboard-content");
const subsystemImpactList = document.getElementById("subsystem-impact-list");
const subsystemModal = document.getElementById("subsystem-modal");
const subsystemModalTitle = document.getElementById("subsystem-modal-title");
const subsystemModalSubtitle = document.getElementById("subsystem-modal-subtitle");
const subsystemModalBody = document.getElementById("subsystem-modal-body");
let impactItems = [];

function text(value) {
    return value === null || value === undefined ? "" : String(value).trim();
}

function asDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value;
    }
    if (typeof value === "number") {
        const date = new Date(1899, 11, 30);
        date.setDate(date.getDate() + Math.floor(value));
        return date;
    }
    const raw = text(value);
    if (!raw) {
        return null;
    }
    const parts = raw.split(/[\/\-.]/).map(Number);
    if (parts.length === 3 && parts.every(Number.isFinite)) {
        return parts[0] > 31
            ? new Date(parts[0], parts[1] - 1, parts[2])
            : new Date(parts[2], parts[1] - 1, parts[0]);
    }
    const parsed = new Date(raw);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function dateLabel(value) {
    const date = asDate(value);
    if (!date) {
        return "";
    }
    return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
}

function cleanRow(row) {
    return Object.fromEntries(Object.entries(row).map(function([key, value]) {
        return [key.trim(), value];
    }));
}

function unique(values) {
    return [...new Set(values.map(text).filter(Boolean))];
}

function splitPO(value) {
    return text(value).split("\n").map(function(item) {
        return item.trim();
    }).filter(function(item) {
        return item && !item.includes("…");
    });
}

function splitETAs(value) {
    let dates;
    if (value instanceof Date || typeof value === "number") {
        const parsed = asDate(value);
        dates = parsed ? [parsed] : [];
    } else {
        dates = text(value).split(/[\n;,]+/).map(asDate).filter(Boolean);
    }
    return dates.map(function(date) {
        return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    });
}

function delivered(value) {
    const normalized = text(value).toLocaleLowerCase();
    return Boolean(normalized) && !["no", "no entregado", "pendiente", "false", "0", "n"].includes(normalized);
}

function buildSubsystemImpact(rows) {
    const subsystems = new Map();
    rows.forEach(function(row) {
        const subsystem = text(row.SubSistema);
        const fmr = text(row.FMR);
        if (!subsystem || !fmr || fmr.includes("…")) {
            return;
        }

        if (!subsystems.has(subsystem)) {
            subsystems.set(subsystem, new Map());
        }
        const fmrs = subsystems.get(subsystem);
        if (!fmrs.has(fmr)) {
            fmrs.set(fmr, { pos: new Map(), rowsWithoutPO: [] });
        }
        const fmrData = fmrs.get(fmr);
        const pos = splitPO(row.PO);
        if (!pos.length) {
            fmrData.rowsWithoutPO.push(row);
            return;
        }
        pos.forEach(function(po) {
            if (!fmrData.pos.has(po)) {
                fmrData.pos.set(po, []);
            }
            fmrData.pos.get(po).push(row);
        });
    });

    let totalFmrs = 0;
    let arrivedFmrs = 0;
    let partialFmrs = 0;
    let pendingFmrs = 0;
    let withoutEtaFmrs = 0;

    const items = [...subsystems.entries()].map(function([subsystem, fmrs]) {
        let coverageScore = 0;
        let arrived = 0;
        let partial = 0;
        let pending = 0;
        let withoutEta = 0;
        const details = [];

        fmrs.forEach(function(fmrData, fmr) {
            const poGroups = fmrData.pos.size
                ? [...fmrData.pos.entries()]
                : [["Sin PO", fmrData.rowsWithoutPO]];
            const scores = poGroups.map(function([po, rowsForPO]) {
                const isArrived = rowsForPO.some(function(row) {
                    return text(row["ON SITE"]).toLocaleLowerCase() === "on site" || delivered(row.ENTREGADO);
                });
                const isPartial = rowsForPO.some(function(row) {
                    return text(row["ON SITE"]).toLocaleLowerCase() === "on site - parcial";
                });
                const hasEta = rowsForPO.some(function(row) { return splitETAs(row.ETA).length > 0; });
                details.push({ fmr, po, category: isArrived ? "onSite" : (isPartial ? "partial" : "pending") });
                return {
                    score: isArrived ? 1 : (isPartial ? 0.5 : 0),
                    withoutEta: !isArrived && !hasEta
                };
            });
            const fmrScore = scores.reduce(function(sum, item) { return sum + item.score; }, 0) / scores.length;
            coverageScore += fmrScore;
            if (fmrScore === 1) {
                arrived += 1;
            } else if (fmrScore > 0) {
                partial += 1;
            } else {
                pending += 1;
            }
            if (scores.some(function(item) { return item.withoutEta; })) {
                withoutEta += 1;
            }
        });

        const total = fmrs.size;
        totalFmrs += total;
        arrivedFmrs += arrived;
        partialFmrs += partial;
        pendingFmrs += pending;
        withoutEtaFmrs += withoutEta;

        const coveragePercent = total ? Math.round((coverageScore / total) * 100) : 0;
        return {
            subsystem,
            total,
            arrived,
            partial,
            pending,
            withoutEta,
            details,
            coverageScore,
            coveragePercent,
            impactPercent: 100 - coveragePercent
        };
    }).sort(function(a, b) {
        return a.impactPercent - b.impactPercent || a.subsystem.localeCompare(b.subsystem);
    });

    const totalScore = items.reduce(function(sum, item) {
        return sum + item.coverageScore;
    }, 0);

    return {
        items,
        totals: {
            totalFmrs,
            arrivedFmrs,
            partialFmrs,
            pendingFmrs,
            withoutEtaFmrs,
            coveragePercent: totalFmrs ? Math.round((totalScore / totalFmrs) * 100) : 0
        }
    };
}

function escapeHtml(value) {
    return text(value).replace(/[&<>\"']/g, function(character) {
        return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[character];
    });
}

function buildModel(workbook) {
    const normalizedSheet = workbook.Sheets["FMR_Normalizado"];
    if (!normalizedSheet) {
        throw new Error("El archivo debe contener la hoja FMR_Normalizado.");
    }

    const rows = XLSX.utils.sheet_to_json(normalizedSheet, {
        defval: null,
        cellDates: true
    }).map(cleanRow);
    const validRows = rows.filter(function(row) {
        return text(row.FMR) && !text(row.FMR).includes("…");
    });
    const fmrs = unique(validRows.map(function(row) { return row.FMR; }));
    const subsystems = unique(validRows.map(function(row) { return row.SubSistema; }));
    const poRows = [];

    validRows.forEach(function(row) {
        splitPO(row.PO).forEach(function(po) {
            poRows.push({ row, po });
        });
    });

    const poRecords = new Map();
    poRows.forEach(function(item) {
        const fmr = text(item.row.FMR);
        const key = `${fmr}|${item.po}`;
        if (!poRecords.has(key)) {
            poRecords.set(key, { fmr, po: item.po, rows: [] });
        }
        poRecords.get(key).rows.push(item.row);
    });

    const etaEntries = [];
    const etaKeys = new Set();
    validRows.forEach(function(row) {
        const pos = splitPO(row.PO);
        splitETAs(row.ETA).forEach(function(eta) {
            (pos.length ? pos : ["Sin PO"]).forEach(function(po) {
                const key = `${text(row.FMR)}|${po}|${eta.getFullYear()}-${eta.getMonth()}-${eta.getDate()}`;
                if (!etaKeys.has(key)) {
                    etaKeys.add(key);
                    etaEntries.push({ eta, fmr: text(row.FMR), po });
                }
            });
        });
    });

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() - (today.getDay() + 1) % 7);
    const weekRange = function(index) {
        const start = new Date(weekStart);
        start.setDate(weekStart.getDate() + index * 7);
        const end = new Date(start);
        end.setDate(start.getDate() + 6);
        end.setHours(23, 59, 59, 999);
        return { start, end };
    };
    const currentWeek = weekRange(0);
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    const nextMonthStart = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const monthName = function(date) {
        return new Intl.DateTimeFormat("es", { month: "long", year: "numeric" }).format(date);
    };
    const makePeriod = function(label, range, start, end, groupByPO = false) {
        const inRange = etaEntries.filter(function(entry) {
            return entry.eta >= start && entry.eta <= end;
        }).sort(function(a, b) {
            return a.eta - b.eta || a.fmr.localeCompare(b.fmr) || a.po.localeCompare(b.po);
        });
        if (!groupByPO) {
            return { label, range, entries: inRange };
        }

        const grouped = new Map();
        inRange.forEach(function(entry) {
            const key = `${entry.fmr}|${entry.po}`;
            if (grouped.has(key)) {
                grouped.get(key).dates.push(entry.eta);
            } else {
                grouped.set(key, { ...entry, dates: [entry.eta] });
            }
        });
        return { label, range, entries: [...grouped.values()] };
    };
    const etaPeriods = [
        makePeriod("Semana presente", `${dateLabel(currentWeek.start)} - ${dateLabel(currentWeek.end)}`, currentWeek.start, currentWeek.end, true),
        makePeriod("Mes presente", monthName(monthStart), monthStart, new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59, 999)),
        makePeriod("Próximo mes", monthName(nextMonthStart), nextMonthStart, new Date(today.getFullYear(), today.getMonth() + 2, 0, 23, 59, 59, 999)),
        makePeriod("Próximos dos meses", `${monthName(nextMonthStart)} - ${monthName(new Date(today.getFullYear(), today.getMonth() + 2, 1))}`, nextMonthStart, new Date(today.getFullYear(), today.getMonth() + 3, 0, 23, 59, 59, 999)),
        {
            label: "PO sin ETA",
            range: "Sin fecha ETA registrada",
            noEta: true,
            entries: [...poRecords.values()].filter(function(record) {
                return !record.rows.some(function(row) { return splitETAs(row.ETA).length > 0; });
            }).map(function(record) {
                return {
                    fmr: record.fmr,
                    po: record.po,
                    description: unique(record.rows.map(function(row) { return row.DESCRIPCION; })).join(" · ")
                };
            }).sort(function(a, b) { return a.fmr.localeCompare(b.fmr) || a.po.localeCompare(b.po); })
        }
    ];

    const weeklyEta = Array.from({ length: 3 }, function(_, index) {
        const { start, end } = weekRange(index);
        return {
            label: index === 0 ? "Semana actual" : (index === 1 ? "Próxima semana" : "Semana subsiguiente"),
            range: `${dateLabel(start)} - ${dateLabel(end)}`,
            count: makePeriod("", "", start, end, true).entries.length
        };
    });
    weeklyEta.push({ label: "Sin ETA", range: "", count: etaPeriods[4].entries.length, noEta: true });

    const poKeys = unique(poRows.map(function(item) {
        return `${text(item.row.FMR)}|${item.po}`;
    }));
    const fmrSummary = fmrs.map(function(fmr) {
        const fmrRows = validRows.filter(function(row) { return text(row.FMR) === fmr; });
        const fmrPOs = poRows.filter(function(item, index, items) {
            return text(item.row.FMR) === fmr && items.findIndex(function(candidate) {
                return text(candidate.row.FMR) === text(item.row.FMR) && candidate.po === item.po;
            }) === index;
        });
        const noEta = fmrPOs.filter(function(item) { return !text(item.row.ETA); }).length;
        const futureEta = fmrPOs.filter(function(item) {
            const eta = asDate(text(item.row.ETA).split(/[\n;,]/)[0]);
            return eta && eta > new Date();
        }).length;
        const statuses = unique(fmrRows.map(function(row) { return row.STATUS; }));
        return {
            fmr,
            subsystems: unique(fmrRows.map(function(row) { return row.SubSistema; })).length,
            po: unique(fmrPOs.map(function(item) { return item.po; })).length,
            noEta,
            futureEta,
            status: statuses[0] || "Sin status"
        };
    });

    const statusCounts = {};
    fmrSummary.forEach(function(item) {
        statusCounts[item.status] = (statusCounts[item.status] || 0) + 1;
    });

    const deliveryDetails = { onSite: [], noEta: [], futureEta: [] };
    poRecords.forEach(function(record) {
        const summary = {
            fmr: record.fmr,
            po: record.po,
            description: unique(record.rows.map(function(row) { return row.DESCRIPCION; })).join(" · ") || "Sin descripción"
        };
        if (record.rows.some(function(row) { return Boolean(text(row["ON SITE"])); })) {
            deliveryDetails.onSite.push(summary);
        }
        if (record.rows.some(function(row) { return !text(row.ETA); })) {
            deliveryDetails.noEta.push(summary);
        }
        if (record.rows.some(function(row) {
            return splitETAs(row.ETA).some(function(eta) { return eta > today; });
        })) {
            deliveryDetails.futureEta.push(summary);
        }
    });
    const delivery = {
        onSite: deliveryDetails.onSite.length,
        noEta: deliveryDetails.noEta.length,
        futureEta: deliveryDetails.futureEta.length
    };

    const subsystemImpact = buildSubsystemImpact(validRows);

    return {
        fileRows: rows.length,
        fmrs,
        subsystems,
        poCount: poKeys.length,
        summary: fmrSummary,
        statusCounts,
        delivery,
        deliveryDetails,
        etaPeriods,
        weeklyEta,
        subsystemImpact: subsystemImpact.items,
        subsystemImpactTotals: subsystemImpact.totals
    };
}

function renderBars(targetId, items, maxValue) {
    const target = document.getElementById(targetId);
    target.innerHTML = items.length ? items.map(function(item) {
        const width = Math.max(3, (item.count / maxValue) * 100);
        return `<div class="bar-row"><span class="bar-label" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span><span class="bar-track"><span class="bar-fill" style="width:${width}%"></span></span><span class="bar-value">${item.count}</span></div>`;
    }).join("") : "<p class=\"panel-caption\">Sin datos disponibles.</p>";
}

function renderDelivery(delivery, deliveryDetails, totalPO) {
    const values = [
        { key: "onSite", label: "Con ON SITE", value: delivery.onSite, color: "var(--teal)" },
        { key: "noEta", label: "Sin ETA", value: delivery.noEta, color: "var(--amber)" },
        { key: "futureEta", label: "ETA futura", value: delivery.futureEta, color: "var(--red)" }
    ];
    const totalSignals = values.reduce(function(total, item) { return total + item.value; }, 0);
    const denominator = Math.max(totalSignals, 1);
    let cursor = 0;
    const stops = values.map(function(item) {
        const start = cursor;
        cursor += (item.value / denominator) * 360;
        return `${item.color} ${start}deg ${cursor}deg`;
    }).join(", ");
    const gradient = totalSignals
        ? `conic-gradient(${stops})`
        : "conic-gradient(#d9e1dc 0deg 360deg)";
    document.getElementById("delivery-chart").innerHTML = `
        <div class="donut" aria-label="Gráfico de riesgos de entrega; seleccione un sector o su leyenda" style="background:${gradient}">
            <div class="donut-hole">${totalPO}<small>PO totales</small></div>
        </div>
        <div class="legend">${values.map(function(item) {
            return `<button type="button" class="legend-item delivery-filter" data-category="${item.key}" aria-pressed="false"><span class="legend-dot" style="background:${item.color}"></span><span>${item.label}</span><span class="legend-value">${item.value}</span></button>`;
        }).join("")}</div>`;

    const detailPanel = document.getElementById("delivery-detail");
    const filterButtons = [...document.querySelectorAll(".delivery-filter")];
    let selectedCategory = "";

    function selectCategory(category) {
        if (selectedCategory === category) {
            selectedCategory = "";
            detailPanel.hidden = true;
            filterButtons.forEach(function(button) { button.setAttribute("aria-pressed", "false"); });
            return;
        }

        selectedCategory = category;
        filterButtons.forEach(function(button) {
            button.setAttribute("aria-pressed", String(button.dataset.category === category));
        });
        const selected = values.find(function(item) { return item.key === category; });
        const records = deliveryDetails[category] || [];
        detailPanel.innerHTML = `
            <div class="delivery-detail-heading">
                <div><h3>${selected.label}</h3><p>${records.length} PO asociadas</p></div>
                <button type="button" class="delivery-detail-close">Cerrar</button>
            </div>
            <div class="delivery-table-scroll"><table>
                <thead><tr><th>FMR</th><th>Descripción</th><th>PO</th></tr></thead>
                <tbody>${records.map(function(record) {
                    return `<tr><td><strong>${escapeHtml(record.fmr)}</strong></td><td><span class="delivery-description" title="${escapeHtml(record.description)}">${escapeHtml(record.description)}</span></td><td>${escapeHtml(record.po)}</td></tr>`;
                }).join("") || '<tr><td colspan="3">No hay PO en esta categoría.</td></tr>'}</tbody>
            </table></div>`;
        detailPanel.hidden = false;
        detailPanel.querySelector(".delivery-detail-close").addEventListener("click", function() {
            selectedCategory = "";
            detailPanel.hidden = true;
            filterButtons.forEach(function(button) { button.setAttribute("aria-pressed", "false"); });
        });
    }

    filterButtons.forEach(function(button) {
        button.addEventListener("click", function() { selectCategory(button.dataset.category); });
    });

    const donut = document.querySelector(".donut");
    donut.addEventListener("click", function(event) {
        if (!totalSignals) {
            return;
        }
        const bounds = donut.getBoundingClientRect();
        const radius = Math.min(bounds.width, bounds.height) / 2;
        const x = event.clientX - (bounds.left + bounds.width / 2);
        const y = event.clientY - (bounds.top + bounds.height / 2);
        if (Math.hypot(x, y) < radius * 0.63) {
            return;
        }
        const angle = (Math.atan2(y, x) * 180 / Math.PI + 450) % 360;
        let boundary = 0;
        for (const item of values) {
            const nextBoundary = boundary + (item.value / totalSignals) * 360;
            if (item.value && angle >= boundary && angle < nextBoundary) {
                selectCategory(item.key);
                break;
            }
            boundary = nextBoundary;
        }
    });
}

function renderEtaPeriods(periods) {
    document.getElementById("eta-periods").innerHTML = periods.map(function(period) {
        const entries = period.entries.map(function(entry) {
            if (period.noEta) {
                return `<li class="no-eta-entry"><span><strong>${escapeHtml(entry.fmr)}</strong><small>PO ${escapeHtml(entry.po)}</small></span></li>`;
            }
            const dates = (entry.dates || [entry.eta]).map(dateLabel).join("<br>");
            return `<li><time>${dates}</time><span><strong>${escapeHtml(entry.fmr)}</strong><small>PO ${escapeHtml(entry.po)}</small></span></li>`;
        }).join("");
        return `<article class="eta-period-card">
            <header><div><h3>${escapeHtml(period.label)}</h3><p>${escapeHtml(period.range)}</p></div><strong class="eta-count">${period.entries.length}</strong></header>
            ${entries ? `<ul>${entries}</ul>` : '<p class="eta-empty">Sin fechas ETA en este periodo.</p>'}
        </article>`;
    }).join("");
}

function renderEtaWeekChart(items) {
    const maximum = Math.max(...items.map(function(item) { return item.count; }), 1);
    const ticks = [...new Set([maximum, Math.ceil(maximum * 0.75), Math.ceil(maximum * 0.5), Math.ceil(maximum * 0.25), 0])];

    document.getElementById("eta-week-chart").innerHTML = `
        <div class="eta-chart-y-title">Total de PO</div>
        <div class="eta-chart-layout">
            <div class="eta-chart-y-axis">${ticks.map(function(value) { return `<span>${value}</span>`; }).join("")}</div>
            <div class="eta-chart-main">
                <div class="eta-chart-plot">
                    <div class="eta-chart-grid-lines" aria-hidden="true">${ticks.map(function() { return "<span></span>"; }).join("")}</div>
                    <div class="eta-chart-columns">${items.map(function(item) {
                        const ratio = item.count / maximum;
                        const height = item.count ? Math.max(3, ratio * 100) : 0;
                        const background = item.noEta ? "#F89880" : "linear-gradient(180deg, #8bb6d5, #285f88)";
                        const opacity = item.noEta ? 1 : 0.45 + ratio * 0.55;
                        return `<div class="eta-chart-column">
                            <strong class="eta-chart-value" style="--bar-height:${height}%">${item.count}</strong>
                            <span class="eta-chart-bar${item.noEta ? " eta-chart-bar-no-eta" : ""}" style="--bar-height:${height}%;--bar-opacity:${opacity};background:${background}"></span>
                        </div>`;
                    }).join("")}</div>
                </div>
                <div class="eta-chart-x-axis">${items.map(function(item) {
                    return `<div><strong>${escapeHtml(item.label)}</strong>${item.range ? `<small>${escapeHtml(item.range)}</small>` : ""}</div>`;
                }).join("")}</div>
            </div>
        </div>`;
}

function renderSubsystemImpact(items, totals) {
    document.getElementById("subsystem-impact-summary").innerHTML = [
        { value: `${totals.coveragePercent}%`, label: "Cobertura estimada" },
        { value: totals.arrivedFmrs, label: "FMR recibidas" },
        { value: totals.partialFmrs, label: "FMR parciales" },
        { value: totals.pendingFmrs, label: "FMR pendientes" }
    ].map(function(item) {
        return `<div class="impact-summary-item"><strong>${item.value}</strong><span>${item.label}</span></div>`;
    }).join("");

    impactItems = items;
    subsystemImpactList.innerHTML = items.map(function(item, index) {
        return `<article class="subsystem-impact-row" role="button" tabindex="0" aria-haspopup="dialog" data-index="${index}">
            <div class="impact-row-heading">
                <div><h3>${escapeHtml(item.subsystem)}</h3><p>${item.total} FMR asociadas</p></div>
                <strong class="impact-percent">${item.coveragePercent}%<small>cobertura</small></strong>
            </div>
            <div class="impact-progress" role="progressbar" aria-label="Cobertura estimada de ${escapeHtml(item.subsystem)}" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${item.coveragePercent}"><span style="width:${item.coveragePercent}%"></span></div>
            <div class="impact-breakdown"><span>${item.arrived} On Site</span><span>${item.partial} parciales</span><span>${item.pending} pendientes</span></div>
            <p class="impact-risk">Impacto pendiente: <strong>${item.impactPercent}%</strong>${item.withoutEta ? ` · ${item.withoutEta} sin ETA` : ""}</p>
        </article>`;
    }).join("") || '<p class="panel-caption">No hay SubSistemas asociados a FMR en el archivo.</p>';
}

function renderDashboard(model, filename) {
    const totalNoEta = model.delivery.noEta;
    const totalFuture = model.delivery.futureEta;
    const statusItems = Object.entries(model.statusCounts).map(function([label, count]) { return { label, count }; }).sort(function(a, b) { return b.count - a.count; });
    document.getElementById("metric-fmr").textContent = model.fmrs.length;
    document.getElementById("metric-fmr-note").textContent = `${model.fileRows} Items`;
    document.getElementById("metric-subsystems").textContent = model.subsystems.length;
    document.getElementById("metric-po").textContent = model.poCount;
    document.getElementById("metric-po-note").textContent = `${totalFuture} con ETA futura`;
    document.getElementById("metric-no-eta").textContent = totalNoEta;
    document.getElementById("loaded-file").textContent = filename;
    document.getElementById("data-scope").textContent = "Datos leídos localmente · FMR_Normalizado";

    const highestStatus = statusItems[0];
    document.getElementById("headline-insight").textContent = totalNoEta
        ? `${totalNoEta} PO no tienen ETA registrada.`
        : "No hay PO sin ETA registrada.";
    document.getElementById("headline-detail").textContent = highestStatus
        ? `El status predominante es ${highestStatus.label} con ${highestStatus.count} FMR.`
        : "";

    renderBars("status-chart", statusItems, Math.max(...statusItems.map(function(item) { return item.count; }), 1));
    renderDelivery(model.delivery, model.deliveryDetails, model.poCount);
    renderEtaPeriods(model.etaPeriods);
    renderEtaWeekChart(model.weeklyEta);
    renderSubsystemImpact(model.subsystemImpact, model.subsystemImpactTotals);
}

function openSubsystemModal(index) {
    const item = impactItems[index];
    if (!item) {
        return;
    }

    const groups = [
        { key: "onSite", label: "On Site" },
        { key: "partial", label: "Parciales" },
        { key: "pending", label: "Pendientes" }
    ];
    subsystemModalTitle.textContent = item.subsystem;
    subsystemModalSubtitle.textContent = `${item.total} FMR asociadas · ${item.coveragePercent}% de cobertura`;
    subsystemModalBody.innerHTML = groups.map(function(group) {
        const records = item.details.filter(function(detail) {
            return detail.category === group.key;
        }).sort(function(a, b) {
            return a.fmr.localeCompare(b.fmr) || a.po.localeCompare(b.po);
        });
        const rows = records.map(function(record) {
            return `<tr><td><strong>${escapeHtml(record.fmr)}</strong></td><td>${escapeHtml(record.po)}</td></tr>`;
        }).join("");
        return `<section class="subsystem-modal-group">
            <h3>${group.label}<span>${records.length} PO</span></h3>
            ${rows ? `<div class="subsystem-modal-scroll"><table><thead><tr><th>FMR</th><th>PO</th></tr></thead><tbody>${rows}</tbody></table></div>` : '<p class="eta-empty">Sin PO en esta categoría.</p>'}
        </section>`;
    }).join("");

    if (!subsystemModal.open) {
        subsystemModal.showModal();
    }
}

subsystemImpactList.addEventListener("click", function(event) {
    const row = event.target.closest(".subsystem-impact-row");
    if (row) {
        openSubsystemModal(Number(row.dataset.index));
    }
});

subsystemImpactList.addEventListener("keydown", function(event) {
    const row = event.target.closest(".subsystem-impact-row");
    if (row && (event.key === "Enter" || event.key === " ")) {
        event.preventDefault();
        openSubsystemModal(Number(row.dataset.index));
    }
});

subsystemModal.addEventListener("click", function(event) {
    if (event.target === subsystemModal) {
        subsystemModal.close();
    }
});

document.getElementById("subsystem-modal-close").addEventListener("click", function() {
    subsystemModal.close();
});

async function procesarDashboardLocal(buffer, filename, guardar = true) {
    statusMessage.textContent = "Leyendo el archivo localmente...";
    try {
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
        const model = buildModel(workbook);
        renderDashboard(model, filename);
        emptyState.hidden = true;
        content.hidden = false;
        statusMessage.textContent = "Archivo procesado en el navegador.";

        if (guardar) {
            try {
                await guardarWorkbookLocal(buffer, filename);
            } catch (storageError) {
                console.warn("No se pudo guardar el archivo en el navegador:", storageError);
            }
        }

        return true;
    } catch (error) {
        content.hidden = true;
        emptyState.hidden = false;
        statusMessage.textContent = error.message || "No se pudo leer el archivo.";
        return false;
    }
}

fileInput.addEventListener("change", async function() {
    const file = this.files[0];
    if (file) {
        await procesarDashboardLocal(await file.arrayBuffer(), file.name);
    }
});

async function restaurarDashboardLocal() {
    try {
        const guardado = await leerWorkbookLocal();
        if (guardado) {
            await procesarDashboardLocal(guardado.buffer, guardado.name, false);
        }
    } catch (error) {
        console.warn("No se pudo restaurar el archivo guardado:", error);
    }
}

restaurarDashboardLocal();
