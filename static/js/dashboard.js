const fileInput = document.getElementById("dashboard-file");
const statusMessage = document.getElementById("dashboard-status");
const emptyState = document.getElementById("dashboard-empty");
const content = document.getElementById("dashboard-content");

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

        fmrs.forEach(function(fmrData) {
            const pos = fmrData.pos.size
                ? [...fmrData.pos.values()]
                : [fmrData.rowsWithoutPO];
            const scores = pos.map(function(rowsForPO) {
                const isArrived = rowsForPO.some(function(row) {
                    return text(row["ON SITE"]).toLocaleLowerCase() === "on site" || delivered(row.ENTREGADO);
                });
                const isPartial = rowsForPO.some(function(row) {
                    return text(row["ON SITE"]).toLocaleLowerCase() === "on site - parcial";
                });
                const hasEta = rowsForPO.some(function(row) { return splitETAs(row.ETA).length > 0; });
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
            coverageScore,
            coveragePercent,
            impactPercent: 100 - coveragePercent
        };
    }).sort(function(a, b) {
        return a.coveragePercent - b.coveragePercent || b.withoutEta - a.withoutEta || a.subsystem.localeCompare(b.subsystem);
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
    weekStart.setDate(today.getDate() - (today.getDay() + 6) % 7);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    const nextMonthStart = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const monthName = function(date) {
        return new Intl.DateTimeFormat("es", { month: "long", year: "numeric" }).format(date);
    };
    const makePeriod = function(label, range, start, end) {
        return {
            label,
            range,
            entries: etaEntries.filter(function(entry) {
                return entry.eta >= start && entry.eta <= end;
            }).sort(function(a, b) {
                return a.eta - b.eta || a.fmr.localeCompare(b.fmr) || a.po.localeCompare(b.po);
            })
        };
    };
    const etaPeriods = [
        makePeriod("Semana presente", `${dateLabel(weekStart)} - ${dateLabel(weekEnd)}`, weekStart, weekEnd),
        makePeriod("Mes presente", monthName(monthStart), monthStart, new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59, 999)),
        makePeriod("Próximo mes", monthName(nextMonthStart), nextMonthStart, new Date(today.getFullYear(), today.getMonth() + 2, 0, 23, 59, 59, 999)),
        makePeriod("Próximos dos meses", `${monthName(nextMonthStart)} - ${monthName(new Date(today.getFullYear(), today.getMonth() + 2, 1))}`, nextMonthStart, new Date(today.getFullYear(), today.getMonth() + 3, 0, 23, 59, 59, 999))
    ];

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

    const delivery = {
        onSite: unique(poRows.filter(function(item) { return text(item.row["ON SITE"]); }).map(function(item) { return `${item.row.FMR}|${item.po}`; })).length,
        noEta: unique(poRows.filter(function(item) { return !text(item.row.ETA); }).map(function(item) { return `${item.row.FMR}|${item.po}`; })).length,
        futureEta: unique(poRows.filter(function(item) {
            const eta = asDate(text(item.row.ETA).split(/[\n;,]/)[0]);
            return eta && eta > new Date();
        }).map(function(item) { return `${item.row.FMR}|${item.po}`; })).length
    };

    const weeks = {};
    validRows.forEach(function(row) {
        const week = dateLabel(row["3week"]) || "Sin fecha";
        weeks[week] = weeks[week] || new Set();
        weeks[week].add(text(row.FMR));
    });

    const subsystemCounts = {};
    validRows.forEach(function(row) {
        const subsystem = text(row.SubSistema);
        if (!subsystem) {
            return;
        }
        subsystemCounts[subsystem] = subsystemCounts[subsystem] || new Set();
        subsystemCounts[subsystem].add(text(row.FMR));
    });
    const subsystemImpact = buildSubsystemImpact(validRows);

    return {
        fileRows: rows.length,
        fmrs,
        subsystems,
        poCount: poKeys.length,
        summary: fmrSummary,
        statusCounts,
        delivery,
        etaPeriods,
        subsystemImpact: subsystemImpact.items,
        subsystemImpactTotals: subsystemImpact.totals,
        weeks: Object.entries(weeks).map(function([label, values]) { return { label, count: values.size }; }),
        subsystemCounts: Object.entries(subsystemCounts).map(function([label, values]) { return { label, count: values.size }; })
    };
}

function renderBars(targetId, items, maxValue) {
    const target = document.getElementById(targetId);
    target.innerHTML = items.length ? items.map(function(item) {
        const width = Math.max(3, (item.count / maxValue) * 100);
        return `<div class="bar-row"><span class="bar-label" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span><span class="bar-track"><span class="bar-fill" style="width:${width}%"></span></span><span class="bar-value">${item.count}</span></div>`;
    }).join("") : "<p class=\"panel-caption\">Sin datos disponibles.</p>";
}

function renderDelivery(delivery, totalPO) {
    const values = [
        { label: "Con ON SITE", value: delivery.onSite, color: "var(--teal)" },
        { label: "Sin ETA", value: delivery.noEta, color: "var(--amber)" },
        { label: "ETA futura", value: delivery.futureEta, color: "var(--red)" }
    ];
    const sum = values.reduce(function(total, item) { return total + item.value; }, 0);
    const denominator = Math.max(totalPO, 1);
    let cursor = 0;
    const stops = values.map(function(item) {
        const start = cursor;
        cursor += (item.value / denominator) * 360;
        return `${item.color} ${start}deg ${cursor}deg`;
    }).join(", ");
    document.getElementById("delivery-chart").innerHTML = `
        <div class="donut" style="background:conic-gradient(${stops || "#d9e1dc 0deg 360deg"})">
            <div class="donut-hole">${totalPO}<small>PO totales</small></div>
        </div>
        <div class="legend">${values.map(function(item) {
            return `<div class="legend-item"><span class="legend-dot" style="background:${item.color}"></span><span>${item.label}</span><span class="legend-value">${item.value}</span></div>`;
        }).join("")}</div>`;
}

function renderEtaPeriods(periods) {
    document.getElementById("eta-periods").innerHTML = periods.map(function(period) {
        const entries = period.entries.map(function(entry) {
            return `<li><time>${dateLabel(entry.eta)}</time><span><strong>${escapeHtml(entry.fmr)}</strong><small>PO ${escapeHtml(entry.po)}</small></span></li>`;
        }).join("");
        return `<article class="eta-period-card">
            <header><div><h3>${escapeHtml(period.label)}</h3><p>${escapeHtml(period.range)}</p></div><strong class="eta-count">${period.entries.length}</strong></header>
            ${entries ? `<ul>${entries}</ul>` : '<p class="eta-empty">Sin fechas ETA en este periodo.</p>'}
        </article>`;
    }).join("");
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

    document.getElementById("subsystem-impact-list").innerHTML = items.map(function(item) {
        return `<article class="subsystem-impact-row">
            <div class="impact-row-heading">
                <div><h3>${escapeHtml(item.subsystem)}</h3><p>${item.total} FMR asociadas</p></div>
                <strong class="impact-percent">${item.coveragePercent}%<small>cobertura</small></strong>
            </div>
            <div class="impact-progress" role="progressbar" aria-label="Cobertura estimada de ${escapeHtml(item.subsystem)}" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${item.coveragePercent}"><span style="width:${item.coveragePercent}%"></span></div>
            <div class="impact-breakdown"><span>${item.arrived} recibidas</span><span>${item.partial} parciales</span><span>${item.pending} pendientes</span></div>
            <p class="impact-risk">Impacto pendiente: <strong>${item.impactPercent}%</strong>${item.withoutEta ? ` · ${item.withoutEta} sin ETA` : ""}</p>
        </article>`;
    }).join("") || '<p class="panel-caption">No hay SubSistemas asociados a FMR en el archivo.</p>';
}

function renderDashboard(model, filename) {
    const totalNoEta = model.delivery.noEta;
    const totalFuture = model.delivery.futureEta;
    const statusItems = Object.entries(model.statusCounts).map(function([label, count]) { return { label, count }; }).sort(function(a, b) { return b.count - a.count; });
    const weekItems = model.weeks.sort(function(a, b) { return asDate(a.label) - asDate(b.label); });
    const subsystemItems = model.subsystemCounts.sort(function(a, b) { return b.count - a.count; }).slice(0, 7);
    document.getElementById("metric-fmr").textContent = model.fmrs.length;
    document.getElementById("metric-fmr-note").textContent = `${model.fileRows} filas normalizadas`;
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
    renderDelivery(model.delivery, model.poCount);
    renderEtaPeriods(model.etaPeriods);
    renderSubsystemImpact(model.subsystemImpact, model.subsystemImpactTotals);
    renderBars("week-chart", weekItems, Math.max(...weekItems.map(function(item) { return item.count; }), 1));
    document.getElementById("subsystem-table").innerHTML = subsystemItems.map(function(item, index) {
        return `<div class="rank-row"><span class="rank-number">${String(index + 1).padStart(2, "0")}</span><span class="rank-name" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span><span class="rank-count">${item.count}</span></div>`;
    }).join("") || "<p class=\"panel-caption\">Sin datos disponibles.</p>";
}

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
