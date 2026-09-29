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
        return new Date(Date.UTC(1899, 11, 30) + value * 86400000);
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
        return text(row.FMR) && !text(row.FMR).includes("…") && text(row.SubSistema);
    });
    const fmrs = unique(validRows.map(function(row) { return row.FMR; }));
    const subsystems = unique(validRows.map(function(row) { return row.SubSistema; }));
    const poRows = [];

    validRows.forEach(function(row) {
        splitPO(row.PO).forEach(function(po) {
            poRows.push({ row, po });
        });
    });

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
        subsystemCounts[subsystem] = subsystemCounts[subsystem] || new Set();
        subsystemCounts[subsystem].add(text(row.FMR));
    });

    return {
        fileRows: rows.length,
        fmrs,
        subsystems,
        poCount: poKeys.length,
        summary: fmrSummary,
        statusCounts,
        delivery,
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

function renderDashboard(model, filename) {
    const totalNoEta = model.delivery.noEta;
    const totalFuture = model.delivery.futureEta;
    const statusItems = Object.entries(model.statusCounts).map(function([label, count]) { return { label, count }; }).sort(function(a, b) { return b.count - a.count; });
    const weekItems = model.weeks.sort(function(a, b) { return asDate(a.label) - asDate(b.label); });
    const subsystemItems = model.subsystemCounts.sort(function(a, b) { return b.count - a.count; }).slice(0, 7);
    const riskItems = model.summary.filter(function(item) { return item.noEta || item.futureEta; }).sort(function(a, b) {
        return (b.noEta + b.futureEta) - (a.noEta + a.futureEta);
    }).slice(0, 10);

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
    renderBars("week-chart", weekItems, Math.max(...weekItems.map(function(item) { return item.count; }), 1));
    document.getElementById("subsystem-table").innerHTML = subsystemItems.map(function(item, index) {
        return `<div class="rank-row"><span class="rank-number">${String(index + 1).padStart(2, "0")}</span><span class="rank-name" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span><span class="rank-count">${item.count}</span></div>`;
    }).join("") || "<p class=\"panel-caption\">Sin datos disponibles.</p>";
    document.getElementById("risk-table").innerHTML = riskItems.map(function(item) {
        const risk = item.noEta >= item.futureEta && item.noEta ? "high" : "";
        const status = item.noEta ? "Sin ETA" : "ETA futura";
        return `<tr><td><strong>${escapeHtml(item.fmr)}</strong></td><td>${item.subsystems}</td><td>${item.po}</td><td>${item.noEta}</td><td>${item.futureEta}</td><td><span class="status-chip ${risk}">${status}</span></td></tr>`;
    }).join("") || "<tr><td colspan=\"6\">No hay FMR con exposición identificada.</td></tr>";
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
