/* LMC Lab — Isothermal Calorimetry page JS */

(function () {
    "use strict";

    // ── File input: show selected filename ──────────────────────────
    var fileInput  = document.getElementById("file-input");
    var filenameEl = document.getElementById("upload-filename");
    var uploadZone = document.getElementById("upload-zone");

    if (fileInput) {
        fileInput.addEventListener("change", function () {
            if (this.files && this.files[0]) {
                filenameEl.textContent = "\u2713 " + this.files[0].name;
                filenameEl.classList.add("visible");
            } else {
                filenameEl.textContent = "";
                filenameEl.classList.remove("visible");
            }
        });
    }

    // Drag-and-drop visual feedback
    if (uploadZone) {
        ["dragenter", "dragover"].forEach(function (evt) {
            uploadZone.addEventListener(evt, function (e) {
                e.preventDefault();
                uploadZone.classList.add("dragover");
            });
        });
        ["dragleave", "drop"].forEach(function (evt) {
            uploadZone.addEventListener(evt, function () {
                uploadZone.classList.remove("dragover");
            });
        });
    }

    // ── Cutoff mode toggle ───────────────────────────────────────────
    var cutoffRadios = document.querySelectorAll('input[name="cutoff_mode"]');
    var manualInput  = document.getElementById("cutoff-manual");
    var cutoffValue  = document.getElementById("cutoff_value");

    function updateCutoffVisibility() {
        var selected = document.querySelector('input[name="cutoff_mode"]:checked');
        if (!selected || !manualInput) return;
        if (selected.value === "manual") {
            manualInput.classList.add("visible");
            if (cutoffValue) cutoffValue.focus();
        } else {
            manualInput.classList.remove("visible");
            if (cutoffValue) cutoffValue.value = "";
        }
    }

    cutoffRadios.forEach(function (radio) {
        radio.addEventListener("change", updateCutoffVisibility);
    });

    // ── Optional components toggle ──────────────────────────────────
    var optToggle = document.getElementById("optional-toggle");
    var optFields = document.getElementById("optional-fields");
    var optArrow  = document.getElementById("optional-arrow");

    if (optToggle && optFields) {
        var optInputs = optFields.querySelectorAll("input");
        var anyFilled = Array.from(optInputs).some(function (inp) {
            return inp.value && inp.value !== "0";
        });
        if (anyFilled) {
            optFields.classList.add("open");
            optArrow.textContent = "\u25BC";
            optToggle.setAttribute("aria-expanded", "true");
        }

        optToggle.addEventListener("click", function () {
            var isOpen = optFields.classList.toggle("open");
            optArrow.textContent = isOpen ? "\u25BC" : "\u25B6";
            optToggle.setAttribute("aria-expanded", String(isOpen));
        });
    }

    // ── Submit button: loading state ────────────────────────────────
    var form      = document.getElementById("isocal-form");
    var submitBtn = document.getElementById("submit-btn");

    if (form && submitBtn) {
        form.addEventListener("submit", function () {
            setTimeout(function () {
                if (form.checkValidity()) {
                    submitBtn.classList.add("loading");
                    submitBtn.disabled = true;
                }
            }, 0);
        });
    }

    // ═══════════════════════════════════════════════════════════════
    // ── Chart Customization: Auto/Custom toggle & Plotly previews ──
    // ═══════════════════════════════════════════════════════════════

    var chartModeRadios = document.querySelectorAll('input[name="chart_mode"]');
    var customPanels    = document.getElementById("chart-custom-panels");

    function updateChartModeVisibility() {
        var selected = document.querySelector('input[name="chart_mode"]:checked');
        if (!selected || !customPanels) return;
        if (selected.value === "custom") {
            customPanels.classList.add("visible");
            renderPreviews();
        } else {
            customPanels.classList.remove("visible");
        }
    }

    chartModeRadios.forEach(function (radio) {
        radio.addEventListener("change", updateChartModeVisibility);
    });

    // Auto-open if already set to custom (e.g. after validation error)
    (function () {
        var sel = document.querySelector('input[name="chart_mode"]:checked');
        if (sel && sel.value === "custom" && customPanels) {
            customPanels.classList.add("visible");
        }
    })();

    // ── Color picker ↔ hex text sync ────────────────────────────────

    var COLOR_PAIRS = [
        ["hf_line_color_picker",   "hf_line_color"],
        ["hf_x_font_color_picker", "hf_x_font_color"],
        ["hf_y_font_color_picker", "hf_y_font_color"],
        ["hf_border_color_picker", "hf_border_color"],
        ["h_line_color_picker",    "h_line_color"],
        ["h_x_font_color_picker",  "h_x_font_color"],
        ["h_y_font_color_picker",  "h_y_font_color"],
        ["h_border_color_picker",  "h_border_color"]
    ];

    COLOR_PAIRS.forEach(function (pair) {
        var picker = document.getElementById(pair[0]);
        var text   = document.getElementById(pair[1]);
        if (!picker || !text) return;

        // Sync picker → text
        picker.addEventListener("input", function () {
            text.value = picker.value.toUpperCase();
            debouncedRender();
        });

        // Sync text → picker
        text.addEventListener("input", function () {
            var v = text.value.trim();
            if (/^#[0-9A-Fa-f]{6}$/.test(v)) {
                picker.value = v;
            }
            debouncedRender();
        });

        // Init: sync text to picker on load
        if (text.value && /^#[0-9A-Fa-f]{6}$/.test(text.value.trim())) {
            picker.value = text.value.trim();
        }
    });

    // ── Sample data for previews ────────────────────────────────────
    // Realistic isothermal calorimetry curves

    function generateSampleHeatFlow() {
        var x = [], y = [];
        for (var i = 0; i <= 700; i++) {
            var t = i * 0.01;  // 0 to 7 days
            x.push(t);
            // Typical heat flow curve: initial spike, induction, acceleration, deceleration
            var hf = 0.3 * Math.exp(-t * 8) +
                     2.8 * Math.exp(-0.5 * Math.pow(t - 0.45, 2) / 0.02) +
                     0.15 * Math.exp(-0.3 * t);
            y.push(Math.max(hf, 0.01));
        }
        return { x: x, y: y };
    }

    function generateSampleHeat() {
        var x = [], y = [];
        for (var i = 0; i <= 700; i++) {
            var t = i * 0.01;  // 0 to 7 days
            x.push(t);
            // Typical cumulative heat curve: sigmoid-like growth
            var h = 280 * (1 - Math.exp(-1.5 * t)) + 20 * t;
            y.push(h);
        }
        return { x: x, y: y };
    }

    var sampleHF = generateSampleHeatFlow();
    var sampleH  = generateSampleHeat();

    // ── Helper: read form value ─────────────────────────────────────

    function val(id, fallback) {
        var el = document.getElementById(id);
        if (!el) return fallback;
        var v = el.value.trim();
        return v === "" ? fallback : v;
    }

    function numVal(id, fallback) {
        var v = parseFloat(val(id, ""));
        return isNaN(v) ? fallback : v;
    }

    // ── Render Plotly previews ───────────────────────────────────────

    var renderTimeout = null;
    function debouncedRender() {
        clearTimeout(renderTimeout);
        renderTimeout = setTimeout(renderPreviews, 150);
    }

    function renderPreviews() {
        if (!customPanels || !customPanels.classList.contains("visible")) return;
        renderHeatFlowPreview();
        renderHeatPreview();
    }

    function renderHeatFlowPreview() {
        var container = document.getElementById("hf-preview");
        if (!container) return;

        var title      = val("hf_chart_title", "Sample (Paste)");
        var xTitle     = val("hf_x_title", "Time (Days)");
        var yTitle     = val("hf_y_title", "Heat Flow (mW/g)");
        var lineW      = numVal("hf_line_thickness", 2.25);
        var lineColor  = val("hf_line_color", "#ED7D31");
        var xFontSz    = numVal("hf_x_font_size", 10);
        var xFontClr   = val("hf_x_font_color", "#000000");
        var yFontSz    = numVal("hf_y_font_size", 10);
        var yFontClr   = val("hf_y_font_color", "#000000");
        var borderClr  = val("hf_border_color", "#FFFFFF");

        var trace = {
            x: sampleHF.x,
            y: sampleHF.y,
            mode: "lines",
            type: "scatter",
            name: title,
            line: { color: lineColor, width: lineW }
        };

        var layout = {
            title: { text: title, font: { size: 14 } },
            xaxis: {
                title: { text: xTitle, font: { size: 12 } },
                tickfont: { size: xFontSz, color: xFontClr },
                showgrid: false,
                zeroline: false,
                linecolor: "#000000",
                linewidth: 1,
                mirror: false
            },
            yaxis: {
                title: { text: yTitle, font: { size: 12 } },
                tickfont: { size: yFontSz, color: yFontClr },
                showgrid: false,
                zeroline: false,
                linecolor: "#000000",
                linewidth: 1,
                mirror: false
            },
            plot_bgcolor: "#FFFFFF",
            paper_bgcolor: "#FFFFFF",
            margin: { l: 60, r: 30, t: 50, b: 50 },
            showlegend: true,
            legend: { orientation: "h", y: -0.2 },
            shapes: [{
                type: "rect",
                xref: "paper", yref: "paper",
                x0: 0, y0: 0, x1: 1, y1: 1,
                line: { color: borderClr, width: 1 }
            }]
        };

        var config = { responsive: true, displayModeBar: false };

        Plotly.react(container, [trace], layout, config);
    }

    function renderHeatPreview() {
        var container = document.getElementById("h-preview");
        if (!container) return;

        var title      = val("h_chart_title", "Sample (Paste)");
        var xTitle     = val("h_x_title", "Time (Days)");
        var yTitle     = val("h_y_title", "Heat (J/g)");
        var lineW      = numVal("h_line_thickness", 2.25);
        var lineColor  = val("h_line_color", "#4472C4");
        var xFontSz    = numVal("h_x_font_size", 10);
        var xFontClr   = val("h_x_font_color", "#000000");
        var yFontSz    = numVal("h_y_font_size", 10);
        var yFontClr   = val("h_y_font_color", "#000000");
        var borderClr  = val("h_border_color", "#FFFFFF");

        var trace = {
            x: sampleH.x,
            y: sampleH.y,
            mode: "lines",
            type: "scatter",
            name: title,
            line: { color: lineColor, width: lineW }
        };

        var layout = {
            title: { text: title, font: { size: 14 } },
            xaxis: {
                title: { text: xTitle, font: { size: 12 } },
                tickfont: { size: xFontSz, color: xFontClr },
                showgrid: false,
                zeroline: false,
                linecolor: "#000000",
                linewidth: 1,
                mirror: false
            },
            yaxis: {
                title: { text: yTitle, font: { size: 12 } },
                tickfont: { size: yFontSz, color: yFontClr },
                showgrid: false,
                zeroline: false,
                linecolor: "#000000",
                linewidth: 1,
                mirror: false
            },
            plot_bgcolor: "#FFFFFF",
            paper_bgcolor: "#FFFFFF",
            margin: { l: 60, r: 30, t: 50, b: 50 },
            showlegend: true,
            legend: { orientation: "h", y: -0.2 },
            shapes: [{
                type: "rect",
                xref: "paper", yref: "paper",
                x0: 0, y0: 0, x1: 1, y1: 1,
                line: { color: borderClr, width: 1 }
            }]
        };

        var config = { responsive: true, displayModeBar: false };

        Plotly.react(container, [trace], layout, config);
    }

    // ── Attach real-time update listeners to all customization inputs ──

    var CUSTOM_FIELDS = [
        "hf_chart_title", "hf_x_title", "hf_y_title",
        "hf_line_thickness", "hf_line_color",
        "hf_x_font_size", "hf_x_font_color",
        "hf_y_font_size", "hf_y_font_color",
        "hf_border_color",
        "h_chart_title", "h_x_title", "h_y_title",
        "h_line_thickness", "h_line_color",
        "h_x_font_size", "h_x_font_color",
        "h_y_font_size", "h_y_font_color",
        "h_border_color"
    ];

    CUSTOM_FIELDS.forEach(function (id) {
        var el = document.getElementById(id);
        if (el) {
            el.addEventListener("input", debouncedRender);
        }
    });

    // Initial render if custom mode is already selected
    if (customPanels && customPanels.classList.contains("visible")) {
        // Slight delay so Plotly CDN has loaded
        setTimeout(renderPreviews, 200);
    }

}());
