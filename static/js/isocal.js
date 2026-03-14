/* LMC Lab — Isothermal Calorimetry page JS */

(function () {
    "use strict";

    // ── File input: show selected filename ──────────────────────────
    const fileInput    = document.getElementById("file-input");
    const filenameEl   = document.getElementById("upload-filename");
    const uploadZone   = document.getElementById("upload-zone");

    if (fileInput) {
        fileInput.addEventListener("change", function () {
            if (this.files && this.files[0]) {
                filenameEl.textContent = "✓ " + this.files[0].name;
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
    const cutoffRadios = document.querySelectorAll('input[name="cutoff_mode"]');
    const manualInput  = document.getElementById("cutoff-manual");
    const cutoffValue  = document.getElementById("cutoff_value");

    function updateCutoffVisibility() {
        const selected = document.querySelector('input[name="cutoff_mode"]:checked');
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
    const optToggle = document.getElementById("optional-toggle");
    const optFields = document.getElementById("optional-fields");
    const optArrow  = document.getElementById("optional-arrow");

    if (optToggle && optFields) {
        // Auto-open if any optional field has a value (e.g. after validation error)
        const optInputs = optFields.querySelectorAll("input");
        const anyFilled = Array.from(optInputs).some(function (inp) {
            return inp.value && inp.value !== "0";
        });
        if (anyFilled) {
            optFields.classList.add("open");
            optArrow.textContent = "▼";
            optToggle.setAttribute("aria-expanded", "true");
        }

        optToggle.addEventListener("click", function () {
            const isOpen = optFields.classList.toggle("open");
            optArrow.textContent = isOpen ? "▼" : "▶";
            optToggle.setAttribute("aria-expanded", String(isOpen));
        });
    }

    // ── Submit button: loading state ────────────────────────────────
    const form      = document.getElementById("isocal-form");
    const submitBtn = document.getElementById("submit-btn");

    if (form && submitBtn) {
        form.addEventListener("submit", function () {
            // Small delay so browser has time to validate required fields first
            setTimeout(function () {
                if (form.checkValidity()) {
                    submitBtn.classList.add("loading");
                    submitBtn.disabled = true;
                }
            }, 0);
        });
    }

}());
