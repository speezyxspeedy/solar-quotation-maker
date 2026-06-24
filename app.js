const STORAGE_KEY = "sgse_six_page_quotations_v1";
const COUNTER_KEY = "sgse_six_page_counter_v1";

const COMPANY = {
    name: "SHREE GANESH SOLAR ENERGY",
    office: "P/N 1764/29 HANUMAN NAGAR BHANDEWADI PARDI NAGPUR 440008",
    email: "rakeshbhude69@gmail.com",
    gst: "27DSQPB6025M1ZC",
    gstFooter: "27DSQPB6025M",
    contact: "+91 7385284731",
    bankName: "INDIAN BANK",
    accountHolder: "SHREE GANESH SOLAR ENERGY",
    accountNo: "8275515270",
    ifsc: "IDIB000R558",
    branch: "DESHPANDE LAYOUT, NAGPUR"
};

const DEFAULT_INTRODUCTION = `SHREE GANESH SOLAR ENERGY is a Govt. of India recognized solar energy solution provider. We provide On-Grid, Off-Grid, Hybrid & all kinds of solar solutions.
Our main goal is to provide simple & smart solar solutions at reasonable rates to our customers.
Our company is led by highly qualified & experienced professionals, who are expert in providing solutions on turnkey basis for solar power plant.
It is proposed to engineer, procurement & installation of solar PV system at roof of building.
This document describes the proposed solution & their details.
Our main objective is to reduce dependency on utility power.
This is our small initiative towards #Atmanirbharbharat #Startupindia #MakeinIndia #CleanIndia #Gogreen #Renewableenergy

Why to go Solar?
We know that, our resources are limited & are going to extinguish within a few decades.
Sun is the ultimate source of energy which can provide us free & limitless amount of power in the form of electricity.
By our recent technologies we can harvest solar energy & this set up is one time investment for 25 years.
The system provides us energy continuously which is emission free, reduces greenhouse effect.
The return of investment is approximately 5.0 years and then we get free energy for the rest of the next 21 years.`;

const DEFAULT_TERMS = `TERMS & CONDITIONS OF PROJECT: -

1. PRIVACY OF CONTRACT: Once the contract of Installation of Solar power plant is given to SHREE GANESH SOLAR ENERGY then from the supply of material to the installation of solar it will be completely & solely scope of our company and no further third party should involve in between without the permission of SGSE.

2. TAXES: As per the norms of Govt. of India GST taxes as applicable on Solar Power Plant, Structure, Solar Meters & other accessories (if any) shall be borne by the customer.

3. DELIVERY OF PROJECT: Standard duration of project completion may vary from 45 days to 90 days, depending on the site & documentation process.

4. SCOPE OF WORK: Supply, Design, Installation, Testing & Commissioning of project will be complete scope of SGSE. All the documentation work will be performed by SGSE except Loan case documentation (If any).

5. WARRANTY: The solar power plant carries 25 years warranty against any manufacturing defects, Solar PV modules carry a performance warranty against excess of 75% degradation in 25 years. Solar inverters carry a warranty of 10 years from manufacturer.

6. PAYMENT TERMS: Standard payment policy of SGSE is 50% payment before supply of material, 40% payment on supply of material & before completion of installation and rest 10% before commissioning of system BUT may vary according to site.

7. SAFETY & SECURITY: SHREE GANESH SOLAR ENERGY works with all the safety precautions required at site. In case of any misshape occurs because of any employee of our company then it will be solely responsibility of SGSE.

8. GOVERNING LAW AND JURISDICTION: All the questions concerning the solar system & this proposal will be governed by the laws of India & the courts at Nagpur, India.

9. HANDOVERING THE PROJECT: After successful completion of work, the project will be handed over to the client. Any physical damage will not be covered by the company.

Warranty Solar panels carry a 12-Year Product Warranty & 25-Year Performance Warranty as per manufacturer terms. Warranty covers manufacturing defects only and excludes damage due to physical impact, fire, theft, flood, lightning, wind/storm/cyclone, animal damage, over-voltage, natural calamities, or unauthorized modifications. INSURANCE: SHREE GANESH SOLAR ENERGY do not provide any kind of insurance policy. Insurance costs & process shall be borne by the customer.`;

const els = {
    form: document.getElementById("quotationForm"),
    quoteNumberDisplay: document.getElementById("quoteNumberDisplay"),
    quotationNumber: document.getElementById("quotationNumber"),
    quotationDate: document.getElementById("quotationDate"),
    customerName: document.getElementById("customerName"),
    customerLocation: document.getElementById("customerLocation"),
    capacityKw: document.getElementById("capacityKw"),
    finalPrice: document.getElementById("finalPrice"),
    panelBrand: document.getElementById("panelBrand"),
    inverterBrand: document.getElementById("inverterBrand"),
    remarkText: document.getElementById("remarkText"),
    specRows: document.getElementById("specRows"),
    priceRows: document.getElementById("priceRows"),
    priceHeaderText: document.getElementById("priceHeaderText"),
    priceNote: document.getElementById("priceNote"),
    paymentNote: document.getElementById("paymentNote"),
    introductionText: document.getElementById("introductionText"),
    termsText: document.getElementById("termsText"),
    previewPages: document.getElementById("previewPages"),
    historyRows: document.getElementById("historyRows"),
    summaryCustomer: document.getElementById("summaryCustomer"),
    summaryLocation: document.getElementById("summaryLocation"),
    summaryCapacity: document.getElementById("summaryCapacity"),
    summaryPrice: document.getElementById("summaryPrice"),
    generateQuotation: document.getElementById("generateQuotation"),
    sideGenerate: document.getElementById("sideGenerate"),
    newQuotation: document.getElementById("newQuotation"),
    addSpecRow: document.getElementById("addSpecRow"),
    resetSpecs: document.getElementById("resetSpecs"),
    resetPriceRows: document.getElementById("resetPriceRows"),
    resetIntro: document.getElementById("resetIntro"),
    resetTerms: document.getElementById("resetTerms"),
    downloadPdf: document.getElementById("downloadPdf"),
    downloadWord: document.getElementById("downloadWord"),
    downloadExcel: document.getElementById("downloadExcel"),
    shareWhatsapp: document.getElementById("shareWhatsapp"),
    clearHistory: document.getElementById("clearHistory")
};

let currentQuotationId = null;

function todayIso() {
    return new Date().toISOString().slice(0, 10);
}

function currentYear() {
    return new Date().getFullYear();
}

function nextQuoteNumber(peek = false) {
    const next = Number(localStorage.getItem(COUNTER_KEY) || "0") + 1;
    if (!peek) {
        localStorage.setItem(COUNTER_KEY, String(next));
    }
    return `SGSE-${currentYear()}-${String(next).padStart(4, "0")}`;
}

function makeId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
        return window.crypto.randomUUID();
    }
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function escapeHtml(value = "") {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function plainPrice(value) {
    const number = Number(value || 0);
    if (!Number.isFinite(number)) return String(value || "0");
    return number.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function formatRupees(value) {
    return `Rs. ${plainPrice(value)}`;
}

function capacityText(value = els.capacityKw.value) {
    const number = Number(value || 0);
    if (!Number.isFinite(number) || number <= 0) return "0KW";
    return `${String(value).trim()}KW`;
}

function moduleQuantity(value = els.capacityKw.value) {
    const number = Number(value || 0);
    if (!Number.isFinite(number) || number <= 0) return "AUTO";
    return `${Math.max(1, Math.ceil((number * 1000) / 550))}NOS.`;
}

function nl2br(value = "") {
    return escapeHtml(value)
        .split(/\n{2,}/)
        .map(block => `<p>${block.replace(/\n/g, "<br>")}</p>`)
        .join("");
}

function getHistory() {
    try {
        const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
        return Array.isArray(value) ? value : [];
    } catch {
        return [];
    }
}

function saveHistory(items) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
}

function defaultSpecificationRows() {
    const cap = capacityText();
    return [
        {
            item: "PV MODULE",
            brand_parameters: els.panelBrand.value || "Adani",
            detail: "BIFACIAL / DCR",
            specification: "MADE IN INDIA",
            quantity: moduleQuantity()
        },
        {
            item: "SOLAR STRING INVERTER",
            brand_parameters: els.inverterBrand.value || "Polycab",
            detail: "STRING INVERTER",
            specification: `${cap.replace("KW", " KW")} 1PHASE`,
            quantity: "1NOS."
        },
        {
            item: "STRUCTURE",
            brand_parameters: "3*3 BOX PIPES APOLLO",
            detail: "AS PER DESIGN",
            specification: "-",
            quantity: "AS PER DESIGN"
        },
        {
            item: "EARTHING LA LIGHTING ARRESTOR",
            brand_parameters: "STD AS PER MSEDCL",
            detail: "MSEDCL APPROVED",
            specification: "EARTHING STREEP AL MAKE",
            quantity: "3NOS."
        },
        {
            item: "SERVICE",
            brand_parameters: "2YEARS",
            detail: "IN SITES",
            specification: "MAINTANANCE",
            quantity: "FREE"
        },
        {
            item: "AC CABLE",
            brand_parameters: "POLYCAB CU MAKE 4SQ MM 4CORE",
            detail: "AS PER MSEB REQUIREMENT",
            specification: "DESIGN",
            quantity: "AS PER DESIGN"
        },
        {
            item: "DC CABLE",
            brand_parameters: "POLYCAB 4SQMM",
            detail: "PANNELS TO INVERTER",
            specification: "6SQ.MM",
            quantity: "AS PER STRUCTURE"
        },
        {
            item: "LIGHTNING ARRESTOR",
            brand_parameters: "STANDERD AS PER MSEB",
            detail: "COPPER BONDED ROD",
            specification: "UPTO EARTH",
            quantity: "1"
        },
        {
            item: "SOLAR METER",
            brand_parameters: "L&T MAKE",
            detail: "BIDIRECTION-AL",
            specification: "-",
            quantity: "2PAIR"
        }
    ];
}

function defaultPriceRows() {
    return [
        { sr_no: "1", system_detail: "RATE", rate: "" },
        { sr_no: "2", system_detail: "DEMAND", rate: "EXTRA" },
        { sr_no: "3", system_detail: "PROCESSING FEES OF DISCOM", rate: "INCLUDED" },
        { sr_no: "4", system_detail: "TOTAL", rate: `${plainPrice(els.finalPrice.value)} /-` }
    ];
}

function createSpecRow(row = {}) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="text" value="${escapeHtml(row.item || "")}" aria-label="Specification item"></td>
        <td><input type="text" value="${escapeHtml(row.brand_parameters || "")}" aria-label="Brand or parameters"></td>
        <td><input type="text" value="${escapeHtml(row.detail || "")}" aria-label="Detail"></td>
        <td><input type="text" value="${escapeHtml(row.specification || "")}" aria-label="Specification"></td>
        <td><input type="text" value="${escapeHtml(row.quantity || "")}" aria-label="Quantity"></td>
        <td><button class="row-remove" type="button">Remove</button></td>
    `;
    els.specRows.appendChild(tr);
}

function renderSpecificationRows(rows = defaultSpecificationRows()) {
    els.specRows.innerHTML = "";
    rows.forEach(createSpecRow);
}

function createPriceRow(row = {}) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="text" value="${escapeHtml(row.sr_no || "")}" aria-label="Serial number"></td>
        <td><input type="text" value="${escapeHtml(row.system_detail || "")}" aria-label="System detail"></td>
        <td><input type="text" value="${escapeHtml(row.rate || "")}" aria-label="Rate"></td>
    `;
    els.priceRows.appendChild(tr);
}

function renderPriceRows(rows = defaultPriceRows()) {
    els.priceRows.innerHTML = "";
    rows.forEach(createPriceRow);
    syncPriceRows();
}

function syncSpecificationRows() {
    const rows = [...els.specRows.querySelectorAll("tr")];
    const pvInputs = rows[0]?.querySelectorAll("input");
    const inverterInputs = rows[1]?.querySelectorAll("input");
    const cap = capacityText().replace("KW", " KW");

    if (pvInputs && pvInputs[0].value.trim().toUpperCase() === "PV MODULE") {
        pvInputs[1].value = els.panelBrand.value;
        pvInputs[4].value = moduleQuantity();
    }
    if (inverterInputs && inverterInputs[0].value.trim().toUpperCase() === "SOLAR STRING INVERTER") {
        inverterInputs[1].value = els.inverterBrand.value;
        inverterInputs[3].value = `${cap} 1PHASE`;
        inverterInputs[4].value = "1NOS.";
    }
}

function syncPriceRows() {
    const cap = capacityText();
    els.priceHeaderText.textContent = `${cap} SANCTION AND ${cap} PV MODULE`;
    [...els.priceRows.querySelectorAll("tr")].forEach(row => {
        const inputs = row.querySelectorAll("input");
        if ((inputs[1]?.value || "").trim().toUpperCase() === "TOTAL") {
            inputs[2].value = `${plainPrice(els.finalPrice.value)} /-`;
        }
    });
}

function collectSpecifications() {
    return [...els.specRows.querySelectorAll("tr")]
        .map(row => {
            const inputs = row.querySelectorAll("input");
            return {
                item: inputs[0].value.trim(),
                brand_parameters: inputs[1].value.trim(),
                detail: inputs[2].value.trim(),
                specification: inputs[3].value.trim(),
                quantity: inputs[4].value.trim()
            };
        })
        .filter(row => row.item || row.brand_parameters || row.detail || row.specification || row.quantity);
}

function collectPriceRows() {
    return [...els.priceRows.querySelectorAll("tr")]
        .map(row => {
            const inputs = row.querySelectorAll("input");
            return {
                sr_no: inputs[0].value.trim(),
                system_detail: inputs[1].value.trim(),
                rate: inputs[2].value.trim()
            };
        })
        .filter(row => row.sr_no || row.system_detail || row.rate);
}

function buildDraft() {
    syncPriceRows();
    return {
        id: currentQuotationId || "",
        quotation_number: els.quotationNumber.value.trim() || nextQuoteNumber(true),
        customer_name: els.customerName.value.trim(),
        location: els.customerLocation.value.trim(),
        capacity_kw: els.capacityKw.value.trim(),
        quote_date: els.quotationDate.value,
        final_price: els.finalPrice.value.trim(),
        panel_brand: els.panelBrand.value,
        inverter_brand: els.inverterBrand.value,
        introduction: els.introductionText.value.trim(),
        specifications: collectSpecifications(),
        price_header: els.priceHeaderText.textContent.trim(),
        price_rows: collectPriceRows(),
        price_note: els.priceNote.value.trim(),
        payment_note: els.paymentNote.value.trim(),
        terms_conditions: els.termsText.value.trim(),
        remark: els.remarkText.value.trim()
    };
}

function collectQuotation(assignNumber = false) {
    if (!els.form.reportValidity()) return null;
    const quotation = buildDraft();
    if (assignNumber && !currentQuotationId) {
        quotation.quotation_number = nextQuoteNumber(false);
        quotation.id = makeId();
        currentQuotationId = quotation.id;
    }
    if (currentQuotationId) {
        quotation.id = currentQuotationId;
    }
    els.quotationNumber.value = quotation.quotation_number;
    els.quoteNumberDisplay.textContent = quotation.quotation_number;
    return quotation;
}

function updateSummary() {
    els.summaryCustomer.textContent = els.customerName.value.trim() || "-";
    els.summaryLocation.textContent = els.customerLocation.value.trim() || "-";
    els.summaryCapacity.textContent = els.capacityKw.value ? `${els.capacityKw.value} KW` : "-";
    els.summaryPrice.textContent = formatRupees(els.finalPrice.value);
    els.quoteNumberDisplay.textContent = els.quotationNumber.value || nextQuoteNumber(true);
}

function pageFooter() {
    return `
        <footer>
            <strong>${COMPANY.name}</strong><br>
            OFFICE ADD. ${COMPANY.office}<br>
            Email: ${COMPANY.email} &nbsp; Gst No: ${COMPANY.gstFooter}
        </footer>
    `;
}

function renderSpecPreview(specifications) {
    return specifications.map(row => `
        <tr>
            <td>${escapeHtml(row.item)}</td>
            <td>${escapeHtml(row.brand_parameters)}</td>
            <td>${escapeHtml(row.detail)}</td>
            <td>${escapeHtml(row.specification)}</td>
            <td>${escapeHtml(row.quantity)}</td>
        </tr>
    `).join("");
}

function renderPricePreview(quotation) {
    return quotation.price_rows.map(row => `
        <tr>
            <td>${escapeHtml(row.sr_no)}</td>
            <td>${escapeHtml(row.system_detail)}</td>
            <td>${escapeHtml(row.rate)}</td>
        </tr>
    `).join("");
}

function renderPreview(quotation = buildDraft()) {
    const customerName = quotation.customer_name || "{customer_name}";
    const location = quotation.location || "{location}";
    const capacity = quotation.capacity_kw || "{capacity}";
    const quoteDate = quotation.quote_date || "{date}";
    const price = quotation.final_price ? formatRupees(quotation.final_price) : "Rs. 0";

    els.previewPages.innerHTML = `
        <article class="preview-page cover-preview">
            <div class="cover-top">
                <div class="logo-box"><img src="assets/logo.png" alt="Company logo"></div>
                <div class="cover-company">
                    <div class="gst-box">GST No: ${COMPANY.gst}</div>
                    <h2>${COMPANY.name}</h2>
                </div>
            </div>
            <h3>QUOTATION</h3>
            <h4>Proposal for installation, Engineering and Procurement of<br>ON GRID Connected Roof-Top Solar PV System.</h4>
            <div class="customer-block">
                <p><strong>NAME -</strong> ${escapeHtml(customerName)}</p>
                <p><strong>LOCATION -</strong> ${escapeHtml(location)}</p>
                <p><strong>CAPACITY -</strong> ${escapeHtml(capacity)}KW</p>
                <p><strong>DATE -</strong> ${escapeHtml(quoteDate)}</p>
            </div>
            ${pageFooter()}
        </article>

        <article class="preview-page">
            <h2>INTRODUCTION</h2>
            <div class="text-flow">${nl2br(quotation.introduction || DEFAULT_INTRODUCTION)}</div>
            ${pageFooter()}
        </article>

        <article class="preview-page">
            <h2>SYSTEM SPECIFICATION</h2>
            <table class="print-table compact-table">
                <thead>
                    <tr>
                        <th>ITEM</th>
                        <th>BRAND/PARAMETERS</th>
                        <th>DETAIL</th>
                        <th>SPECIFICATION</th>
                        <th>QUANTITY</th>
                    </tr>
                </thead>
                <tbody>${renderSpecPreview(quotation.specifications)}</tbody>
            </table>
            ${pageFooter()}
        </article>

        <article class="preview-page">
            <h2>PRICE QUOTION</h2>
            <div class="price-headline preview-price-headline">${escapeHtml(quotation.price_header)}</div>
            <table class="print-table">
                <thead>
                    <tr><th>SR NO</th><th>SYSTEM DETAIL</th><th>RATE</th></tr>
                </thead>
                <tbody>${renderPricePreview(quotation)}</tbody>
            </table>
            <p class="bullet-note">* ${escapeHtml(quotation.price_note)}</p>
            <p class="payment-note">${escapeHtml(quotation.payment_note)}</p>
            <p class="price-total">Final Amount: ${escapeHtml(price)}</p>
            ${pageFooter()}
        </article>

        <article class="preview-page terms-preview">
            <h2>TERMS & CONDITIONS</h2>
            <div class="text-flow two-column-text">${nl2br(quotation.terms_conditions || DEFAULT_TERMS)}</div>
            ${pageFooter()}
        </article>

        <article class="preview-page">
            <h2>BANK DETAILS + SIGNATURE</h2>
            <section class="remark-box">
                <strong>REMARK IF ANY</strong>
                <p>${escapeHtml(quotation.remark || "-")}</p>
            </section>
            <table class="print-table bank-table">
                <tbody>
                    <tr><th>MAIL ID</th><td>${COMPANY.email}</td></tr>
                    <tr><th>CONTACT NO</th><td>${COMPANY.contact}</td></tr>
                    <tr><th>ACCOUNT HOLDER</th><td>${COMPANY.accountHolder}</td></tr>
                    <tr><th>BANK NAME</th><td>${COMPANY.bankName}</td></tr>
                    <tr><th>CC ACCOUNT NO</th><td>${COMPANY.accountNo}</td></tr>
                    <tr><th>IFSC CODE</th><td>${COMPANY.ifsc}</td></tr>
                    <tr><th>GST</th><td>${COMPANY.gst}</td></tr>
                    <tr><th>BRANCH</th><td>${COMPANY.branch}</td></tr>
                </tbody>
            </table>
            <div class="signature-line">SIGNATURE & SEAL OF DIRECTOR</div>
            ${pageFooter()}
        </article>
    `;
}

function saveCurrentQuotation() {
    const quotation = collectQuotation(true);
    if (!quotation) return null;
    const history = getHistory().filter(item => item.id !== quotation.id);
    saveHistory([quotation, ...history]);
    renderHistory();
    renderPreview(quotation);
    updateSummary();
    return quotation;
}

function loadQuotation(quotation, tab = "details") {
    currentQuotationId = quotation.id || null;
    els.quotationNumber.value = quotation.quotation_number || nextQuoteNumber(true);
    els.quotationDate.value = quotation.quote_date || todayIso();
    els.customerName.value = quotation.customer_name || "";
    els.customerLocation.value = quotation.location || "";
    els.capacityKw.value = quotation.capacity_kw || "";
    els.finalPrice.value = quotation.final_price || "";
    els.panelBrand.value = quotation.panel_brand || "Adani";
    els.inverterBrand.value = quotation.inverter_brand || "Polycab";
    els.remarkText.value = quotation.remark || "";
    els.introductionText.value = quotation.introduction || DEFAULT_INTRODUCTION;
    els.termsText.value = quotation.terms_conditions || DEFAULT_TERMS;
    els.priceNote.value = quotation.price_note || "DEMAND CHARGES FOR LOAD EXTENSION OR REDUCTION IS NOT INCLUDED.";
    els.paymentNote.value = quotation.payment_note || "PAYMENT TERMS CAN BE ALTERED AS PER MUTUAL DECISION OF CLIENT AND COMPANY.";
    renderSpecificationRows(quotation.specifications?.length ? quotation.specifications : defaultSpecificationRows());
    renderPriceRows(quotation.price_rows?.length ? quotation.price_rows : defaultPriceRows());
    updateSummary();
    renderPreview(buildDraft());
    setActiveTab(tab);
}

function resetQuotation() {
    currentQuotationId = null;
    els.form.reset();
    els.quotationDate.value = todayIso();
    els.quotationNumber.value = nextQuoteNumber(true);
    els.quoteNumberDisplay.textContent = els.quotationNumber.value;
    els.panelBrand.value = "Adani";
    els.inverterBrand.value = "Polycab";
    els.introductionText.value = DEFAULT_INTRODUCTION;
    els.termsText.value = DEFAULT_TERMS;
    els.priceNote.value = "DEMAND CHARGES FOR LOAD EXTENSION OR REDUCTION IS NOT INCLUDED.";
    els.paymentNote.value = "PAYMENT TERMS CAN BE ALTERED AS PER MUTUAL DECISION OF CLIENT AND COMPANY.";
    renderSpecificationRows();
    renderPriceRows();
    updateSummary();
    renderPreview(buildDraft());
    setActiveTab("details");
}

function renderHistory() {
    const history = getHistory();
    if (!history.length) {
        els.historyRows.innerHTML = `<tr><td colspan="7" class="empty-row">No quotations saved yet.</td></tr>`;
        return;
    }

    els.historyRows.innerHTML = history.map(item => `
        <tr>
            <td>${escapeHtml(item.quotation_number || "-")}</td>
            <td>${escapeHtml(item.customer_name || "-")}</td>
            <td>${escapeHtml(item.location || "-")}</td>
            <td>${escapeHtml(item.capacity_kw || "-")} KW</td>
            <td>${formatRupees(item.final_price)}</td>
            <td>${escapeHtml(item.quote_date || "-")}</td>
            <td>
                <div class="history-actions">
                    <button type="button" data-action="view" data-id="${escapeHtml(item.id)}">View</button>
                    <button type="button" data-action="edit" data-id="${escapeHtml(item.id)}">Edit</button>
                    <button type="button" data-action="delete" data-id="${escapeHtml(item.id)}">Delete</button>
                    <button type="button" data-action="pdf" data-id="${escapeHtml(item.id)}">PDF</button>
                    <button type="button" data-action="word" data-id="${escapeHtml(item.id)}">Word</button>
                    <button type="button" data-action="excel" data-id="${escapeHtml(item.id)}">Excel</button>
                    <button class="whatsapp-action small" type="button" data-action="whatsapp" data-id="${escapeHtml(item.id)}">WhatsApp</button>
                </div>
            </td>
        </tr>
    `).join("");
}

function setActiveTab(tabName) {
    document.querySelectorAll(".tab-button").forEach(button => {
        button.classList.toggle("active", button.dataset.tab === tabName);
    });
    document.querySelectorAll(".tab-panel").forEach(panel => {
        panel.classList.toggle("active", panel.dataset.tab === tabName);
    });
    if (tabName === "preview") {
        renderPreview(buildDraft());
    }
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function apiCandidates() {
    const urls = [];
    if (window.location.protocol !== "file:") {
        urls.push("");
    }
    urls.push("http://127.0.0.1:8000");
    return [...new Set(urls)];
}

async function generateServerExport(type, quotation) {
    const endpoint = {
        pdf: "generate-pdf",
        word: "generate-word",
        excel: "generate-excel"
    }[type];

    let lastError = null;
    for (const baseUrl of apiCandidates()) {
        try {
            const response = await fetch(`${baseUrl}/${endpoint}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(quotation)
            });
            if (response.ok) return response.blob();
            lastError = new Error(`Export failed with status ${response.status}`);
        } catch (error) {
            lastError = error;
        }
    }
    throw lastError || new Error("Export failed");
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

async function downloadExport(type, quotation = null) {
    const payload = quotation || saveCurrentQuotation();
    if (!payload) return;

    try {
        const blob = await generateServerExport(type, payload);
        const extension = type === "word" ? "docx" : type === "excel" ? "xlsx" : "pdf";
        downloadBlob(blob, `${payload.quotation_number}.${extension}`);
    } catch {
        alert(`Unable to generate ${type.toUpperCase()} file. Please start the FastAPI backend and try again.`);
    }
}

function openWhatsapp(quotation = null) {
    const payload = quotation || saveCurrentQuotation();
    if (!payload) return;

    const message = `Hello ${payload.customer_name},
Your solar quotation for ${payload.capacity_kw}KW is ready.
Final Amount: ₹${plainPrice(payload.final_price)}
Regards,
${COMPANY.name}`;

    window.open(`https://wa.me/?text=${encodeURIComponent(message)}`, "_blank");
}

function handleGenerate() {
    const quotation = saveCurrentQuotation();
    if (!quotation) return;
    setActiveTab("preview");
}

document.querySelectorAll(".tab-button").forEach(button => {
    button.addEventListener("click", () => setActiveTab(button.dataset.tab));
});

document.querySelectorAll("[data-go-tab]").forEach(button => {
    button.addEventListener("click", () => setActiveTab(button.dataset.goTab));
});

["input", "change"].forEach(eventName => {
    [
        els.customerName,
        els.customerLocation,
        els.capacityKw,
        els.finalPrice,
        els.panelBrand,
        els.inverterBrand,
        els.quotationDate,
        els.remarkText,
        els.introductionText,
        els.termsText,
        els.priceNote,
        els.paymentNote
    ].forEach(input => {
        input.addEventListener(eventName, () => {
            if ([els.capacityKw, els.panelBrand, els.inverterBrand].includes(input)) {
                syncSpecificationRows();
            }
            syncPriceRows();
            updateSummary();
            renderPreview(buildDraft());
        });
    });
});

els.generateQuotation.addEventListener("click", handleGenerate);
els.sideGenerate.addEventListener("click", handleGenerate);
els.newQuotation.addEventListener("click", resetQuotation);
els.addSpecRow.addEventListener("click", () => {
    createSpecRow();
    renderPreview(buildDraft());
});
els.resetSpecs.addEventListener("click", () => {
    renderSpecificationRows();
    renderPreview(buildDraft());
});
els.resetPriceRows.addEventListener("click", () => {
    renderPriceRows();
    renderPreview(buildDraft());
});
els.resetIntro.addEventListener("click", () => {
    els.introductionText.value = DEFAULT_INTRODUCTION;
    renderPreview(buildDraft());
});
els.resetTerms.addEventListener("click", () => {
    els.termsText.value = DEFAULT_TERMS;
    renderPreview(buildDraft());
});
els.downloadPdf.addEventListener("click", () => downloadExport("pdf"));
els.downloadWord.addEventListener("click", () => downloadExport("word"));
els.downloadExcel.addEventListener("click", () => downloadExport("excel"));
els.shareWhatsapp.addEventListener("click", () => openWhatsapp());

els.specRows.addEventListener("click", event => {
    const button = event.target.closest(".row-remove");
    if (!button) return;
    button.closest("tr").remove();
    renderPreview(buildDraft());
});

els.specRows.addEventListener("input", () => renderPreview(buildDraft()));
els.priceRows.addEventListener("input", () => renderPreview(buildDraft()));

els.historyRows.addEventListener("click", event => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const quotation = getHistory().find(item => item.id === button.dataset.id);
    if (!quotation) return;

    if (button.dataset.action === "view") loadQuotation(quotation, "preview");
    if (button.dataset.action === "edit") loadQuotation(quotation, "details");
    if (button.dataset.action === "delete") {
        if (!confirm(`Delete quotation ${quotation.quotation_number}?`)) return;
        saveHistory(getHistory().filter(item => item.id !== quotation.id));
        if (currentQuotationId === quotation.id) resetQuotation();
        renderHistory();
    }
    if (button.dataset.action === "pdf") downloadExport("pdf", quotation);
    if (button.dataset.action === "word") downloadExport("word", quotation);
    if (button.dataset.action === "excel") downloadExport("excel", quotation);
    if (button.dataset.action === "whatsapp") openWhatsapp(quotation);
});

els.clearHistory.addEventListener("click", () => {
    if (!confirm("Clear all saved quotations from this browser?")) return;
    localStorage.removeItem(STORAGE_KEY);
    renderHistory();
});

resetQuotation();
renderHistory();
