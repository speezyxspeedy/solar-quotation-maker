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

const INVOICE_COUNTER_KEY = "sgse_invoice_counter_v1";
const SELLER_SETTINGS_KEY = "sgse_invoice_seller_settings_v1";

const DEFAULT_SELLER = {
    name: "SHREE GANESH SOLAR ENERGY",
    address: "Ground, Plot No 1764/29, C/O Mitaram Bhude, Bhandewadi Pardi, Near Hanuman Mandir, Hanuman Nagar, Nagpur",
    gstin: "27DSQPB6025M1ZC",
    state: "Maharashtra",
    state_code: "27"
};

const DEFAULT_INVOICE_DECLARATION = "We declare that this invoice shows the actual price of the goods and services described and that all particulars are true and correct.";
const DEFAULT_TAX_RATES = {
    cgst: 9,
    sgst: 9,
    igst: 18
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

const invoiceEls = {
    form: document.getElementById("invoiceForm"),
    invoiceNumberDisplay: document.getElementById("invoiceNumberDisplay"),
    invoiceSummaryBuyer: document.getElementById("invoiceSummaryBuyer"),
    invoiceSummaryDate: document.getElementById("invoiceSummaryDate"),
    invoiceSummaryQty: document.getElementById("invoiceSummaryQty"),
    invoiceSummaryTotal: document.getElementById("invoiceSummaryTotal"),
    invoiceDbStatus: document.getElementById("invoiceDbStatus"),
    newInvoice: document.getElementById("newInvoice"),
    saveInvoice: document.getElementById("saveInvoice"),
    printInvoice: document.getElementById("printInvoice"),
    downloadInvoicePdf: document.getElementById("downloadInvoicePdf"),
    saveSellerDefaults: document.getElementById("saveSellerDefaults"),
    sellerCompanyName: document.getElementById("sellerCompanyName"),
    sellerAddress: document.getElementById("sellerAddress"),
    sellerGstin: document.getElementById("sellerGstin"),
    sellerState: document.getElementById("sellerState"),
    sellerStateCode: document.getElementById("sellerStateCode"),
    useIgst: document.getElementById("useIgst"),
    invoiceNo: document.getElementById("invoiceNo"),
    invoiceDate: document.getElementById("invoiceDate"),
    deliveryNote: document.getElementById("deliveryNote"),
    invoicePaymentTerms: document.getElementById("invoicePaymentTerms"),
    dispatchThrough: document.getElementById("dispatchThrough"),
    destination: document.getElementById("destination"),
    vehicleNo: document.getElementById("vehicleNo"),
    termsOfDelivery: document.getElementById("termsOfDelivery"),
    cgstRate: document.getElementById("cgstRate"),
    sgstRate: document.getElementById("sgstRate"),
    igstRate: document.getElementById("igstRate"),
    buyerName: document.getElementById("buyerName"),
    buyerAddress: document.getElementById("buyerAddress"),
    buyerGstin: document.getElementById("buyerGstin"),
    buyerState: document.getElementById("buyerState"),
    buyerStateCode: document.getElementById("buyerStateCode"),
    shipToName: document.getElementById("shipToName"),
    shipToAddress: document.getElementById("shipToAddress"),
    shipToGstin: document.getElementById("shipToGstin"),
    shipToState: document.getElementById("shipToState"),
    shipToStateCode: document.getElementById("shipToStateCode"),
    invoiceItemRows: document.getElementById("invoiceItemRows"),
    addInvoiceRow: document.getElementById("addInvoiceRow"),
    invoiceTaxSummaryRows: document.getElementById("invoiceTaxSummaryRows"),
    invoiceTotalQuantity: document.getElementById("invoiceTotalQuantity"),
    invoiceTaxCgstHeader: document.getElementById("invoiceTaxCgstHeader"),
    invoiceTaxSgstHeader: document.getElementById("invoiceTaxSgstHeader"),
    invoiceTaxIgstHeader: document.getElementById("invoiceTaxIgstHeader"),
    invoiceSubtotal: document.getElementById("invoiceSubtotal"),
    invoiceCgstLabel: document.getElementById("invoiceCgstLabel"),
    invoiceSgstLabel: document.getElementById("invoiceSgstLabel"),
    invoiceIgstLabel: document.getElementById("invoiceIgstLabel"),
    invoiceCgst: document.getElementById("invoiceCgst"),
    invoiceSgst: document.getElementById("invoiceSgst"),
    invoiceIgst: document.getElementById("invoiceIgst"),
    invoiceRoundOff: document.getElementById("invoiceRoundOff"),
    invoiceGrandTotal: document.getElementById("invoiceGrandTotal"),
    invoiceAmountWords: document.getElementById("invoiceAmountWords"),
    invoiceAccountHolder: document.getElementById("invoiceAccountHolder"),
    invoiceBankName: document.getElementById("invoiceBankName"),
    invoiceAccountNo: document.getElementById("invoiceAccountNo"),
    invoiceIfsc: document.getElementById("invoiceIfsc"),
    invoiceBranch: document.getElementById("invoiceBranch"),
    invoiceDeclaration: document.getElementById("invoiceDeclaration"),
    invoicePreview: document.getElementById("invoicePreview"),
    refreshInvoiceHistory: document.getElementById("refreshInvoiceHistory"),
    invoiceHistoryRows: document.getElementById("invoiceHistoryRows")
};

let currentQuotationId = null;
let currentInvoiceId = null;
let invoiceHistoryLoaded = false;
let invoiceHistoryCache = [];

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

function round2(value) {
    return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function numberValue(value) {
    const number = Number(String(value || "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(number) ? number : 0;
}

function taxRateValue(value, fallback) {
    const raw = String(value ?? "").trim();
    if (!raw) return fallback;
    const rate = Number(raw.replace(/[^0-9.-]/g, ""));
    return Number.isFinite(rate) && rate >= 0 ? rate : fallback;
}

function currentTaxRates() {
    return {
        cgst: taxRateValue(invoiceEls.cgstRate.value, DEFAULT_TAX_RATES.cgst),
        sgst: taxRateValue(invoiceEls.sgstRate.value, DEFAULT_TAX_RATES.sgst),
        igst: taxRateValue(invoiceEls.igstRate.value, DEFAULT_TAX_RATES.igst)
    };
}

function formatTaxRate(rate) {
    return round2(rate).toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

function taxLabel(label, rate) {
    return `${label} ${formatTaxRate(rate)}%`;
}

function formatInvoiceMoney(value) {
    return `Rs. ${round2(value).toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    })}`;
}

function safeFileName(value) {
    return String(value || "invoice").replace(/[\\/:*?"<>|]+/g, "-");
}

function nextInvoiceNumber() {
    const next = Number(localStorage.getItem(INVOICE_COUNTER_KEY) || "0") + 1;
    localStorage.setItem(INVOICE_COUNTER_KEY, String(next));
    return `SGSE/INV/${currentYear()}/${String(next).padStart(4, "0")}`;
}

const WORDS_UNDER_TWENTY = [
    "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen"
];
const WORDS_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

function wordsBelowThousand(number) {
    const n = Math.floor(number);
    const parts = [];
    if (n >= 100) {
        parts.push(`${WORDS_UNDER_TWENTY[Math.floor(n / 100)]} Hundred`);
    }
    const rest = n % 100;
    if (rest >= 20) {
        const tens = WORDS_TENS[Math.floor(rest / 10)];
        const ones = rest % 10 ? ` ${WORDS_UNDER_TWENTY[rest % 10]}` : "";
        parts.push(`${tens}${ones}`);
    } else if (rest > 0 || !parts.length) {
        parts.push(WORDS_UNDER_TWENTY[rest]);
    }
    return parts.join(" ");
}

function numberToIndianWords(number) {
    let n = Math.floor(Math.abs(number));
    if (n === 0) return "Zero";

    const units = [
        ["Crore", 10000000],
        ["Lakh", 100000],
        ["Thousand", 1000],
        ["", 1]
    ];
    const parts = [];
    for (const [label, divisor] of units) {
        const chunk = Math.floor(n / divisor);
        if (!chunk) continue;
        parts.push(`${wordsBelowThousand(chunk)}${label ? ` ${label}` : ""}`);
        n %= divisor;
    }
    return parts.join(" ");
}

function amountToWords(value) {
    const total = round2(value);
    const rupees = Math.floor(Math.abs(total));
    const paise = Math.round((Math.abs(total) - rupees) * 100);
    const prefix = total < 0 ? "Minus " : "";
    const paiseText = paise ? ` and ${numberToIndianWords(paise)} Paise` : "";
    return `${prefix}Rupees ${numberToIndianWords(rupees)}${paiseText} Only`;
}

function invoiceLineBreaks(value = "") {
    return escapeHtml(value).replace(/\n/g, "<br>");
}

function setInvoiceStatus(message = "", type = "") {
    invoiceEls.invoiceDbStatus.textContent = message;
    invoiceEls.invoiceDbStatus.className = `status-line ${type}`.trim();
}

function sellerDefaults() {
    try {
        return { ...DEFAULT_SELLER, ...JSON.parse(localStorage.getItem(SELLER_SETTINGS_KEY) || "{}") };
    } catch {
        return { ...DEFAULT_SELLER };
    }
}

function applySellerDefaults() {
    const seller = sellerDefaults();
    invoiceEls.sellerCompanyName.value = seller.name || DEFAULT_SELLER.name;
    invoiceEls.sellerAddress.value = seller.address || DEFAULT_SELLER.address;
    invoiceEls.sellerGstin.value = seller.gstin || DEFAULT_SELLER.gstin;
    invoiceEls.sellerState.value = seller.state || DEFAULT_SELLER.state;
    invoiceEls.sellerStateCode.value = seller.state_code || DEFAULT_SELLER.state_code;
}

function saveSellerDefaults() {
    localStorage.setItem(SELLER_SETTINGS_KEY, JSON.stringify({
        name: invoiceEls.sellerCompanyName.value.trim() || DEFAULT_SELLER.name,
        address: invoiceEls.sellerAddress.value.trim() || DEFAULT_SELLER.address,
        gstin: invoiceEls.sellerGstin.value.trim() || DEFAULT_SELLER.gstin,
        state: invoiceEls.sellerState.value.trim() || DEFAULT_SELLER.state,
        state_code: invoiceEls.sellerStateCode.value.trim() || DEFAULT_SELLER.state_code
    }));
    setInvoiceStatus("Seller defaults saved.", "success");
}

function defaultInvoiceRows() {
    return [
        {
            sl_no: "1",
            description: "Solar Power Plant",
            hsn_sac: "8541",
            quantity: 1,
            unit: "SET",
            rate: 0,
            amount: 0
        }
    ];
}

function createInvoiceRow(row = {}) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="text" data-field="sl_no" value="${escapeHtml(row.sl_no || "")}" aria-label="Serial number"></td>
        <td><input type="text" data-field="description" value="${escapeHtml(row.description || "")}" aria-label="Description"></td>
        <td><input type="text" data-field="hsn_sac" value="${escapeHtml(row.hsn_sac || "")}" aria-label="HSN or SAC"></td>
        <td><input type="number" min="0" step="0.01" data-field="quantity" value="${escapeHtml(row.quantity ?? 0)}" aria-label="Quantity"></td>
        <td><input type="text" data-field="unit" value="${escapeHtml(row.unit || "")}" aria-label="Unit"></td>
        <td><input type="number" min="0" step="0.01" data-field="rate" value="${escapeHtml(row.rate ?? 0)}" aria-label="Rate"></td>
        <td><input type="text" data-field="amount" value="${round2(row.amount || 0).toFixed(2)}" aria-label="Amount" readonly></td>
        <td><button class="row-remove" type="button">Delete</button></td>
    `;
    invoiceEls.invoiceItemRows.appendChild(tr);
}

function renderInvoiceRows(rows = defaultInvoiceRows()) {
    invoiceEls.invoiceItemRows.innerHTML = "";
    rows.forEach(createInvoiceRow);
}

function invoiceRowInputs(row) {
    return {
        sl_no: row.querySelector('[data-field="sl_no"]'),
        description: row.querySelector('[data-field="description"]'),
        hsn_sac: row.querySelector('[data-field="hsn_sac"]'),
        quantity: row.querySelector('[data-field="quantity"]'),
        unit: row.querySelector('[data-field="unit"]'),
        rate: row.querySelector('[data-field="rate"]'),
        amount: row.querySelector('[data-field="amount"]')
    };
}

function syncInvoiceRowAmounts() {
    [...invoiceEls.invoiceItemRows.querySelectorAll("tr")].forEach((row, index) => {
        const inputs = invoiceRowInputs(row);
        if (!inputs.sl_no.value.trim()) inputs.sl_no.value = String(index + 1);
        const amount = round2(numberValue(inputs.quantity.value) * numberValue(inputs.rate.value));
        inputs.amount.value = amount.toFixed(2);
    });
}

function collectInvoiceItems() {
    return [...invoiceEls.invoiceItemRows.querySelectorAll("tr")]
        .map(row => {
            const inputs = invoiceRowInputs(row);
            return {
                sl_no: inputs.sl_no.value.trim(),
                description: inputs.description.value.trim(),
                hsn_sac: inputs.hsn_sac.value.trim(),
                quantity: numberValue(inputs.quantity.value),
                unit: inputs.unit.value.trim(),
                rate: numberValue(inputs.rate.value),
                amount: numberValue(inputs.amount.value)
            };
        })
        .filter(row => row.description || row.hsn_sac || row.quantity || row.rate || row.amount);
}

function calculateInvoiceTotals(items, useIgst, rates) {
    const subtotal = round2(items.reduce((sum, item) => sum + item.amount, 0));
    const totalQuantity = round2(items.reduce((sum, item) => sum + item.quantity, 0));
    const cgstRate = taxRateValue(rates?.cgst, DEFAULT_TAX_RATES.cgst);
    const sgstRate = taxRateValue(rates?.sgst, DEFAULT_TAX_RATES.sgst);
    const igstRate = taxRateValue(rates?.igst, DEFAULT_TAX_RATES.igst);
    const cgst = useIgst ? 0 : round2(subtotal * (cgstRate / 100));
    const sgst = useIgst ? 0 : round2(subtotal * (sgstRate / 100));
    const igst = useIgst ? round2(subtotal * (igstRate / 100)) : 0;
    const rawTotal = round2(subtotal + cgst + sgst + igst);
    const grandTotal = Math.round(rawTotal);
    const roundOff = round2(grandTotal - rawTotal);
    const grouped = new Map();

    items.forEach(item => {
        const key = item.hsn_sac || "-";
        grouped.set(key, (grouped.get(key) || 0) + item.amount);
    });

    const taxSummary = [...grouped.entries()].map(([hsnSac, taxable]) => {
        const rowCgst = useIgst ? 0 : round2(taxable * (cgstRate / 100));
        const rowSgst = useIgst ? 0 : round2(taxable * (sgstRate / 100));
        const rowIgst = useIgst ? round2(taxable * (igstRate / 100)) : 0;
        return {
            hsn_sac: hsnSac,
            taxable_value: round2(taxable),
            cgst: rowCgst,
            sgst: rowSgst,
            igst: rowIgst,
            total_tax: round2(rowCgst + rowSgst + rowIgst)
        };
    });

    return { subtotal, totalQuantity, cgst, sgst, igst, roundOff, grandTotal, taxSummary };
}

function invoiceParty(name, address, gstin, state, stateCode) {
    return {
        name: name.value.trim(),
        address: address.value.trim(),
        gstin: gstin.value.trim(),
        state: state.value.trim(),
        state_code: stateCode.value.trim()
    };
}

function buildInvoiceDraft() {
    syncInvoiceRowAmounts();
    const items = collectInvoiceItems();
    const useIgst = invoiceEls.useIgst.checked;
    const rates = currentTaxRates();
    const totals = calculateInvoiceTotals(items, useIgst, rates);
    const buyer = invoiceParty(
        invoiceEls.buyerName,
        invoiceEls.buyerAddress,
        invoiceEls.buyerGstin,
        invoiceEls.buyerState,
        invoiceEls.buyerStateCode
    );
    const shipTo = {
        name: invoiceEls.shipToName.value.trim() || buyer.name,
        address: invoiceEls.shipToAddress.value.trim() || buyer.address,
        gstin: invoiceEls.shipToGstin.value.trim() || buyer.gstin,
        state: invoiceEls.shipToState.value.trim() || buyer.state,
        state_code: invoiceEls.shipToStateCode.value.trim() || buyer.state_code
    };

    return {
        id: currentInvoiceId || "",
        invoice_no: invoiceEls.invoiceNo.value.trim(),
        invoice_date: invoiceEls.invoiceDate.value,
        seller: invoiceParty(
            invoiceEls.sellerCompanyName,
            invoiceEls.sellerAddress,
            invoiceEls.sellerGstin,
            invoiceEls.sellerState,
            invoiceEls.sellerStateCode
        ),
        buyer,
        ship_to: shipTo,
        delivery_note: invoiceEls.deliveryNote.value.trim(),
        payment_terms: invoiceEls.invoicePaymentTerms.value.trim(),
        dispatch_through: invoiceEls.dispatchThrough.value.trim(),
        destination: invoiceEls.destination.value.trim(),
        vehicle_no: invoiceEls.vehicleNo.value.trim(),
        terms_of_delivery: invoiceEls.termsOfDelivery.value.trim(),
        items,
        use_igst: useIgst,
        cgst_rate: rates.cgst,
        sgst_rate: rates.sgst,
        igst_rate: rates.igst,
        subtotal: totals.subtotal,
        cgst: totals.cgst,
        sgst: totals.sgst,
        igst: totals.igst,
        round_off: totals.roundOff,
        grand_total: totals.grandTotal,
        total_quantity: totals.totalQuantity,
        amount_in_words: amountToWords(totals.grandTotal),
        tax_summary: totals.taxSummary,
        bank_details: {
            account_holder: invoiceEls.invoiceAccountHolder.value.trim(),
            bank_name: invoiceEls.invoiceBankName.value.trim(),
            account_no: invoiceEls.invoiceAccountNo.value.trim(),
            ifsc: invoiceEls.invoiceIfsc.value.trim(),
            branch: invoiceEls.invoiceBranch.value.trim()
        },
        declaration: invoiceEls.invoiceDeclaration.value.trim()
    };
}

function renderInvoiceTaxSummary(rows) {
    if (!rows.length) {
        invoiceEls.invoiceTaxSummaryRows.innerHTML = `<tr><td colspan="6" class="empty-row">No item rows.</td></tr>`;
        return;
    }

    invoiceEls.invoiceTaxSummaryRows.innerHTML = rows.map(row => `
        <tr>
            <td>${escapeHtml(row.hsn_sac || "-")}</td>
            <td>${formatInvoiceMoney(row.taxable_value)}</td>
            <td>${formatInvoiceMoney(row.cgst)}</td>
            <td>${formatInvoiceMoney(row.sgst)}</td>
            <td>${formatInvoiceMoney(row.igst)}</td>
            <td>${formatInvoiceMoney(row.total_tax)}</td>
        </tr>
    `).join("");
}

function renderInvoiceItemsPreview(items) {
    if (!items.length) {
        return `<tr><td colspan="7" class="empty-row">No item rows.</td></tr>`;
    }

    return items.map((item, index) => `
        <tr>
            <td>${escapeHtml(item.sl_no || String(index + 1))}</td>
            <td>${escapeHtml(item.description || "-")}</td>
            <td>${escapeHtml(item.hsn_sac || "-")}</td>
            <td>${escapeHtml(String(item.quantity || 0))}</td>
            <td>${escapeHtml(item.unit || "-")}</td>
            <td>${formatInvoiceMoney(item.rate)}</td>
            <td>${formatInvoiceMoney(item.amount)}</td>
        </tr>
    `).join("");
}

function renderInvoiceTaxPreview(rows) {
    if (!rows.length) {
        return `<tr><td colspan="6" class="empty-row">No tax summary.</td></tr>`;
    }

    return rows.map(row => `
        <tr>
            <td>${escapeHtml(row.hsn_sac || "-")}</td>
            <td>${formatInvoiceMoney(row.taxable_value)}</td>
            <td>${formatInvoiceMoney(row.cgst)}</td>
            <td>${formatInvoiceMoney(row.sgst)}</td>
            <td>${formatInvoiceMoney(row.igst)}</td>
            <td>${formatInvoiceMoney(row.total_tax)}</td>
        </tr>
    `).join("");
}

function renderInvoicePreview(invoice) {
    const cgstLabel = taxLabel("CGST", taxRateValue(invoice.cgst_rate, DEFAULT_TAX_RATES.cgst));
    const sgstLabel = taxLabel("SGST", taxRateValue(invoice.sgst_rate, DEFAULT_TAX_RATES.sgst));
    const igstLabel = taxLabel("IGST", taxRateValue(invoice.igst_rate, DEFAULT_TAX_RATES.igst));

    invoiceEls.invoicePreview.innerHTML = `
        <article class="invoice-sheet">
            <header class="invoice-sheet-header">
                <h2>${escapeHtml(invoice.seller.name || DEFAULT_SELLER.name)}</h2>
                <p>TAX INVOICE / SOLAR INVOICE</p>
            </header>

            <section class="invoice-head-grid">
                <div class="invoice-seller-box">
                    <strong>${escapeHtml(invoice.seller.name || DEFAULT_SELLER.name)}</strong>
                    <span>${invoiceLineBreaks(invoice.seller.address || DEFAULT_SELLER.address)}</span>
                    <span>GSTIN/UIN: ${escapeHtml(invoice.seller.gstin || DEFAULT_SELLER.gstin)}</span>
                    <span>State: ${escapeHtml(invoice.seller.state || DEFAULT_SELLER.state)} &nbsp; Code: ${escapeHtml(invoice.seller.state_code || DEFAULT_SELLER.state_code)}</span>
                </div>
                <table class="invoice-print-table invoice-meta-table">
                    <tbody>
                        <tr><th>Invoice No</th><td>${escapeHtml(invoice.invoice_no || "-")}</td></tr>
                        <tr><th>Date</th><td>${escapeHtml(invoice.invoice_date || "-")}</td></tr>
                        <tr><th>Delivery Note</th><td>${escapeHtml(invoice.delivery_note || "-")}</td></tr>
                        <tr><th>Payment Terms</th><td>${escapeHtml(invoice.payment_terms || "-")}</td></tr>
                    </tbody>
                </table>
            </section>

            <section class="invoice-party-grid">
                <div>
                    <h3>Buyer (Bill To)</h3>
                    <strong>${escapeHtml(invoice.buyer.name || "-")}</strong>
                    <span>${invoiceLineBreaks(invoice.buyer.address || "-")}</span>
                    <span>GSTIN: ${escapeHtml(invoice.buyer.gstin || "-")}</span>
                    <span>State: ${escapeHtml(invoice.buyer.state || "-")} &nbsp; Code: ${escapeHtml(invoice.buyer.state_code || "-")}</span>
                </div>
                <div>
                    <h3>Ship To</h3>
                    <strong>${escapeHtml(invoice.ship_to.name || "-")}</strong>
                    <span>${invoiceLineBreaks(invoice.ship_to.address || "-")}</span>
                    <span>GSTIN: ${escapeHtml(invoice.ship_to.gstin || "-")}</span>
                    <span>State: ${escapeHtml(invoice.ship_to.state || "-")} &nbsp; Code: ${escapeHtml(invoice.ship_to.state_code || "-")}</span>
                </div>
            </section>

            <table class="invoice-print-table">
                <tbody>
                    <tr>
                        <th>Dispatch Through</th><td>${escapeHtml(invoice.dispatch_through || "-")}</td>
                        <th>Destination</th><td>${escapeHtml(invoice.destination || "-")}</td>
                    </tr>
                    <tr>
                        <th>Vehicle No</th><td>${escapeHtml(invoice.vehicle_no || "-")}</td>
                        <th>Terms of Delivery</th><td>${escapeHtml(invoice.terms_of_delivery || "-")}</td>
                    </tr>
                </tbody>
            </table>

            <table class="invoice-print-table invoice-line-table">
                <thead>
                    <tr>
                        <th>Sl No</th>
                        <th>Description</th>
                        <th>HSN/SAC</th>
                        <th>Quantity</th>
                        <th>Unit</th>
                        <th>Rate</th>
                        <th>Amount</th>
                    </tr>
                </thead>
                <tbody>${renderInvoiceItemsPreview(invoice.items)}</tbody>
            </table>

            <section class="invoice-bottom-grid">
                <div>
                    <table class="invoice-print-table">
                        <thead>
                            <tr>
                                <th>HSN/SAC</th>
                                <th>Taxable Value</th>
                                <th>${cgstLabel}</th>
                                <th>${sgstLabel}</th>
                                <th>${igstLabel}</th>
                                <th>Total Tax</th>
                            </tr>
                        </thead>
                        <tbody>${renderInvoiceTaxPreview(invoice.tax_summary)}</tbody>
                    </table>
                    <div class="invoice-words">
                        <strong>Amount in Words</strong>
                        <span>${escapeHtml(invoice.amount_in_words)}</span>
                    </div>
                </div>
                <table class="invoice-print-table invoice-total-table">
                    <tbody>
                        <tr><th>Total Quantity</th><td>${escapeHtml(String(invoice.total_quantity))}</td></tr>
                        <tr><th>Subtotal</th><td>${formatInvoiceMoney(invoice.subtotal)}</td></tr>
                        <tr><th>${cgstLabel}</th><td>${formatInvoiceMoney(invoice.cgst)}</td></tr>
                        <tr><th>${sgstLabel}</th><td>${formatInvoiceMoney(invoice.sgst)}</td></tr>
                        <tr><th>${igstLabel}</th><td>${formatInvoiceMoney(invoice.igst)}</td></tr>
                        <tr><th>Round Off</th><td>${formatInvoiceMoney(invoice.round_off)}</td></tr>
                        <tr class="grand-row"><th>Grand Total</th><td>${formatInvoiceMoney(invoice.grand_total)}</td></tr>
                    </tbody>
                </table>
            </section>

            <section class="invoice-bank-grid">
                <div>
                    <h3>Bank Details</h3>
                    <p><strong>Account Holder:</strong> ${escapeHtml(invoice.bank_details.account_holder || "-")}</p>
                    <p><strong>Bank Name:</strong> ${escapeHtml(invoice.bank_details.bank_name || "-")}</p>
                    <p><strong>Account No:</strong> ${escapeHtml(invoice.bank_details.account_no || "-")}</p>
                    <p><strong>IFSC:</strong> ${escapeHtml(invoice.bank_details.ifsc || "-")}</p>
                    <p><strong>Branch:</strong> ${escapeHtml(invoice.bank_details.branch || "-")}</p>
                </div>
                <div>
                    <h3>Declaration</h3>
                    <p>${escapeHtml(invoice.declaration || DEFAULT_INVOICE_DECLARATION)}</p>
                    <strong class="invoice-signature">For ${escapeHtml(invoice.seller.name || DEFAULT_SELLER.name)}<br>Authorised Signatory</strong>
                </div>
            </section>
        </article>
    `;
}

function updateInvoiceUi() {
    const invoice = buildInvoiceDraft();
    const cgstLabel = taxLabel("CGST", invoice.cgst_rate);
    const sgstLabel = taxLabel("SGST", invoice.sgst_rate);
    const igstLabel = taxLabel("IGST", invoice.igst_rate);
    invoiceEls.invoiceNumberDisplay.textContent = invoice.invoice_no || "-";
    invoiceEls.invoiceSummaryBuyer.textContent = invoice.buyer.name || "-";
    invoiceEls.invoiceSummaryDate.textContent = invoice.invoice_date || "-";
    invoiceEls.invoiceSummaryQty.textContent = String(invoice.total_quantity);
    invoiceEls.invoiceSummaryTotal.textContent = formatInvoiceMoney(invoice.grand_total);
    invoiceEls.invoiceTaxCgstHeader.textContent = cgstLabel;
    invoiceEls.invoiceTaxSgstHeader.textContent = sgstLabel;
    invoiceEls.invoiceTaxIgstHeader.textContent = igstLabel;
    invoiceEls.invoiceTotalQuantity.value = String(invoice.total_quantity);
    invoiceEls.invoiceSubtotal.value = formatInvoiceMoney(invoice.subtotal);
    invoiceEls.invoiceCgstLabel.textContent = cgstLabel;
    invoiceEls.invoiceSgstLabel.textContent = sgstLabel;
    invoiceEls.invoiceIgstLabel.textContent = igstLabel;
    invoiceEls.invoiceCgst.value = formatInvoiceMoney(invoice.cgst);
    invoiceEls.invoiceSgst.value = formatInvoiceMoney(invoice.sgst);
    invoiceEls.invoiceIgst.value = formatInvoiceMoney(invoice.igst);
    invoiceEls.invoiceRoundOff.value = formatInvoiceMoney(invoice.round_off);
    invoiceEls.invoiceGrandTotal.value = formatInvoiceMoney(invoice.grand_total);
    invoiceEls.invoiceAmountWords.value = invoice.amount_in_words;
    renderInvoiceTaxSummary(invoice.tax_summary);
    renderInvoicePreview(invoice);
    return invoice;
}

async function invoiceApi(endpoint, options = {}) {
    let lastError = null;
    for (const baseUrl of apiCandidates()) {
        try {
            const response = await fetch(`${baseUrl}/${endpoint}`, options);
            if (response.ok) return response;
            const details = await response.text();
            lastError = new Error(`${endpoint} failed with status ${response.status}: ${details.slice(0, 180)}`);
        } catch (error) {
            lastError = error;
        }
    }
    throw lastError || new Error(`${endpoint} failed`);
}

async function saveInvoiceToDatabase({ silent = false } = {}) {
    if (!invoiceEls.form.reportValidity()) return null;
    if (!currentInvoiceId) currentInvoiceId = makeId();
    const invoice = buildInvoiceDraft();
    invoice.id = currentInvoiceId;

    try {
        if (!silent) setInvoiceStatus("Saving invoice to database...", "");
        const response = await invoiceApi("save-invoice", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(invoice)
        });
        const result = await response.json();
        if (!result.success) throw new Error(result.message || "Invoice save failed");
        currentInvoiceId = result.invoice?.id || currentInvoiceId;
        setInvoiceStatus("Invoice saved to database.", "success");
        await loadInvoiceHistory(false);
        return result.invoice || invoice;
    } catch (error) {
        const message = `Invoice was not saved: ${error.message}`;
        setInvoiceStatus(message, "error");
        if (!silent) alert(message);
        return null;
    }
}

async function downloadInvoicePdf(invoice = null) {
    const payload = invoice || await saveInvoiceToDatabase();
    if (!payload) return;

    try {
        setInvoiceStatus("Preparing invoice PDF...", "");
        const response = await invoiceApi("generate-invoice-pdf", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });
        const blob = await response.blob();
        downloadBlob(blob, `${safeFileName(payload.invoice_no)}.pdf`);
        setInvoiceStatus("Invoice PDF downloaded.", "success");
    } catch (error) {
        const message = `Invoice PDF was not generated: ${error.message}`;
        setInvoiceStatus(message, "error");
        alert(message);
    }
}

async function printInvoice() {
    const saved = await saveInvoiceToDatabase();
    if (!saved) return;

    const cleanup = () => {
        document.body.classList.remove("print-invoice-mode");
        window.removeEventListener("afterprint", cleanup);
    };
    document.body.classList.add("print-invoice-mode");
    window.addEventListener("afterprint", cleanup);
    window.print();
}

function renderInvoiceHistory(invoices) {
    invoiceHistoryCache = invoices;
    if (!invoices.length) {
        invoiceEls.invoiceHistoryRows.innerHTML = `<tr><td colspan="5" class="empty-row">No invoices saved in database.</td></tr>`;
        return;
    }

    invoiceEls.invoiceHistoryRows.innerHTML = invoices.map(invoice => `
        <tr>
            <td>${escapeHtml(invoice.invoice_no || "-")}</td>
            <td>${escapeHtml(invoice.buyer?.name || "-")}</td>
            <td>${escapeHtml(invoice.invoice_date || "-")}</td>
            <td>${formatInvoiceMoney(invoice.grand_total || 0)}</td>
            <td>
                <div class="history-actions">
                    <button type="button" data-invoice-action="edit" data-id="${escapeHtml(invoice.id)}">Edit</button>
                    <button type="button" data-invoice-action="pdf" data-id="${escapeHtml(invoice.id)}">PDF</button>
                </div>
            </td>
        </tr>
    `).join("");
}

async function loadInvoiceHistory(showErrors = true) {
    try {
        const response = await invoiceApi("invoices");
        const result = await response.json();
        renderInvoiceHistory(Array.isArray(result.invoices) ? result.invoices : []);
        invoiceHistoryLoaded = true;
    } catch (error) {
        if (showErrors) {
            setInvoiceStatus(`Invoice database unavailable: ${error.message}`, "error");
        }
    }
}

function loadInvoice(invoice) {
    currentInvoiceId = invoice.id || null;
    invoiceEls.invoiceNo.value = invoice.invoice_no || nextInvoiceNumber();
    invoiceEls.invoiceDate.value = invoice.invoice_date || todayIso();
    invoiceEls.deliveryNote.value = invoice.delivery_note || "";
    invoiceEls.invoicePaymentTerms.value = invoice.payment_terms || "Immediate";
    invoiceEls.dispatchThrough.value = invoice.dispatch_through || "";
    invoiceEls.destination.value = invoice.destination || "";
    invoiceEls.vehicleNo.value = invoice.vehicle_no || "";
    invoiceEls.termsOfDelivery.value = invoice.terms_of_delivery || "";
    invoiceEls.useIgst.checked = Boolean(invoice.use_igst);
    invoiceEls.cgstRate.value = taxRateValue(invoice.cgst_rate, DEFAULT_TAX_RATES.cgst);
    invoiceEls.sgstRate.value = taxRateValue(invoice.sgst_rate, DEFAULT_TAX_RATES.sgst);
    invoiceEls.igstRate.value = taxRateValue(invoice.igst_rate, DEFAULT_TAX_RATES.igst);

    invoiceEls.sellerCompanyName.value = invoice.seller?.name || DEFAULT_SELLER.name;
    invoiceEls.sellerAddress.value = invoice.seller?.address || DEFAULT_SELLER.address;
    invoiceEls.sellerGstin.value = invoice.seller?.gstin || DEFAULT_SELLER.gstin;
    invoiceEls.sellerState.value = invoice.seller?.state || DEFAULT_SELLER.state;
    invoiceEls.sellerStateCode.value = invoice.seller?.state_code || DEFAULT_SELLER.state_code;

    invoiceEls.buyerName.value = invoice.buyer?.name || "";
    invoiceEls.buyerAddress.value = invoice.buyer?.address || "";
    invoiceEls.buyerGstin.value = invoice.buyer?.gstin || "";
    invoiceEls.buyerState.value = invoice.buyer?.state || "Maharashtra";
    invoiceEls.buyerStateCode.value = invoice.buyer?.state_code || "27";
    invoiceEls.shipToName.value = invoice.ship_to?.name || "";
    invoiceEls.shipToAddress.value = invoice.ship_to?.address || "";
    invoiceEls.shipToGstin.value = invoice.ship_to?.gstin || "";
    invoiceEls.shipToState.value = invoice.ship_to?.state || "Maharashtra";
    invoiceEls.shipToStateCode.value = invoice.ship_to?.state_code || "27";

    renderInvoiceRows(invoice.items?.length ? invoice.items : defaultInvoiceRows());

    invoiceEls.invoiceAccountHolder.value = invoice.bank_details?.account_holder || COMPANY.accountHolder;
    invoiceEls.invoiceBankName.value = invoice.bank_details?.bank_name || COMPANY.bankName;
    invoiceEls.invoiceAccountNo.value = invoice.bank_details?.account_no || COMPANY.accountNo;
    invoiceEls.invoiceIfsc.value = invoice.bank_details?.ifsc || COMPANY.ifsc;
    invoiceEls.invoiceBranch.value = invoice.bank_details?.branch || COMPANY.branch;
    invoiceEls.invoiceDeclaration.value = invoice.declaration || DEFAULT_INVOICE_DECLARATION;
    updateInvoiceUi();
    setActiveModule("invoice");
}

function resetInvoice() {
    currentInvoiceId = null;
    invoiceEls.form.reset();
    applySellerDefaults();
    invoiceEls.invoiceNo.value = nextInvoiceNumber();
    invoiceEls.invoiceDate.value = todayIso();
    invoiceEls.invoicePaymentTerms.value = "Immediate";
    invoiceEls.buyerState.value = "Maharashtra";
    invoiceEls.buyerStateCode.value = "27";
    invoiceEls.shipToState.value = "Maharashtra";
    invoiceEls.shipToStateCode.value = "27";
    invoiceEls.cgstRate.value = DEFAULT_TAX_RATES.cgst;
    invoiceEls.sgstRate.value = DEFAULT_TAX_RATES.sgst;
    invoiceEls.igstRate.value = DEFAULT_TAX_RATES.igst;
    invoiceEls.invoiceAccountHolder.value = COMPANY.accountHolder;
    invoiceEls.invoiceBankName.value = COMPANY.bankName;
    invoiceEls.invoiceAccountNo.value = COMPANY.accountNo;
    invoiceEls.invoiceIfsc.value = COMPANY.ifsc;
    invoiceEls.invoiceBranch.value = COMPANY.branch;
    invoiceEls.invoiceDeclaration.value = DEFAULT_INVOICE_DECLARATION;
    renderInvoiceRows();
    setInvoiceStatus("");
    updateInvoiceUi();
}

function setActiveModule(moduleName) {
    document.querySelectorAll(".module-button").forEach(button => {
        button.classList.toggle("active", button.dataset.module === moduleName);
    });
    document.querySelectorAll(".module-panel").forEach(panel => {
        panel.classList.toggle("active", panel.dataset.module === moduleName);
    });

    const quotationOnly = moduleName === "quotation";
    document.querySelector(".quotation-tab-nav").hidden = !quotationOnly;
    document.querySelector(".sidebar-actions").hidden = !quotationOnly;

    if (moduleName === "invoice") {
        updateInvoiceUi();
        if (!invoiceHistoryLoaded) loadInvoiceHistory(false);
    }
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function handleGenerate() {
    const quotation = saveCurrentQuotation();
    if (!quotation) return;
    setActiveTab("preview");
}

document.querySelectorAll(".tab-button").forEach(button => {
    button.addEventListener("click", () => setActiveTab(button.dataset.tab));
});

document.querySelectorAll(".module-button").forEach(button => {
    button.addEventListener("click", () => setActiveModule(button.dataset.module));
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

["input", "change"].forEach(eventName => {
    [
        invoiceEls.sellerCompanyName,
        invoiceEls.sellerAddress,
        invoiceEls.sellerGstin,
        invoiceEls.sellerState,
        invoiceEls.sellerStateCode,
        invoiceEls.useIgst,
        invoiceEls.invoiceNo,
        invoiceEls.invoiceDate,
        invoiceEls.deliveryNote,
        invoiceEls.invoicePaymentTerms,
        invoiceEls.dispatchThrough,
        invoiceEls.destination,
        invoiceEls.vehicleNo,
        invoiceEls.termsOfDelivery,
        invoiceEls.cgstRate,
        invoiceEls.sgstRate,
        invoiceEls.igstRate,
        invoiceEls.buyerName,
        invoiceEls.buyerAddress,
        invoiceEls.buyerGstin,
        invoiceEls.buyerState,
        invoiceEls.buyerStateCode,
        invoiceEls.shipToName,
        invoiceEls.shipToAddress,
        invoiceEls.shipToGstin,
        invoiceEls.shipToState,
        invoiceEls.shipToStateCode,
        invoiceEls.invoiceAccountHolder,
        invoiceEls.invoiceBankName,
        invoiceEls.invoiceAccountNo,
        invoiceEls.invoiceIfsc,
        invoiceEls.invoiceBranch,
        invoiceEls.invoiceDeclaration
    ].forEach(input => input.addEventListener(eventName, updateInvoiceUi));
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

invoiceEls.newInvoice.addEventListener("click", resetInvoice);
invoiceEls.saveInvoice.addEventListener("click", () => saveInvoiceToDatabase());
invoiceEls.printInvoice.addEventListener("click", printInvoice);
invoiceEls.downloadInvoicePdf.addEventListener("click", () => downloadInvoicePdf());
invoiceEls.saveSellerDefaults.addEventListener("click", saveSellerDefaults);
invoiceEls.addInvoiceRow.addEventListener("click", () => {
    createInvoiceRow({ sl_no: String(invoiceEls.invoiceItemRows.querySelectorAll("tr").length + 1), unit: "NOS", quantity: 1, rate: 0 });
    updateInvoiceUi();
});
invoiceEls.refreshInvoiceHistory.addEventListener("click", () => loadInvoiceHistory(true));

els.specRows.addEventListener("click", event => {
    const button = event.target.closest(".row-remove");
    if (!button) return;
    button.closest("tr").remove();
    renderPreview(buildDraft());
});

els.specRows.addEventListener("input", () => renderPreview(buildDraft()));
els.priceRows.addEventListener("input", () => renderPreview(buildDraft()));

invoiceEls.invoiceItemRows.addEventListener("click", event => {
    const button = event.target.closest(".row-remove");
    if (!button) return;
    const rows = invoiceEls.invoiceItemRows.querySelectorAll("tr");
    if (rows.length === 1) {
        renderInvoiceRows();
    } else {
        button.closest("tr").remove();
    }
    updateInvoiceUi();
});

invoiceEls.invoiceItemRows.addEventListener("input", updateInvoiceUi);

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

invoiceEls.invoiceHistoryRows.addEventListener("click", event => {
    const button = event.target.closest("button[data-invoice-action]");
    if (!button) return;
    const invoice = invoiceHistoryCache.find(item => item.id === button.dataset.id);
    if (!invoice) return;

    if (button.dataset.invoiceAction === "edit") loadInvoice(invoice);
    if (button.dataset.invoiceAction === "pdf") downloadInvoicePdf(invoice);
});

els.clearHistory.addEventListener("click", () => {
    if (!confirm("Clear all saved quotations from this browser?")) return;
    localStorage.removeItem(STORAGE_KEY);
    renderHistory();
});

resetQuotation();
resetInvoice();
renderHistory();
