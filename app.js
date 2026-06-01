const EXCHANGE_RATE = 3.67;
const HOURS_PER_DAY = 8;
const FISCAL_QUARTER_END_MONTHS = {
  1: 7,
  2: 10,
  3: 1,
  4: 4
};
const FISCAL_QUARTER_START_MONTHS = {
  1: 5,
  2: 8,
  3: 11,
  4: 2
};
const MONTH_INDEX = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11
};
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const LOCAL_PERSON_IMAGES = {
  "ali omar": "assets/pics/ali-omar.png",
  "amr hussam": "assets/pics/amr-hussam.png",
  "karim khaled": "assets/pics/karim-khaled.png",
  "manal murad": "assets/pics/manal-murad.png",
  "manal mourad": "assets/pics/manal-murad.png",
  "khaled selim": "assets/pics/khaled-selim.png"
};

const SAMPLE_PROJECTS = [
  buildProjectFromEstimatorRows(
    {
      sourceName: "TAQA.xlsx sample",
      number: "400379228",
      contractNumber: "301018758",
      name: "A7ZNQZ_ FP_TECH TAQA Fusion Analytics Accelerator CLOUD",
      customerLogoImage: "assets/customer-logos/taqa.png",
      fundUsd: 499932,
      fundAed: 1834750.44,
      accountNumber: "11105402",
      totalExpenses: 23123,
      utilizedExpenses: 0,
      origBidHours: 6194,
      origBidRevenue: 500340,
      origBidCost: 262894,
      origBidMargin: 0.525430707119159,
      estimateAtClosureHours: 7481,
      estimateAtClosureRevenue: 500340,
      estimateAtClosureCost: 238492,
      estimateAtClosureMargin: 0.47665987128752446,
      hoursWithoutContingency: 6194,
      hoursWithContingency: 7744,
      actualToDateHours: 4599.9,
      projectOwner: "Amr Hussam",
      portfolioLead: "Ali Omar",
      milestones: [
        { sno: 1, milestone: "Kick-Off Meeting Notes", percentage: 0.1, feeAed: 183625, expectedDate: excelSerialDate(46078), certificateStatus: "Submitted", invoiceStatus: "Paid", notes: "" },
        { sno: 2, milestone: "Project Plan", percentage: 0.15, feeAed: 275437.5, expectedDate: excelSerialDate(46108), certificateStatus: "Submitted", invoiceStatus: "Paid", notes: "" },
        { sno: 3, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - Finance - Ledger 1", percentage: 0.1, feeAed: 183625, expectedDate: excelSerialDate(46108), certificateStatus: "Submitted", invoiceStatus: "Paid", notes: "" },
        { sno: 4, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - Finance - Ledger 2", percentage: 0.1, feeAed: 183625, expectedDate: excelSerialDate(46192), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 5, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - Finance - Ledger 3", percentage: 0.1, feeAed: 183625, expectedDate: excelSerialDate(46192), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 6, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - SCM - Ledger 1", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46108), certificateStatus: "Submitted", invoiceStatus: "Paid", notes: "" },
        { sno: 7, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - SCM - Ledger 2", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46192), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 8, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - HR - Operating Unit 1", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46108), certificateStatus: "Submitted", invoiceStatus: "Paid", notes: "" },
        { sno: 9, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - HR - Operating Unit 2", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46192), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 10, milestone: "Submission of Technical Configuration Document for Enterprise Analytics - HR - Operating Unit 3", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46192), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 11, milestone: "Submission of Functional Requirement Document for Oracle Data Warehouse for Maximo", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46127), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Delayed, expected new date 15-6-2026" },
        { sno: 12, milestone: "Submission of Technical Design Document for Oracle Data Warehouse for Maximo", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46136), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Delayed, expected new date 15-6-2027" },
        { sno: 13, milestone: "Submission of UAT Issue log tracker for Enterprise Analytics and Oracle Data Warehouse for Maximo", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46247), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" },
        { sno: 14, milestone: "Submission of PPS Issue log tracker for Enterprise Analytics and Oracle Data Warehouse for Maximo", percentage: 0.05, feeAed: 91812.5, expectedDate: excelSerialDate(46288), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "" }
      ],
      forecastRevenue: 502062.39204545447,
      chartQuarterlyRows: [
        { label: "Q3-26", revenue: 144998.705722071, cumulativeRevenue: 144998.705722071 },
        { label: "Q4-26", revenue: 152546.04995974718, cumulativeRevenue: 297544.75568181818 },
        { label: "Q1-27", revenue: 185150.43595041323, cumulativeRevenue: 482695.19163223141 },
        { label: "Q2-27", revenue: 19367.200413223065, cumulativeRevenue: 502062.39204545447 }
      ],
      chartMonthlyRows: [
        { label: "Q3-26", revenue: 144998.705722071, cumulativeRevenue: 144998.705722071 },
        { label: "Q4-26", revenue: 152546.04995974718, cumulativeRevenue: 297544.75568181818 },
        { label: "Jun 2026", revenue: 58876.28925619839, cumulativeRevenue: 356421.04493801657 },
        { label: "Jul 2026", revenue: 73595.361570247915, cumulativeRevenue: 430016.40650826448 },
        { label: "Aug 2026", revenue: 52678.785123966925, cumulativeRevenue: 482695.19163223141 },
        { label: "Q2-27", revenue: 19367.200413223065, cumulativeRevenue: 502062.39204545447 }
      ]
    },
    [
      { label: "Q3-26", revenue: 144998.705722071, cumulativeRevenue: 144998.705722071, days: 286.5, hours: 2292, cumulativeDays: 286.5, eacDays: 968, eacHours: 7744, poc: 0.29597107438016529, funding: 499932 },
      { label: "Q4-26", revenue: 152546.04995974718, cumulativeRevenue: 297544.75568181818, days: 289.625, hours: 2317, cumulativeDays: 576.125, eacDays: 968, eacHours: 7744, poc: 0.59517045454545459, funding: 499932 },
      { label: "Jun 2026", revenue: 58876.28925619839, cumulativeRevenue: 356421.04493801657, days: 114, hours: 912, cumulativeDays: 690.125, eacDays: 968, eacHours: 7744, poc: 0.7129390495867769, funding: 499932 },
      { label: "Jul 2026", revenue: 73595.361570247915, cumulativeRevenue: 430016.40650826448, days: 142.5, hours: 1140, cumulativeDays: 832.625, eacDays: 968, eacHours: 7744, poc: 0.86014979338842978, funding: 499932 },
      { label: "Aug 2026", revenue: 52678.785123966925, cumulativeRevenue: 482695.19163223141, days: 102, hours: 816, cumulativeDays: 934.625, eacDays: 968, eacHours: 7744, poc: 0.96552169421487599, funding: 499932 }
    ],
    "Workbook estimator"
  ),
  buildProjectFromEstimatorRows(
    {
      sourceName: "Master Sheet.xlsx sample",
      number: "400398318",
      contractNumber: "301020783",
      accountNumber: "15426325",
      name: "A8XQVP_FP_Solution+_Workload Migration",
      customerLogoImage: "assets/customer-logos/solution-plus.png",
      fundUsd: 620000,
      fundAed: 2277880,
      totalExpenses: 117453,
      utilizedExpenses: 0,
      origBidHours: 3914,
      origBidRevenue: 620000,
      origBidCost: 309299,
      origBidMargin: 0.4988693548387097,
      estimateAtClosureHours: 5592,
      estimateAtClosureRevenue: 620000,
      estimateAtClosureCost: 269619,
      estimateAtClosureMargin: 0.4348693548387097,
      hoursWithoutContingency: 4888,
      hoursWithContingency: 5592,
      actualToDateHours: 3404,
      projectOwner: "Manal Murad",
      portfolioLead: "Khaled Selim",
      milestones: [
        { sno: 1, milestone: "Completion of deliverable D02: Project Management Plan including Governance setup.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46142), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 2, milestone: "Completion of deliverable D05: Data Collection and Discovery Workshops", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46142), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 3, milestone: "Completion of deliverable D09: High level Architecture covering Target Network Design, Target Security principles, Supporting Infrastructure.", percentage: 0.05, feeAed: 113863, expectedDate: excelSerialDate(46142), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 4, milestone: "Completion of deliverable D10: Candidate Application List (CAL) for Each Move Groups (Tools extract).", percentage: 0.05, feeAed: 113863, expectedDate: excelSerialDate(46151), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 5, milestone: "Completion of deliverable D19: Execution of Trail Run.", percentage: 0.05, feeAed: 113863, expectedDate: excelSerialDate(46151), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 6, milestone: "Completion of deliverable D21: Technical Verification document for first (1st) batch of fifty (50) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46158), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 7, milestone: "Completion of deliverable D22: Technical Verification document for second (2nd) batch of fifty (50) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46172), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 8, milestone: "Completion of deliverable D23: Technical Verification document for third (3rd) batch of fifty (50) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46186), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 9, milestone: "Completion of deliverable D24: Technical Verification document for fourth (4th) batch of fifty (50) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46200), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 10, milestone: "Completion of deliverable D25: Technical Verification document for fifth (5th) batch of fifty (50) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46214), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 11, milestone: "Completion of deliverable D26: Technical Verification document for sixth (6th) batch of up to forty-seven (47) migrated workloads.", percentage: 0.1, feeAed: 227726, expectedDate: excelSerialDate(46242), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" },
        { sno: 12, milestone: "Completion of deliverable D27: Hypercare and Handover", percentage: 0.05, feeAed: 113863, expectedDate: excelSerialDate(46256), certificateStatus: "Not Yet Submitted", invoiceStatus: "Not Paid Yet", notes: "Project Delayed by 3 months due to Pre-Sales Discussion" }
      ],
      forecastRevenue: 616946,
      chartQuarterlyRows: [
        { label: "Q4-26", revenue: 426684, cumulativeRevenue: 426684 },
        { label: "Q1-27", revenue: 150262, cumulativeRevenue: 576946 },
        { label: "Q2-27", revenue: 40000, cumulativeRevenue: 616946 }
      ],
      chartMonthlyRows: [
        { label: "Mar 2026", revenue: 18290, cumulativeRevenue: 18290, hours: 336, days: 42, cumulativeDays: 42, eacDays: 699, poc: 0.0295 },
        { label: "Apr 2026", revenue: 121210, cumulativeRevenue: 139500, hours: 1340, days: 167.5, cumulativeDays: 209.5, eacDays: 699, poc: 0.225 },
        { label: "May 2026", revenue: 287184, cumulativeRevenue: 426684, hours: 1728, days: 216, cumulativeDays: 425.5, eacDays: 699, poc: 0.6882 },
        { label: "Jun 2026", revenue: 50000, cumulativeRevenue: 476684, hours: 451, days: 56.375, cumulativeDays: 481.875, eacDays: 699, poc: 0.7688 },
        { label: "Jul 2026", revenue: 50262, cumulativeRevenue: 526946, hours: 453.3, days: 56.6625, cumulativeDays: 538.5375, eacDays: 699, poc: 0.8499 },
        { label: "Aug 2026", revenue: 50000, cumulativeRevenue: 576946, hours: 451, days: 56.375, cumulativeDays: 594.9125, eacDays: 699, poc: 0.9306 },
        { label: "Sep 2026", revenue: 40000, cumulativeRevenue: 616946, hours: 360.8, days: 45.1, cumulativeDays: 640.0125, eacDays: 699, poc: 0.9951 }
      ]
    },
    [
      { label: "Mar 2026", revenue: 18290, cumulativeRevenue: 18290, hours: 336, days: 42, cumulativeDays: 42, eacDays: 699, poc: 0.0295 },
      { label: "Apr 2026", revenue: 121210, cumulativeRevenue: 139500, hours: 1340, days: 167.5, cumulativeDays: 209.5, eacDays: 699, poc: 0.225 },
      { label: "May 2026", revenue: 287184, cumulativeRevenue: 426684, hours: 1728, days: 216, cumulativeDays: 425.5, eacDays: 699, poc: 0.6882 },
      { label: "Jun 2026", revenue: 50000, cumulativeRevenue: 476684, hours: 451, days: 56.375, cumulativeDays: 481.875, eacDays: 699, poc: 0.7688 },
      { label: "Jul 2026", revenue: 50262, cumulativeRevenue: 526946, hours: 453.3, days: 56.6625, cumulativeDays: 538.5375, eacDays: 699, poc: 0.8499 },
      { label: "Aug 2026", revenue: 50000, cumulativeRevenue: 576946, hours: 451, days: 56.375, cumulativeDays: 594.9125, eacDays: 699, poc: 0.9306 },
      { label: "Sep 2026", revenue: 40000, cumulativeRevenue: 616946, hours: 360.8, days: 45.1, cumulativeDays: 640.0125, eacDays: 699, poc: 0.9951 }
    ],
    "Master Sheet revenue forecast"
  )
];

const state = {
  projects: SAMPLE_PROJECTS,
  selectedIndex: 0,
  activeTab: "summary",
  selectedQuarterKey: null,
  theme: localStorage.getItem("dashboardTheme") || "day",
  sourceStatus: "Sample Data",
  sourceWarning: false
};

const elements = {
  themeButtons: document.querySelectorAll("[data-theme-choice]"),
  workbookInput: document.getElementById("workbookInput"),
  sampleButton: document.getElementById("sampleButton"),
  projectSelect: document.getElementById("projectSelect"),
  projectTitle: document.getElementById("projectTitle"),
  customerLogoCard: document.getElementById("customerLogoCard"),
  peopleStrip: document.getElementById("peopleStrip"),
  sourceStatus: document.getElementById("sourceStatus"),
  summaryGrid: document.getElementById("summaryGrid"),
  quickProjectFund: document.getElementById("quickProjectFund"),
  quickEacHours: document.getElementById("quickEacHours"),
  quickTargetHours: document.getElementById("quickTargetHours"),
  quickRevenueCard: document.getElementById("quickRevenueCard"),
  quickRevenueResult: document.getElementById("quickRevenueResult"),
  quickRevenueFormula: document.getElementById("quickRevenueFormula"),
  projectFacts: document.getElementById("projectFacts"),
  hoursGrid: document.getElementById("hoursGrid"),
  readoutGrid: document.getElementById("readoutGrid"),
  summaryChart: document.getElementById("summaryChart"),
  methodTitle: document.getElementById("methodTitle"),
  methodGrid: document.getElementById("methodGrid"),
  quarterRevenueRows: document.getElementById("quarterRevenueRows"),
  quarterRevenueTotals: document.getElementById("quarterRevenueTotals"),
  monthlyRevenueRows: document.getElementById("monthlyRevenueRows"),
  monthlyRevenueTotals: document.getElementById("monthlyRevenueTotals"),
  monthlyTableTitle: document.getElementById("monthlyTableTitle"),
  unbilledSummary: document.getElementById("unbilledSummary"),
  unbilledCharts: document.getElementById("unbilledCharts"),
  unbilledRows: document.getElementById("unbilledRows"),
  unbilledTotals: document.getElementById("unbilledTotals"),
  milestoneSummary: document.getElementById("milestoneSummary"),
  milestoneCharts: document.getElementById("milestoneCharts"),
  milestoneRows: document.getElementById("milestoneRows"),
  milestoneTotals: document.getElementById("milestoneTotals"),
  formulaGrid: document.getElementById("formulaGrid"),
  auditList: document.getElementById("auditList"),
  dropZone: document.getElementById("dropZone")
};

init();

function init() {
  applyTheme(state.theme);
  elements.themeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      applyTheme(button.dataset.themeChoice);
    });
  });

  elements.workbookInput.addEventListener("change", handleFileInput);
  elements.sampleButton.addEventListener("click", () => {
    state.projects = SAMPLE_PROJECTS;
    state.selectedIndex = 0;
    state.selectedQuarterKey = null;
    state.sourceStatus = "Sample Data";
    state.sourceWarning = false;
    render();
  });

  elements.projectSelect.addEventListener("change", (event) => {
    state.selectedIndex = Number(event.target.value);
    state.selectedQuarterKey = null;
    render();
  });

  [elements.quickProjectFund, elements.quickEacHours, elements.quickTargetHours].forEach((input) => {
    if (input) input.addEventListener("input", updateQuickRevenueCalculator);
  });

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTab = button.dataset.tab;
      renderTabs();
    });
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropZone.classList.add("dragging");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropZone.classList.remove("dragging");
    });
  });

  elements.dropZone.addEventListener("drop", async (event) => {
    const file = event.dataTransfer.files[0];
    if (file) {
      await loadWorkbookFile(file);
    }
  });

  render();
}

function applyTheme(theme) {
  state.theme = theme === "dark" ? "dark" : "day";
  document.body.dataset.theme = state.theme;
  localStorage.setItem("dashboardTheme", state.theme);
  elements.themeButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.themeChoice === state.theme);
  });
}

async function handleFileInput(event) {
  const file = event.target.files[0];
  if (!file) return;
  await loadWorkbookFile(file);
  event.target.value = "";
}

async function loadWorkbookFile(file) {
  try {
    setStatus("", false);
    const workbook = await parseXlsx(file);
    const projects = extractProjects(workbook, file.name);
    if (!projects.length) {
      throw new Error("No project rows or estimator rows were found.");
    }
    state.projects = projects;
    state.selectedIndex = 0;
    state.sourceStatus = "";
    state.sourceWarning = false;
    render();
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Workbook could not be loaded", true);
  }
}

function setStatus(text, warning) {
  state.sourceStatus = text;
  state.sourceWarning = warning;
  render();
}

function render() {
  const project = currentProject();
  renderTabs();
  renderProjectSelect();
  elements.projectTitle.textContent = "Project Monitoring Dashboard";
  if (elements.sourceStatus) {
    const visibleStatus = state.sourceStatus === "Sample Data" ? "" : state.sourceStatus;
    elements.sourceStatus.textContent = visibleStatus;
    elements.sourceStatus.hidden = !visibleStatus;
    elements.sourceStatus.classList.toggle("warning", state.sourceWarning);
  }
  renderPeople(project);
  renderCustomerLogo(project);

  if (!project) {
    elements.summaryGrid.innerHTML = emptyState("No project loaded");
    renderQuickRevenueCalculator(null);
    return;
  }

  renderSummary(project);
  renderQuickRevenueCalculator(project);
  renderFacts(project);
  renderHours(project);
  renderReadout(project);
  renderChart(project);
  renderMethod(project);
  renderTable(project);
  renderUnbilled(project);
  renderMilestones(project);
  renderFormulaTrace(project);
  renderAudit(project);
}

function renderTabs() {
  document.querySelectorAll(".tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === state.activeTab);
  });
  document.querySelectorAll(".panel").forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.panel === state.activeTab);
  });
}

function renderProjectSelect() {
  elements.projectSelect.innerHTML = state.projects
    .map((project, index) => `<option value="${index}">${escapeHtml(project.name || project.number || `Project ${index + 1}`)}</option>`)
    .join("");
  elements.projectSelect.value = String(state.selectedIndex);
}

function renderPeople(project) {
  if (!elements.peopleStrip) return;
  if (!project) {
    elements.peopleStrip.innerHTML = "";
    return;
  }

  const people = [
    personCard("Project Owner", project.projectOwner, project.projectOwnerImage),
    personCard("Portfolio Lead", project.portfolioLead, project.portfolioLeadImage)
  ].filter(Boolean);
  elements.peopleStrip.innerHTML = people.join("");
}

function renderCustomerLogo(project) {
  if (!elements.customerLogoCard) return;
  const logo = cleanText(project?.customerLogoImage);
  if (!logo) {
    elements.customerLogoCard.hidden = true;
    elements.customerLogoCard.innerHTML = "";
    return;
  }

  elements.customerLogoCard.hidden = false;
  elements.customerLogoCard.innerHTML = `
    <span>Customer</span>
    <img src="${escapeHtml(logo)}" alt="${escapeHtml(project?.name || "Customer logo")}">
  `;
}

function renderSummary(project) {
  const latestPoc = currentProjectPoc(project);
  const forecastVariance = fundingVariance(project.fundUsd, project.forecastRevenue || 0);
  const forecastTone = forecastVariance.key === "on" ? "accent" : "risk";
  const cards = [
    metric("Project Fund", money(project.fundUsd), "USD baseline", "metric-funding"),
    metric("Project Fund AED", money(project.fundAed, "AED"), "Converted at 3.67", "metric-funding-aed"),
    metric("Revenue To Date", money(project.revenueToDate), project.revenueToDateNote, "metric-revenue"),
    metric("Remaining", money(project.remainingRevenueToDate), remainingTone(project.remainingRevenueToDate), "metric-balance"),
    metric("POC", percent(latestPoc), project.currentQuarterLabel, "metric-progress"),
    metric("Forecast Total", money(project.forecastRevenue), forecastFundingNote(forecastVariance), `metric-forecast ${forecastTone}`)
  ];
  elements.summaryGrid.innerHTML = cards.join("");
}

function renderQuickRevenueCalculator(project) {
  if (!elements.quickProjectFund || !elements.quickEacHours || !elements.quickTargetHours) return;
  if (!project) {
    elements.quickProjectFund.value = "";
    elements.quickEacHours.value = "";
    elements.quickTargetHours.value = "";
    updateQuickRevenueCalculator();
    return;
  }

  elements.quickProjectFund.value = quickInputValue(project.fundUsd);
  elements.quickEacHours.value = quickInputValue(project.eacHours || project.estimateAtClosureHours || project.hoursWithContingency || project.hoursWithoutContingency);
  elements.quickTargetHours.value = quickInputValue(defaultTargetHours(project));
  updateQuickRevenueCalculator();
}

function updateQuickRevenueCalculator() {
  if (!elements.quickRevenueResult || !elements.quickRevenueFormula) return;
  const projectFund = toNullableNumber(elements.quickProjectFund?.value) ?? 0;
  const eacHours = toNullableNumber(elements.quickEacHours?.value) ?? 0;
  const targetHours = toNullableNumber(elements.quickTargetHours?.value) ?? 0;
  const valid = projectFund > 0 && eacHours > 0 && targetHours >= 0;
  const revenue = valid ? (targetHours / eacHours) * projectFund : null;
  const targetRatio = valid ? targetHours / eacHours : null;

  elements.quickRevenueResult.textContent = revenue === null ? "-" : money(revenue);
  elements.quickRevenueFormula.textContent = valid
    ? `${number(targetHours)} / ${number(eacHours)} = ${percent(targetRatio)}`
    : "Enter fund, EAC hours, and target hours";
  if (elements.quickRevenueCard) {
    elements.quickRevenueCard.classList.toggle("risk", valid && targetHours > eacHours);
    elements.quickRevenueCard.classList.toggle("accent", valid && targetHours <= eacHours);
  }
}

function defaultTargetHours(project) {
  const latest = latestRevenueToDateRow(project) || lastRow(project);
  if (Number.isFinite(latest?.cumulativeDays)) return latest.cumulativeDays * HOURS_PER_DAY;
  if (Number.isFinite(latest?.hours)) return latest.hours;
  if (Number.isFinite(project.actualToDateHours) && project.actualToDateHours > 0) return project.actualToDateHours;
  return null;
}

function quickInputValue(value) {
  if (!Number.isFinite(value)) return "";
  return String(Math.round(value * 100) / 100);
}

function renderFacts(project) {
  const items = [
    fact("Project Number", project.number || "-"),
    fact("Contract Number", project.contractNumber || "-"),
    fact("Account Number", project.accountNumber || "-"),
    projectExpenseFact(project),
    projectComparisonFact("Hours EAC vs Bid", project.estimateAtClosureHours, project.origBidHours, number, "hours"),
    projectComparisonFact("Revenue EAC vs Bid", project.estimateAtClosureRevenue, project.origBidRevenue, money, "revenue"),
    projectComparisonFact("Cost EAC vs Bid", project.estimateAtClosureCost, project.origBidCost, money, "cost", true),
    projectMarginFact(project)
  ];
  elements.projectFacts.innerHTML = items.filter(Boolean).join("");
}

function projectExpenseFact(project) {
  if (!Number.isFinite(project.totalExpenses) && !Number.isFinite(project.utilizedExpenses)) return "";
  const total = toNumber(project.totalExpenses);
  const utilized = toNumber(project.utilizedExpenses);
  const remaining = total - utilized;
  const used = total ? utilized / total : 0;
  const className = used > 1 ? "risk" : used > 0.85 ? "warn" : "accent";
  return fact("Expenses", money(total), `Utilized ${money(utilized)} (${percent(used)}), remaining ${money(remaining)}`, className);
}

function projectComparisonFact(label, current, baseline, formatter, noun, lowerIsBetter = false) {
  if (!Number.isFinite(current) && !Number.isFinite(baseline)) return "";
  const currentValue = toNumber(current);
  const baselineValue = toNumber(baseline);
  const delta = currentValue - baselineValue;
  const deltaPct = baselineValue ? delta / Math.abs(baselineValue) : null;
  const value = Number.isFinite(current) ? formatter(currentValue) : "-";
  const baselineText = Number.isFinite(baseline) ? formatter(baselineValue) : "-";
  const direction = Math.abs(delta) < 0.000001 ? "on plan" : delta > 0 ? "higher" : "lower";
  const note = `Orig bid ${baselineText}; ${Math.abs(delta) < 0.000001 ? "no change" : `${formatter(Math.abs(delta))} ${direction}${Number.isFinite(deltaPct) ? ` (${signedPercent(deltaPct)})` : ""}`}`;
  const className = comparisonTone(delta, lowerIsBetter);
  return fact(label, value, note, className);
}

function projectMarginFact(project) {
  if (!Number.isFinite(project.estimateAtClosureMargin) && !Number.isFinite(project.origBidMargin)) return "";
  const eac = toNumber(project.estimateAtClosureMargin);
  const bid = toNumber(project.origBidMargin);
  const delta = eac - bid;
  const points = Math.abs(delta) * 100;
  const note = `Orig bid ${percent(bid)}; ${Math.abs(delta) < 0.000001 ? "no change" : `${delta > 0 ? "+" : "-"}${number(points)} pts`}`;
  return fact("Margin EAC vs Bid", percent(eac), note, delta < -0.03 ? "risk" : delta < 0 ? "warn" : "accent");
}

function comparisonTone(delta, lowerIsBetter = false) {
  if (!Number.isFinite(delta) || Math.abs(delta) < 0.000001) return "accent";
  if (lowerIsBetter) return delta > 0 ? "risk" : "accent";
  return delta < 0 ? "warn" : "accent";
}

function renderHours(project) {
  const cards = [
    hoursCard("Actual To Date", number(project.actualToDateHours), "Timesheet total", "accent"),
    hoursCard("With Contingency", number(project.hoursWithContingency), `${number(project.eacDays)} days`, "neutral"),
    hoursCard("Without Contingency", number(project.hoursWithoutContingency), `${number(project.hoursWithoutContingency / HOURS_PER_DAY)} days`, "neutral"),
    hoursComparisonCard("Actual vs With Contingency", project.actualToDateHours, project.hoursWithContingency),
    hoursComparisonCard("Actual vs Without Contingency", project.actualToDateHours, project.hoursWithoutContingency)
  ];
  elements.hoursGrid.innerHTML = cards.join("");
}

function hoursComparisonCard(label, actualHours, baselineHours) {
  if (!Number.isFinite(actualHours) || !Number.isFinite(baselineHours) || baselineHours <= 0) {
    return hoursCard(label, "-", "Baseline not loaded", "warn");
  }

  const variance = baselineHours - actualHours;
  const used = actualHours / baselineHours;
  const value = Math.abs(variance) < 0.05
    ? "On baseline"
    : variance < 0
      ? `${number(Math.abs(variance))} over`
      : `${number(variance)} remaining`;
  const className = variance < 0 ? "risk" : used >= 0.9 ? "warn" : "accent";
  return hoursCard(label, value, `${percent(used)} used of ${number(baselineHours)}`, className);
}

function renderReadout(project) {
  const latestPoc = currentProjectPoc(project);
  const burn = project.fundUsd ? project.revenueToDate / project.fundUsd : 0;
  const statusClass = latestPoc > 1 ? "risk" : project.remainingRevenueToDate < 0 ? "risk" : burn > 0.85 ? "warn" : "accent";
  const cards = [
    readout("Commercial Status", project.remainingRevenueToDate < 0 ? "Over funding" : "Within funding", statusClass),
    readout("Revenue Burn", percent(burn), burn > 0.85 ? "warn" : "accent"),
    readout("To-Date Cutoff", project.currentQuarterLabel, "accent"),
    readout("Forecast Total", money(project.forecastRevenue), "accent")
  ];
  elements.readoutGrid.innerHTML = cards.join("");
}

function renderChart(project) {
  const chartData = chartDataForProject(project);
  const quarterRows = chartData.quarterRows;
  if (!quarterRows.length) {
    elements.summaryChart.innerHTML = emptyState("No chart data available");
    return;
  }

  const selectedQuarter = selectedQuarterForChart(project, chartData);
  const monthlyRows = chartData.monthlyGroups.get(selectedQuarter.key) || [selectedQuarter];
  const latestQuarter = quarterRows[quarterRows.length - 1] || {};
  const chartVariance = fundingVariance(project.fundUsd, latestQuarter.cumulativeRevenue || 0);
  elements.summaryChart.innerHTML = `
    <div class="chart-stack">
      <div class="chart-block">
        <div class="chart-heading">
          <div>
            <span>Quarterly view</span>
            <strong>Click a quarter to drill down</strong>
          </div>
          ${chartVarianceBadge(chartVariance)}
        </div>
        ${chartSvg(quarterRows, { activeKey: selectedQuarter.key, interactive: true, title: "Quarterly revenue forecast", fundingVariance: chartVariance })}
      </div>
      <div class="chart-block">
        <div class="chart-heading">
          <div>
            <span>Monthly view</span>
            <strong>${escapeHtml(selectedQuarter.displayLabel || selectedQuarter.label)}</strong>
          </div>
        </div>
        ${chartSvg(monthlyRows, { title: `${selectedQuarter.displayLabel || selectedQuarter.label} monthly revenue` })}
      </div>
    </div>
  `;

  elements.summaryChart.querySelectorAll("[data-quarter-key]").forEach((element) => {
    element.addEventListener("click", () => {
      state.selectedQuarterKey = element.getAttribute("data-quarter-key");
      renderChart(project);
      renderTable(project);
    });
  });
}

function renderMethod(project) {
  elements.methodTitle.textContent = project.modeLabel;
  const latestPoc = currentProjectPoc(project);
  const items = [
    fact("Primary Basis", project.mode === "hours" ? "Period hours" : project.mode === "estimator" ? "Workbook estimator rows" : "Imported revenue"),
    fact("Funding", money(project.fundUsd)),
    fact("Latest POC", percent(latestPoc)),
    fact("Revenue Formula", project.mode === "hours" ? "(POC * Funding) - Prior Cumulative" : "Imported period revenue")
  ];
  elements.methodGrid.innerHTML = items.join("");
}

function renderTable(project) {
  const chartData = chartDataForProject(project);
  const quarterRows = chartData.quarterRows;
  if (!quarterRows.length) {
    elements.quarterRevenueRows.innerHTML = `<tr><td colspan="8">${emptyState("No quarterly revenue detected")}</td></tr>`;
    elements.quarterRevenueTotals.innerHTML = "";
    elements.monthlyRevenueRows.innerHTML = `<tr><td colspan="9">${emptyState("No monthly revenue detected")}</td></tr>`;
    elements.monthlyRevenueTotals.innerHTML = "";
    return;
  }

  const selectedQuarter = selectedQuarterForChart(project, chartData);
  const monthlyRows = chartData.monthlyGroups.get(selectedQuarter.key) || [selectedQuarter];
  elements.monthlyTableTitle.textContent = selectedQuarter.displayLabel || selectedQuarter.label;
  elements.quarterRevenueRows.innerHTML = quarterRows
    .map((row, index) => {
      const trend = trendForRows(quarterRows, index);
      const active = selectedQuarter.key === row.key;
      return `
        <tr class="${active ? "selected-row" : ""}" data-quarter-row="${escapeHtml(row.key || row.label)}">
          <td>${escapeHtml(row.displayLabel || row.label)}</td>
          <td>${statusBadge(periodStatus(row))}</td>
          <td>${money(row.revenue)}</td>
          <td>${trendBadge(trend)}</td>
          <td class="${trend.className}">${trendPercent(trend)}</td>
          <td>${money(row.cumulativeRevenue)}</td>
          <td>${percent(pocForRow(row, project))}</td>
          <td>${money(project.fundUsd - (row.cumulativeRevenue || 0))}</td>
        </tr>
      `;
    })
    .join("");
  elements.quarterRevenueTotals.innerHTML = quarterTotalsRow(quarterRows, project);

  elements.monthlyRevenueRows.innerHTML = monthlyRows
    .map((row, index) => {
      const trend = trendForRows(monthlyRows, index);
      return `
      <tr>
        <td>${escapeHtml(row.label)}</td>
        <td>${optionalNumber(row.hours)}</td>
        <td>${optionalNumber(row.days)}</td>
        <td>${percent(pocForRow(row, project))}</td>
        <td>${money(row.revenue)}</td>
        <td>${trendBadge(trend)}</td>
        <td class="${trend.className}">${trendPercent(trend)}</td>
        <td>${money(row.cumulativeRevenue)}</td>
        <td>${money(project.fundUsd - (row.cumulativeRevenue || 0))}${isFuturePeriod(row) ? " - future" : ""}</td>
      </tr>
    `;
    })
    .join("");
  elements.monthlyRevenueTotals.innerHTML = monthlyTotalsRow(monthlyRows, project);

  elements.quarterRevenueRows.querySelectorAll("[data-quarter-row]").forEach((row) => {
    row.addEventListener("click", () => {
      state.selectedQuarterKey = row.getAttribute("data-quarter-row");
      renderChart(project);
      renderTable(project);
    });
  });
}

function quarterTotalsRow(rows, project) {
  const latest = rows[rows.length - 1] || {};
  const variance = fundingVariance(project.fundUsd, latest.cumulativeRevenue || 0);
  return `
    <tr class="total-row">
      <td>Total</td>
      <td>${rows.length} quarters</td>
      <td>${money(sum(rows.map((row) => row.revenue)))}</td>
      <td></td>
      <td></td>
      <td>${money(latest.cumulativeRevenue || 0)}</td>
      <td>${percent(pocForRow(latest, project))}</td>
      <td>${varianceCell(variance)}</td>
    </tr>
  `;
}

function monthlyTotalsRow(rows, project) {
  const latest = rows[rows.length - 1] || {};
  return `
    <tr class="total-row">
      <td>Total</td>
      <td>${optionalNumber(sum(rows.map((row) => row.hours)))}</td>
      <td>${optionalNumber(sum(rows.map((row) => row.days)))}</td>
      <td>${percent(pocForRow(latest, project))}</td>
      <td>${money(sum(rows.map((row) => row.revenue)))}</td>
      <td></td>
      <td></td>
      <td>${money(latest.cumulativeRevenue || 0)}</td>
      <td>${money(project.fundUsd - (latest.cumulativeRevenue || 0))}${isFuturePeriod(latest) ? " <span class=\"future-note\">future</span>" : ""}</td>
    </tr>
  `;
}

function renderUnbilled(project) {
  if (!elements.unbilledSummary || !elements.unbilledCharts || !elements.unbilledRows || !elements.unbilledTotals) return;
  const rows = unbilledRowsForProject(project);
  if (!rows.length) {
    elements.unbilledSummary.innerHTML = emptyState("No unbilled data available");
    elements.unbilledCharts.innerHTML = emptyState("No unbilled charts available");
    elements.unbilledRows.innerHTML = `<tr><td colspan="8">${emptyState("No unbilled data available")}</td></tr>`;
    elements.unbilledTotals.innerHTML = "";
    return;
  }

  const summary = unbilledSummaryForRows(rows);
  const rate = projectCurrencyRate(project);
  const baseline = unbilledBaseline(project, rows);
  elements.unbilledSummary.innerHTML = [
    billedVsUnbilledCard(summary, rate),
    unbilledCard("Latest Forecast Unbilled", summary.latestUnbilled, escapeHtml(summary.latestLabel), unbilledTone(summary.latestUnbilled, summary.latestAged), rate),
    unbilledCard(">180 Days Risk", summary.maxAged, escapeHtml(summary.maxAgedLabel), summary.maxAged > 0 ? "risk" : "accent", rate),
    unbilledCard("Invoice Outlook", summary.totalInvoicing, `Revenue ${dualMoneyMini(summary.totalRevenue, rate)}`, summary.totalInvoicing < summary.totalRevenue ? "warn" : "accent", rate)
  ].join("");

  elements.unbilledCharts.innerHTML = renderUnbilledCharts(rows, project, rate);
  elements.unbilledRows.innerHTML = rows
    .map((row) => {
      const status = unbilledStatus(row);
      return `
        <tr class="unbilled-row unbilled-${status.key}">
          <td>${escapeHtml(row.displayLabel || row.label || "-")}</td>
          <td>${moneyWithPercent(row.revenue, baseline)}</td>
          <td>${moneyWithPercent(row.invoicing, baseline)}</td>
          <td>${moneyWithPercent(row.cumulativeRevenue, baseline)}</td>
          <td>${moneyWithPercent(row.cumulativeInvoicing, baseline)}</td>
          <td>${moneyWithPercent(row.unbilledRevenue, baseline)}</td>
          <td>${moneyWithPercent(row.agedUnbilled, baseline)}</td>
          <td>${unbilledStatusBadge(status)}</td>
        </tr>
      `;
    })
    .join("");

  const latest = rows[rows.length - 1];
  elements.unbilledTotals.innerHTML = `
    <tr class="total-row">
      <td>Total</td>
      <td>${moneyWithPercent(sum(rows.map((row) => row.revenue)), baseline)}</td>
      <td>${moneyWithPercent(sum(rows.map((row) => row.invoicing)), baseline)}</td>
      <td>${moneyWithPercent(latest.cumulativeRevenue, baseline)}</td>
      <td>${moneyWithPercent(latest.cumulativeInvoicing, baseline)}</td>
      <td>${moneyWithPercent(latest.unbilledRevenue, baseline)}</td>
      <td>${moneyWithPercent(latest.agedUnbilled, baseline)}</td>
      <td>${escapeHtml(unbilledStatus(latest).label)}</td>
    </tr>
  `;
}

function unbilledRowsForProject(project) {
  if (project.unbilledRows?.length) return prepareUnbilledRows(project.unbilledRows);
  return calculateUnbilledRows(project);
}

function prepareUnbilledRows(rows) {
  let previousEndDate = null;
  let cumulativeRevenue = 0;
  let cumulativeInvoicing = 0;
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  const prepared = [];
  rows.forEach((row) => {
    const periodEndDate = row.periodEndDate || inferPeriodEndDate(row.label, previousEndDate, currentQuarter);
    if (periodEndDate) previousEndDate = periodEndDate;
    const quarter = periodEndDate ? fiscalQuarterForDate(periodEndDate) : quarterFromLabel(row.label, currentQuarter);
    const revenue = toNumber(row.revenue);
    const invoicing = toNumber(row.invoicing);
    cumulativeRevenue = toNullableNumber(row.cumulativeRevenue) ?? cumulativeRevenue + revenue;
    cumulativeInvoicing = toNullableNumber(row.cumulativeInvoicing) ?? cumulativeInvoicing + invoicing;
    const unbilledRevenue = toNullableNumber(row.unbilledRevenue) ?? cumulativeRevenue - cumulativeInvoicing;
    const agedBase = prepared.length >= 2 ? prepared[prepared.length - 2].cumulativeRevenue : 0;
    const agedUnbilled = toNullableNumber(row.agedUnbilled) ?? (prepared.length >= 2 ? agedBase - cumulativeInvoicing : 0);
    prepared.push({
      ...row,
      key: row.key || quarter.key || cleanText(row.label),
      displayLabel: row.displayLabel || quarter.displayLabel || cleanText(row.label),
      periodEndDate,
      revenue,
      invoicing,
      cumulativeRevenue,
      cumulativeInvoicing,
      unbilledRevenue,
      agedUnbilled
    });
  });
  return prepared;
}

function calculateUnbilledRows(project) {
  const revenueRows = chartDataForProject(project).quarterRows;
  const revenueByQuarter = new Map();
  revenueRows.forEach((row) => {
    const quarter = quarterInfoForRow(row);
    if (!quarter) return;
    const current = revenueByQuarter.get(quarter.key) || 0;
    revenueByQuarter.set(quarter.key, current + toNumber(row.revenue));
  });

  const invoicingByQuarter = invoiceScheduleByQuarter(project);
  const quarterKeys = new Set([...revenueByQuarter.keys(), ...invoicingByQuarter.keys()]);
  if (!quarterKeys.size) return [];

  const sorted = Array.from(quarterKeys)
    .map(parseFiscalQuarterKey)
    .filter(Boolean)
    .sort((a, b) => fiscalQuarterOrdinal(a) - fiscalQuarterOrdinal(b));
  if (!sorted.length) return [];

  const start = addFiscalQuarters(sorted[0], -2);
  const end = sorted[sorted.length - 1];
  const quarters = fiscalQuarterSequence(start, end);
  const rows = [];
  let cumulativeRevenue = 0;
  let cumulativeInvoicing = 0;

  quarters.forEach((quarter) => {
    const revenue = revenueByQuarter.get(quarter.key) || 0;
    const invoicing = invoicingByQuarter.get(quarter.key) || 0;
    cumulativeRevenue += revenue;
    cumulativeInvoicing += invoicing;
    rows.push({
      key: quarter.key,
      label: quarter.displayLabel,
      displayLabel: quarter.displayLabel,
      periodEndDate: fiscalQuarterEndDate(quarter.quarter, quarter.fiscalYear),
      revenue,
      invoicing,
      cumulativeRevenue,
      cumulativeInvoicing,
      unbilledRevenue: cumulativeRevenue - cumulativeInvoicing,
      agedUnbilled: 0
    });
  });

  const firstRevenueIndex = rows.findIndex((row) => Math.abs(row.revenue) > 0.000001);
  rows.forEach((row, index) => {
    if (firstRevenueIndex === -1 || index < firstRevenueIndex + 2 || index < 2) {
      row.agedUnbilled = 0;
      return;
    }
    row.agedUnbilled = rows[index - 2].cumulativeRevenue - row.cumulativeInvoicing;
  });

  return rows;
}

function invoiceScheduleByQuarter(project) {
  const schedule = new Map();
  const milestones = Array.isArray(project.milestones) ? project.milestones : [];
  if (!milestones.length) return schedule;

  const currentQuarter = currentFiscalQuarterInfo(new Date());
  const nextQuarter = addFiscalQuarters(currentQuarter, 1);
  const rate = project.fundAed && project.fundUsd ? project.fundAed / project.fundUsd : EXCHANGE_RATE;

  milestones.forEach((item) => {
    const feeUsd = milestoneFeeUsd(item, project, rate);
    if (!Number.isFinite(feeUsd) || Math.abs(feeUsd) < 0.000001) return;

    let quarter = currentQuarter;
    if (!isPaid(item.invoiceStatus)) {
      const dueQuarter = item.expectedDate ? fiscalQuarterForDate(item.expectedDate) : currentQuarter;
      const dueEnd = dueQuarter ? fiscalQuarterEndDate(dueQuarter.quarter, dueQuarter.fiscalYear) : null;
      quarter = dueEnd && dueEnd <= currentQuarter.endDate ? nextQuarter : dueQuarter;
    }
    const key = quarter.key;
    schedule.set(key, (schedule.get(key) || 0) + feeUsd);
  });

  return schedule;
}

function milestoneFeeUsd(item, project, rate) {
  const feeAed = toNullableNumber(item.feeAed);
  if (feeAed !== null && rate > 0) return feeAed / rate;
  const percentage = toNullableNumber(item.percentage);
  if (percentage !== null && project.fundUsd) return percentage * project.fundUsd;
  return null;
}

function quarterInfoForRow(row) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  const periodEndDate = row.periodEndDate || inferPeriodEndDate(row.label, null, currentQuarter);
  const quarter = periodEndDate ? fiscalQuarterForDate(periodEndDate) : quarterFromLabel(row.label, currentQuarter);
  return quarter?.key ? quarter : null;
}

function unbilledSummaryForRows(rows) {
  const currentQuarterKey = currentFiscalQuarterInfo(new Date()).key;
  const current = rows.find((row) => row.key === currentQuarterKey) || rows.find((row) => row.periodEndDate && row.periodEndDate <= currentFiscalQuarterInfo(new Date()).endDate) || rows[0];
  const latest = rows[rows.length - 1];
  const maxAgedRow = rows.reduce((best, row) => row.agedUnbilled > best.agedUnbilled ? row : best, rows[0]);
  return {
    currentUnbilled: current.unbilledRevenue,
    currentAged: current.agedUnbilled,
    currentLabel: current.displayLabel || current.label,
    latestUnbilled: latest.unbilledRevenue,
    latestAged: latest.agedUnbilled,
    latestLabel: latest.displayLabel || latest.label,
    maxAged: Math.max(0, maxAgedRow.agedUnbilled),
    maxAgedLabel: maxAgedRow.agedUnbilled > 0 ? `${maxAgedRow.displayLabel || maxAgedRow.label} exposure` : "No positive aged exposure",
    totalRevenue: sum(rows.map((row) => row.revenue)),
    totalInvoicing: sum(rows.map((row) => row.invoicing))
  };
}

function renderUnbilledCharts(rows, project, rate) {
  const baseline = project.fundUsd || Math.max(...rows.map((row) => Math.max(row.cumulativeRevenue, row.cumulativeInvoicing)), 1);
  return `
    <div class="unbilled-chart-layout">
      <div class="unbilled-chart-card">
        <div class="chart-heading">
          <div>
            <span>Quarterly Comparison</span>
            <strong>Revenue vs invoicing</strong>
          </div>
        </div>
        ${unbilledRevenueInvoiceSvg(rows, rate, baseline)}
      </div>
      <div class="unbilled-chart-card">
        <div class="chart-heading">
          <div>
            <span>Aging View</span>
            <strong>Unbilled and &gt;180 days</strong>
          </div>
        </div>
        ${unbilledAgingSvg(rows, rate, baseline)}
      </div>
    </div>
  `;
}

function unbilledRevenueInvoiceSvg(rows, rate, baseline) {
  const width = 920;
  const height = 390;
  const pad = { top: 88, right: 28, bottom: 72, left: 82 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const maxValue = Math.max(...rows.map((row) => Math.max(0, row.revenue, row.invoicing)), 1);
  const groupGap = 14;
  const groupW = Math.max(38, (innerW - groupGap * Math.max(0, rows.length - 1)) / rows.length);
  const barW = Math.max(10, (groupW - 8) / 2);

  const bars = rows.map((row, index) => {
    const x = pad.left + index * (groupW + groupGap);
    const revenueH = (Math.max(0, row.revenue) / maxValue) * innerH;
    const invoiceH = (Math.max(0, row.invoicing) / maxValue) * innerH;
    const revenueY = pad.top + innerH - revenueH;
    const invoiceY = pad.top + innerH - invoiceH;
    const revenueLabelY = Math.max(26, revenueY - 34);
    const invoiceLabelY = Math.max(26, invoiceY - 34);
    return `
      <g>
        <rect x="${x}" y="${revenueY}" width="${barW}" height="${revenueH}" rx="4" fill="var(--chart-actual-bar)"></rect>
        <rect x="${x + barW + 8}" y="${invoiceY}" width="${barW}" height="${invoiceH}" rx="4" fill="var(--amber)"></rect>
        ${Math.abs(row.revenue) > 0.000001 ? unbilledChartValueLabel(x + barW / 2, revenueLabelY, row.revenue, rate, baseline, "var(--chart-actual-bar)") : ""}
        ${Math.abs(row.invoicing) > 0.000001 ? unbilledChartValueLabel(x + barW + 8 + barW / 2, invoiceLabelY, row.invoicing, rate, baseline, "var(--amber)") : ""}
        <text x="${x + groupW / 2}" y="${height - 40}" text-anchor="middle" font-size="10" font-weight="800" fill="var(--muted)">${escapeHtml(shortQuarterLabel(row.displayLabel || row.label))}</text>
      </g>
    `;
  }).join("");

  return `
    <svg class="unbilled-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Revenue and invoicing comparison chart">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="var(--surface)"></rect>
      ${currencyGridLines(maxValue, width, height, pad, innerH, rate)}
      <line x1="${pad.left}" y1="${pad.top + innerH}" x2="${width - pad.right}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      ${bars}
      ${unbilledLegend([["Revenue", "var(--chart-actual-bar)"], ["Invoicing", "var(--amber)"]], pad.left, height - 16)}
    </svg>
  `;
}

function unbilledAgingSvg(rows, rate, baseline) {
  const width = 920;
  const height = 390;
  const pad = { top: 76, right: 28, bottom: 72, left: 82 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const maxAbs = Math.max(...rows.flatMap((row) => [Math.abs(row.unbilledRevenue), Math.abs(row.agedUnbilled)]), 1);
  const groupGap = 14;
  const groupW = Math.max(38, (innerW - groupGap * Math.max(0, rows.length - 1)) / rows.length);
  const barW = Math.max(10, (groupW - 8) / 2);
  const zeroY = pad.top + innerH / 2;
  const halfH = innerH / 2;

  const bars = rows.map((row, index) => {
    const x = pad.left + index * (groupW + groupGap);
    const values = [
      { value: row.unbilledRevenue, color: "var(--berry)", x },
      { value: row.agedUnbilled, color: "var(--teal)", x: x + barW + 8 }
    ];
    return `
      <g>
        ${values.map((item) => {
          const h = (Math.abs(item.value) / maxAbs) * halfH;
          const y = item.value >= 0 ? zeroY - h : zeroY;
          const labelY = item.value >= 0 ? Math.max(26, y - 34) : Math.min(height - 108, y + h + 10);
          return `
            <rect x="${item.x}" y="${y}" width="${barW}" height="${h}" rx="4" fill="${item.color}"></rect>
            ${Math.abs(item.value) > 0.000001 ? unbilledChartValueLabel(item.x + barW / 2, labelY, item.value, rate, baseline, item.color) : ""}
          `;
        }).join("")}
        <text x="${x + groupW / 2}" y="${height - 40}" text-anchor="middle" font-size="10" font-weight="800" fill="var(--muted)">${escapeHtml(shortQuarterLabel(row.displayLabel || row.label))}</text>
      </g>
    `;
  }).join("");

  return `
    <svg class="unbilled-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Unbilled aging chart">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="var(--surface)"></rect>
      <line x1="${pad.left}" y1="${zeroY}" x2="${width - pad.right}" y2="${zeroY}" stroke="var(--line)" stroke-width="2"></line>
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      ${axisDualLabel(pad.left - 10, pad.top + 2, maxAbs, rate, "end")}
      <text x="${pad.left - 10}" y="${zeroY + 4}" text-anchor="end" font-size="10" font-weight="700" fill="var(--muted)">0</text>
      ${axisDualLabel(pad.left - 10, pad.top + innerH - 8, -maxAbs, rate, "end")}
      ${bars}
      ${unbilledLegend([["Unbilled", "var(--berry)"], [">180 days", "var(--teal)"]], pad.left, height - 16)}
    </svg>
  `;
}

function currencyGridLines(maxValue, width, height, pad, innerH, rate) {
  return [0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = pad.top + innerH - innerH * ratio;
    return `
      <line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" stroke="var(--line)" stroke-width="1"></line>
      ${axisDualLabel(pad.left - 10, y - 6, maxValue * ratio, rate, "end")}
    `;
  }).join("");
}

function unbilledLegend(items, startX, y) {
  return `
    <g class="unbilled-chart-legend">
      ${items.map(([label, color], index) => {
        const x = startX + index * 120;
        return `
          <rect x="${x}" y="${y - 10}" width="12" height="12" rx="3" fill="${color}"></rect>
          <text x="${x + 18}" y="${y}" font-size="11" font-weight="800" fill="var(--muted)">${escapeHtml(label)}</text>
        `;
      }).join("")}
    </g>
  `;
}

function unbilledChartValueLabel(x, y, value, rate, baseline, color) {
  return `
    <text class="unbilled-chart-value" x="${x}" y="${y}" text-anchor="middle" font-size="9" font-weight="850" fill="${color}">
      <tspan x="${x}">${escapeHtml(compactDollar(value))}</tspan>
      <tspan x="${x}" dy="11">${escapeHtml(signedPercentOf(value, baseline))}</tspan>
    </text>
  `;
}

function axisDualLabel(x, y, value, rate, anchor = "end") {
  return `
    <text x="${x}" y="${y}" text-anchor="${anchor}" font-size="10" font-weight="750" fill="var(--muted)">
      <tspan x="${x}">${escapeHtml(compactDollar(value))}</tspan>
    </text>
  `;
}

function billedVsUnbilledCard(summary, rate) {
  const billed = summary.totalInvoicing;
  const unbilled = summary.latestUnbilled;
  const denominator = Math.abs(summary.totalRevenue) > 0.000001
    ? Math.abs(summary.totalRevenue)
    : Math.max(Math.abs(billed) + Math.abs(unbilled), 1);
  const billedShare = billed / denominator;
  const unbilledShare = unbilled / denominator;
  return `
    <div class="unbilled-card billed-vs-card ${unbilledTone(unbilled, summary.latestAged)}">
      <span>Total Billed vs Unbilled</span>
      <div class="billed-vs-grid">
        <div>
          <em>Billed</em>
          <strong>${escapeHtml(money(billed))}</strong>
          <small class="aed-money">${escapeHtml(money(billed * rate, "AED"))}</small>
        </div>
        <div>
          <em>Unbilled</em>
          <strong>${escapeHtml(money(unbilled))}</strong>
          <small class="aed-money">${escapeHtml(money(unbilled * rate, "AED"))}</small>
        </div>
      </div>
      <small class="billed-vs-percent">Billed ${escapeHtml(percent(billedShare))} / Unbilled ${escapeHtml(percent(unbilledShare))}</small>
    </div>
  `;
}

function unbilledCard(label, value, note, className, rate) {
  return `
    <div class="unbilled-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(money(value))}</strong>
      <small class="aed-money">${escapeHtml(money(value * rate, "AED"))}</small>
      <small>${note}</small>
    </div>
  `;
}

function dualMoney(value, rate) {
  return `
    <span class="dual-money">
      <span class="dual-money-main">${escapeHtml(money(value))}</span>
      <span class="dual-money-aed">${escapeHtml(money(value * rate, "AED"))}</span>
    </span>
  `;
}

function moneyWithPercent(value, baseline) {
  return `
    <span class="dual-money">
      <span class="dual-money-main">${escapeHtml(money(value))}</span>
      <span class="dual-money-percent">${escapeHtml(signedPercentOf(value, baseline))}</span>
    </span>
  `;
}

function dualMoneyMini(value, rate) {
  return `
    <span class="dual-money-mini">
      <span>${escapeHtml(money(value))}</span>
      <span>${escapeHtml(money(value * rate, "AED"))}</span>
    </span>
  `;
}

function projectCurrencyRate(project) {
  const rate = project?.fundUsd && project?.fundAed ? project.fundAed / project.fundUsd : EXCHANGE_RATE;
  return Number.isFinite(rate) && rate > 0 ? rate : EXCHANGE_RATE;
}

function unbilledBaseline(project, rows) {
  const maxCumulative = Math.max(...rows.map((row) => Math.max(row.cumulativeRevenue || 0, row.cumulativeInvoicing || 0)), 0);
  const baseline = project?.fundUsd || maxCumulative || 1;
  return Number.isFinite(baseline) && baseline !== 0 ? baseline : 1;
}

function compactDollar(value) {
  if (!Number.isFinite(value)) return "$0";
  const absolute = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (absolute >= 1000000) return `${sign}$${(absolute / 1000000).toFixed(1)}M`;
  if (absolute >= 1000) return `${sign}$${Math.round(absolute / 1000)}K`;
  return `${sign}$${Math.round(absolute)}`;
}

function signedPercentOf(value, baseline) {
  if (!Number.isFinite(value) || !Number.isFinite(baseline) || baseline === 0) return "-";
  return percent(value / Math.abs(baseline));
}

function unbilledTone(unbilled, aged) {
  if (aged > 0) return "risk";
  if (unbilled > 0) return "warn";
  return "accent";
}

function unbilledStatus(row) {
  if (row.agedUnbilled > 0) return { key: "risk", label: ">180 days exposure" };
  if (row.unbilledRevenue > 0) return { key: "warn", label: "Unbilled open" };
  if (row.unbilledRevenue < 0) return { key: "neutral", label: "Invoice ahead" };
  return { key: "accent", label: "Balanced" };
}

function unbilledStatusBadge(status) {
  return `<span class="unbilled-badge unbilled-badge-${status.key}">${escapeHtml(status.label)}</span>`;
}

function shortQuarterLabel(label) {
  const match = cleanText(label).match(/\bQ([1-4])\s*(?:FY)?(\d{2,4})\b/i);
  if (!match) return shortLabel(label);
  return `Q${match[1]}-${String(normalizeFiscalYear(match[2])).slice(-2)}`;
}

function renderMilestones(project) {
  if (!elements.milestoneSummary || !elements.milestoneRows || !elements.milestoneTotals) return;
  const milestones = Array.isArray(project.milestones) ? project.milestones : [];
  if (!milestones.length) {
    elements.milestoneSummary.innerHTML = emptyState("No milestones detected for this project");
    if (elements.milestoneCharts) elements.milestoneCharts.innerHTML = emptyState("No milestone charts available");
    elements.milestoneRows.innerHTML = `<tr><td colspan="8">${emptyState("No milestones detected")}</td></tr>`;
    elements.milestoneTotals.innerHTML = "";
    return;
  }

  const summary = milestoneSummary(milestones);
  elements.milestoneSummary.innerHTML = [
    milestoneCard("Total Milestones", String(summary.total), `${summary.open} open`, "accent"),
    milestoneCard("Certificate Progress", percent(summary.submittedWeight), `${summary.submitted} of ${summary.total} submitted`, summary.submittedWeight >= 0.8 ? "accent" : "warn"),
    milestoneCard("Invoice Paid", money(summary.paidFee, "AED"), `${percent(summary.paidFeeRatio)} of milestone value`, summary.paidFeeRatio >= 0.8 ? "accent" : "warn"),
    milestoneCard("Risk Watch", `${summary.overdue} overdue`, `${summary.dueSoon} due in 30 days`, summary.overdue ? "risk" : summary.dueSoon ? "warn" : "accent")
  ].join("");
  if (elements.milestoneCharts) {
    elements.milestoneCharts.innerHTML = renderMilestoneCharts(milestones, summary);
  }

  elements.milestoneRows.innerHTML = milestones
    .map((item) => {
      const health = milestoneHealth(item);
      return `
        <tr class="milestone-row milestone-${health.key}">
          <td>${escapeHtml(item.sno || "-")}</td>
          <td>${escapeHtml(item.milestone || "-")}</td>
          <td>${percent(item.percentage)}</td>
          <td>${money(item.feeAed, "AED")}</td>
          <td>${formatDate(item.expectedDate)}</td>
          <td>${milestoneStatusBadge(item.certificateStatus, health.key)}</td>
          <td>${invoiceStatusBadge(item.invoiceStatus)}</td>
          <td>${escapeHtml(item.notes || "-")}</td>
        </tr>
      `;
    })
    .join("");
  elements.milestoneTotals.innerHTML = `
    <tr class="total-row">
      <td>Total</td>
      <td>${summary.total} milestones</td>
      <td>${percent(summary.totalWeight)}</td>
      <td>${money(summary.totalFee, "AED")}</td>
      <td></td>
      <td>${summary.submitted} submitted</td>
      <td>${money(summary.paidFee, "AED")} paid</td>
      <td>${summary.overdue} overdue</td>
    </tr>
  `;
}

function milestoneSummary(milestones) {
  const total = milestones.length;
  const submittedItems = milestones.filter((item) => isSubmitted(item.certificateStatus));
  const paidItems = milestones.filter((item) => isPaid(item.invoiceStatus));
  const totalWeight = sum(milestones.map((item) => item.percentage));
  const submittedWeight = sum(submittedItems.map((item) => item.percentage));
  const totalFee = sum(milestones.map((item) => item.feeAed));
  const paidFee = sum(paidItems.map((item) => item.feeAed));
  const healthCounts = milestones.reduce((counts, item) => {
    const key = milestoneHealth(item).key;
    counts[key] = (counts[key] || 0) + 1;
    return counts;
  }, { complete: 0, due: 0, overdue: 0, pending: 0 });
  const overdue = healthCounts.overdue || 0;
  const dueSoon = healthCounts.due || 0;
  return {
    total,
    open: total - submittedItems.length,
    submitted: submittedItems.length,
    totalWeight,
    submittedWeight: totalWeight ? submittedWeight / totalWeight : 0,
    totalFee,
    paidFee,
    paidFeeRatio: totalFee ? paidFee / totalFee : 0,
    overdue,
    dueSoon,
    healthCounts
  };
}

function renderMilestoneCharts(milestones, summary) {
  return `
    <div class="milestone-chart-layout">
      <div class="milestone-chart-card wide">
        <div class="chart-heading">
          <div>
            <span>Comparison</span>
            <strong>Certificates, invoices, and schedule health</strong>
          </div>
        </div>
        ${milestoneProgressComparisonSvg(summary)}
      </div>
      <div class="milestone-chart-card wide">
        <div class="chart-heading">
          <div>
            <span>Milestone Values</span>
            <strong>Fee distribution by deliverable</strong>
          </div>
        </div>
        ${milestoneValueComparisonSvg(milestones)}
      </div>
    </div>
  `;
}

function milestoneProgressComparisonSvg(summary) {
  const width = 860;
  const height = 370;
  const groups = milestonePieGroups(summary);
  const centers = [144, 430, 716];
  const pies = groups.map((group, index) => milestonePieGroupSvg(group, centers[index], 138, 72)).join("");

  return `
    <svg class="milestone-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Milestones progress comparison chart">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="var(--surface)"></rect>
      ${pies}
    </svg>
  `;
}

function milestonePieGroups(summary) {
  const certificateRatio = clampRatio(summary.submittedWeight);
  const invoiceRatio = clampRatio(summary.paidFeeRatio);
  const healthCounts = summary.healthCounts || {};
  const totalCount = Math.max(summary.total || 0, 1);
  const pendingCount = Math.max(0, healthCounts.pending || 0);
  const dueCount = healthCounts.due || 0;
  const overdueCount = healthCounts.overdue || 0;
  const completeCount = healthCounts.complete || 0;
  const paidFee = Math.max(0, summary.paidFee || 0);
  const unpaidFee = Math.max(0, (summary.totalFee || 0) - paidFee);

  return [
    {
      label: "Certificates",
      note: "By milestone weight",
      centerValue: percent(certificateRatio),
      centerNote: `${summary.submitted} of ${summary.total} submitted`,
      segments: [
        { key: "complete", label: "Submitted", ratio: certificateRatio, value: `${summary.submitted} submitted` },
        { key: "pending", label: "Open", ratio: 1 - certificateRatio, value: `${summary.open} open` }
      ]
    },
    {
      label: "Invoices",
      note: "By AED value",
      centerValue: compactMoney(paidFee, "AED"),
      centerNote: `${percent(invoiceRatio)} paid`,
      segments: [
        { key: "complete", label: "Paid", ratio: invoiceRatio, value: compactMoney(paidFee, "AED") },
        { key: "overdue", label: "Unpaid", ratio: 1 - invoiceRatio, value: compactMoney(unpaidFee, "AED") }
      ]
    },
    {
      label: "Schedule",
      note: "By milestone count",
      centerValue: `${overdueCount} overdue`,
      centerNote: `${dueCount} due soon`,
      segments: [
        { key: "complete", label: "Submitted", ratio: completeCount / totalCount, value: String(completeCount) },
        { key: "due", label: "Due soon", ratio: dueCount / totalCount, value: String(dueCount) },
        { key: "overdue", label: "Overdue", ratio: overdueCount / totalCount, value: String(overdueCount) },
        { key: "pending", label: "Future", ratio: pendingCount / totalCount, value: String(pendingCount) }
      ]
    }
  ];
}

function milestonePieGroupSvg(group, cx, cy, radius) {
  const normalizedSegments = normalizePieSegments(group.segments);
  const slices = normalizedSegments
    .filter((segment) => segment.ratio > 0)
    .reduce((items, segment) => {
      const start = items.cursor;
      const end = start + segment.ratio;
      items.markup.push(milestonePieSlice(segment, cx, cy, radius, start, end));
      items.cursor = end;
      return items;
    }, { cursor: -0.25, markup: [] }).markup.join("");
  const labels = group.segments.map((segment, index) => {
    const ratio = clampRatio(segment.ratio);
    const labelY = 244 + index * 25;
    return `
      <g>
        <rect x="${cx - 116}" y="${labelY - 11}" width="10" height="10" rx="3" fill="${milestoneChartColor(segment.key)}"></rect>
        <text x="${cx - 100}" y="${labelY - 2}" font-size="10" font-weight="850" fill="var(--ink)">${escapeHtml(segment.label)}</text>
        <text x="${cx - 100}" y="${labelY + 11}" font-size="10" font-weight="750" fill="var(--muted)">${escapeHtml(`${segment.value} | ${percent(ratio)}`)}</text>
      </g>
    `;
  }).join("");

  return `
    <g>
      <text x="${cx}" y="28" text-anchor="middle" font-size="14" font-weight="900" fill="var(--ink)">${escapeHtml(group.label)}</text>
      <text x="${cx}" y="46" text-anchor="middle" font-size="11" font-weight="750" fill="var(--muted)">${escapeHtml(group.note)}</text>
      <circle cx="${cx}" cy="${cy}" r="${radius}" fill="var(--surface-soft)" stroke="var(--line)" stroke-width="1"></circle>
      ${slices}
      <circle cx="${cx}" cy="${cy}" r="39" fill="var(--surface)" stroke="var(--line)" stroke-width="1"></circle>
      <text x="${cx}" y="${cy - 3}" text-anchor="middle" font-size="13" font-weight="900" fill="var(--ink)">${escapeHtml(group.centerValue)}</text>
      <text x="${cx}" y="${cy + 15}" text-anchor="middle" font-size="10" font-weight="750" fill="var(--muted)">${escapeHtml(group.centerNote)}</text>
      ${labels}
    </g>
  `;
}

function normalizePieSegments(segments) {
  const total = sum(segments.map((segment) => Math.max(0, segment.ratio)));
  if (!total) {
    return [{ key: "pending", label: "No data", ratio: 1, value: "0" }];
  }
  return segments.map((segment) => ({
    ...segment,
    ratio: Math.max(0, segment.ratio) / total
  }));
}

function milestonePieSlice(segment, cx, cy, radius, start, end) {
  const ratio = end - start;
  if (ratio >= 0.9999) {
    return `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="${milestoneChartColor(segment.key)}"></circle>`;
  }
  const startPoint = pointOnCircle(cx, cy, radius, start);
  const endPoint = pointOnCircle(cx, cy, radius, end);
  const largeArc = ratio > 0.5 ? 1 : 0;
  return `
    <path d="M ${cx} ${cy} L ${startPoint.x} ${startPoint.y} A ${radius} ${radius} 0 ${largeArc} 1 ${endPoint.x} ${endPoint.y} Z" fill="${milestoneChartColor(segment.key)}" stroke="var(--surface)" stroke-width="2">
      <title>${escapeHtml(`${segment.label}: ${segment.value} | ${percent(segment.ratio)}`)}</title>
    </path>
  `;
}

function pointOnCircle(cx, cy, radius, turn) {
  const angle = turn * Math.PI * 2;
  return {
    x: roundChartNumber(cx + Math.cos(angle) * radius),
    y: roundChartNumber(cy + Math.sin(angle) * radius)
  };
}

function roundChartNumber(value) {
  return Math.round(value * 100) / 100;
}

function milestoneValueComparisonSvg(milestones) {
  const rows = [...milestones].sort((a, b) => toNumber(a.sno) - toNumber(b.sno));
  const width = 1120;
  const height = 330;
  const pad = { top: 46, right: 26, bottom: 72, left: 70 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const maxFee = Math.max(...rows.map((item) => Math.max(0, toNumber(item.feeAed))), 1);
  const gap = rows.length > 16 ? 5 : 8;
  const barW = Math.max(14, (innerW - gap * Math.max(0, rows.length - 1)) / Math.max(rows.length, 1));

  const grid = [0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = pad.top + innerH - innerH * ratio;
    return `
      <line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" stroke="var(--line)" stroke-width="1"></line>
      <text x="${pad.left - 10}" y="${y + 4}" text-anchor="end" font-size="10" font-weight="700" fill="var(--muted)">${escapeHtml(compactMoney(maxFee * ratio, "AED"))}</text>
    `;
  }).join("");

  const bars = rows.map((item, index) => {
    const fee = Math.max(0, toNumber(item.feeAed));
    const h = (fee / maxFee) * innerH;
    const x = pad.left + index * (barW + gap);
    const y = pad.top + innerH - h;
    const health = milestoneHealth(item);
    const label = `M${item.sno || index + 1}`;
    const valueLabel = compactMoney(fee, "AED");
    const showValue = barW >= 28;
    return `
      <g>
        <title>${escapeHtml(`${label}: ${item.milestone || "Milestone"} - ${money(fee, "AED")} - ${health.label}`)}</title>
        <rect x="${x}" y="${y}" width="${barW}" height="${h}" rx="5" fill="${milestoneChartColor(health.key)}"></rect>
        ${showValue ? `<text x="${x + barW / 2}" y="${Math.max(14, y - 7)}" text-anchor="middle" font-size="10" font-weight="850" fill="${milestoneChartColor(health.key)}">${escapeHtml(valueLabel)}</text>` : ""}
        <text x="${x + barW / 2}" y="${height - 44}" text-anchor="middle" font-size="10" font-weight="800" fill="var(--muted)">${escapeHtml(label)}</text>
      </g>
    `;
  }).join("");

  const legend = milestoneChartLegend([
    ["Submitted", "complete"],
    ["Due soon", "due"],
    ["Overdue", "overdue"],
    ["Future / open", "pending"]
  ], pad.left, height - 18);

  return `
    <svg class="milestone-chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Milestone fee comparison chart">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="var(--surface)"></rect>
      <line x1="${pad.left}" y1="${pad.top + innerH}" x2="${width - pad.right}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      ${grid}
      ${bars}
      ${legend}
    </svg>
  `;
}

function milestoneChartLegend(items, startX, y) {
  return `
    <g class="milestone-chart-legend">
      ${items.map(([label, key], index) => {
        const x = startX + index * 142;
        return `
          <rect x="${x}" y="${y - 10}" width="12" height="12" rx="3" fill="${milestoneChartColor(key)}"></rect>
          <text x="${x + 18}" y="${y}" font-size="11" font-weight="800" fill="var(--muted)">${escapeHtml(label)}</text>
        `;
      }).join("")}
    </g>
  `;
}

function milestoneChartColor(key) {
  if (key === "complete" || key === "paid") return "var(--green)";
  if (key === "due") return "var(--amber)";
  if (key === "overdue" || key === "unpaid") return "var(--berry)";
  return "var(--teal)";
}

function milestoneChartTextColor(key) {
  return key === "due" ? "var(--ink)" : "#ffffff";
}

function compactMoney(value, currency = "AED") {
  if (!Number.isFinite(value)) return "-";
  const absolute = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (absolute >= 1000000) return `${sign}${currency} ${(absolute / 1000000).toFixed(1)}M`;
  if (absolute >= 1000) return `${sign}${currency} ${Math.round(absolute / 1000)}K`;
  return `${sign}${currency} ${Math.round(absolute)}`;
}

function clampRatio(value) {
  if (!Number.isFinite(value)) return 0;
  return Math.min(1, Math.max(0, value));
}

function milestoneHealth(item) {
  if (isSubmitted(item.certificateStatus)) return { key: "complete", label: "Submitted" };
  const dueDate = item.expectedDate;
  if (!dueDate) return { key: "pending", label: "Pending" };
  const today = startOfDay(new Date());
  const dueSoonDate = new Date(today);
  dueSoonDate.setDate(dueSoonDate.getDate() + 30);
  if (dueDate < today) return { key: "overdue", label: "Overdue" };
  if (dueDate <= dueSoonDate) return { key: "due", label: "Due soon" };
  return { key: "pending", label: "Pending" };
}

function milestoneCard(label, value, note, className) {
  return `
    <div class="milestone-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(note)}</small>
    </div>
  `;
}

function milestoneStatusBadge(status, healthKey) {
  const label = cleanText(status) || "Unknown";
  return `<span class="milestone-badge milestone-badge-${healthKey}">${escapeHtml(label)}</span>`;
}

function invoiceStatusBadge(status) {
  const label = cleanText(status) || "Unknown";
  const key = isPaid(status) ? "paid" : "unpaid";
  return `<span class="milestone-badge milestone-badge-${key}">${escapeHtml(label)}</span>`;
}

function isSubmitted(status) {
  return normalizeText(status).includes("submitted") && !normalizeText(status).includes("not yet");
}

function isPaid(status) {
  const text = normalizeText(status);
  return text.includes("paid") && !text.includes("not paid");
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function fundingVariance(fund, cumulativeRevenue) {
  const variance = fund - cumulativeRevenue;
  if (Math.abs(variance) < 0.5) {
    return { key: "on", label: "On fund", variance };
  }
  return variance < 0
    ? { key: "over", label: "Over fund", variance }
    : { key: "under", label: "Under fund", variance };
}

function forecastFundingNote(variance) {
  if (variance.key === "on") return "On fund";
  const amount = money(Math.abs(variance.variance));
  return variance.key === "over" ? `Over fund by ${amount}` : `Less than fund by ${amount}`;
}

function varianceTone(variance) {
  if (variance.key === "over") return "risk";
  if (variance.key === "under") return "warn";
  return "accent";
}

function varianceCell(variance) {
  return `
    <span class="variance-cell">
      ${money(variance.variance)}
      <span class="variance-badge variance-${variance.key}">${escapeHtml(variance.label)}</span>
    </span>
  `;
}

function chartVarianceBadge(variance) {
  return `
    <div class="chart-variance">
      <span>${money(variance.variance)}</span>
      <span class="variance-badge variance-${variance.key}">${escapeHtml(variance.label)}</span>
    </div>
  `;
}

function trendForRows(rows, index) {
  if (index === 0) {
    return { direction: "flat", className: "trend-flat", percent: null };
  }
  const previous = toNumber(rows[index - 1]?.revenue);
  const current = toNumber(rows[index]?.revenue);
  const delta = current - previous;
  const percentChange = previous ? delta / Math.abs(previous) : null;
  if (Math.abs(delta) < 0.5) {
    return { direction: "flat", className: "trend-flat", percent: percentChange };
  }
  return {
    direction: delta > 0 ? "up" : "down",
    className: delta > 0 ? "trend-up" : "trend-down",
    percent: percentChange
  };
}

function trendBadge(trend) {
  const symbol = trend.direction === "up" ? "↑" : trend.direction === "down" ? "↓" : "→";
  const label = trend.direction === "up" ? "Up" : trend.direction === "down" ? "Down" : "Base";
  return `<span class="trend-badge ${trend.className}"><span>${symbol}</span>${label}</span>`;
}

function statusBadge(status) {
  return `<span class="status-badge status-${status.key}">${escapeHtml(status.label)}</span>`;
}

function trendPercent(trend) {
  if (!Number.isFinite(trend.percent)) return "-";
  const value = percent(trend.percent);
  return trend.percent > 0 ? `+${value}` : value;
}

function pocForRow(row, project) {
  if (Number.isFinite(row.poc)) return row.poc;
  if (project.fundUsd && Number.isFinite(row.cumulativeRevenue)) return row.cumulativeRevenue / project.fundUsd;
  return null;
}

function isFuturePeriod(row) {
  return periodStatus(row).key === "future";
}

function periodStatus(row) {
  if (!row.periodEndDate) return { key: "unknown", label: "Unknown" };
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  if (row.periodEndDate > currentQuarter.endDate) return { key: "future", label: "Future" };
  if (row.periodEndDate >= currentQuarter.startDate && row.periodEndDate <= currentQuarter.endDate) {
    return { key: "progress", label: "In progress" };
  }
  return { key: "closed", label: "Closed" };
}

function renderFormulaTrace(project) {
  const formulas = [
    ["Days", "Hours / 8"],
    ["Cumulative Days", "Previous cumulative days + current days"],
    ["EAC Days", "EAC hours / 8"],
    ["POC", "Cumulative days / EAC days"],
    ["Period Revenue", "(POC * funding) - prior cumulative revenue"],
    ["Cumulative Revenue", "Previous cumulative revenue + period revenue"]
  ];
  elements.formulaGrid.innerHTML = formulas
    .map(([title, formula]) => `
      <div class="formula-card">
        <span>${escapeHtml(title)}</span>
        <strong>${escapeHtml(formula)}</strong>
      </div>
    `)
    .join("");
}

function renderAudit(project) {
  const items = [
    ["Workbook", project.sourceName || "-"],
    ["Parsed Projects", String(state.projects.length)],
    ["Selected Project", project.number || project.name || "-"],
    ["Estimator Periods", String(project.rows.length)],
    ["Mode", project.modeLabel],
    ["Parser", "Local XLSX XML parser"]
  ];
  elements.auditList.innerHTML = items
    .map(([label, value]) => `
      <div class="audit-item">
        <strong>${escapeHtml(label)}</strong>
        <span>${escapeHtml(value)}</span>
      </div>
    `)
    .join("");
}

function metric(label, value, note, className = "") {
  return `
    <div class="metric-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong class="${metricValueClass(value)}">${escapeHtml(value)}</strong>
      <small>${escapeHtml(note)}</small>
    </div>
  `;
}

function metricValueClass(value) {
  const length = cleanText(value).length;
  if (length >= 15) return "metric-value metric-value-compact";
  if (length >= 12) return "metric-value metric-value-tight";
  return "metric-value";
}

function fact(label, value, note = "", className = "") {
  return `
    <div class="fact-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      ${note ? `<small>${escapeHtml(note)}</small>` : ""}
    </div>
  `;
}

function readout(label, value, className) {
  return `
    <div class="readout-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `;
}

function hoursCard(label, value, note, className) {
  return `
    <div class="hours-card ${className}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(note)}</small>
    </div>
  `;
}

function personCard(role, name, image) {
  if (!cleanText(name)) return "";
  const imageSource = image || localPersonImage(name);
  const avatar = imageSource
    ? `<img src="${escapeHtml(imageSource)}" alt="${escapeHtml(name)}">`
    : `<span class="person-initials">${escapeHtml(initials(name))}</span>`;
  return `
    <div class="person-card welcoming-person">
      <div class="person-photo">${avatar}</div>
      <div>
        <span>${escapeHtml(role)}</span>
        <strong>${escapeHtml(name)}</strong>
      </div>
    </div>
  `;
}

function localPersonImage(name) {
  return LOCAL_PERSON_IMAGES[normalizePersonName(name)] || "";
}

function localCustomerLogo(projectName) {
  const name = normalizeProjectKey(projectName);
  if (name.includes("solution") || name.includes("workload migration")) {
    return "assets/customer-logos/solution-plus.png";
  }
  if (name.includes("taqa")) {
    return "assets/customer-logos/taqa.png";
  }
  return "";
}

function initials(name) {
  return cleanText(name)
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() || "")
    .join("");
}

function emptyState(text) {
  return `<div class="empty-state">${escapeHtml(text)}</div>`;
}

function chartDataForProject(project) {
  const quarterRows = project.chartQuarterlyRows?.length ? prepareChartRows(project.chartQuarterlyRows) : aggregateRowsByQuarter(project.rows);
  const monthlyGroups = project.chartMonthlyGroups instanceof Map ? project.chartMonthlyGroups : groupRowsByQuarter(monthlyRowsForProject(project, quarterRows));
  return { quarterRows, monthlyGroups };
}

function monthlyRowsForProject(project, quarterRows) {
  const rows = project.rows?.length ? [...project.rows] : [...(project.chartMonthlyRows || [])];
  const groupKeys = new Set(Array.from(groupRowsByQuarter(rows).keys()));
  quarterRows.forEach((quarterRow) => {
    if (!groupKeys.has(quarterRow.key)) {
      rows.push(quarterRow);
      groupKeys.add(quarterRow.key);
    }
  });
  return rows;
}

function selectedQuarterForChart(project, chartData) {
  const currentKey = currentFiscalQuarterInfo(new Date()).key;
  const preferredKey = state.selectedQuarterKey || currentKey;
  return chartData.quarterRows.find((row) => row.key === preferredKey) || chartData.quarterRows[0];
}

function chartSvg(rows, options = {}) {
  if (!rows.length) return emptyState("No chart data available");
  const width = 980;
  const height = 390;
  const pad = { top: 68, right: 48, bottom: 66, left: 72 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const maxRevenue = Math.max(...rows.map((row) => Math.max(0, row.revenue || 0)), 1);
  const maxCum = Math.max(...rows.map((row) => Math.max(0, row.cumulativeRevenue || 0)), 1);
  const barGap = 14;
  const barW = Math.max(22, (innerW - barGap * (rows.length - 1)) / rows.length);

  const bars = rows.map((row, index) => {
    const x = pad.left + index * (barW + barGap);
    const h = ((row.revenue || 0) / maxRevenue) * innerH;
    const y = pad.top + innerH - h;
    const active = options.activeKey && row.key === options.activeKey;
    const palette = chartPeriodPalette(row, active);
    const attrs = options.interactive ? ` data-quarter-key="${escapeHtml(row.key || row.label)}" tabindex="0" role="button"` : "";
    return `<rect${attrs} class="${options.interactive ? "chart-quarter-hit" : ""}" x="${x}" y="${y}" width="${barW}" height="${h}" rx="4" fill="${palette.bar}" stroke="${active ? "var(--ink)" : "none"}" stroke-width="${active ? "3" : "0"}"></rect>`;
  }).join("");

  const points = rows.map((row, index) => {
    const x = pad.left + index * (barW + barGap) + barW / 2;
    const y = pad.top + innerH - ((row.cumulativeRevenue || 0) / maxCum) * innerH;
    return { x, y, value: row.cumulativeRevenue || 0, row };
  });

  const lineSegments = points.slice(1).map((point, index) => {
    const previous = points[index];
    const palette = chartPeriodPalette(point.row);
    const dash = isFuturePeriod(point.row) ? ` stroke-dasharray="8 6"` : "";
    return `<line x1="${previous.x}" y1="${previous.y}" x2="${point.x}" y2="${point.y}" stroke="${palette.line}" stroke-width="3"${dash}></line>`;
  }).join("");

  const labels = rows.map((row, index) => {
    const x = pad.left + index * (barW + barGap) + barW / 2;
    return `<text x="${x}" y="${height - 24}" text-anchor="middle" font-size="12" fill="var(--muted)">${escapeHtml(shortLabel(row.label))}</text>`;
  }).join("");

  const barValueLabels = rows.map((row, index) => {
    const x = pad.left + index * (barW + barGap) + barW / 2;
    const h = ((row.revenue || 0) / maxRevenue) * innerH;
    const y = pad.top + innerH - h;
    const inside = h > 30;
    const labelY = inside ? pad.top + innerH - 10 : y - 8;
    const palette = chartPeriodPalette(row);
    return `
      <text x="${x}" y="${labelY}" text-anchor="middle" font-size="11" font-weight="800" fill="${inside ? "#ffffff" : palette.bar}">${escapeHtml(chartMoney(row.revenue || 0))}</text>
    `;
  }).join("");

  const markers = points.map((point) => {
    const palette = chartPeriodPalette(point.row);
    return `<circle cx="${point.x}" cy="${point.y}" r="4" fill="${palette.line}"></circle>`;
  }).join("");

  const lineValueLabels = points.map((point) => {
    const value = chartMoney(point.value);
    const widthEstimate = value.length * 7 + 14;
    const preferredOffset = point.y < pad.top + 34 ? 30 : -24;
    const labelY = Math.min(pad.top + innerH - 34, Math.max(28, point.y + preferredOffset));
    const labelX = Math.min(width - pad.right - widthEstimate / 2, Math.max(pad.left + widthEstimate / 2, point.x));
    const palette = chartPeriodPalette(point.row);
    return `
      <line x1="${point.x}" y1="${point.y}" x2="${labelX}" y2="${labelY - 8}" stroke="var(--line)" stroke-width="1"></line>
      <rect x="${labelX - widthEstimate / 2}" y="${labelY - 16}" width="${widthEstimate}" height="20" rx="5" fill="var(--surface)" stroke="var(--line)"></rect>
      <text x="${labelX}" y="${labelY - 2}" text-anchor="middle" font-size="11" font-weight="800" fill="${palette.line}">${escapeHtml(value)}</text>
    `;
  }).join("");
  const fundingAlert = options.fundingVariance ? chartFundingAlert(options.fundingVariance, width, pad) : "";
  const legend = chartLegend(width, pad);

  return `
    <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(options.title || "Revenue forecast chart")}">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="var(--surface)"></rect>
      <line x1="${pad.left}" y1="${pad.top + innerH}" x2="${width - pad.right}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + innerH}" stroke="var(--line)"></line>
      <text x="${pad.left}" y="18" font-size="12" font-weight="700" fill="var(--muted)">Period revenue</text>
      <text x="${width - pad.right}" y="18" text-anchor="end" font-size="12" font-weight="700" fill="var(--berry)">Cumulative revenue</text>
      ${fundingAlert}
      ${legend}
      ${bars}
      ${barValueLabels}
      ${points.length === 1 ? "" : lineSegments}
      ${markers}
      ${lineValueLabels}
      ${labels}
    </svg>
  `;
}

function chartPeriodPalette(row, active = false) {
  if (isFuturePeriod(row)) {
    return {
      bar: active ? "var(--chart-future-bar-strong)" : "var(--chart-future-bar)",
      line: "var(--chart-future-line)"
    };
  }
  return {
    bar: active ? "var(--chart-actual-bar-strong)" : "var(--chart-actual-bar)",
    line: "var(--chart-actual-line)"
  };
}

function chartLegend(width, pad) {
  const x = pad.left;
  const y = 36;
  return `
    <g class="chart-legend">
      <rect x="${x}" y="${y - 10}" width="12" height="12" rx="3" fill="var(--chart-actual-bar)"></rect>
      <line x1="${x + 4}" y1="${y + 12}" x2="${x + 28}" y2="${y + 12}" stroke="var(--chart-actual-line)" stroke-width="3"></line>
      <text x="${x + 38}" y="${y + 2}" font-size="11" font-weight="800">Actual</text>
      <rect x="${x + 104}" y="${y - 10}" width="12" height="12" rx="3" fill="var(--chart-future-bar)"></rect>
      <line x1="${x + 108}" y1="${y + 12}" x2="${x + 132}" y2="${y + 12}" stroke="var(--chart-future-line)" stroke-width="3" stroke-dasharray="8 6"></line>
      <text x="${x + 142}" y="${y + 2}" font-size="11" font-weight="800">Future</text>
    </g>
  `;
}

function chartFundingAlert(variance, width, pad) {
  const label = `${variance.label}: ${chartMoney(variance.variance)}`;
  const alertWidth = Math.max(138, label.length * 7 + 18);
  const x = width - pad.right - alertWidth;
  const y = 34;
  return `
    <g class="chart-funding-alert chart-funding-alert-${variance.key}">
      <rect x="${x}" y="${y}" width="${alertWidth}" height="24" rx="6"></rect>
      <text x="${x + alertWidth / 2}" y="${y + 16}" text-anchor="middle" font-size="11" font-weight="850">${escapeHtml(label)}</text>
    </g>
  `;
}

function chartMoney(value) {
  if (!Number.isFinite(value)) return "$0";
  const rounded = Math.round(Math.abs(value));
  const digits = String(rounded).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return `${value < 0 ? "-$" : "$"}${digits}`;
}

function shortLabel(label) {
  const text = String(label || "");
  return text.length > 9 ? `${text.slice(0, 9)}` : text;
}

function currentProject() {
  return state.projects[state.selectedIndex] || state.projects[0];
}

function lastRow(project) {
  return project.rows[project.rows.length - 1];
}

function latestRevenueToDateRow(project) {
  return project.revenueToDateRows[project.revenueToDateRows.length - 1];
}

function currentProjectPoc(project) {
  const latest = latestRevenueToDateRow(project) || lastRow(project);
  if (!latest) return null;
  const rowPoc = pocForRow(latest, project);
  if (Number.isFinite(rowPoc)) return rowPoc;
  return project.fundUsd && Number.isFinite(project.revenueToDate) ? project.revenueToDate / project.fundUsd : null;
}

function buildProjectFromHours(base, periodInputs, modeLabel = "Formula mode") {
  const project = normalizeProjectBase(base);
  const eacHours = project.hoursWithContingency || project.hoursWithoutContingency || sum(periodInputs.map((item) => item.hours));
  const eacDays = eacHours / HOURS_PER_DAY;
  let cumulativeDays = 0;
  let cumulativeRevenue = 0;

  const rows = periodInputs
    .filter((item) => isFiniteNumber(item.hours))
    .map((item) => {
      const hours = toNumber(item.hours);
      const days = hours / HOURS_PER_DAY;
      cumulativeDays += days;
      const poc = eacDays ? cumulativeDays / eacDays : 0;
      const targetCumulativeRevenue = poc * project.fundUsd;
      const revenue = targetCumulativeRevenue - cumulativeRevenue;
      cumulativeRevenue = targetCumulativeRevenue;
      return {
        label: item.label,
        hours,
        days,
        cumulativeDays,
        eacDays,
        poc,
        funding: project.fundUsd,
        revenue,
        cumulativeRevenue,
        remainingRevenue: project.fundUsd - cumulativeRevenue
      };
    });

  return finalizeProject(project, rows, "hours", modeLabel);
}

function buildProjectFromRevenue(base, periodInputs, modeLabel = "Imported revenue") {
  const project = normalizeProjectBase(base);
  let cumulativeRevenue = 0;
  const rows = periodInputs
    .filter((item) => isFiniteNumber(item.revenue))
    .map((item) => {
      const revenue = toNumber(item.revenue);
      cumulativeRevenue += revenue;
      const poc = project.fundUsd ? cumulativeRevenue / project.fundUsd : 0;
      return {
        label: item.label,
        hours: null,
        days: null,
        cumulativeDays: null,
        eacDays: project.eacDays,
        poc,
        funding: project.fundUsd,
        revenue,
        cumulativeRevenue,
        remainingRevenue: project.fundUsd - cumulativeRevenue
      };
    });

  return finalizeProject(project, rows, "revenue", modeLabel);
}

function buildProjectFromEstimatorRows(base, periodInputs, modeLabel = "Workbook estimator") {
  const project = normalizeProjectBase(base);
  let previousCumulativeRevenue = 0;
  const rows = periodInputs.map((item) => {
    const sourceRevenue = sourceEstimatorRevenue(item, previousCumulativeRevenue);
    previousCumulativeRevenue = sourceRevenue.cumulativeRevenue;
    return {
      label: item.label,
      hours: toNullableNumber(item.hours),
      days: toNullableNumber(item.days),
      cumulativeDays: toNullableNumber(item.cumulativeDays),
      eacDays: toNullableNumber(item.eacDays),
      poc: toNullableNumber(item.poc),
      funding: toNullableNumber(item.funding) || project.fundUsd,
      revenue: sourceRevenue.revenue,
      cumulativeRevenue: sourceRevenue.cumulativeRevenue,
      remainingRevenue: project.fundUsd - sourceRevenue.cumulativeRevenue
    };
  });
  return finalizeProject(project, rows, "estimator", modeLabel);
}

function sourceEstimatorRevenue(item, previousCumulativeRevenue) {
  const cachedCumulative = toNullableNumber(item.cumulativeRevenue);
  const cachedRevenue = toNullableNumber(item.revenue);

  if (cachedCumulative !== null) {
    return {
      revenue: cachedRevenue ?? cachedCumulative - previousCumulativeRevenue,
      cumulativeRevenue: cachedCumulative
    };
  }

  const revenue = cachedRevenue ?? 0;
  return {
    revenue,
    cumulativeRevenue: previousCumulativeRevenue + revenue
  };
}

function normalizeProjectBase(base) {
  const fundUsd = toNumber(base.fundUsd);
  const fundAed = toNumber(base.fundAed) || fundUsd * EXCHANGE_RATE;
  const hoursWithContingency = toNumber(base.hoursWithContingency);
  const hoursWithoutContingency = toNumber(base.hoursWithoutContingency);
  const eacHours = hoursWithContingency || hoursWithoutContingency;
  return {
    sourceName: base.sourceName || "",
    number: cleanText(base.number),
    contractNumber: cleanText(base.contractNumber),
    name: cleanText(base.name) || "Unnamed Project",
    customerLogoImage: cleanText(base.customerLogoImage) || localCustomerLogo(base.name),
    projectOwner: cleanText(base.projectOwner),
    portfolioLead: cleanText(base.portfolioLead),
    projectOwnerImage: cleanText(base.projectOwnerImage),
    portfolioLeadImage: cleanText(base.portfolioLeadImage),
    accountNumber: cleanText(base.accountNumber),
    totalExpenses: toNullableNumber(base.totalExpenses),
    utilizedExpenses: toNullableNumber(base.utilizedExpenses),
    origBidHours: toNullableNumber(base.origBidHours),
    origBidRevenue: toNullableNumber(base.origBidRevenue),
    origBidCost: toNullableNumber(base.origBidCost),
    origBidMargin: toNullableNumber(base.origBidMargin),
    estimateAtClosureHours: toNullableNumber(base.estimateAtClosureHours),
    estimateAtClosureRevenue: toNullableNumber(base.estimateAtClosureRevenue),
    estimateAtClosureCost: toNullableNumber(base.estimateAtClosureCost),
    estimateAtClosureMargin: toNullableNumber(base.estimateAtClosureMargin),
    milestones: Array.isArray(base.milestones) ? base.milestones : [],
    unbilledRows: normalizeUnbilledRows(base.unbilledRows),
    fundUsd,
    fundAed,
    hoursWithoutContingency,
    hoursWithContingency,
    actualToDateHours: toNullableNumber(base.actualToDateHours),
    eacHours,
    eacDays: eacHours ? eacHours / HOURS_PER_DAY : null,
    forecastRevenueOverride: toNullableNumber(base.forecastRevenue),
    chartQuarterlyRows: normalizeChartRows(base.chartQuarterlyRows),
    chartMonthlyRows: normalizeChartRows(base.chartMonthlyRows)
  };
}

function normalizeChartRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows
    .map((row) => ({
      label: cleanText(row.label),
      revenue: toNumber(row.revenue),
      cumulativeRevenue: toNumber(row.cumulativeRevenue),
      hours: toNullableNumber(row.hours),
      days: toNullableNumber(row.days),
      cumulativeDays: toNullableNumber(row.cumulativeDays),
      eacDays: toNullableNumber(row.eacDays),
      poc: toNullableNumber(row.poc),
      periodEndDate: row.periodEndDate || null
    }))
    .filter((row) => row.label && Number.isFinite(row.revenue));
}

function normalizeUnbilledRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows
    .map((row) => ({
      label: cleanText(row.label),
      revenue: toNumber(row.revenue),
      invoicing: toNumber(row.invoicing),
      cumulativeRevenue: toNullableNumber(row.cumulativeRevenue),
      cumulativeInvoicing: toNullableNumber(row.cumulativeInvoicing),
      unbilledRevenue: toNullableNumber(row.unbilledRevenue),
      agedUnbilled: toNullableNumber(row.agedUnbilled)
    }))
    .filter((row) => row.label);
}

function finalizeProject(project, rows, mode, modeLabel) {
  const annotatedRows = annotateRowsWithFiscalCutoff(rows);
  const revenueToDateRows = annotatedRows.filter((row) => row.inRevenueToDate);
  const forecastRevenue = project.forecastRevenueOverride ?? sum(annotatedRows.map((row) => row.revenue));
  const revenueToDate = sum(revenueToDateRows.map((row) => row.revenue));
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  return {
    ...project,
    rows: annotatedRows,
    revenueToDateRows,
    mode,
    modeLabel,
    forecastRevenue,
    totalRevenue: revenueToDate,
    revenueToDate,
    revenueToDateNote: `Till ${currentQuarter.label}`,
    remainingRevenue: project.fundUsd - revenueToDate,
    remainingRevenueToDate: project.fundUsd - revenueToDate,
    currentQuarterLabel: `${currentQuarter.label} ending ${formatDate(currentQuarter.endDate)}`
  };
}

function annotateRowsWithFiscalCutoff(rows) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  let previousEndDate = null;

  return rows.map((row) => {
    const periodEndDate = inferPeriodEndDate(row.label, previousEndDate, currentQuarter);
    if (periodEndDate) previousEndDate = periodEndDate;
    return {
      ...row,
      periodEndDate,
      inRevenueToDate: periodEndDate ? periodEndDate <= currentQuarter.endDate : true
    };
  });
}

function aggregateRowsByQuarter(rows) {
  const groups = groupRowsByQuarter(rows);
  return Array.from(groups.values()).map((groupRows) => {
    const last = groupRows[groupRows.length - 1];
    return {
      key: last.key,
      label: last.displayLabel || last.label,
      displayLabel: last.displayLabel,
      revenue: sum(groupRows.map((row) => row.revenue)),
      cumulativeRevenue: last.cumulativeRevenue,
      periodEndDate: last.periodEndDate
    };
  });
}

function prepareChartRows(rows) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  let previousEndDate = null;
  return rows.map((row) => {
    const periodEndDate = row.periodEndDate || inferPeriodEndDate(row.label, previousEndDate, currentQuarter);
    if (periodEndDate) previousEndDate = periodEndDate;
    const quarter = periodEndDate ? fiscalQuarterForDate(periodEndDate) : quarterFromLabel(row.label, currentQuarter);
    return {
      ...row,
      key: row.key || quarter.key || cleanText(row.label),
      displayLabel: row.displayLabel || quarter.displayLabel || cleanText(row.label),
      periodEndDate
    };
  });
}

function groupRowsByQuarter(rows) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  const groups = new Map();
  let previousEndDate = null;

  rows.forEach((row) => {
    const periodEndDate = row.periodEndDate || inferPeriodEndDate(row.label, previousEndDate, currentQuarter);
    if (periodEndDate) previousEndDate = periodEndDate;
    const quarter = periodEndDate ? fiscalQuarterForDate(periodEndDate) : quarterFromLabel(row.label, currentQuarter);
    const key = quarter.key || cleanText(row.label);
    const displayLabel = quarter.displayLabel || cleanText(row.label);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({
      ...row,
      key,
      displayLabel,
      periodEndDate
    });
  });

  return groups;
}

function quarterFromLabel(label, fallbackQuarter) {
  const text = cleanText(label);
  const match = text.match(/\bQ([1-4])\s*[-_ ]?\s*(?:FY)?(\d{2,4})\b/i);
  if (match) {
    const quarter = Number(match[1]);
    const fiscalYear = normalizeFiscalYear(match[2]);
    return {
      quarter,
      fiscalYear,
      key: fiscalQuarterKey(quarter, fiscalYear),
      displayLabel: fiscalQuarterDisplayLabel(quarter, fiscalYear)
    };
  }
  return fallbackQuarter;
}

function currentFiscalQuarterInfo(date) {
  const month = date.getMonth();
  const year = date.getFullYear();
  const fiscalYear = month >= 5 ? year + 1 : year;
  let quarter = 4;
  if (month >= 5 && month <= 7) quarter = 1;
  else if (month >= 8 && month <= 10) quarter = 2;
  else if (month === 11 || month <= 1) quarter = 3;

  return {
    quarter,
    fiscalYear,
    key: fiscalQuarterKey(quarter, fiscalYear),
    label: `Q${quarter} FY${String(fiscalYear).slice(-2)}`,
    startDate: fiscalQuarterStartDate(quarter, fiscalYear),
    endDate: fiscalQuarterEndDate(quarter, fiscalYear)
  };
}

function fiscalQuarterKey(quarter, fiscalYear) {
  return `Q${quarter}-FY${fiscalYear}`;
}

function fiscalQuarterDisplayLabel(quarter, fiscalYear) {
  return `Q${quarter} FY${String(fiscalYear).slice(-2)}`;
}

function fiscalQuarterInfo(quarter, fiscalYear) {
  return {
    quarter,
    fiscalYear,
    key: fiscalQuarterKey(quarter, fiscalYear),
    displayLabel: fiscalQuarterDisplayLabel(quarter, fiscalYear)
  };
}

function fiscalQuarterEndDate(quarter, fiscalYear) {
  const month = FISCAL_QUARTER_END_MONTHS[quarter];
  const year = quarter <= 2 ? fiscalYear - 1 : fiscalYear;
  return endOfMonth(year, month);
}

function fiscalQuarterStartDate(quarter, fiscalYear) {
  const month = FISCAL_QUARTER_START_MONTHS[quarter];
  const year = quarter <= 3 ? fiscalYear - 1 : fiscalYear;
  return new Date(year, month, 1);
}

function fiscalQuarterForDate(date) {
  const info = currentFiscalQuarterInfo(new Date(date.getFullYear(), date.getMonth(), 1));
  return fiscalQuarterInfo(info.quarter, info.fiscalYear);
}

function parseFiscalQuarterKey(keyValue) {
  const match = cleanText(keyValue).match(/^Q([1-4])-FY(\d{4})$/i);
  if (!match) return null;
  return fiscalQuarterInfo(Number(match[1]), Number(match[2]));
}

function fiscalQuarterOrdinal(info) {
  return info.fiscalYear * 4 + info.quarter;
}

function addFiscalQuarters(info, offset) {
  const base = fiscalQuarterOrdinal(info) + offset;
  const fiscalYear = Math.floor((base - 1) / 4);
  const quarter = ((base - 1) % 4) + 1;
  return fiscalQuarterInfo(quarter, fiscalYear);
}

function fiscalQuarterSequence(start, end) {
  const sequence = [];
  let current = start;
  while (fiscalQuarterOrdinal(current) <= fiscalQuarterOrdinal(end)) {
    sequence.push(current);
    current = addFiscalQuarters(current, 1);
  }
  return sequence;
}

function inferPeriodEndDate(label, previousEndDate, currentQuarter) {
  const text = cleanText(label);
  if (!text) return null;

  const quarterYear = text.match(/\bQ([1-4])\s*[-_ ]?\s*(?:FY)?(\d{2,4})\b/i);
  if (quarterYear) {
    return fiscalQuarterEndDate(Number(quarterYear[1]), normalizeFiscalYear(quarterYear[2]));
  }

  const month = extractMonthIndex(text);
  if (month === null) return null;

  const explicitYear = extractYear(text);
  if (explicitYear) return endOfMonth(explicitYear, month);

  let year = previousEndDate ? previousEndDate.getFullYear() : inferredYearForMonth(month, currentQuarter);
  let date = endOfMonth(year, month);
  while (previousEndDate && date <= previousEndDate) {
    year += 1;
    date = endOfMonth(year, month);
  }
  return date;
}

function normalizeFiscalYear(value) {
  const numberValue = Number(value);
  return numberValue < 100 ? 2000 + numberValue : numberValue;
}

function extractMonthIndex(text) {
  const normalized = text.toLowerCase();
  for (const [name, index] of Object.entries(MONTH_INDEX)) {
    if (new RegExp(`\\b${name}\\b`, "i").test(normalized)) return index;
  }
  return null;
}

function extractYear(text) {
  const match = text.match(/\b(?:FY)?(\d{2,4})\b/i);
  if (!match) return null;
  const value = Number(match[1]);
  if (!Number.isFinite(value)) return null;
  return value < 100 ? 2000 + value : value;
}

function inferredYearForMonth(month, currentQuarter) {
  const fiscalYearStart = new Date(currentQuarter.fiscalYear - 1, 5, 1);
  const fiscalYearEnd = currentQuarter.endDate;
  let year = month >= 5 ? currentQuarter.fiscalYear - 1 : currentQuarter.fiscalYear;
  const date = endOfMonth(year, month);
  if (date < fiscalYearStart) return year + 1;
  if (date > fiscalYearEnd && month >= 5) return year - 1;
  return year;
}

function endOfMonth(year, month) {
  return new Date(year, month + 1, 0, 23, 59, 59, 999);
}

function formatDate(date) {
  if (!date) return "-";
  return `${MONTH_NAMES[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

function parseWorkbookDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const serial = toNullableNumber(value);
  if (serial !== null) return excelSerialDate(serial);
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function excelSerialDate(serial) {
  if (!Number.isFinite(serial)) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
}

function extractProjects(workbook, sourceName) {
  const actualHours = extractActualHoursByProject(workbook);
  const personPictures = extractPersonPictures(workbook);
  const milestones = extractMilestonesByProject(workbook);
  const unbilled = extractUnbilledByProject(workbook);
  const structuredMasterProjects = [];
  for (const sheet of workbook.sheets) {
    structuredMasterProjects.push(...extractStructuredMasterProjects(sheet, sourceName));
  }
  if (structuredMasterProjects.length) return enrichProjects(structuredMasterProjects, actualHours, personPictures, milestones, unbilled);

  const estimatorProjects = [];
  for (const sheet of workbook.sheets) {
    estimatorProjects.push(...extractEstimatorProjects(sheet, sourceName));
  }
  if (estimatorProjects.length) return enrichProjects(estimatorProjects, actualHours, personPictures, milestones, unbilled);

  const masterProjects = [];
  for (const sheet of workbook.sheets) {
    masterProjects.push(...extractMasterProjects(sheet, sourceName));
  }
  return enrichProjects(masterProjects, actualHours, personPictures, milestones, unbilled);
}

function extractStructuredMasterProjects(sheet, sourceName) {
  const header = findProjectHeader(sheet);
  if (!header) return [];

  const forecastStartCol = findSectionStartColumn(sheet, "Revenue Forecast", header.row) || findSectionStartColumn(sheet, "Forecast", header.row);
  if (!forecastStartCol) return [];

  const forecastEndCol = findNextTopSectionColumn(sheet, forecastStartCol, header.row) - 1 || sheet.maxCol;
  const sections = [{
    label: "forecast",
    groups: findMasterQuarterGroups(sheet, header.row, forecastStartCol, forecastEndCol)
  }];

  const actualStartCol = findSectionStartColumn(sheet, "Actuals", header.row) || findSectionStartColumn(sheet, "Actual", header.row);
  if (actualStartCol) {
    const actualEndCol = findNextTopSectionColumn(sheet, actualStartCol, header.row) - 1 || sheet.maxCol;
    const actualGroups = findMasterQuarterGroups(sheet, header.row, actualStartCol, actualEndCol);
    if (actualGroups.length) sections.push({ label: "actual", groups: actualGroups });
  }

  if (!sections.some((section) => section.groups.length)) return [];

  const projects = [];
  const firstProjectRow = findFirstProjectDataRow(sheet, header) || header.row + 3;
  for (let row = firstProjectRow; row <= sheet.maxRow; row += 1) {
    const name = getCellValue(sheet, row, header.columns.name);
    const numberValue = getCellValue(sheet, row, header.columns.number);
    if (!cleanText(name) || !cleanText(numberValue)) continue;

    const revenueRow = findMasterTypeRowBeforeProject(sheet, header, "Revenue", row) || header.row + 1;
    const invoicingRow = findMasterTypeRowBeforeProject(sheet, header, "Invoicing", row);
    const monthHeaderRow = findMasterTypeRowBeforeProject(sheet, header, "Month", row) || header.row + 2;

    const projectCardValues = extractProjectCardValues(sheet, header, row);
    const base = {
      ...projectCardValues,
      sourceName,
      number: numberValue,
      contractNumber: getCellValue(sheet, row, header.columns.contractNumber),
      accountNumber: getCellValue(sheet, row, header.columns.accountNumber) || projectCardValues.accountNumber,
      name,
      customerLogoImage: getCellImage(sheet, row, header.columns.customerLogoImage) || projectCardValues.customerLogoImage,
      projectOwner: projectPersonValue(sheet, header, row, "projectOwner", "Project Owner"),
      portfolioLead: projectPersonValue(sheet, header, row, "portfolioLead", "Portfolio Lead"),
      fundUsd: getCellValue(sheet, row, header.columns.fundUsd),
      fundAed: getCellValue(sheet, row, header.columns.fundAed),
      hoursWithoutContingency: getCellValue(sheet, row, header.columns.hoursWithoutContingency),
      hoursWithContingency: getCellValue(sheet, row, header.columns.hoursWithContingency)
    };

    const revenue = masterRevenueRowsForProject(sheet, row, sections, revenueRow, monthHeaderRow, invoicingRow);
    if (!revenue.periodRows.length) continue;

    projects.push(
      buildProjectFromEstimatorRows(
        {
          ...base,
          forecastRevenue: revenue.forecastRevenue,
          chartQuarterlyRows: revenue.quarterRows,
          chartMonthlyRows: revenue.monthlyRows,
          unbilledRows: revenue.unbilledRows
        },
        revenue.periodRows,
        "Master Sheet revenue forecast"
      )
    );
  }

  return projects;
}

function findSectionStartColumn(sheet, sectionLabel, maxHeaderRow) {
  const wanted = normalizeText(sectionLabel);
  for (let row = 1; row <= Math.max(1, maxHeaderRow); row += 1) {
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === wanted) return col;
    }
  }
  return null;
}

function findNextTopSectionColumn(sheet, startCol, maxHeaderRow) {
  let nextCol = null;
  for (let row = 1; row <= Math.max(1, maxHeaderRow); row += 1) {
    for (let col = startCol + 1; col <= sheet.maxCol; col += 1) {
      const text = normalizeText(getCellValue(sheet, row, col));
      if (!text) continue;
      if (text.includes("forecast") || text.includes("actual") || text.includes("revenue") || text.includes("project card")) {
        nextCol = nextCol === null ? col : Math.min(nextCol, col);
      }
    }
  }
  return nextCol || sheet.maxCol + 1;
}

function findFirstProjectDataRow(sheet, header) {
  for (let row = header.row + 1; row <= sheet.maxRow; row += 1) {
    const name = cleanText(getCellValue(sheet, row, header.columns.name));
    const numberValue = cleanText(getCellValue(sheet, row, header.columns.number));
    if (name && numberValue) return row;
  }
  return null;
}

function findMasterTypeRow(sheet, header, label, firstProjectRow) {
  const wanted = normalizeText(label);
  const { startCol, endCol } = masterTypeSearchBounds(sheet, header);
  for (let row = header.row + 1; row < firstProjectRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === wanted) return row;
    }
  }
  return null;
}

function findMasterTypeRowBeforeProject(sheet, header, label, projectRow) {
  const wanted = normalizeText(label);
  const { startCol, endCol } = masterTypeSearchBounds(sheet, header);
  let foundRow = null;
  for (let row = header.row + 1; row < projectRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === wanted) foundRow = row;
    }
  }
  return foundRow;
}

function masterTypeSearchBounds(sheet, header) {
  const typeCol = findSectionStartColumn(sheet, "Type", header.row);
  if (typeCol) {
    return {
      startCol: Math.max(1, typeCol - 1),
      endCol: Math.min(sheet.maxCol, typeCol + 3)
    };
  }

  const startCol = Math.max(1, Math.max(...Object.values(header.columns)) + 1);
  return {
    startCol,
    endCol: Math.min(sheet.maxCol, startCol + 12)
  };
}

function extractProjectCardValues(sheet, header, projectRow) {
  const values = {};
  const startCol = findSectionStartColumn(sheet, "Project Card", header.row) || Math.min(...Object.values(header.columns));
  const typeCol = findSectionStartColumn(sheet, "Type", header.row) || sheet.maxCol + 1;
  const endCol = Math.min(typeCol - 1, sheet.maxCol);
  let groupLabel = "";

  for (let col = startCol; col <= endCol; col += 1) {
    const topLabel = cleanText(getCellValue(sheet, header.row, col));
    const detailLabel = cleanText(getCellValue(sheet, header.row + 1, col));
    if (topLabel) groupLabel = topLabel;
    const label = projectCardMetricLabel(groupLabel, detailLabel || topLabel);
    const field = projectCardFieldName(label);
    if (!field) continue;
    if (field.endsWith("Image")) {
      const image = getCellImage(sheet, projectRow, col);
      if (image) values[field] = image;
      continue;
    }
    const value = getCellValue(sheet, projectRow, col);
    if (cleanText(value) || Number.isFinite(toNullableNumber(value))) {
      values[field] = value;
    }
  }

  return values;
}

function projectCardMetricLabel(groupLabel, detailLabel) {
  const group = normalizeText(groupLabel);
  const detail = cleanText(detailLabel);
  if (!detail) return cleanText(groupLabel);
  if (group.includes("orig bid")) return `Orig Bid ${detail}`;
  if (group.includes("eac")) return `EAC ${detail}`;
  if (normalizeText(detail) === "utilized expenses") return "Utilized Expenses";
  return detail;
}

function projectCardFieldName(label) {
  const text = normalizeText(label);
  const map = {
    "account number": "accountNumber",
    "customer logo": "customerLogoImage",
    "customer logo image": "customerLogoImage",
    "logo": "customerLogoImage",
    "total expenses": "totalExpenses",
    "utilized expenses": "utilizedExpenses",
    "orig bid hours": "origBidHours",
    "orig bid revenue": "origBidRevenue",
    "orig bid cost": "origBidCost",
    "orig bid margin": "origBidMargin",
    "eac hours": "estimateAtClosureHours",
    "eac revenue": "estimateAtClosureRevenue",
    "eac cost": "estimateAtClosureCost",
    "eac margin": "estimateAtClosureMargin"
  };
  return map[text] || null;
}

function findMasterQuarterGroups(sheet, quarterHeaderRow, startCol, endCol) {
  const groups = [];
  for (let col = startCol; col <= endCol; col += 1) {
    const label = cleanText(getCellValue(sheet, quarterHeaderRow, col));
    if (!isQuarterLabel(label)) continue;
    const match = label.match(/\bQ([1-4])\s*[-_ ]?\s*(?:FY)?(\d{2,4})\b/i);
    groups.push({
      label,
      quarter: Number(match[1]),
      fiscalYear: normalizeFiscalYear(match[2]),
      startCol: col,
      endCol
    });
  }

  return groups.map((group, index) => ({
    ...group,
    endCol: index + 1 < groups.length ? groups[index + 1].startCol - 1 : endCol
  }));
}

function masterRevenueRowsForProject(sheet, projectRow, sections, revenueRow, monthHeaderRow, invoicingRow) {
  const financialByQuarter = new Map();
  sections.forEach((section) => {
    section.groups.forEach((group) => {
      const key = fiscalQuarterKey(group.quarter, group.fiscalYear);
      const existing = financialByQuarter.get(key) || {
        key,
        label: group.label,
        displayLabel: fiscalQuarterDisplayLabel(group.quarter, group.fiscalYear),
        quarter: group.quarter,
        fiscalYear: group.fiscalYear,
        revenue: null,
        invoicing: null,
        monthlyValues: []
      };
      const monthlyValues = masterMonthlyValuesForGroup(sheet, projectRow, group, monthHeaderRow);
      const quarterlyValue = toNullableNumber(getCellValue(sheet, revenueRow, group.startCol));
      const monthlyTotal = sum(monthlyValues.map((item) => item.revenue));
      const revenue = quarterlyValue ?? (monthlyTotal || null);
      const invoicing = invoicingRow ? toNullableNumber(getCellValue(sheet, invoicingRow, group.startCol)) : null;

      if (revenue !== null && Math.abs(revenue) > 0.000001 && (section.label === "actual" || existing.revenue === null)) {
        existing.revenue = revenue;
        existing.monthlyValues = monthlyValues;
      }
      if (invoicing !== null && Math.abs(invoicing) > 0.000001 && (section.label === "actual" || existing.invoicing === null)) {
        existing.invoicing = invoicing;
      }
      financialByQuarter.set(key, existing);
    });
  });

  const financialRows = Array.from(financialByQuarter.values())
    .filter((row) => row.revenue !== null || row.invoicing !== null)
    .sort((a, b) => fiscalQuarterOrdinal(a) - fiscalQuarterOrdinal(b));

  const quarterRows = [];
  const monthlyRows = [];
  const periodRows = [];
  let cumulativeRevenue = 0;

  financialRows.forEach((group) => {
    const revenue = toNumber(group.revenue);
    if (Math.abs(revenue) < 0.000001) return;

    const startingCumulative = cumulativeRevenue;
    const quarterCumulative = cumulativeRevenue + revenue;
    const quarterRow = {
      label: group.label,
      revenue,
      cumulativeRevenue: quarterCumulative,
      periodEndDate: fiscalQuarterEndDate(group.quarter, group.fiscalYear)
    };
    quarterRows.push(quarterRow);

    if (group.monthlyValues.length) {
      let monthlyCumulative = startingCumulative;
      group.monthlyValues.forEach((item) => {
        monthlyCumulative += item.revenue;
        const monthlyRow = {
          label: item.label,
          revenue: item.revenue,
          cumulativeRevenue: monthlyCumulative
        };
        monthlyRows.push(monthlyRow);
        periodRows.push(monthlyRow);
      });
      cumulativeRevenue = quarterCumulative;
    } else {
      cumulativeRevenue = quarterCumulative;
      monthlyRows.push(quarterRow);
      periodRows.push(quarterRow);
    }
  });

  const latestQuarter = quarterRows[quarterRows.length - 1];
  return {
    quarterRows,
    monthlyRows,
    periodRows,
    forecastRevenue: latestQuarter ? latestQuarter.cumulativeRevenue : null,
    unbilledRows: masterUnbilledRowsFromFinancialRows(financialRows)
  };
}

function masterUnbilledRowsFromFinancialRows(financialRows) {
  if (!financialRows.length) return [];
  const first = financialRows[0];
  const last = financialRows[financialRows.length - 1];
  const rowsByQuarter = new Map(financialRows.map((row) => [row.key, row]));
  const rows = [];
  let cumulativeRevenue = 0;
  let cumulativeInvoicing = 0;

  fiscalQuarterSequence(addFiscalQuarters(first, -2), last).forEach((quarter) => {
    const source = rowsByQuarter.get(quarter.key);
    const revenue = toNumber(source?.revenue);
    const invoicing = toNumber(source?.invoicing);
    cumulativeRevenue += revenue;
    cumulativeInvoicing += invoicing;
    rows.push({
      label: quarter.displayLabel,
      revenue,
      invoicing,
      cumulativeRevenue,
      cumulativeInvoicing,
      unbilledRevenue: cumulativeRevenue - cumulativeInvoicing,
      agedUnbilled: 0
    });
  });

  const firstRevenueIndex = rows.findIndex((row) => Math.abs(row.revenue) > 0.000001);
  rows.forEach((row, index) => {
    if (firstRevenueIndex === -1 || index < firstRevenueIndex + 2 || index < 2) {
      row.agedUnbilled = 0;
      return;
    }
    row.agedUnbilled = rows[index - 2].cumulativeRevenue - row.cumulativeInvoicing;
  });

  return rows;
}

function masterMonthlyValuesForGroup(sheet, projectRow, group, monthHeaderRow) {
  const values = [];
  for (let col = group.startCol; col <= group.endCol; col += 1) {
    const revenue = toNullableNumber(getCellValue(sheet, projectRow, col));
    if (revenue === null || Math.abs(revenue) < 0.000001) continue;
    const month = extractMonthIndex(cleanText(getCellValue(sheet, monthHeaderRow, col)));
    const label = month === null
      ? group.label
      : `${MONTH_NAMES[month]} ${calendarYearForFiscalMonth(month, group.fiscalYear)}`;
    values.push({ label, revenue });
  }
  return values;
}

function calendarYearForFiscalMonth(month, fiscalYear) {
  return month >= 5 ? fiscalYear - 1 : fiscalYear;
}

function extractActualHoursByProject(workbook) {
  const entries = [];
  for (const sheet of workbook.sheets) {
    const sheetName = normalizeText(sheet.name);
    const looksLikeTimesheet = sheetName.includes("timesheet") || sheetName.includes("time sheet") || hasLabel(sheet, "Sum of Quantity");
    if (!looksLikeTimesheet) continue;
    const projectBlocks = extractTimesheetProjectBlocks(sheet);
    if (projectBlocks.length) {
      entries.push(...projectBlocks);
      continue;
    }

    const hours = extractGrandTotalHours(sheet);
    if (hours !== null) {
      entries.push({
        projectName: cleanText(valueRightOf(sheet, "Project Name")),
        hours
      });
    }
  }
  return entries;
}

function extractTimesheetProjectBlocks(sheet) {
  const projectCells = findLabelCells(sheet).filter((cell) => normalizeText(cell.value) === "project name");
  const entries = [];

  projectCells.forEach((cell, index) => {
    const projectName = cleanText(getCellValue(sheet, cell.row, cell.col + 1));
    if (!projectName) return;
    const nextProjectRow = projectCells[index + 1]?.row || sheet.maxRow + 1;
    const totals = extractTimesheetHoursInRange(sheet, cell.row + 1, nextProjectRow - 1);
    if (totals.hours !== null) entries.push({ projectName, hours: totals.hours, monthlyHours: totals.monthlyHours });
  });

  return entries;
}

function extractTimesheetHoursInRange(sheet, startRow, endRow) {
  const cells = findLabelCells(sheet).filter((cell) => {
    return cell.row >= startRow && cell.row <= endRow && normalizeText(cell.value) === "grand total";
  });
  for (const cell of cells) {
    for (let col = sheet.maxCol; col > cell.col; col -= 1) {
      const value = toNullableNumber(getCellValue(sheet, cell.row, col));
      if (value !== null) {
        return {
          hours: value,
          monthlyHours: extractMonthlyTimesheetHours(sheet, startRow, cell.row, cell.col, col)
        };
      }
    }
  }
  return { hours: null, monthlyHours: [] };
}

function extractGrandTotalHoursInRange(sheet, startRow, endRow) {
  return extractTimesheetHoursInRange(sheet, startRow, endRow).hours;
}

function extractMonthlyTimesheetHours(sheet, blockStartRow, totalRow, labelCol, grandTotalCol) {
  const headerRow = findTimesheetMonthHeaderRow(sheet, blockStartRow, totalRow - 1, labelCol, grandTotalCol);
  if (!headerRow) return [];

  const monthColumns = new Map();
  for (let col = labelCol + 1; col < grandTotalCol; col += 1) {
    const header = cleanText(getCellValue(sheet, headerRow, col));
    const month = extractMonthIndex(header);
    if (month === null) continue;
    const isTotal = normalizeText(header).includes("total");
    const existing = monthColumns.get(month);
    if (!existing || isTotal || !existing.isTotal) {
      monthColumns.set(month, { col, header, isTotal });
    }
  }

  return Array.from(monthColumns.entries())
    .map(([month, item]) => {
      const hours = toNullableNumber(getCellValue(sheet, totalRow, item.col));
      if (hours === null) return null;
      const year = timesheetCalendarYear(month, item.header);
      const periodEndDate = endOfMonth(year, month);
      return {
        label: `${MONTH_NAMES[month]} ${year}`,
        month,
        year,
        hours,
        days: hours / HOURS_PER_DAY,
        periodEndDate
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.periodEndDate - b.periodEndDate);
}

function findTimesheetMonthHeaderRow(sheet, startRow, endRow, labelCol, grandTotalCol) {
  let bestRow = null;
  let bestScore = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    let score = 0;
    for (let col = labelCol + 1; col <= grandTotalCol; col += 1) {
      const text = cleanText(getCellValue(sheet, row, col));
      if (!text) continue;
      const normalized = normalizeText(text);
      if (normalized === "grand total") score += 3;
      if (extractMonthIndex(text) !== null) score += normalized.includes("total") ? 2 : 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }
  return bestRow;
}

function timesheetCalendarYear(month, headerText) {
  const explicitYear = extractYear(cleanText(headerText));
  if (explicitYear && explicitYear >= 2000) return explicitYear;
  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  return month > currentMonth ? currentYear - 1 : currentYear;
}

function extractPersonPictures(workbook) {
  const pictures = new Map();
  for (const sheet of workbook.sheets) {
    if (!isPicSheet(sheet)) continue;
    const header = findPicHeader(sheet);
    if (!header) continue;

    for (let row = header.row + 1; row <= sheet.maxRow; row += 1) {
      addPersonPicture(pictures, getCellValue(sheet, row, header.ownerCol), getCellImage(sheet, row, header.ownerPicCol));
      addPersonPicture(pictures, getCellValue(sheet, row, header.leadCol), getCellImage(sheet, row, header.leadPicCol));
    }
  }
  return pictures;
}

function extractMilestonesByProject(workbook) {
  const milestonesByProject = new Map();
  for (const sheet of workbook.sheets) {
    const header = findMilestoneHeader(sheet);
    if (!header) continue;

    for (let row = header.row + 1; row <= sheet.maxRow; row += 1) {
      const projectName = cleanText(getCellValue(sheet, row, header.columns.projectName));
      const milestone = cleanText(getCellValue(sheet, row, header.columns.milestone));
      if (!projectName || !milestone) continue;

      const item = {
        projectName,
        sno: getCellValue(sheet, row, header.columns.sno),
        milestone,
        percentage: toNullableNumber(getCellValue(sheet, row, header.columns.percentage)),
        feeAed: toNullableNumber(getCellValue(sheet, row, header.columns.feeAed)),
        expectedDate: parseWorkbookDate(getCellValue(sheet, row, header.columns.expectedDate)),
        certificateStatus: cleanText(getCellValue(sheet, row, header.columns.certificateStatus)),
        invoiceStatus: cleanText(getCellValue(sheet, row, header.columns.invoiceStatus)),
        notes: cleanText(getCellValue(sheet, row, header.columns.notes))
      };
      const key = normalizeProjectKey(projectName);
      if (!milestonesByProject.has(key)) milestonesByProject.set(key, []);
      milestonesByProject.get(key).push(item);
    }
  }
  return milestonesByProject;
}

function extractUnbilledByProject(workbook) {
  const unbilledByProject = new Map();
  for (const sheet of workbook.sheets) {
    const sections = findUnbilledSections(sheet);
    sections.forEach((section) => {
      const rows = unbilledRowsFromSection(sheet, section);
      if (!rows.length) return;
      const projectName = cleanText(valueRightOf(sheet, "Project Name")) || cleanText(valueRightOf(sheet, "Project")) || "__fallback__";
      const key = projectName === "__fallback__" ? projectName : normalizeProjectKey(projectName);
      unbilledByProject.set(key, rows);
    });
  }
  return unbilledByProject;
}

function findUnbilledSections(sheet) {
  return findLabelCells(sheet)
    .filter((cell) => normalizeText(cell.value) === "unbilled revenue analysis")
    .map((cell) => {
      const revenueRow = findUnbilledSectionRow(sheet, "Revenue", cell.row + 1, cell.col);
      const invoicingRow = findUnbilledSectionRow(sheet, "Invoicing", cell.row + 1, cell.col);
      const cumulativeRevenueRow = findUnbilledSectionRow(sheet, "Cumulative Revenue", cell.row + 1, cell.col);
      const cumulativeInvoicingRow = findUnbilledSectionRow(sheet, "Cumulative Invoicing", cell.row + 1, cell.col);
      const unbilledRevenueRow = findUnbilledSectionRow(sheet, "Unbilled Revenue", cell.row + 1, cell.col);
      const agedUnbilledRow = findUnbilledSectionRow(sheet, "Unbilled Revenue >180 days", cell.row + 1, cell.col);
      if (!revenueRow || !invoicingRow || !cumulativeRevenueRow || !cumulativeInvoicingRow || !unbilledRevenueRow || !agedUnbilledRow) return null;
      return {
        labelCol: revenueRow.col,
        headerRow: revenueRow.row - 1,
        revenueRow: revenueRow.row,
        invoicingRow: invoicingRow.row,
        cumulativeRevenueRow: cumulativeRevenueRow.row,
        cumulativeInvoicingRow: cumulativeInvoicingRow.row,
        unbilledRevenueRow: unbilledRevenueRow.row,
        agedUnbilledRow: agedUnbilledRow.row
      };
    })
    .filter(Boolean);
}

function findUnbilledSectionRow(sheet, label, startRow, labelCol) {
  const normalized = normalizeText(label);
  for (let row = startRow; row <= Math.min(sheet.maxRow, startRow + 20); row += 1) {
    for (let col = Math.max(1, labelCol - 1); col <= Math.min(sheet.maxCol, labelCol + 2); col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === normalized) return { row, col };
    }
  }
  return null;
}

function unbilledRowsFromSection(sheet, section) {
  const rows = [];
  for (let col = section.labelCol + 1; col <= sheet.maxCol; col += 1) {
    const label = cleanText(getCellValue(sheet, section.headerRow, col));
    if (!label) continue;
    rows.push({
      label,
      revenue: toNumber(getCellValue(sheet, section.revenueRow, col)),
      invoicing: toNumber(getCellValue(sheet, section.invoicingRow, col)),
      cumulativeRevenue: toNumber(getCellValue(sheet, section.cumulativeRevenueRow, col)),
      cumulativeInvoicing: toNumber(getCellValue(sheet, section.cumulativeInvoicingRow, col)),
      unbilledRevenue: toNumber(getCellValue(sheet, section.unbilledRevenueRow, col)),
      agedUnbilled: toNumber(getCellValue(sheet, section.agedUnbilledRow, col))
    });
  }

  let lastUsefulIndex = rows.length - 1;
  while (lastUsefulIndex >= 0) {
    const row = rows[lastUsefulIndex];
    const values = [row.revenue, row.invoicing, row.cumulativeRevenue, row.cumulativeInvoicing, row.unbilledRevenue, row.agedUnbilled];
    if (values.some((value) => Math.abs(value) > 0.000001)) break;
    lastUsefulIndex -= 1;
  }
  return rows.slice(0, lastUsefulIndex + 1);
}

function findMilestoneHeader(sheet) {
  const wanted = {
    projectName: ["project name"],
    sno: ["sno", "s no", "serial no"],
    milestone: ["milestone"],
    percentage: ["percentage"],
    feeAed: ["fee in aed", "fee aed"],
    expectedDate: ["expected date to submit the certificate", "expected date"],
    certificateStatus: ["certificate status"],
    invoiceStatus: ["invoice status"],
    notes: ["notes"]
  };

  for (let row = 1; row <= Math.min(sheet.maxRow, 12); row += 1) {
    const columns = {};
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      const text = normalizeText(getCellValue(sheet, row, col));
      for (const [key, labels] of Object.entries(wanted)) {
        if (labels.includes(text)) columns[key] = col;
      }
    }
    if (columns.projectName && columns.milestone && columns.certificateStatus && columns.invoiceStatus) {
      return { row, columns };
    }
  }
  return null;
}

function isPicSheet(sheet) {
  const name = normalizeText(sheet.name);
  return name === "pic" || name.includes("pic") || hasLabel(sheet, "Pic");
}

function findPicHeader(sheet) {
  for (let row = 1; row <= Math.min(sheet.maxRow, 8); row += 1) {
    let ownerCol = null;
    let leadCol = null;
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      const text = normalizeText(getCellValue(sheet, row, col));
      if (text === "project owner") ownerCol = col;
      if (text === "portfolio lead") leadCol = col;
    }
    if (ownerCol && leadCol) {
      return {
        row,
        ownerCol,
        ownerPicCol: findPicColumn(sheet, row, ownerCol, leadCol),
        leadCol,
        leadPicCol: findPicColumn(sheet, row, leadCol, sheet.maxCol + 1)
      };
    }
  }
  return null;
}

function findPicColumn(sheet, row, startCol, stopCol) {
  for (let col = startCol + 1; col < stopCol; col += 1) {
    if (normalizeText(getCellValue(sheet, row, col)) === "pic") return col;
  }
  return startCol + 1;
}

function addPersonPicture(pictures, name, image) {
  const key = normalizePersonName(name);
  if (key && image) pictures.set(key, image);
}

function extractGrandTotalHours(sheet) {
  const cells = findLabelCells(sheet).filter((cell) => normalizeText(cell.value) === "grand total");
  for (const cell of cells) {
    for (let col = sheet.maxCol; col > cell.col; col -= 1) {
      const value = toNullableNumber(getCellValue(sheet, cell.row, col));
      if (value !== null) return value;
    }
  }
  return null;
}

function enrichProjects(projects, actualHoursEntries, personPictures, milestonesByProject, unbilledByProject) {
  return applyUnbilledToProjects(
    applyMilestonesToProjects(
      applyPersonPicturesToProjects(applyActualHoursToProjects(projects, actualHoursEntries), personPictures),
      milestonesByProject
    ),
    unbilledByProject
  );
}

function applyActualHoursToProjects(projects, actualHoursEntries) {
  return projects.map((project) => {
    const match = actualHoursEntries.find((entry) => sameProjectName(project.name, entry.projectName));
    const fallback = projects.length === 1 && actualHoursEntries.length === 1 ? actualHoursEntries[0] : null;
    const hours = match?.hours ?? fallback?.hours;
    const actualEntry = match || fallback || null;
    const updatedProject = hours === undefined ? project : { ...project, actualToDateHours: hours };
    return enrichRevenueRowsWithHours(updatedProject, actualEntry);
  });
}

function enrichRevenueRowsWithHours(project, actualEntry) {
  if (!project.rows?.length) return project;
  const monthlyHours = Array.isArray(actualEntry?.monthlyHours) ? actualEntry.monthlyHours : [];
  const rows = decorateRowsWithReadableHours(project.rows, project, monthlyHours);
  const monthlyRows = buildReadableMonthlyRows(rows, project, monthlyHours);
  return finalizeProject(
    {
      ...project,
      chartMonthlyGroups: groupRowsByQuarter(monthlyRows)
    },
    rows,
    project.mode,
    project.modeLabel
  );
}

function decorateRowsWithReadableHours(rows, project, monthlyHours) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  let previousEndDate = null;
  let cumulativeHours = 0;

  return rows.map((row) => {
    const periodEndDate = row.periodEndDate || inferPeriodEndDate(row.label, previousEndDate, currentQuarter);
    if (periodEndDate) previousEndDate = periodEndDate;
    const quarter = periodEndDate ? fiscalQuarterForDate(periodEndDate) : quarterFromLabel(row.label, currentQuarter);
    const existingHours = toNullableNumber(row.hours);
    const periodHours = existingHours
      ?? timesheetHoursForRow(row, periodEndDate, quarter, monthlyHours)
      ?? derivedHoursForRevenue(row, project);
    const existingCumulativeDays = toNullableNumber(row.cumulativeDays);

    if (periodHours !== null) {
      cumulativeHours += periodHours;
    } else if (existingCumulativeDays !== null) {
      cumulativeHours = existingCumulativeDays * HOURS_PER_DAY;
    }

    return {
      ...row,
      periodEndDate,
      hours: periodHours,
      days: periodHours !== null ? periodHours / HOURS_PER_DAY : toNullableNumber(row.days),
      cumulativeDays: periodHours !== null ? cumulativeHours / HOURS_PER_DAY : existingCumulativeDays,
      eacDays: toNullableNumber(row.eacDays) ?? project.eacDays
    };
  });
}

function buildReadableMonthlyRows(rows, project, monthlyHours) {
  const monthlyRows = [];

  rows.forEach((row) => {
    const rowQuarter = row.periodEndDate ? fiscalQuarterForDate(row.periodEndDate) : quarterFromLabel(row.label, currentFiscalQuarterInfo(new Date()));
    const quarterHours = monthlyHours.filter((item) => fiscalQuarterForDate(item.periodEndDate).key === rowQuarter.key);

    if (isQuarterLabel(row.label) && quarterHours.length) {
      const totalHours = sum(quarterHours.map((item) => item.hours));
      const startRevenue = (row.cumulativeRevenue || 0) - (row.revenue || 0);
      const startHours = (toNumber(row.cumulativeDays) * HOURS_PER_DAY) - toNumber(row.hours);
      let cumulativeRevenue = startRevenue;
      let cumulativeHours = startHours;

      quarterHours.forEach((item, index) => {
        const isLast = index === quarterHours.length - 1;
        const revenue = isLast
          ? (row.cumulativeRevenue || 0) - cumulativeRevenue
          : (row.revenue || 0) * (totalHours ? item.hours / totalHours : 1 / quarterHours.length);
        cumulativeRevenue += revenue;
        cumulativeHours += item.hours;
        monthlyRows.push({
          ...row,
          label: item.label,
          revenue,
          cumulativeRevenue,
          hours: item.hours,
          days: item.days,
          cumulativeDays: cumulativeHours / HOURS_PER_DAY,
          periodEndDate: item.periodEndDate,
          poc: project.fundUsd ? cumulativeRevenue / project.fundUsd : row.poc
        });
      });
      return;
    }

    monthlyRows.push(row);
  });

  return monthlyRows;
}

function timesheetHoursForRow(row, periodEndDate, quarter, monthlyHours) {
  if (!monthlyHours.length) return null;
  if (isQuarterLabel(row.label)) {
    const hours = sum(monthlyHours
      .filter((item) => fiscalQuarterForDate(item.periodEndDate).key === quarter.key)
      .map((item) => item.hours));
    return hours || null;
  }

  const direct = periodEndDate
    ? monthlyHours.find((item) => monthPeriodKey(item.periodEndDate) === monthPeriodKey(periodEndDate))
    : null;
  return direct?.hours ?? null;
}

function derivedHoursForRevenue(row, project) {
  if (!project.fundUsd || !project.eacHours) return null;
  const revenue = toNullableNumber(row.revenue);
  if (revenue === null) return null;
  return (revenue / project.fundUsd) * project.eacHours;
}

function monthPeriodKey(date) {
  if (!date) return "";
  return `${date.getFullYear()}-${date.getMonth()}`;
}

function applyPersonPicturesToProjects(projects, personPictures) {
  if (!personPictures?.size) return projects;
  return projects.map((project) => ({
    ...project,
    projectOwnerImage: personPictures.get(normalizePersonName(project.projectOwner)) || project.projectOwnerImage,
    portfolioLeadImage: personPictures.get(normalizePersonName(project.portfolioLead)) || project.portfolioLeadImage
  }));
}

function applyMilestonesToProjects(projects, milestonesByProject) {
  if (!milestonesByProject?.size) return projects;
  return projects.map((project) => {
    const direct = milestonesByProject.get(normalizeProjectKey(project.name));
    const fuzzy = direct || Array.from(milestonesByProject.entries()).find(([key]) => sameProjectName(project.name, key))?.[1];
    const fallback = projects.length === 1 && milestonesByProject.size === 1 ? Array.from(milestonesByProject.values())[0] : null;
    return { ...project, milestones: fuzzy || fallback || project.milestones || [] };
  });
}

function applyUnbilledToProjects(projects, unbilledByProject) {
  if (!unbilledByProject?.size) return projects;
  return projects.map((project) => {
    const direct = unbilledByProject.get(normalizeProjectKey(project.name));
    const fuzzy = direct || Array.from(unbilledByProject.entries()).find(([key]) => key !== "__fallback__" && sameProjectName(project.name, key))?.[1];
    const fallback = unbilledByProject.get("__fallback__") || (projects.length === 1 && unbilledByProject.size === 1 ? Array.from(unbilledByProject.values())[0] : null);
    return { ...project, unbilledRows: fuzzy || fallback || project.unbilledRows || [] };
  });
}

function projectPersonValue(sheet, header, projectRow, columnKey, label) {
  if (header.columns[columnKey]) {
    const value = getCellValue(sheet, projectRow, header.columns[columnKey]);
    if (cleanText(value)) return value;
  }
  return valueBelowOf(sheet, label);
}

function valueBelowOf(sheet, label) {
  const normalized = normalizeText(label);
  for (let row = 1; row <= sheet.maxRow; row += 1) {
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === normalized) {
        return getCellValue(sheet, row + 1, col);
      }
    }
  }
  return "";
}

function sameProjectName(left, right) {
  const a = normalizeProjectKey(left);
  const b = normalizeProjectKey(right);
  return Boolean(a && b && (a === b || a.includes(b) || b.includes(a)));
}

function normalizeProjectKey(value) {
  return normalizeText(value).replace(/\s+/g, " ");
}

function normalizePersonName(value) {
  return normalizeText(value).replace(/\s+/g, " ");
}

function hasLabel(sheet, label) {
  const wanted = normalizeText(label);
  for (let row = 1; row <= sheet.maxRow; row += 1) {
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === wanted) return true;
    }
  }
  return false;
}

function extractEstimatorProjects(sheet, sourceName) {
  const labelCells = findLabelCells(sheet);
  const hasRevenueForecast = labelCells.some((cell) => normalizeText(cell.value) === "revenue forecast");
  if (!hasRevenueForecast) return [];

  const base = {
    sourceName,
    name: valueRightOf(sheet, "Project Name"),
    number: valueRightOf(sheet, "Project #") || valueRightOf(sheet, "Project Number"),
    contractNumber: valueRightOf(sheet, "Contract Number"),
    fundUsd: valueRightOf(sheet, "Project Fund ($)"),
    fundAed: valueRightOf(sheet, "Project Fund (AED)"),
    hoursWithoutContingency: valueRightOf(sheet, "Total Hours (Without Contingency)"),
    hoursWithContingency: valueRightOf(sheet, "Total Hours (With Contingency)")
  };

  const revenueRows = labelCells.filter((cell) => normalizeText(cell.value) === "revenue");
  if (!revenueRows.length) return [];
  const sections = revenueRows
    .map((cell) => parseEstimatorSection(sheet, cell.row))
    .filter((section) => section.periods.length);
  if (!sections.length) return [];

  const quarterlySection = sections.find((section) => section.periods.every((period) => isQuarterLabel(period.label))) || sections[0];
  const displaySection = sections[sections.length - 1];
  const forecastRevenue = forecastRevenueFromEstimatorSections(sections);
  const chartQuarterlyRows = chartRowsFromEstimatorPeriods(quarterlySection.periods);
  const chartMonthlyRows = chartMonthlyRowsFromEstimatorSections(displaySection.periods, quarterlySection.periods);
  return [buildProjectFromEstimatorRows({ ...base, forecastRevenue, chartQuarterlyRows, chartMonthlyRows }, displaySection.periods, "Workbook estimator")];
}

function isQuarterLabel(label) {
  return /^Q[1-4]\s*[-_ ]?\s*(?:FY)?\d{2,4}$/i.test(cleanText(label));
}

function chartRowsFromEstimatorPeriods(periods) {
  let previousCumulativeRevenue = 0;
  return periods.map((period) => {
    const sourceRevenue = sourceEstimatorRevenue(period, previousCumulativeRevenue);
    previousCumulativeRevenue = sourceRevenue.cumulativeRevenue;
    return {
      label: period.label,
      revenue: sourceRevenue.revenue,
      cumulativeRevenue: sourceRevenue.cumulativeRevenue
    };
  });
}

function chartMonthlyRowsFromEstimatorSections(displayPeriods, quarterlyPeriods) {
  const monthlyRows = chartRowsFromEstimatorPeriods(displayPeriods);
  const monthlyKeys = new Set(Array.from(groupRowsByQuarter(monthlyRows).keys()));
  chartRowsFromEstimatorPeriods(quarterlyPeriods).forEach((quarterRow) => {
    const prepared = prepareChartRows([quarterRow])[0];
    if (!monthlyKeys.has(prepared.key)) {
      monthlyRows.push(quarterRow);
      monthlyKeys.add(prepared.key);
    }
  });
  return monthlyRows;
}

function parseEstimatorSection(sheet, revenueRow) {
  const headerRow = revenueRow - 1;
  const rowMap = {
    revenue: revenueRow,
    cumulativeRevenue: findRowWithLabel(sheet, "Cumulative Revenue", revenueRow),
    days: findRowWithLabel(sheet, "Days", revenueRow),
    hours: findRowWithLabel(sheet, "Hours", revenueRow),
    cumulativeDays: findRowWithLabel(sheet, "Cumulative Days", revenueRow),
    eacDays: findRowWithLabel(sheet, "EAC ( Days )", revenueRow),
    eacHours: findRowWithLabel(sheet, "EAC ( Hours )", revenueRow),
    poc: findRowWithLabel(sheet, "POC", revenueRow),
    funding: findRowWithLabel(sheet, "Funding", revenueRow)
  };

  const periods = [];
  for (let col = 3; col <= sheet.maxCol; col += 1) {
    const label = cleanText(getCellValue(sheet, headerRow, col));
    const revenue = toNullableNumber(getCellValue(sheet, rowMap.revenue, col));
    if (!label || revenue === null) continue;
    periods.push({
      label,
      revenue,
      cumulativeRevenue: getCellValue(sheet, rowMap.cumulativeRevenue, col),
      days: getCellValue(sheet, rowMap.days, col),
      hours: getCellValue(sheet, rowMap.hours, col),
      cumulativeDays: getCellValue(sheet, rowMap.cumulativeDays, col),
      eacDays: getCellValue(sheet, rowMap.eacDays, col),
      eacHours: getCellValue(sheet, rowMap.eacHours, col),
      poc: getCellValue(sheet, rowMap.poc, col),
      funding: getCellValue(sheet, rowMap.funding, col)
    });
  }

  return { revenueRow, headerRow, periods };
}

function forecastRevenueFromEstimatorSections(sections) {
  const currentQuarter = currentFiscalQuarterInfo(new Date());
  const candidates = sections
    .map((section) => {
      let previousEndDate = null;
      let latestDate = null;
      let latestCumulative = null;
      let revenueTotal = 0;
      let previousCumulativeRevenue = 0;

      section.periods.forEach((period) => {
        const periodEndDate = inferPeriodEndDate(period.label, previousEndDate, currentQuarter);
        if (periodEndDate) {
          previousEndDate = periodEndDate;
          latestDate = periodEndDate;
        }

        const sourceRevenue = sourceEstimatorRevenue(period, previousCumulativeRevenue);
        previousCumulativeRevenue = sourceRevenue.cumulativeRevenue;
        latestCumulative = sourceRevenue.cumulativeRevenue;
        revenueTotal += sourceRevenue.revenue;
      });

      return {
        date: latestDate,
        value: latestCumulative ?? revenueTotal,
        periodCount: section.periods.length
      };
    })
    .filter((candidate) => Number.isFinite(candidate.value));

  if (!candidates.length) return null;

  const datedCandidates = candidates.filter((candidate) => candidate.date);
  if (datedCandidates.length) {
    datedCandidates.sort((a, b) => a.date - b.date || a.value - b.value);
    return datedCandidates[datedCandidates.length - 1].value;
  }

  candidates.sort((a, b) => a.periodCount - b.periodCount || a.value - b.value);
  return candidates[candidates.length - 1].value;
}

function extractMasterProjects(sheet, sourceName) {
  const header = findProjectHeader(sheet);
  if (!header) return [];

  const projects = [];
  const periodHeaderRow = findPeriodHeaderRow(sheet, header.row);
  const periodStartCol = Math.max(...Object.values(header.columns)) + 1;

  for (let row = header.row + 1; row <= sheet.maxRow; row += 1) {
    const name = getCellValue(sheet, row, header.columns.name);
    const numberValue = getCellValue(sheet, row, header.columns.number);
    if (!cleanText(name) || !cleanText(numberValue)) continue;

    const projectCardValues = extractProjectCardValues(sheet, header, row);
    const base = {
      ...projectCardValues,
      sourceName,
      number: numberValue,
      contractNumber: getCellValue(sheet, row, header.columns.contractNumber),
      accountNumber: getCellValue(sheet, row, header.columns.accountNumber) || projectCardValues.accountNumber,
      name,
      customerLogoImage: getCellImage(sheet, row, header.columns.customerLogoImage) || projectCardValues.customerLogoImage,
      projectOwner: projectPersonValue(sheet, header, row, "projectOwner", "Project Owner"),
      portfolioLead: projectPersonValue(sheet, header, row, "portfolioLead", "Portfolio Lead"),
      fundUsd: getCellValue(sheet, row, header.columns.fundUsd),
      fundAed: getCellValue(sheet, row, header.columns.fundAed),
      hoursWithoutContingency: getCellValue(sheet, row, header.columns.hoursWithoutContingency),
      hoursWithContingency: getCellValue(sheet, row, header.columns.hoursWithContingency)
    };

    const rawPeriods = [];
    for (let col = periodStartCol; col <= sheet.maxCol; col += 1) {
      const value = toNullableNumber(getCellValue(sheet, row, col));
      if (value === null || value === 0) continue;
      const label = periodLabel(sheet, header.row, periodHeaderRow, col);
      rawPeriods.push({ label, value });
    }

    if (!rawPeriods.length) {
      projects.push(buildProjectFromRevenue(base, [], "No period values"));
      continue;
    }

    const mode = inferPeriodMode(base, rawPeriods.map((item) => item.value));
    if (mode === "hours") {
      projects.push(buildProjectFromHours(base, rawPeriods.map((item) => ({ label: item.label, hours: item.value })), "TAQA formula forecast"));
    } else {
      projects.push(buildProjectFromRevenue(base, rawPeriods.map((item) => ({ label: item.label, revenue: item.value })), "Imported revenue plan"));
    }
  }
  return projects;
}

function findProjectHeader(sheet) {
  const wanted = {
    projectOwner: ["project owner"],
    portfolioLead: ["portfolio lead"],
    number: ["project number", "project #"],
    contractNumber: ["contract number"],
    accountNumber: ["account number"],
    name: ["project name"],
    customerLogoImage: ["customer logo", "customer logo pic", "customer logo image", "logo"],
    fundUsd: ["project fund", "project fund $", "project fund usd"],
    fundAed: ["project fund aed"],
    hoursWithoutContingency: ["total hours without contingency"],
    hoursWithContingency: ["total hours with contingency"]
  };

  for (let row = 1; row <= Math.min(sheet.maxRow, 12); row += 1) {
    const columns = {};
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      const text = normalizeText(getCellValue(sheet, row, col));
      for (const [key, labels] of Object.entries(wanted)) {
        if (labels.includes(text)) columns[key] = col;
      }
    }
    if (columns.number && columns.name && columns.fundUsd) {
      return { row, columns };
    }
  }
  return null;
}

function findPeriodHeaderRow(sheet, headerRow) {
  let bestRow = headerRow;
  let bestCount = 0;
  for (let row = headerRow + 1; row <= Math.min(headerRow + 5, sheet.maxRow); row += 1) {
    let count = 0;
    for (let col = 8; col <= sheet.maxCol; col += 1) {
      const value = cleanText(getCellValue(sheet, row, col));
      if (value && /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|q\d/i.test(value)) count += 1;
    }
    if (count > bestCount) {
      bestCount = count;
      bestRow = row;
    }
  }
  return bestRow;
}

function periodLabel(sheet, headerRow, periodHeaderRow, col) {
  const month = cleanText(getCellValue(sheet, periodHeaderRow, col));
  const group = cleanText(getCellValue(sheet, headerRow, col));
  return month || group || columnName(col);
}

function inferPeriodMode(base, values) {
  const project = normalizeProjectBase(base);
  const eacHours = project.hoursWithContingency || project.hoursWithoutContingency;
  const total = sum(values);
  const maxValue = Math.max(...values.map(Math.abs));
  if (eacHours && total <= eacHours * 1.35 && maxValue <= eacHours) return "hours";
  return "revenue";
}

function findLabelCells(sheet) {
  const cells = [];
  for (let row = 1; row <= sheet.maxRow; row += 1) {
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      const value = getCellValue(sheet, row, col);
      if (cleanText(value)) cells.push({ row, col, value: cleanText(value) });
    }
  }
  return cells;
}

function valueRightOf(sheet, label) {
  const normalized = normalizeText(label);
  for (let row = 1; row <= sheet.maxRow; row += 1) {
    for (let col = 1; col <= sheet.maxCol; col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === normalized) {
        return getCellValue(sheet, row, col + 1);
      }
    }
  }
  return "";
}

function findRowWithLabel(sheet, label, minRow) {
  const normalized = normalizeText(label);
  for (let row = minRow; row <= sheet.maxRow; row += 1) {
    for (let col = 1; col <= Math.min(sheet.maxCol, 4); col += 1) {
      if (normalizeText(getCellValue(sheet, row, col)) === normalized) return row;
    }
  }
  return null;
}

async function parseXlsx(file) {
  const buffer = await file.arrayBuffer();
  const zip = await readZip(buffer);
  const workbookXml = await zip.text("xl/workbook.xml");
  const workbookRelsXml = await zip.text("xl/_rels/workbook.xml.rels");
  const sharedXml = zip.has("xl/sharedStrings.xml") ? await zip.text("xl/sharedStrings.xml") : "";
  const sharedStrings = sharedXml ? parseSharedStrings(sharedXml) : [];
  const richImages = await parseRichDataImages(zip);
  const workbookDoc = parseXml(workbookXml);
  const rels = parseRelationships(workbookRelsXml);

  const sheets = [];
  for (const sheetNode of Array.from(workbookDoc.getElementsByTagName("sheet"))) {
    const name = sheetNode.getAttribute("name") || "Sheet";
    const relId = sheetNode.getAttribute("r:id");
    const target = rels.get(relId);
    if (!target) continue;
    const path = normalizeXlPath(target);
    const sheetXml = await zip.text(path);
    sheets.push(parseSheet(sheetXml, name, sharedStrings, richImages));
  }

  return { sheets, richImages };
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error("Workbook XML could not be parsed.");
  return doc;
}

function parseRelationships(xml) {
  const doc = parseXml(xml);
  const rels = new Map();
  for (const rel of Array.from(doc.getElementsByTagName("Relationship"))) {
    rels.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
  }
  return rels;
}

function normalizeXlPath(target) {
  const cleaned = target.replace(/^\/+/, "");
  if (cleaned.startsWith("xl/")) return normalizePackagePath("", cleaned);
  return normalizePackagePath("xl", cleaned);
}

function normalizePackagePath(baseDirectory, target) {
  const raw = String(target || "").replace(/\\/g, "/");
  const combined = raw.startsWith("/")
    ? raw.slice(1)
    : [baseDirectory.replace(/\/+$/, ""), raw].filter(Boolean).join("/");
  const parts = [];
  combined.split("/").forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") parts.pop();
    else parts.push(part);
  });
  return parts.join("/");
}

function parseSharedStrings(xml) {
  const doc = parseXml(xml);
  return Array.from(doc.getElementsByTagName("si")).map((item) => {
    return Array.from(item.getElementsByTagName("t"))
      .map((node) => node.textContent || "")
      .join("");
  });
}

async function parseRichDataImages(zip) {
  const required = [
    "xl/metadata.xml",
    "xl/richData/rdrichvalue.xml",
    "xl/richData/richValueRel.xml",
    "xl/richData/_rels/richValueRel.xml.rels"
  ];
  if (!required.every((name) => zip.has(name))) return new Map();

  const metadataDoc = parseXml(await zip.text("xl/metadata.xml"));
  const richValueDoc = parseXml(await zip.text("xl/richData/rdrichvalue.xml"));
  const relOrderDoc = parseXml(await zip.text("xl/richData/richValueRel.xml"));
  const relTargets = parseRelationships(await zip.text("xl/richData/_rels/richValueRel.xml.rels"));
  const relIds = Array.from(relOrderDoc.getElementsByTagName("rel")).map((rel) => rel.getAttribute("r:id"));
  const richValues = Array.from(richValueDoc.getElementsByTagName("rv")).map((rv) => {
    const firstValue = rv.getElementsByTagName("v")[0]?.textContent;
    return Number(firstValue);
  });

  const imagesByMetadata = new Map();
  const valueMetadata = metadataDoc.getElementsByTagName("valueMetadata")[0];
  if (!valueMetadata) return imagesByMetadata;

  const metadataBlocks = Array.from(valueMetadata.getElementsByTagName("bk"));
  for (let index = 0; index < metadataBlocks.length; index += 1) {
    const richValueIndex = Number(metadataBlocks[index].getElementsByTagName("rc")[0]?.getAttribute("v"));
    const relIndex = richValues[richValueIndex];
    const relId = relIds[relIndex];
    const target = relTargets.get(relId);
    if (!target) continue;
    const imagePath = normalizePackagePath("xl/richData", target);
    if (!zip.has(imagePath)) continue;
    const bytes = await zip.bytes(imagePath);
    imagesByMetadata.set(index + 1, `data:${mimeTypeForPath(imagePath)};base64,${bytesToBase64(bytes)}`);
  }

  return imagesByMetadata;
}

function mimeTypeForPath(path) {
  const lower = path.toLowerCase();
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".webp")) return "image/webp";
  return "image/png";
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

function parseSheet(xml, name, sharedStrings, richImages = new Map()) {
  const doc = parseXml(xml);
  const cells = new Map();
  let maxRow = 0;
  let maxCol = 0;

  for (const rowNode of Array.from(doc.getElementsByTagName("row"))) {
    const rowNumber = Number(rowNode.getAttribute("r"));
    maxRow = Math.max(maxRow, rowNumber);
    for (const cellNode of Array.from(rowNode.getElementsByTagName("c"))) {
      const ref = cellNode.getAttribute("r");
      const { row, col } = parseCellRef(ref);
      const type = cellNode.getAttribute("t");
      const valueMetadata = Number(cellNode.getAttribute("vm"));
      const valueNode = cellNode.getElementsByTagName("v")[0];
      const formulaNode = cellNode.getElementsByTagName("f")[0];
      let value = valueNode ? valueNode.textContent : "";
      if (type === "s") value = sharedStrings[Number(value)] || "";
      if (type === "inlineStr") {
        value = Array.from(cellNode.getElementsByTagName("t")).map((node) => node.textContent || "").join("");
      }
      const parsed = parseCellPrimitive(value, type);
      cells.set(key(row, col), {
        ref,
        row,
        col,
        value: parsed,
        formula: formulaNode ? formulaNode.textContent : "",
        image: Number.isFinite(valueMetadata) ? richImages.get(valueMetadata) || "" : ""
      });
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
  }

  return { name, cells, maxRow, maxCol };
}

async function readZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const eocd = findEndOfCentralDirectory(view);
  const entryCount = view.getUint16(eocd + 10, true);
  const centralDirOffset = view.getUint32(eocd + 16, true);
  const entries = new Map();
  let offset = centralDirOffset;

  for (let index = 0; index < entryCount; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("Invalid XLSX central directory.");
    }
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const name = decodeBytes(bytes.slice(offset + 46, offset + 46 + nameLength));
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const data = bytes.slice(dataStart, dataStart + compressedSize);
    entries.set(name, { name, method, data });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return {
    has(name) {
      return entries.has(name);
    },
    async bytes(name) {
      const entry = entries.get(name);
      if (!entry) throw new Error(`${name} was not found in workbook.`);
      return entry.method === 0 ? entry.data : await inflateRaw(entry.data);
    },
    async text(name) {
      const output = await this.bytes(name);
      return decodeBytes(output);
    }
  };
}

function findEndOfCentralDirectory(view) {
  const min = Math.max(0, view.byteLength - 66000);
  for (let offset = view.byteLength - 22; offset >= min; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("Invalid XLSX file.");
}

async function inflateRaw(data) {
  if ("DecompressionStream" in window) {
    for (const format of ["deflate-raw", "deflate"]) {
      try {
        const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream(format));
        return new Uint8Array(await new Response(stream).arrayBuffer());
      } catch (error) {
        if (format === "deflate") break;
      }
    }
  }
  return inflateRawFallback(data);
}

function decodeBytes(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

function inflateRawFallback(input) {
  const reader = createBitReader(input);
  const output = createInflateOutput();
  let isFinal = false;

  while (!isFinal) {
    isFinal = reader.readBits(1) === 1;
    const blockType = reader.readBits(2);

    if (blockType === 0) {
      reader.alignByte();
      const length = reader.readBits(16);
      const check = reader.readBits(16);
      if (((length ^ 0xffff) & 0xffff) !== check) {
        throw new Error("Invalid XLSX deflate block.");
      }
      for (let index = 0; index < length; index += 1) {
        output.push(reader.readBits(8));
      }
    } else if (blockType === 1) {
      decodeCompressedBlock(reader, output, fixedHuffmanTables());
    } else if (blockType === 2) {
      decodeCompressedBlock(reader, output, dynamicHuffmanTables(reader));
    } else {
      throw new Error("Unsupported XLSX deflate block.");
    }
  }

  return output.bytes();
}

function createBitReader(bytes) {
  let bitOffset = 0;
  return {
    readBits(count) {
      let value = 0;
      for (let bit = 0; bit < count; bit += 1) {
        const byte = bytes[bitOffset >> 3];
        if (byte === undefined) throw new Error("Unexpected end of XLSX data.");
        value |= ((byte >> (bitOffset & 7)) & 1) << bit;
        bitOffset += 1;
      }
      return value;
    },
    alignByte() {
      bitOffset = Math.ceil(bitOffset / 8) * 8;
    }
  };
}

function createInflateOutput() {
  let buffer = new Uint8Array(32768);
  let length = 0;

  function ensure(extra) {
    if (length + extra <= buffer.length) return;
    let nextLength = buffer.length;
    while (length + extra > nextLength) nextLength *= 2;
    const next = new Uint8Array(nextLength);
    next.set(buffer);
    buffer = next;
  }

  return {
    push(value) {
      ensure(1);
      buffer[length] = value & 255;
      length += 1;
    },
    copy(distance, count) {
      if (distance <= 0 || distance > length) throw new Error("Invalid XLSX deflate distance.");
      ensure(count);
      for (let index = 0; index < count; index += 1) {
        buffer[length] = buffer[length - distance];
        length += 1;
      }
    },
    bytes() {
      return buffer.slice(0, length);
    }
  };
}

function decodeCompressedBlock(reader, output, tables) {
  while (true) {
    const symbol = decodeSymbol(reader, tables.literalLength);
    if (symbol < 256) {
      output.push(symbol);
    } else if (symbol === 256) {
      return;
    } else {
      const lengthIndex = symbol - 257;
      const baseLength = LENGTH_BASE[lengthIndex];
      const extraLengthBits = LENGTH_EXTRA[lengthIndex];
      if (baseLength === undefined) throw new Error("Invalid XLSX deflate length.");
      const copyLength = baseLength + (extraLengthBits ? reader.readBits(extraLengthBits) : 0);
      const distanceSymbol = decodeSymbol(reader, tables.distance);
      const baseDistance = DISTANCE_BASE[distanceSymbol];
      const extraDistanceBits = DISTANCE_EXTRA[distanceSymbol];
      if (baseDistance === undefined) throw new Error("Invalid XLSX deflate distance.");
      const distance = baseDistance + (extraDistanceBits ? reader.readBits(extraDistanceBits) : 0);
      output.copy(distance, copyLength);
    }
  }
}

function fixedHuffmanTables() {
  const literalLengths = new Array(288).fill(0);
  for (let index = 0; index <= 143; index += 1) literalLengths[index] = 8;
  for (let index = 144; index <= 255; index += 1) literalLengths[index] = 9;
  for (let index = 256; index <= 279; index += 1) literalLengths[index] = 7;
  for (let index = 280; index <= 287; index += 1) literalLengths[index] = 8;
  return {
    literalLength: buildHuffman(literalLengths),
    distance: buildHuffman(new Array(32).fill(5))
  };
}

function dynamicHuffmanTables(reader) {
  const literalLengthCount = reader.readBits(5) + 257;
  const distanceCount = reader.readBits(5) + 1;
  const codeLengthCount = reader.readBits(4) + 4;
  const order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
  const codeLengthLengths = new Array(19).fill(0);

  for (let index = 0; index < codeLengthCount; index += 1) {
    codeLengthLengths[order[index]] = reader.readBits(3);
  }

  const codeLengthTable = buildHuffman(codeLengthLengths);
  const lengths = [];
  const total = literalLengthCount + distanceCount;

  while (lengths.length < total) {
    const symbol = decodeSymbol(reader, codeLengthTable);
    if (symbol <= 15) {
      lengths.push(symbol);
    } else if (symbol === 16) {
      const previous = lengths[lengths.length - 1] || 0;
      const repeat = reader.readBits(2) + 3;
      for (let index = 0; index < repeat; index += 1) lengths.push(previous);
    } else if (symbol === 17) {
      const repeat = reader.readBits(3) + 3;
      for (let index = 0; index < repeat; index += 1) lengths.push(0);
    } else if (symbol === 18) {
      const repeat = reader.readBits(7) + 11;
      for (let index = 0; index < repeat; index += 1) lengths.push(0);
    }
  }

  return {
    literalLength: buildHuffman(lengths.slice(0, literalLengthCount)),
    distance: buildHuffman(lengths.slice(literalLengthCount, total))
  };
}

function buildHuffman(lengths) {
  const maxLength = Math.max(...lengths, 0);
  const counts = new Array(maxLength + 1).fill(0);
  lengths.forEach((length) => {
    if (length) counts[length] += 1;
  });

  let code = 0;
  const nextCodes = new Array(maxLength + 1).fill(0);
  for (let bits = 1; bits <= maxLength; bits += 1) {
    code = (code + counts[bits - 1]) << 1;
    nextCodes[bits] = code;
  }

  const tables = Array.from({ length: maxLength + 1 }, () => new Map());
  lengths.forEach((length, symbol) => {
    if (!length) return;
    const canonical = nextCodes[length];
    nextCodes[length] += 1;
    tables[length].set(reverseBits(canonical, length), symbol);
  });

  return { maxLength, tables };
}

function decodeSymbol(reader, huffman) {
  let code = 0;
  for (let length = 1; length <= huffman.maxLength; length += 1) {
    code |= reader.readBits(1) << (length - 1);
    const symbol = huffman.tables[length].get(code);
    if (symbol !== undefined) return symbol;
  }
  throw new Error("Invalid XLSX deflate code.");
}

function reverseBits(value, length) {
  let reversed = 0;
  for (let index = 0; index < length; index += 1) {
    reversed = (reversed << 1) | (value & 1);
    value >>= 1;
  }
  return reversed;
}

const LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DISTANCE_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DISTANCE_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

function parseCellPrimitive(value, type) {
  if (type === "s" || type === "str" || type === "inlineStr") return value;
  if (value === "") return "";
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : value;
}

function getCellValue(sheet, row, col) {
  if (!row || !col) return "";
  return sheet.cells.get(key(row, col))?.value ?? "";
}

function getCellImage(sheet, row, col) {
  if (!row || !col) return "";
  return sheet.cells.get(key(row, col))?.image ?? "";
}

function parseCellRef(ref) {
  const match = String(ref || "").match(/^([A-Z]+)(\d+)$/);
  if (!match) return { row: 0, col: 0 };
  return { col: columnIndex(match[1]), row: Number(match[2]) };
}

function columnIndex(letters) {
  return letters.split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function columnName(index) {
  let name = "";
  let value = index;
  while (value > 0) {
    const mod = (value - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function key(row, col) {
  return `${row}:${col}`;
}

function toNumber(value) {
  const parsed = toNullableNumber(value);
  return parsed === null ? 0 : parsed;
}

function toNullableNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const cleaned = value.replace(/,/g, "").trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function isFiniteNumber(value) {
  return toNullableNumber(value) !== null;
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeText(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/\([^)]*\)/g, (match) => match.replace(/[$#]/g, ""))
    .replace(/[$#]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function sum(values) {
  return values.reduce((total, value) => total + (toNumber(value) || 0), 0);
}

function money(value, currency = "USD") {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency,
    maximumFractionDigits: 0
  }).format(value);
}

function number(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 }).format(value);
}

function optionalNumber(value) {
  return Number.isFinite(value) ? number(value) : "-";
}

function percent(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("en-US", {
    style: "percent",
    maximumFractionDigits: 1
  }).format(value);
}

function signedPercent(value) {
  if (!Number.isFinite(value)) return "-";
  const output = percent(value);
  return value > 0 ? `+${output}` : output;
}

function remainingTone(value) {
  if (!Number.isFinite(value)) return "-";
  if (value < 0) return "Over funding";
  if (value === 0) return "Fully consumed";
  return "Available balance";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
