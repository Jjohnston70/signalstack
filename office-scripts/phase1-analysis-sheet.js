// Phase 1: Create Analysis sheet and build header + summary stats
const sheets = context.workbook.worksheets;
sheets.load("items/name");
await context.sync();

// Check if Analysis sheet already exists
const existing = sheets.items.find(s => s.name === "Analysis");
if (existing) {
  existing.delete();
  await context.sync();
}

const ws = sheets.add("Analysis");
ws.activate();

// ===== TITLE BLOCK =====
ws.getRange("A1").values = [["TNDS SALES DATA — COMPREHENSIVE ANALYSIS"]];
const titleRange = ws.getRange("A1:L1");
titleRange.merge();
titleRange.format.fill.color = "#1F4E79";
titleRange.format.font.color = "#FFFFFF";
titleRange.format.font.bold = true;
titleRange.format.font.size = 14;
titleRange.format.horizontalAlignment = "Center";
titleRange.format.rowHeight = 30;

ws.getRange("A2").values = [["Source: RAW_INPUT sheet | All figures in USD unless noted"]];
const subtitleRange = ws.getRange("A2:L2");
subtitleRange.merge();
subtitleRange.format.font.size = 10;
subtitleRange.format.font.italic = true;
subtitleRange.format.font.color = "#666666";
subtitleRange.format.horizontalAlignment = "Center";

// ===== SECTION 1: KEY METRICS SUMMARY (Row 4-8) =====
ws.getRange("A4").values = [["KEY METRICS SUMMARY"]];
const sec1Header = ws.getRange("A4:L4");
sec1Header.merge();
sec1Header.format.fill.color = "#1F4E79";
sec1Header.format.font.color = "#FFFFFF";
sec1Header.format.font.bold = true;
sec1Header.format.font.size = 11;
sec1Header.format.rowHeight = 25;

// Summary labels and formulas
ws.getRange("A6:A8").values = [["Total Revenue"], ["Total Transactions"], ["Total Fuel Gallons"]];
ws.getRange("B6:B8").formulas = [
  ["=SUM(RAW_INPUT!G5:G445)"],
  ["=COUNTA(RAW_INPUT!A5:A445)"],
  ["=SUMPRODUCT((RAW_INPUT!F5:F445<20)*(RAW_INPUT!E5:E445))"]
];

ws.getRange("D6:D8").values = [["Avg Sale Size"], ["Avg Fuel Price/Gal"], ["Service Transactions"]];
ws.getRange("E6:E8").formulas = [
  ["=B6/B7"],
  ["=SUMPRODUCT((RAW_INPUT!F5:F445<20)*(RAW_INPUT!G5:G445))/SUMPRODUCT((RAW_INPUT!F5:F445<20)*(RAW_INPUT!E5:E445))"],
  ["=SUMPRODUCT((RAW_INPUT!E5:E445=1)*(RAW_INPUT!F5:F445>=20)*1)"]
];

ws.getRange("G6:G8").values = [["Date Range Start"], ["Date Range End"], ["Unique Customers"]];
ws.getRange("H6:H8").formulas = [
  ["=MIN(RAW_INPUT!A5:A445)"],
  ["=MAX(RAW_INPUT!A5:A445)"],
  ["=SUMPRODUCT(1/COUNTIF(RAW_INPUT!B5:B445,RAW_INPUT!B5:B445))"]
];

// Format summary values
ws.getRange("B6").numberFormat = [["$#,##0"]];
ws.getRange("B7").numberFormat = [["#,##0"]];
ws.getRange("B8").numberFormat = [["#,##0"]];
ws.getRange("E6").numberFormat = [["$#,##0"]];
ws.getRange("E7").numberFormat = [["$#,##0.00"]];
ws.getRange("E8").numberFormat = [["#,##0"]];
ws.getRange("H6:H7").numberFormat = [["mm/dd/yyyy"], ["mm/dd/yyyy"]];
ws.getRange("H8").numberFormat = [["#,##0"]];

// Bold labels
ws.getRange("A6:A8").format.font.bold = true;
ws.getRange("D6:D8").format.font.bold = true;
ws.getRange("G6:G8").format.font.bold = true;

// Light background for summary area
ws.getRange("A5:L8").format.fill.color = "#F2F7FB";

await context.sync();

// Verify key metrics
const checkVals = ws.getRange("B6:B8");
checkVals.load("values");
const checkVals2 = ws.getRange("E6:E8");
checkVals2.load("values");
const checkVals3 = ws.getRange("H6:H8");
checkVals3.load("values");
await context.sync();

return {
  phase: "Phase 1 - Key Metrics",
  totalRevenue: checkVals.values[0][0],
  totalTransactions: checkVals.values[1][0],
  totalGallons: checkVals.values[2][0],
  avgSaleSize: checkVals2.values[0][0],
  avgFuelPrice: checkVals2.values[1][0],
  serviceTransactions: checkVals2.values[2][0],
  dateStart: checkVals3.values[0][0],
  dateEnd: checkVals3.values[1][0],
  uniqueCustomers: checkVals3.values[2][0]
};

---
result

{
  "success": true,
  "result": {
    "phase": "Phase 1 - Key Metrics",
    "totalRevenue": 790660.08,
    "totalTransactions": 441,
    "totalGallons": 181158,
    "avgSaleSize": 1792.88,
    "avgFuelPrice": 3.69867253999271,
    "serviceTransactions": 72,
    "dateStart": 45600,
    "dateEnd": 46224,
    "uniqueCustomers": 15
  }

---

// Phase 2: Monthly Revenue Trend (Row 10+)
const ws = context.workbook.worksheets.getItem("Analysis");

// ===== SECTION 2: MONTHLY REVENUE TREND =====
ws.getRange("A10").values = [["MONTHLY REVENUE TREND"]];
const sec2Header = ws.getRange("A10:L10");
sec2Header.merge();
sec2Header.format.fill.color = "#1F4E79";
sec2Header.format.font.color = "#FFFFFF";
sec2Header.format.font.bold = true;
sec2Header.format.font.size = 11;
sec2Header.format.rowHeight = 25;

// Column headers
ws.getRange("A11:F11").values = [["Month", "Revenue", "# Transactions", "Fuel Gallons", "Avg Price/Gal", "Avg Sale Size"]];
const colHeaders = ws.getRange("A11:F11");
colHeaders.format.fill.color = "#D9E2F3";
colHeaders.format.font.bold = true;
colHeaders.format.horizontalAlignment = "Center";

// We need to figure out the months. Data spans serial 45600 to 46224.
// 45600 = approx Nov 2024. Let's use EOMONTH-based approach.
// We'll create 21 month rows (Nov 2024 through ~Jul 2026) and use SUMPRODUCT

// First, let's determine start/end months using a helper approach
// Row 12 onwards: month start dates, then SUMPRODUCT formulas

// Put month start dates using formula
ws.getRange("A12").formulas = [["=DATE(YEAR(MIN(RAW_INPUT!A5:A445)),MONTH(MIN(RAW_INPUT!A5:A445)),1)"]];
ws.getRange("A12").numberFormat = [["mmm yyyy"]];

// Fill next months - we'll do 21 rows to be safe
for (let i = 1; i <= 20; i++) {
  ws.getRange(`A${12+i}`).formulas = [[`=EDATE(A12,${i})`]];
  ws.getRange(`A${12+i}`).numberFormat = [["mmm yyyy"]];
}

// Revenue per month: SUMPRODUCT where date >= month start and < next month start
for (let i = 0; i <= 20; i++) {
  const row = 12 + i;
  const nextMonthRef = i < 20 ? `A${row+1}` : `EDATE(A${row},1)`;
  
  // Revenue
  ws.getRange(`B${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=A${row})*(RAW_INPUT!A5:A445<${nextMonthRef})*RAW_INPUT!G5:G445)`
  ]];
  
  // Transaction count
  ws.getRange(`C${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=A${row})*(RAW_INPUT!A5:A445<${nextMonthRef})*1)`
  ]];
  
  // Fuel gallons (where price < 20, meaning fuel not service)
  ws.getRange(`D${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=A${row})*(RAW_INPUT!A5:A445<${nextMonthRef})*(RAW_INPUT!F5:F445<20)*RAW_INPUT!E5:E445)`
  ]];
  
  // Avg price/gal
  ws.getRange(`E${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=A${row})*(RAW_INPUT!A5:A445<${nextMonthRef})*(RAW_INPUT!F5:F445<20)*RAW_INPUT!G5:G445)/D${row},0)`
  ]];
  
  // Avg sale size
  ws.getRange(`F${row}`).formulas = [[
    `=IFERROR(B${row}/C${row},0)`
  ]];
}

// Format columns
ws.getRange("B12:B32").numberFormat = [["$#,##0"]].concat(Array(20).fill(["$#,##0"]));
ws.getRange("C12:C32").numberFormat = [["#,##0"]].concat(Array(20).fill(["#,##0"]));
ws.getRange("D12:D32").numberFormat = [["#,##0"]].concat(Array(20).fill(["#,##0"]));
ws.getRange("E12:E32").numberFormat = [["$#,##0.00"]].concat(Array(20).fill(["$#,##0.00"]));
ws.getRange("F12:F32").numberFormat = [["$#,##0"]].concat(Array(20).fill(["$#,##0"]));

// Center-align data columns
ws.getRange("B11:F32").format.horizontalAlignment = "Center";

await context.sync();

// Check which rows actually have data (non-zero revenue) to know the active range
const revCheck = ws.getRange("B12:B32");
revCheck.load("values");
await context.sync();

const activeMonths = revCheck.values.filter(r => r[0] > 0).length;

return { 
  phase: "Phase 2 - Monthly Trend",
  activeMonths,
  note: "Monthly trend built with SUMPRODUCT formulas referencing RAW_INPUT"
};

---
result

{
  "success": true,
  "result": {
    "phase": "Phase 2 - Monthly Trend",
    "activeMonths": 21,
    "note": "Monthly trend built with SUMPRODUCT formulas referencing RAW_INPUT"
  }
}
---
// Phase 3: Customer Analysis (fixed - no INDIRECT)
const ws = context.workbook.worksheets.getItem("Analysis");

const rawSheet = context.workbook.worksheets.getItem("RAW_INPUT");
const custRange = rawSheet.getRange("B5:B445");
custRange.load("values");
await context.sync();

const allCustomers = custRange.values.map(r => r[0]).filter(v => v && v !== "");
const uniqueCustomers = [...new Set(allCustomers)].sort();

const startRow = 35;

// ===== SECTION 3: CUSTOMER ANALYSIS =====
ws.getRange(`A${startRow}`).values = [["CUSTOMER ANALYSIS"]];
const sec3Header = ws.getRange(`A${startRow}:L${startRow}`);
sec3Header.merge();
sec3Header.format.fill.color = "#1F4E79";
sec3Header.format.font.color = "#FFFFFF";
sec3Header.format.font.bold = true;
sec3Header.format.font.size = 11;
sec3Header.format.rowHeight = 25;

const hdrRow = startRow + 1;
ws.getRange(`A${hdrRow}:E${hdrRow}`).values = [["Customer", "Revenue", "% of Total", "# Transactions", "Avg Sale"]];
const custColHeaders = ws.getRange(`A${hdrRow}:E${hdrRow}`);
custColHeaders.format.fill.color = "#D9E2F3";
custColHeaders.format.font.bold = true;
custColHeaders.format.horizontalAlignment = "Center";

const dataStartRow = hdrRow + 1;
for (let i = 0; i < uniqueCustomers.length; i++) {
  const row = dataStartRow + i;
  const cust = uniqueCustomers[i];
  
  ws.getRange(`A${row}`).values = [[cust]];
  ws.getRange(`B${row}`).formulas = [[`=SUMIF(RAW_INPUT!B5:B445,A${row},RAW_INPUT!G5:G445)`]];
  ws.getRange(`C${row}`).formulas = [[`=B${row}/B$6`]];
  ws.getRange(`D${row}`).formulas = [[`=COUNTIF(RAW_INPUT!B5:B445,A${row})`]];
  ws.getRange(`E${row}`).formulas = [[`=IFERROR(B${row}/D${row},0)`]];
}

const dataEndRow = dataStartRow + uniqueCustomers.length - 1;

// Format
ws.getRange(`B${dataStartRow}:B${dataEndRow}`).numberFormat = Array(uniqueCustomers.length).fill(["$#,##0"]);
ws.getRange(`C${dataStartRow}:C${dataEndRow}`).numberFormat = Array(uniqueCustomers.length).fill(["0.0%"]);
ws.getRange(`D${dataStartRow}:D${dataEndRow}`).numberFormat = Array(uniqueCustomers.length).fill(["#,##0"]);
ws.getRange(`E${dataStartRow}:E${dataEndRow}`).numberFormat = Array(uniqueCustomers.length).fill(["$#,##0"]);
ws.getRange(`B${dataStartRow}:E${dataEndRow}`).format.horizontalAlignment = "Center";

// Summary stats
const statsRow = dataEndRow + 2;
ws.getRange(`A${statsRow}:A${statsRow + 2}`).values = [["Top 5 Revenue Share"], ["Median Revenue/Customer"], ["Customer Count"]];
ws.getRange(`A${statsRow}:A${statsRow + 2}`).format.font.bold = true;
ws.getRange(`A${statsRow}:E${statsRow + 2}`).format.fill.color = "#F2F2F2";

// Top 5 share: LARGE(1)+LARGE(2)+...LARGE(5) / total
ws.getRange(`B${statsRow}`).formulas = [[
  `=(LARGE(B${dataStartRow}:B${dataEndRow},1)+LARGE(B${dataStartRow}:B${dataEndRow},2)+LARGE(B${dataStartRow}:B${dataEndRow},3)+LARGE(B${dataStartRow}:B${dataEndRow},4)+LARGE(B${dataStartRow}:B${dataEndRow},5))/B6`
]];
ws.getRange(`B${statsRow}`).numberFormat = [["0.0%"]];

ws.getRange(`B${statsRow + 1}`).formulas = [[`=MEDIAN(B${dataStartRow}:B${dataEndRow})`]];
ws.getRange(`B${statsRow + 1}`).numberFormat = [["$#,##0"]];

ws.getRange(`B${statsRow + 2}`).formulas = [[`=COUNTA(A${dataStartRow}:A${dataEndRow})`]];

await context.sync();

return {
  phase: "Phase 3 - Customer Analysis",
  customerCount: uniqueCustomers.length,
  dataRange: `A${dataStartRow}:E${dataEndRow}`,
  nextAvailableRow: statsRow + 4
};

---

result
{
  "success": true,
  "result": {
    "phase": "Phase 3 - Customer Analysis",
    "customerCount": 15,
    "dataRange": "A37:E51",
    "nextAvailableRow": 57
  }
}

---

// Phase 4: Product/Service Breakdown
const ws = context.workbook.worksheets.getItem("Analysis");

// Get unique item descriptions
const rawSheet = context.workbook.worksheets.getItem("RAW_INPUT");
const itemRange = rawSheet.getRange("D5:D445");
itemRange.load("values");
await context.sync();

const allItems = itemRange.values.map(r => r[0]).filter(v => v && v !== "");
const uniqueItems = [...new Set(allItems)].sort();

const startRow = 57;

// ===== SECTION 4: PRODUCT/SERVICE BREAKDOWN =====
ws.getRange(`A${startRow}`).values = [["PRODUCT & SERVICE BREAKDOWN"]];
const sec4Header = ws.getRange(`A${startRow}:L${startRow}`);
sec4Header.merge();
sec4Header.format.fill.color = "#1F4E79";
sec4Header.format.font.color = "#FFFFFF";
sec4Header.format.font.bold = true;
sec4Header.format.font.size = 11;
sec4Header.format.rowHeight = 25;

const hdrRow = startRow + 1;
ws.getRange(`A${hdrRow}:F${hdrRow}`).values = [["Product / Service", "Revenue", "% of Total", "# Transactions", "Avg Qty", "Avg Price"]];
const prodColHeaders = ws.getRange(`A${hdrRow}:F${hdrRow}`);
prodColHeaders.format.fill.color = "#D9E2F3";
prodColHeaders.format.font.bold = true;
prodColHeaders.format.horizontalAlignment = "Center";

const dataStartRow = hdrRow + 1;
for (let i = 0; i < uniqueItems.length; i++) {
  const row = dataStartRow + i;
  const item = uniqueItems[i];
  
  ws.getRange(`A${row}`).values = [[item]];
  
  // Revenue
  ws.getRange(`B${row}`).formulas = [[`=SUMIF(RAW_INPUT!D5:D445,A${row},RAW_INPUT!G5:G445)`]];
  
  // % of Total
  ws.getRange(`C${row}`).formulas = [[`=B${row}/B$6`]];
  
  // Transaction count
  ws.getRange(`D${row}`).formulas = [[`=COUNTIF(RAW_INPUT!D5:D445,A${row})`]];
  
  // Average Qty
  ws.getRange(`E${row}`).formulas = [[`=IFERROR(SUMIF(RAW_INPUT!D5:D445,A${row},RAW_INPUT!E5:E445)/D${row},0)`]];
  
  // Average Price
  ws.getRange(`F${row}`).formulas = [[`=IFERROR(SUMIF(RAW_INPUT!D5:D445,A${row},RAW_INPUT!F5:F445)/D${row},0)`]];
}

const dataEndRow = dataStartRow + uniqueItems.length - 1;

// Format
ws.getRange(`B${dataStartRow}:B${dataEndRow}`).numberFormat = Array(uniqueItems.length).fill(["$#,##0"]);
ws.getRange(`C${dataStartRow}:C${dataEndRow}`).numberFormat = Array(uniqueItems.length).fill(["0.0%"]);
ws.getRange(`D${dataStartRow}:D${dataEndRow}`).numberFormat = Array(uniqueItems.length).fill(["#,##0"]);
ws.getRange(`E${dataStartRow}:E${dataEndRow}`).numberFormat = Array(uniqueItems.length).fill(["#,##0"]);
ws.getRange(`F${dataStartRow}:F${dataEndRow}`).numberFormat = Array(uniqueItems.length).fill(["$#,##0.00"]);
ws.getRange(`B${dataStartRow}:F${dataEndRow}`).format.horizontalAlignment = "Center";

// Category summary (Fuel vs Services vs Specialty)
const catRow = dataEndRow + 2;
ws.getRange(`A${catRow}`).values = [["CATEGORY SUMMARY"]];
ws.getRange(`A${catRow}:F${catRow}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${catRow}`).format.font.bold = true;

// Fuel deliveries: items containing "Delivery", "Fuel Drop", "Bulk Diesel"
// Services: Line Flush, Pump Maintenance, Tank Inspection, Filter Service
// Specialty: Motor Oil, Lubricant, Synthetic Blend

ws.getRange(`A${catRow+1}:A${catRow+3}`).values = [["Fuel Deliveries"], ["Services"], ["Specialty Products"]];
ws.getRange(`A${catRow+1}:A${catRow+3}`).format.font.bold = true;

// Fuel: SUMPRODUCT with OR conditions for fuel-related items
ws.getRange(`B${catRow+1}`).formulas = [[
  `=SUMPRODUCT((ISNUMBER(SEARCH("Delivery",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Fuel Drop",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Bulk Diesel",RAW_INPUT!D5:D445))>0)*RAW_INPUT!G5:G445)`
]];

// Services
ws.getRange(`B${catRow+2}`).formulas = [[
  `=SUMPRODUCT((ISNUMBER(SEARCH("Line Flush",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Pump Maintenance",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Tank Inspection",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Filter Service",RAW_INPUT!D5:D445))>0)*RAW_INPUT!G5:G445)`
]];

// Specialty
ws.getRange(`B${catRow+3}`).formulas = [[
  `=SUMPRODUCT((ISNUMBER(SEARCH("Motor Oil",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Lubricant",RAW_INPUT!D5:D445))+ISNUMBER(SEARCH("Synthetic",RAW_INPUT!D5:D445))>0)*RAW_INPUT!G5:G445)`
]];

// % of total for categories
ws.getRange(`C${catRow+1}:C${catRow+3}`).formulas = [
  [`=B${catRow+1}/B$6`],
  [`=B${catRow+2}/B$6`],
  [`=B${catRow+3}/B$6`]
];

ws.getRange(`B${catRow+1}:B${catRow+3}`).numberFormat = [["$#,##0"], ["$#,##0"], ["$#,##0"]];
ws.getRange(`C${catRow+1}:C${catRow+3}`).numberFormat = [["0.0%"], ["0.0%"], ["0.0%"]];
ws.getRange(`B${catRow+1}:C${catRow+3}`).format.horizontalAlignment = "Center";
ws.getRange(`A${catRow+1}:F${catRow+3}`).format.fill.color = "#F2F2F2";

await context.sync();

// Verify
const catCheck = ws.getRange(`B${catRow+1}:C${catRow+3}`);
catCheck.load("values");
await context.sync();

return {
  phase: "Phase 4 - Product/Service Breakdown",
  products: uniqueItems,
  productCount: uniqueItems.length,
  categories: {
    fuel: catCheck.values[0],
    services: catCheck.values[1],
    specialty: catCheck.values[2]
  },
  nextRow: catRow + 6
};

---

result

{
  "success": true,
  "result": {
    "phase": "Phase 4 - Product/Service Breakdown",
    "products": [
      "Bulk Diesel",
      "Diesel Delivery",
      "Filter Service",
      "Fleet Fuel Drop",
      "Line Flush Service",
      "Lubricant Package",
      "Motor Oil Case",
      "Premium Fuel Delivery",
      "Pump Maintenance",
      "Synthetic Blend Package",
      "Tank Inspection",
      "Unleaded Delivery"
    ],
    "productCount": 12,
    "categories": {
      "fuel": [
        670044.12,
        0.847449032712009
      ],
      "services": [
        87690.96,
        0.110908546185865
      ],
      "specialty": [
        32925,
        0.0416424211021252
      ]
    },
    "nextRow": 78
  }
}


---
// Phase 5: Set column widths and clean up formatting
const ws = context.workbook.worksheets.getItem("Analysis");

// Set uniform column widths
ws.getRange("A:A").format.columnWidth = 175;
ws.getRange("B:F").format.columnWidth = 110;
ws.getRange("G:G").format.columnWidth = 140;
ws.getRange("H:H").format.columnWidth = 110;

// Clean up: hide zero-value monthly rows (rows where revenue = 0)
// Actually, let's just leave them - they show the full timeline

// Add alternating row shading to customer data for readability
for (let i = 37; i <= 51; i++) {
  if ((i - 37) % 2 === 1) {
    ws.getRange(`A${i}:E${i}`).format.fill.color = "#F8F9FA";
  }
}

// Add alternating row shading to product data
for (let i = 59; i <= 70; i++) {
  if ((i - 59) % 2 === 1) {
    ws.getRange(`A${i}:F${i}`).format.fill.color = "#F8F9FA";
  }
}

// Bold the "Analysis" tab
ws.tabColor = "#1F4E79";

await context.sync();
return { phase: "Phase 5 - Formatting complete" };

---

result

{
  "success": true,
  "result": {
    "phase": "Phase 5 - Formatting complete"
  }
}

---
// Final verification: check for formula errors and key values
const ws = context.workbook.worksheets.getItem("Analysis");

// Check key metrics
const metrics = ws.getRange("B6:B8");
metrics.load("values");
const metrics2 = ws.getRange("E6:E8");
metrics2.load("values");
const metrics3 = ws.getRange("H6:H8");
metrics3.load("values");

// Check a few monthly rows
const monthlyCheck = ws.getRange("A12:F14");
monthlyCheck.load("values");

// Check customer top/bottom
const custTop = ws.getRange("A37:E39");
custTop.load("values");

// Check category totals
const catTotals = ws.getRange("B73:C75");
catTotals.load("values");

// Check top 5 share
const top5 = ws.getRange("B53");
top5.load("values");

// Scan for errors in all data ranges
const allFormulas = ws.getRange("B6:F75");
allFormulas.load("values");

await context.sync();

// Check for #errors
const allVals = allFormulas.values.flat();
const errors = allVals.filter(v => typeof v === "string" && (v.startsWith("#") || v === ""));

// Sanity check: fuel + services + specialty should ≈ total revenue
const fuelRev = catTotals.values[0][0];
const servRev = catTotals.values[1][0];
const specRev = catTotals.values[2][0];
const catSum = fuelRev + servRev + specRev;
const totalRev = metrics.values[0][0];

return {
  keyMetrics: {
    totalRevenue: totalRev,
    totalTransactions: metrics.values[1][0],
    totalGallons: metrics.values[2][0],
    avgSaleSize: metrics2.values[0][0],
    avgFuelPrice: metrics2.values[1][0],
    serviceTransactions: metrics2.values[2][0],
    dateStart: metrics3.values[0][0],
    dateEnd: metrics3.values[1][0],
    uniqueCustomers: metrics3.values[2][0]
  },
  sampleMonthly: monthlyCheck.values,
  sampleCustomers: custTop.values,
  categories: {
    fuel: fuelRev,
    services: servRev,
    specialty: specRev,
    sum: catSum,
    totalRevenue: totalRev,
    coveragePercent: (catSum / totalRev * 100).toFixed(1) + "%"
  },
  top5Share: top5.values[0][0],
  errorCount: errors.length,
  errorSamples: errors.slice(0, 5)
};

---

result {
  "success": true,
  "result": {
    "keyMetrics": {
      "totalRevenue": 790660.08,
      "totalTransactions": 441,
      "totalGallons": 181158,
      "avgSaleSize": 1792.88,
      "avgFuelPrice": 3.69867253999271,
      "serviceTransactions": 72,
      "dateStart": 45600,
      "dateEnd": 46224,
      "uniqueCustomers": 15
    },
    "sampleMonthly": [
      [
        45597,
        30502.09,
        18,
        7660,
        3.81881070496084,
        1694.56055555556
      ],
      [
        45627,
        28816.9,
        21,
        6999,
        3.65293613373339,
        1372.23333333333
      ],
      [
        45658,
        33836.2,
        21,
        8338,
        3.45241065003598,
        1611.24761904762
      ]
    ],
    "sampleCustomers": [
      [
        "Acme Petroleum",
        41501.67,
        0.0524899018551689,
        24,
        1729.23625
      ],
      [
        "Arkansas Valley Oil",
        61909.02,
        0.0783004246274834,
        31,
        1997.06516129032
      ],
      [
        "Centennial Hauling",
        47913.34,
        0.0605991641819073,
        30,
        1597.11133333333
      ]
    ],
    "categories": {
      "fuel": 670044.12,
      "services": 87690.96,
      "specialty": 32925,
      "sum": 790660.08,
      "totalRevenue": 790660.08,
      "coveragePercent": "100.0%"
    },
    "top5Share": 0.423472701442066,
    "errorCount": 96,
    "errorSamples": [
      "",
      "",
      "",
      "",
      ""
    ]
  }
}

---

const ws = context.workbook.worksheets.getItem("Analysis");

// Fix dates - they may need explicit date format
const dateStart = ws.getRange("H6");
const dateEnd = ws.getRange("H7");
dateStart.load("values");
dateEnd.load("values");
await context.sync();

// The dates are serial numbers, format them properly
dateStart.numberFormat = [["mm/dd/yyyy"]];
dateEnd.numberFormat = [["mm/dd/yyyy"]];

// Also hide the helper data for pie chart (H57:I59) by making font white
// Actually let's just format it cleanly
ws.getRange("H57:I59").format.font.color = "#FFFFFF";
ws.getRange("I57:I59").numberFormat = [["$#,##0"], ["$#,##0"], ["$#,##0"]];

await context.sync();

// Verify dates
dateStart.load("text");
dateEnd.load("text");
await context.sync();

return { 
  dateStart: dateStart.text,
  dateEnd: dateEnd.text,
  note: "Dates formatted, pie chart helper data hidden"
};

---

result

{
  "success": true,
  "result": {
    "dateStart": [
      [
        "11/04/2024"
      ]
    ],
    "dateEnd": [
      [
        "07/21/2026"
      ]
    ],
    "note": "Dates formatted, pie chart helper data hidden"
  }
}

---

// ===== SECTION 5: PRICING TRENDS (Row 78+) =====
const ws = context.workbook.worksheets.getItem("Analysis");

const startRow = 78;

ws.getRange(`A${startRow}`).values = [["PRICING TRENDS"]];
const sec5Header = ws.getRange(`A${startRow}:L${startRow}`);
sec5Header.merge();
sec5Header.format.fill.color = "#1F4E79";
sec5Header.format.font.color = "#FFFFFF";
sec5Header.format.font.bold = true;
sec5Header.format.font.size = 11;
sec5Header.format.rowHeight = 25;

const hdrRow = startRow + 1;
ws.getRange(`A${hdrRow}:H${hdrRow}`).values = [[
  "Month", "Avg Diesel $/Gal", "Avg Unleaded $/Gal", "Avg Premium $/Gal", 
  "Avg Bulk $/Gal", "Avg Fleet $/Gal", "Overall Avg $/Gal", "Price Δ MoM"
]];
const colHdr = ws.getRange(`A${hdrRow}:H${hdrRow}`);
colHdr.format.fill.color = "#D9E2F3";
colHdr.format.font.bold = true;
colHdr.format.horizontalAlignment = "Center";

// Use same month dates from monthly trend section (A12:A32)
const dataStart = hdrRow + 1;
for (let i = 0; i <= 20; i++) {
  const row = dataStart + i;
  const monthRef = `A${12 + i}`; // references existing month dates in monthly trend
  const nextMonthRef = i < 20 ? `A${13 + i}` : `EDATE(A${12 + i},1)`;

  // Month reference
  ws.getRange(`A${row}`).formulas = [[`=${monthRef}`]];
  ws.getRange(`A${row}`).numberFormat = [["mmm yyyy"]];

  // Helper function pattern for avg price by product type containing keyword
  // Avg Diesel $/Gal
  ws.getRange(`B${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Diesel Delivery",RAW_INPUT!D5:D445)))*RAW_INPUT!F5:F445)/SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Diesel Delivery",RAW_INPUT!D5:D445)))*1),"-")`
  ]];

  // Avg Unleaded $/Gal
  ws.getRange(`C${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Unleaded",RAW_INPUT!D5:D445)))*RAW_INPUT!F5:F445)/SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Unleaded",RAW_INPUT!D5:D445)))*1),"-")`
  ]];

  // Avg Premium $/Gal
  ws.getRange(`D${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Premium",RAW_INPUT!D5:D445)))*RAW_INPUT!F5:F445)/SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Premium",RAW_INPUT!D5:D445)))*1),"-")`
  ]];

  // Avg Bulk $/Gal
  ws.getRange(`E${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Bulk",RAW_INPUT!D5:D445)))*RAW_INPUT!F5:F445)/SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Bulk",RAW_INPUT!D5:D445)))*1),"-")`
  ]];

  // Avg Fleet $/Gal
  ws.getRange(`F${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Fleet",RAW_INPUT!D5:D445)))*RAW_INPUT!F5:F445)/SUMPRODUCT((RAW_INPUT!A5:A445>=${monthRef})*(RAW_INPUT!A5:A445<${nextMonthRef})*(ISNUMBER(SEARCH("Fleet",RAW_INPUT!D5:D445)))*1),"-")`
  ]];

  // Overall avg price (reuse from monthly trend)
  ws.getRange(`G${row}`).formulas = [[`=E${12 + i}`]]; // links to existing Avg Price/Gal in monthly trend

  // MoM change
  if (i === 0) {
    ws.getRange(`H${row}`).values = [["-"]];
  } else {
    ws.getRange(`H${row}`).formulas = [[`=IFERROR((G${row}-G${row - 1})/G${row - 1},"-")`]];
  }
}

const dataEnd = dataStart + 20;

// Formatting
const rowCount = 21;
ws.getRange(`B${dataStart}:G${dataEnd}`).numberFormat = Array(rowCount).fill(["$#,##0.00","$#,##0.00","$#,##0.00","$#,##0.00","$#,##0.00","$#,##0.00"]);
ws.getRange(`H${dataStart}:H${dataEnd}`).numberFormat = Array(rowCount).fill(["0.0%"]);
ws.getRange(`B${dataStart}:H${dataEnd}`).format.horizontalAlignment = "Center";

// Alternating shading
for (let i = dataStart; i <= dataEnd; i++) {
  if ((i - dataStart) % 2 === 1) {
    ws.getRange(`A${i}:H${i}`).format.fill.color = "#F8F9FA";
  }
}

// Stats row
const statsRow = dataEnd + 2;
ws.getRange(`A${statsRow}:A${statsRow + 2}`).values = [["Max Price (Overall)"], ["Min Price (Overall)"], ["Price Range"]];
ws.getRange(`A${statsRow}:A${statsRow + 2}`).format.font.bold = true;
ws.getRange(`A${statsRow}:H${statsRow + 2}`).format.fill.color = "#F2F2F2";

ws.getRange(`B${statsRow}`).formulas = [[`=MAX(G${dataStart}:G${dataEnd})`]];
ws.getRange(`B${statsRow + 1}`).formulas = [[`=MIN(G${dataStart}:G${dataEnd})`]];
ws.getRange(`B${statsRow + 2}`).formulas = [[`=B${statsRow}-B${statsRow + 1}`]];
ws.getRange(`B${statsRow}:B${statsRow + 2}`).numberFormat = [["$#,##0.00"],["$#,##0.00"],["$#,##0.00"]];

// Pricing trend chart - line chart with all fuel types
const chart = ws.charts.add(
  Excel.ChartType.lineMarkers,
  ws.getRange(`A${hdrRow}:G${dataEnd}`),
  Excel.ChartSeriesBy.columns
);
chart.title.text = "Monthly Fuel Price by Type ($/Gal)";
chart.title.format.font.size = 12;
chart.title.format.font.bold = true;
chart.setPosition(ws.getRange(`I${startRow}`), ws.getRange(`P${statsRow + 2}`));
chart.legend.visible = true;
chart.legend.position = Excel.ChartLegendPosition.bottom;

await context.sync();

return { 
  phase: "Pricing Trends complete",
  dataRange: `A${dataStart}:H${dataEnd}`,
  nextRow: statsRow + 5
};

---
result
{
  "success": true,
  "result": {
    "phase": "Pricing Trends complete",
    "dataRange": "A80:H100",
    "nextRow": 107
  }
}

---

// ===== SECTION 6: SEASONAL PATTERNS (Row 108+) =====
const ws = context.workbook.worksheets.getItem("Analysis");

const startRow = 108;

ws.getRange(`A${startRow}`).values = [["SEASONAL PATTERNS"]];
const sec6Header = ws.getRange(`A${startRow}:L${startRow}`);
sec6Header.merge();
sec6Header.format.fill.color = "#1F4E79";
sec6Header.format.font.color = "#FFFFFF";
sec6Header.format.font.bold = true;
sec6Header.format.font.size = 11;
sec6Header.format.rowHeight = 25;

// --- Part A: Day of Week Analysis ---
const dowHdr = startRow + 1;
ws.getRange(`A${dowHdr}`).values = [["DAY OF WEEK ANALYSIS"]];
ws.getRange(`A${dowHdr}:F${dowHdr}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${dowHdr}`).format.font.bold = true;

ws.getRange(`A${dowHdr + 1}:E${dowHdr + 1}`).values = [["Day", "Revenue", "# Transactions", "Avg Sale", "% of Weekly Revenue"]];
ws.getRange(`A${dowHdr + 1}:E${dowHdr + 1}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${dowHdr + 1}:E${dowHdr + 1}`).format.font.bold = true;
ws.getRange(`A${dowHdr + 1}:E${dowHdr + 1}`).format.horizontalAlignment = "Center";

const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const dowDataStart = dowHdr + 2;

for (let i = 0; i < days.length; i++) {
  const row = dowDataStart + i;
  ws.getRange(`A${row}`).values = [[days[i]]];
  
  // WEEKDAY returns 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu, 6=Fri, 7=Sat
  const weekdayNum = i + 2; // Mon=2, Tue=3, Wed=4, Thu=5, Fri=6
  
  // Revenue
  ws.getRange(`B${row}`).formulas = [[
    `=SUMPRODUCT((WEEKDAY(RAW_INPUT!A5:A445,1)=${weekdayNum})*RAW_INPUT!G5:G445)`
  ]];
  
  // Count
  ws.getRange(`C${row}`).formulas = [[
    `=SUMPRODUCT((WEEKDAY(RAW_INPUT!A5:A445,1)=${weekdayNum})*1)`
  ]];
  
  // Avg sale
  ws.getRange(`D${row}`).formulas = [[`=IFERROR(B${row}/C${row},0)`]];
  
  // % of total
  ws.getRange(`E${row}`).formulas = [[`=B${row}/SUM(B${dowDataStart}:B${dowDataStart + 4})`]];
}

const dowDataEnd = dowDataStart + 4;
ws.getRange(`B${dowDataStart}:B${dowDataEnd}`).numberFormat = Array(5).fill(["$#,##0"]);
ws.getRange(`C${dowDataStart}:C${dowDataEnd}`).numberFormat = Array(5).fill(["#,##0"]);
ws.getRange(`D${dowDataStart}:D${dowDataEnd}`).numberFormat = Array(5).fill(["$#,##0"]);
ws.getRange(`E${dowDataStart}:E${dowDataEnd}`).numberFormat = Array(5).fill(["0.0%"]);
ws.getRange(`B${dowDataStart}:E${dowDataEnd}`).format.horizontalAlignment = "Center";

// Alternating shading
for (let i = dowDataStart; i <= dowDataEnd; i++) {
  if ((i - dowDataStart) % 2 === 1) {
    ws.getRange(`A${i}:E${i}`).format.fill.color = "#F8F9FA";
  }
}

// --- Part B: Quarter Analysis ---
const qtrHdr = dowDataEnd + 2;
ws.getRange(`A${qtrHdr}`).values = [["QUARTERLY PERFORMANCE"]];
ws.getRange(`A${qtrHdr}:F${qtrHdr}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${qtrHdr}`).format.font.bold = true;

ws.getRange(`A${qtrHdr + 1}:F${qtrHdr + 1}`).values = [["Quarter", "Revenue", "# Transactions", "Fuel Gallons", "Avg Price/Gal", "Avg Sale"]];
ws.getRange(`A${qtrHdr + 1}:F${qtrHdr + 1}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${qtrHdr + 1}:F${qtrHdr + 1}`).format.font.bold = true;
ws.getRange(`A${qtrHdr + 1}:F${qtrHdr + 1}`).format.horizontalAlignment = "Center";

// Quarters: Q4 2024, Q1-Q4 2025, Q1-Q2 2026 (7 quarters)
const quarters = [
  { label: "Q4 2024", startMonth: 10, startYear: 2024, endMonth: 1, endYear: 2025 },
  { label: "Q1 2025", startMonth: 1, startYear: 2025, endMonth: 4, endYear: 2025 },
  { label: "Q2 2025", startMonth: 4, startYear: 2025, endMonth: 7, endYear: 2025 },
  { label: "Q3 2025", startMonth: 7, startYear: 2025, endMonth: 10, endYear: 2025 },
  { label: "Q4 2025", startMonth: 10, startYear: 2025, endMonth: 1, endYear: 2026 },
  { label: "Q1 2026", startMonth: 1, startYear: 2026, endMonth: 4, endYear: 2026 },
  { label: "Q2 2026", startMonth: 4, startYear: 2026, endMonth: 7, endYear: 2026 },
  { label: "Q3 2026", startMonth: 7, startYear: 2026, endMonth: 10, endYear: 2026 },
];

const qtrDataStart = qtrHdr + 2;
for (let i = 0; i < quarters.length; i++) {
  const row = qtrDataStart + i;
  const q = quarters[i];
  
  ws.getRange(`A${row}`).values = [[q.label]];
  
  const startDate = `DATE(${q.startYear},${q.startMonth},1)`;
  const endDate = `DATE(${q.endYear},${q.endMonth},1)`;
  
  // Revenue
  ws.getRange(`B${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=${startDate})*(RAW_INPUT!A5:A445<${endDate})*RAW_INPUT!G5:G445)`
  ]];
  
  // Count
  ws.getRange(`C${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=${startDate})*(RAW_INPUT!A5:A445<${endDate})*1)`
  ]];
  
  // Fuel gallons
  ws.getRange(`D${row}`).formulas = [[
    `=SUMPRODUCT((RAW_INPUT!A5:A445>=${startDate})*(RAW_INPUT!A5:A445<${endDate})*(RAW_INPUT!F5:F445<20)*RAW_INPUT!E5:E445)`
  ]];
  
  // Avg price/gal
  ws.getRange(`E${row}`).formulas = [[
    `=IFERROR(SUMPRODUCT((RAW_INPUT!A5:A445>=${startDate})*(RAW_INPUT!A5:A445<${endDate})*(RAW_INPUT!F5:F445<20)*RAW_INPUT!G5:G445)/D${row},0)`
  ]];
  
  // Avg sale
  ws.getRange(`F${row}`).formulas = [[`=IFERROR(B${row}/C${row},0)`]];
}

const qtrDataEnd = qtrDataStart + quarters.length - 1;

ws.getRange(`B${qtrDataStart}:B${qtrDataEnd}`).numberFormat = Array(quarters.length).fill(["$#,##0"]);
ws.getRange(`C${qtrDataStart}:C${qtrDataEnd}`).numberFormat = Array(quarters.length).fill(["#,##0"]);
ws.getRange(`D${qtrDataStart}:D${qtrDataEnd}`).numberFormat = Array(quarters.length).fill(["#,##0"]);
ws.getRange(`E${qtrDataStart}:E${qtrDataEnd}`).numberFormat = Array(quarters.length).fill(["$#,##0.00"]);
ws.getRange(`F${qtrDataStart}:F${qtrDataEnd}`).numberFormat = Array(quarters.length).fill(["$#,##0"]);
ws.getRange(`B${qtrDataStart}:F${qtrDataEnd}`).format.horizontalAlignment = "Center";

for (let i = qtrDataStart; i <= qtrDataEnd; i++) {
  if ((i - qtrDataStart) % 2 === 1) {
    ws.getRange(`A${i}:F${i}`).format.fill.color = "#F8F9FA";
  }
}

// Day of Week chart
const dowChart = ws.charts.add(
  Excel.ChartType.columnClustered,
  ws.getRange(`A${dowHdr + 1}:B${dowDataEnd}`),
  Excel.ChartSeriesBy.columns
);
dowChart.title.text = "Revenue by Day of Week";
dowChart.title.format.font.size = 12;
dowChart.title.format.font.bold = true;
dowChart.setPosition(ws.getRange(`G${startRow}`), ws.getRange(`N${qtrHdr - 1}`));
dowChart.legend.visible = false;
const dowSeries = dowChart.series.getItemAt(0);
dowSeries.format.fill.setSolidColor("#2E75B6");

// Quarterly chart
const qtrChart = ws.charts.add(
  Excel.ChartType.columnClustered,
  ws.getRange(`A${qtrHdr + 1}:B${qtrDataEnd}`),
  Excel.ChartSeriesBy.columns
);
qtrChart.title.text = "Quarterly Revenue";
qtrChart.title.format.font.size = 12;
qtrChart.title.format.font.bold = true;
qtrChart.setPosition(ws.getRange(`G${qtrHdr}`), ws.getRange(`N${qtrDataEnd + 2}`));
qtrChart.legend.visible = false;
const qtrSeries = qtrChart.series.getItemAt(0);
qtrSeries.format.fill.setSolidColor("#1F4E79");

await context.sync();

return { 
  phase: "Seasonal Patterns complete",
  dowRange: `A${dowDataStart}:E${dowDataEnd}`,
  qtrRange: `A${qtrDataStart}:F${qtrDataEnd}`,
  nextRow: qtrDataEnd + 4
};

---

result 

{
  "success": true,
  "result": {
    "phase": "Seasonal Patterns complete",
    "dowRange": "A111:E115",
    "qtrRange": "A119:F126",
    "nextRow": 130
  }
}

---

// ===== SECTION 7: CUSTOMER GROWTH TRACKING =====
const ws = context.workbook.worksheets.getItem("Analysis");

// Get unique customers again
const rawSheet = context.workbook.worksheets.getItem("RAW_INPUT");
const custRange = rawSheet.getRange("B5:B445");
custRange.load("values");
await context.sync();
const uniqueCustomers = [...new Set(custRange.values.map(r => r[0]).filter(v => v && v !== ""))].sort();

const startRow = 130;

ws.getRange(`A${startRow}`).values = [["CUSTOMER GROWTH TRACKING"]];
const sec7Header = ws.getRange(`A${startRow}:L${startRow}`);
sec7Header.merge();
sec7Header.format.fill.color = "#1F4E79";
sec7Header.format.font.color = "#FFFFFF";
sec7Header.format.font.bold = true;
sec7Header.format.font.size = 11;
sec7Header.format.rowHeight = 25;

// --- Part A: Customer Revenue by Quarter ---
const qtrHdr = startRow + 1;
ws.getRange(`A${qtrHdr}`).values = [["QUARTERLY REVENUE BY CUSTOMER"]];
ws.getRange(`A${qtrHdr}:I${qtrHdr}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${qtrHdr}`).format.font.bold = true;

const quarters = [
  { label: "Q4 2024", startMonth: 10, startYear: 2024, endMonth: 1, endYear: 2025 },
  { label: "Q1 2025", startMonth: 1, startYear: 2025, endMonth: 4, endYear: 2025 },
  { label: "Q2 2025", startMonth: 4, startYear: 2025, endMonth: 7, endYear: 2025 },
  { label: "Q3 2025", startMonth: 7, startYear: 2025, endMonth: 10, endYear: 2025 },
  { label: "Q4 2025", startMonth: 10, startYear: 2025, endMonth: 1, endYear: 2026 },
  { label: "Q1 2026", startMonth: 1, startYear: 2026, endMonth: 4, endYear: 2026 },
  { label: "Q2 2026", startMonth: 4, startYear: 2026, endMonth: 7, endYear: 2026 },
  { label: "Q3 2026", startMonth: 7, startYear: 2026, endMonth: 10, endYear: 2026 },
];

// Headers
const colHdrRow = qtrHdr + 1;
const qLabels = quarters.map(q => q.label);
ws.getRange(`A${colHdrRow}`).values = [["Customer"]];
// Quarters in B through I
for (let q = 0; q < quarters.length; q++) {
  const col = String.fromCharCode(66 + q); // B, C, D, E, F, G, H, I
  ws.getRange(`${col}${colHdrRow}`).values = [[qLabels[q]]];
}
ws.getRange(`A${colHdrRow}:I${colHdrRow}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${colHdrRow}:I${colHdrRow}`).format.font.bold = true;
ws.getRange(`A${colHdrRow}:I${colHdrRow}`).format.horizontalAlignment = "Center";

// Data rows
const dataStart = colHdrRow + 1;
for (let c = 0; c < uniqueCustomers.length; c++) {
  const row = dataStart + c;
  ws.getRange(`A${row}`).values = [[uniqueCustomers[c]]];
  
  for (let q = 0; q < quarters.length; q++) {
    const col = String.fromCharCode(66 + q);
    const qtr = quarters[q];
    const startDate = `DATE(${qtr.startYear},${qtr.startMonth},1)`;
    const endDate = `DATE(${qtr.endYear},${qtr.endMonth},1)`;
    
    ws.getRange(`${col}${row}`).formulas = [[
      `=SUMPRODUCT((RAW_INPUT!B5:B445=A${row})*(RAW_INPUT!A5:A445>=${startDate})*(RAW_INPUT!A5:A445<${endDate})*RAW_INPUT!G5:G445)`
    ]];
  }
}

const dataEnd = dataStart + uniqueCustomers.length - 1;

// Format
for (let q = 0; q < quarters.length; q++) {
  const col = String.fromCharCode(66 + q);
  ws.getRange(`${col}${dataStart}:${col}${dataEnd}`).numberFormat = Array(uniqueCustomers.length).fill(["$#,##0"]);
}
ws.getRange(`B${dataStart}:I${dataEnd}`).format.horizontalAlignment = "Center";

// Alternating rows
for (let i = dataStart; i <= dataEnd; i++) {
  if ((i - dataStart) % 2 === 1) {
    ws.getRange(`A${i}:I${i}`).format.fill.color = "#F8F9FA";
  }
}

// --- Part B: Customer Growth Metrics ---
const growthRow = dataEnd + 2;
ws.getRange(`A${growthRow}`).values = [["CUSTOMER GROWTH METRICS"]];
ws.getRange(`A${growthRow}:F${growthRow}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${growthRow}`).format.font.bold = true;

ws.getRange(`A${growthRow + 1}:F${growthRow + 1}`).values = [[
  "Customer", "First Quarter Rev", "Latest Quarter Rev", "Growth ($)", "Growth (%)", "Trend"
]];
ws.getRange(`A${growthRow + 1}:F${growthRow + 1}`).format.fill.color = "#D9E2F3";
ws.getRange(`A${growthRow + 1}:F${growthRow + 1}`).format.font.bold = true;
ws.getRange(`A${growthRow + 1}:F${growthRow + 1}`).format.horizontalAlignment = "Center";

const gDataStart = growthRow + 2;
for (let c = 0; c < uniqueCustomers.length; c++) {
  const row = gDataStart + c;
  const custDataRow = dataStart + c; // Row in quarterly grid
  
  ws.getRange(`A${row}`).values = [[uniqueCustomers[c]]];
  
  // First non-zero quarter revenue (look left to right in B-I of custDataRow)
  // Use a formula approach: find the first non-zero
  // AGGREGATE(15,6,...,1) = SMALL ignoring errors, kth=1
  // Simpler: just use the first quarter (B) and last active quarter
  // Actually, let's find first and last non-zero quarter programmatically via formula
  // First non-zero: use LOOKUP trick
  ws.getRange(`B${row}`).formulas = [[
    `=IFERROR(INDEX(B${custDataRow}:I${custDataRow},MATCH(TRUE,B${custDataRow}:I${custDataRow}>0,0)),0)`
  ]];
  
  // Latest non-zero: LOOKUP(2,1/(B:I>0),B:I)
  ws.getRange(`C${row}`).formulas = [[
    `=IFERROR(LOOKUP(2,1/(B${custDataRow}:I${custDataRow}>0),B${custDataRow}:I${custDataRow}),0)`
  ]];
  
  // Growth $
  ws.getRange(`D${row}`).formulas = [[`=C${row}-B${row}`]];
  
  // Growth %
  ws.getRange(`E${row}`).formulas = [[`=IFERROR(D${row}/B${row},0)`]];
  
  // Trend indicator (simple: positive = ▲, negative = ▼, flat = ►)
  ws.getRange(`F${row}`).formulas = [[`=IF(D${row}>0,"▲ Growing",IF(D${row}<0,"▼ Declining","► Flat"))`]];
}

const gDataEnd = gDataStart + uniqueCustomers.length - 1;

ws.getRange(`B${gDataStart}:D${gDataEnd}`).numberFormat = Array(uniqueCustomers.length).fill(["$#,##0","$#,##0","$#,##0"]);
ws.getRange(`E${gDataStart}:E${gDataEnd}`).numberFormat = Array(uniqueCustomers.length).fill(["0.0%"]);
ws.getRange(`B${gDataStart}:F${gDataEnd}`).format.horizontalAlignment = "Center";

for (let i = gDataStart; i <= gDataEnd; i++) {
  if ((i - gDataStart) % 2 === 1) {
    ws.getRange(`A${i}:F${i}`).format.fill.color = "#F8F9FA";
  }
}

// Summary stats
const sumRow = gDataEnd + 2;
ws.getRange(`A${sumRow}:A${sumRow + 2}`).values = [["Customers Growing"], ["Customers Declining"], ["Avg Growth %"]];
ws.getRange(`A${sumRow}:A${sumRow + 2}`).format.font.bold = true;
ws.getRange(`A${sumRow}:F${sumRow + 2}`).format.fill.color = "#F2F2F2";

ws.getRange(`B${sumRow}`).formulas = [[`=COUNTIF(E${gDataStart}:E${gDataEnd},">"&0)`]];
ws.getRange(`B${sumRow + 1}`).formulas = [[`=COUNTIF(E${gDataStart}:E${gDataEnd},"<"&0)`]];
ws.getRange(`B${sumRow + 2}`).formulas = [[`=AVERAGE(E${gDataStart}:E${gDataEnd})`]];
ws.getRange(`B${sumRow + 2}`).numberFormat = [["0.0%"]];

// Customer growth chart - stacked bar of quarterly revenue for top 5 customers
// We'll chart the quarterly grid for visual impact
const custChart = ws.charts.add(
  Excel.ChartType.barStacked,
  ws.getRange(`A${colHdrRow}:I${dataEnd}`),
  Excel.ChartSeriesBy.columns
);
custChart.title.text = "Quarterly Revenue by Customer";
custChart.title.format.font.size = 12;
custChart.title.format.font.bold = true;
custChart.setPosition(ws.getRange(`K${startRow}`), ws.getRange(`R${dataEnd + 2}`));
custChart.legend.visible = true;
custChart.legend.position = Excel.ChartLegendPosition.bottom;

await context.sync();

return {
  phase: "Customer Growth Tracking complete",
  qtrGrid: `A${dataStart}:I${dataEnd}`,
  growthMetrics: `A${gDataStart}:F${gDataEnd}`,
  finalRow: sumRow + 4
};

---

result

{
  "success": true,
  "result": {
    "phase": "Customer Growth Tracking complete",
    "qtrGrid": "A133:I147",
    "growthMetrics": "A151:F165",
    "finalRow": 171
  }
}

---

// Final verification across all new sections
const ws = context.workbook.worksheets.getItem("Analysis");

// Check pricing trends sample
const priceSample = ws.getRange("A80:H82");
priceSample.load("values");

// Check day of week
const dowSample = ws.getRange("A111:E115");
dowSample.load("values");

// Check quarterly
const qtrSample = ws.getRange("A119:F122");
qtrSample.load("values");

// Check customer quarterly grid (first 3 customers)
const custQtrSample = ws.getRange("A133:I135");
custQtrSample.load("values");

// Check growth metrics (first 3)
const growthSample = ws.getRange("A151:F153");
growthSample.load("values");

// Check growth summary
const growthStats = ws.getRange("A167:B169");
growthStats.load("values");

// Scan for #errors in all new sections
const section5 = ws.getRange("B80:H100");
section5.load("values");
const section6a = ws.getRange("B111:E115");
section6a.load("values");
const section6b = ws.getRange("B119:F126");
section6b.load("values");
const section7a = ws.getRange("B133:I147");
section7a.load("values");
const section7b = ws.getRange("B151:F165");
section7b.load("values");

await context.sync();

// Check all sections for #errors
const allSections = [
  section5.values, section6a.values, section6b.values, 
  section7a.values, section7b.values
];
const errors = [];
for (const section of allSections) {
  for (const row of section) {
    for (const cell of row) {
      if (typeof cell === "string" && cell.startsWith("#")) {
        errors.push(cell);
      }
    }
  }
}

return {
  verification: "COMPLETE",
  pricingSample: priceSample.values,
  dayOfWeek: dowSample.values,
  quarterlySample: qtrSample.values,
  custQuarterly: custQtrSample.values.map(r => [r[0], r[1], r[8]]), // first & last quarter
  growthSample: growthSample.values,
  growthStats: growthStats.values,
  formulaErrors: errors.length > 0 ? errors.slice(0, 10) : "None — all clean ✓",
  totalNewSections: 3
};

---

result

{
  "success": true,
  "result": {
    "verification": "COMPLETE",
    "pricingSample": [
      [
        45597,
        3.82285714285714,
        3.85,
        3.77,
        "-",
        3.80333333333333,
        3.81881070496084,
        "-"
      ],
      [
        45627,
        3.658,
        3.61333333333333,
        3.64,
        3.62333333333333,
        3.66,
        3.65293613373339,
        -0.0434361857768872
      ],
      [
        45658,
        "-",
        3.46666666666667,
        3.46,
        3.4,
        3.455,
        3.45241065003598,
        -0.0548943305757906
      ]
    ],
    "dayOfWeek": [
      [
        "Monday",
        163108.25,
        88,
        1853.50284090909,
        0.206293771654691
      ],
      [
        "Tuesday",
        171035.56,
        90,
        1900.39511111111,
        0.216319963947086
      ],
      [
        "Wednesday",
        146387.14,
        87,
        1682.6108045977,
        0.185145479964032
      ],
      [
        "Thursday",
        153388.14,
        88,
        1743.04704545454,
        0.194000106847433
      ],
      [
        "Friday",
        156740.99,
        88,
        1781.14761363636,
        0.198240677586758
      ]
    ],
    "quarterlySample": [
      [
        "Q4 2024",
        59318.99,
        39,
        14659,
        3.73961320690361,
        1520.99974358974
      ],
      [
        "Q1 2025",
        109849.28,
        61,
        25671,
        3.70532040045187,
        1800.80786885246
      ],
      [
        "Q2 2025",
        106783.1,
        65,
        25664,
        3.64292355049875,
        1642.81692307692
      ],
      [
        "Q3 2025",
        116385.27,
        66,
        26055,
        3.71235002878526,
        1763.41318181818
      ]
    ],
    "custQuarterly": [
      [
        "Acme Petroleum",
        3528.7,
        4339.7
      ],
      [
        "Arkansas Valley Oil",
        6720,
        0
      ],
      [
        "Centennial Hauling",
        3400.45,
        0
      ]
    ],
    "growthSample": [
      [
        "Acme Petroleum",
        3528.7,
        4339.7,
        811,
        0.229829682319268,
        "▲ Growing"
      ],
      [
        "Arkansas Valley Oil",
        6720,
        6959.94,
        239.94,
        0.0357053571428571,
        "▲ Growing"
      ],
      [
        "Centennial Hauling",
        3400.45,
        7676.26,
        4275.81,
        1.25742475260627,
        "▲ Growing"
      ]
    ],
    "growthStats": [
      [
        "Customers Growing",
        9
      ],
      [
        "Customers Declining",
        6
      ],
      [
        "Avg Growth %",
        0.828921789595938
      ]
    ],
    "formulaErrors": "None — all clean ✓",
    "totalNewSections": 3
  }
}






