/**
 * @file GomagOrderSum.gs
 * @description Script Gomag cu Debugging Avansat pentru a depana structura răspunsului.
 */

// ==========================================
// 1. CONFIGURARE GOMAG
// ==========================================

const GOMAG_CONFIG = {
  // URL-ul magazinului (fără slash la final)
  shopUrl: 'https://www.mt.ro', 
  
  // Cheia API
  apiKey: '6e15b5ec0db4d50efc9ebe6', 
  
  // Perioada de analiză
  timeframeDays: 30,

  // Mapare câmpuri
  fields: {
    total: 'total',       // Câmp valoare
    status: 'status',     // Câmp status
    date: 'date',         // Câmp dată
    items_list: 'orders'  // Câmp listă (opțional)
  }
};

const GOMAG_SHEET_NAME = "GomagOrderTotals"; 

// ==========================================
// 2. FUNCȚIA PRINCIPALĂ
// ==========================================

function runGomagOrderCalculator() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - GOMAG_CONFIG.timeframeDays);
  startDate.setHours(0, 0, 0, 0); 
  
  Logger.log(`[Gomag] Start analiză. Dată de referință: ${startDate.toISOString().split('T')[0]}`);

  try {
    const data = fetchGomagOrders_(startDate);
    
    let totalValue = 0;
    let orderCount = 0;

    data.orders.forEach(order => {
      const status = String(order[GOMAG_CONFIG.fields.status] || "").toLowerCase();
      if (status.includes('anul') || status.includes('cancel') || status.includes('retur')) return;

      let priceRaw = order[GOMAG_CONFIG.fields.total];
      if (typeof priceRaw === 'string') {
          priceRaw = priceRaw.replace(/[^0-9.-]+/g,"");
      }
      const price = parseFloat(priceRaw);

      if (!isNaN(price) && price > 0) {
        totalValue += price;
        orderCount++;
      }
    });

    Logger.log(`[Gomag] Succes: ${orderCount} comenzi. Total: ${totalValue}`);
    writeGomagResultToSheet_(spreadsheet, startDate, totalValue, orderCount);

  } catch (e) {
    Logger.log(`[Gomag] EROARE CRITICĂ: ${e.message}`);
  }
}

// ==========================================
// 3. FUNCȚII AUXILIARE (Fetch & Debug)
// ==========================================

function fetchGomagOrders_(startDate) {
  const allOrders = [];
  let page = 1;
  let keepFetching = true;

  while (keepFetching) {
    const url = `https://api.gomag.ro/api/v1/order/read/json?limit=50&page=${page}`;
    
    Logger.log(`[Gomag] Cerere pagina ${page}...`);
    
    const response = fetchGomagUrl_(url);

    if (!response) {
      Logger.log("[Gomag] Răspuns nul de la API. Verifică erorile anterioare.");
      break;
    }
    
    // --- DEBUGGING: Afișăm ce am primit ---
    if (page === 1) {
      if (Array.isArray(response)) {
         Logger.log(`[DEBUG] Răspunsul este o LISTĂ (Array) cu ${response.length} elemente.`);
         if (response.length > 0) Logger.log(`[DEBUG] Exemplu element 1 chei: ${Object.keys(response[0]).join(', ')}`);
      } else {
         Logger.log(`[DEBUG] Răspunsul este un OBIECT.`);
         Logger.log(`[DEBUG] Cheile obiectului primit: ${Object.keys(response).join(', ')}`);
         
         // Verificăm dacă e mesaj de eroare
         if (response.message || response.error) {
           Logger.log(`[DEBUG] POSIBILĂ EROARE API: ${response.message || response.error}`);
         }
      }
    }
    // --------------------------------------

    let ordersBatch = [];

    // Logică detecție listă
    if (Array.isArray(response)) {
      ordersBatch = response;
    } else if (response[GOMAG_CONFIG.fields.items_list]) {
      // Dacă există cheia 'orders'
      const val = response[GOMAG_CONFIG.fields.items_list];
      ordersBatch = Array.isArray(val) ? val : Object.values(val);
    } else if (response.data) {
       // Dacă există cheia 'data'
       if (Array.isArray(response.data)) {
         ordersBatch = response.data;
       } else if (response.data[GOMAG_CONFIG.fields.items_list]) {
         const val = response.data[GOMAG_CONFIG.fields.items_list];
         ordersBatch = Array.isArray(val) ? val : Object.values(val);
       }
    } 
    // Tentativă de a trata obiectul rădăcină ca o listă de comenzi (dacă cheile sunt ID-uri)
    else if (typeof response === 'object') {
       const vals = Object.values(response);
       // Dacă primul element pare a fi o comandă (are dată), atunci probabil e o listă keyed by ID
       if (vals.length > 0 && vals[0] && typeof vals[0] === 'object' && vals[0][GOMAG_CONFIG.fields.date]) {
         Logger.log("[DEBUG] Detectat format obiect cu chei ID. Convertim în listă.");
         ordersBatch = vals;
       }
    }

    if (!ordersBatch || ordersBatch.length === 0) {
      Logger.log("[Gomag] Nu am găsit o listă validă de comenzi în răspuns. Structura nu corespunde așteptărilor. Stop.");
      break;
    }

    let ordersInBatchAdded = 0;
    for (const order of ordersBatch) {
      const orderDateStr = order[GOMAG_CONFIG.fields.date];
      
      if (!orderDateStr) continue; // Sărim peste cele fără dată

      const orderDate = new Date(orderDateStr);
      
      if (orderDate < startDate) {
        Logger.log(`[Gomag] Data limită atinsă (${orderDateStr}). Oprire.`);
        keepFetching = false;
        break; 
      }
      
      allOrders.push(order);
      ordersInBatchAdded++;
    }

    if (ordersInBatchAdded === 0 && keepFetching && ordersBatch.length > 0) {
       const firstOrderDate = new Date(ordersBatch[0][GOMAG_CONFIG.fields.date]);
       if (firstOrderDate < startDate) {
         keepFetching = false;
       }
    }

    if (keepFetching) {
      page++;
      if (page > 50) keepFetching = false;
    }
  }

  return { orders: allOrders };
}

function fetchGomagUrl_(url) {
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Apikey': GOMAG_CONFIG.apiKey,
      'ApiShop': GOMAG_CONFIG.shopUrl,
      'Content-Type': 'application/json'
    }
  };

  for (let i = 0; i < 3; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const contentText = response.getContentText();
      
      if (code === 200) {
        return JSON.parse(contentText);
      } else if (code === 429) {
        Utilities.sleep(5000);
        continue;
      } else {
        Logger.log(`[API ERROR] Cod: ${code}. Răspuns: ${contentText}`);
        return null; // Returnăm null explicit la eroare
      }
    } catch (e) {
      if (i === 2) {
        Logger.log(`[FETCH ERROR] ${e.message}`);
        return null;
      }
      Utilities.sleep(2000);
    }
  }
  return null;
}

function writeGomagResultToSheet_(spreadsheet, startDate, totalValue, orderCount) {
  let sheet = spreadsheet.getSheetByName(GOMAG_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(GOMAG_SHEET_NAME);
  } else {
    sheet.clear();
  }

  const headers = ["Data Calcul", "Dată Start", "Zile", "Nr. Comenzi", "Valoare Totală (RON)"];
  sheet.getRange("A1:E1").setValues([headers]).setFontWeight("bold").setBackground("#fff2cc");

  const rowData = [
    new Date(),
    startDate,
    GOMAG_CONFIG.timeframeDays,
    orderCount,
    totalValue
  ];

  sheet.getRange("A2:E2").setValues([rowData]);
  sheet.getRange("E2").setNumberFormat("#,##0.00");
  sheet.autoResizeColumns(1, 5);
}