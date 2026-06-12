(function () {
  "use strict";

  var APPROVAL_STAGES = [
    "Draft",
    "RGM Review",
    "Finance Review",
    "Sales Review",
    "Legal Review",
    "Approved",
    "Rejected",
    "Revised"
  ];

  var CHANNEL_BY_RETAILER_ID = {
    "1": "Amazon / Marketplace",
    "2": "Mass / Digital",
    "4": "Mass / Digital",
    "9": "Club",
    "25": "Club",
    "43": "Grocery / Digital",
    "48": "Grocery",
    "95": "Drug",
    "96": "Drug"
  };

  var RETAILER_BY_ID = {
    "1": "Amazon",
    "2": "Walmart",
    "4": "Target",
    "9": "Costco",
    "25": "Sam's Club",
    "43": "Amazon Fresh",
    "48": "Kroger",
    "95": "CVS",
    "96": "Walgreens"
  };

  var state = {
    fileName: "",
    workbook: null,
    sheets: [],
    productMap: [],
    productLookup: new Map(),
    rollupLookup: new Map(),
    priceRows: [],
    ppcRows: [],
    ppcGaps: [],
    retailerScores: [],
    recommendations: [],
    selectedPriceRowId: "",
    selectedRecId: "",
    activeRecTab: "Summary",
    activeTab: "cockpit",
    filters: {
      brand: "All",
      retailer: "All",
      channel: "All",
      severity: "All",
      audience: "executive"
    },
    chat: [],
    dateRange: ""
  };

  var els = {};

  document.addEventListener("DOMContentLoaded", function () {
    cacheElements();
    bindEvents();
    renderEmptyChat();
  });

  function cacheElements() {
    els.fileInput = document.getElementById("file-input");
    els.dropFileInput = document.getElementById("drop-file-input");
    els.dropZone = document.getElementById("drop-zone");
    els.fileStatus = document.getElementById("file-status");
    els.loadMessage = document.getElementById("load-message");
    els.uploadPanel = document.getElementById("upload-panel");
    els.workbench = document.getElementById("workbench");
    els.cockpitView = document.getElementById("cockpit-view");
    els.recommendationsView = document.getElementById("recommendations-view");
    els.skuView = document.getElementById("sku-view");
    els.strategistView = document.getElementById("strategist-view");
    els.approvalsView = document.getElementById("approvals-view");
    els.dataView = document.getElementById("data-view");
    els.brandFilter = document.getElementById("brand-filter");
    els.retailerFilter = document.getElementById("retailer-filter");
    els.channelFilter = document.getElementById("channel-filter");
    els.severityFilter = document.getElementById("severity-filter");
    els.audienceFilter = document.getElementById("audience-filter");
  }

  function bindEvents() {
    els.fileInput.addEventListener("change", function (event) {
      if (event.target.files && event.target.files[0]) {
        loadWorkbookFile(event.target.files[0]);
      }
    });

    els.dropFileInput.addEventListener("change", function (event) {
      if (event.target.files && event.target.files[0]) {
        loadWorkbookFile(event.target.files[0]);
      }
    });

    els.dropZone.addEventListener("click", function () {
      els.dropFileInput.click();
    });

    els.dropZone.addEventListener("dragover", function (event) {
      event.preventDefault();
      els.dropZone.classList.add("dragover");
    });

    els.dropZone.addEventListener("dragleave", function () {
      els.dropZone.classList.remove("dragover");
    });

    els.dropZone.addEventListener("drop", function (event) {
      event.preventDefault();
      els.dropZone.classList.remove("dragover");
      var file = event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) {
        loadWorkbookFile(file);
      }
    });

    document.querySelectorAll(".tab").forEach(function (button) {
      button.addEventListener("click", function () {
        state.activeTab = button.getAttribute("data-tab");
        document.querySelectorAll(".tab").forEach(function (tab) {
          tab.classList.toggle("active", tab === button);
        });
        render();
      });
    });

    [
      ["brand", els.brandFilter],
      ["retailer", els.retailerFilter],
      ["channel", els.channelFilter],
      ["severity", els.severityFilter],
      ["audience", els.audienceFilter]
    ].forEach(function (pair) {
      pair[1].addEventListener("change", function () {
        state.filters[pair[0]] = pair[1].value;
        render();
      });
    });
  }

  async function loadWorkbookFile(file) {
    setLoadMessage("Reading workbook...", false);
    try {
      var workbook = await readXlsx(file);
      var model = buildPricingModel(workbook, file.name);
      Object.assign(state, model);
      state.fileName = file.name;
      state.workbook = workbook;
      state.selectedPriceRowId = pickDefaultSelectedRow(state.priceRows);
      state.selectedRecId = state.recommendations[0] ? state.recommendations[0].id : "";
      state.activeRecTab = "Summary";
      state.chat = [
        {
          role: "agent",
          text: "Workbook loaded. I can summarize guardrail flags, retailer discipline, price-per-count gaps, and RGM approval needs."
        }
      ];
      els.fileStatus.textContent = file.name;
      els.uploadPanel.classList.add("hidden");
      els.workbench.classList.remove("hidden");
      populateFilters();
      render();
      setLoadMessage("", false);
    } catch (error) {
      console.error(error);
      setLoadMessage(error.message || "The workbook could not be read.", true);
    }
  }

  function setLoadMessage(message, isError) {
    els.loadMessage.textContent = message;
    els.loadMessage.classList.toggle("error", Boolean(isError));
  }

  async function readXlsx(file) {
    if (!("DecompressionStream" in window)) {
      throw new Error("This browser cannot read compressed Excel files locally. Use a current version of Microsoft Edge or Chrome.");
    }

    var buffer = await file.arrayBuffer();
    var zip = await parseZip(buffer);
    var workbookXml = await zip.text("xl/workbook.xml");
    var relsXml = await zip.text("xl/_rels/workbook.xml.rels");
    var sharedStrings = [];

    if (zip.has("xl/sharedStrings.xml")) {
      sharedStrings = parseSharedStrings(await zip.text("xl/sharedStrings.xml"));
    }

    var sheetRefs = parseWorkbookSheets(workbookXml, relsXml);
    var sheets = {};
    var sheetNames = [];

    for (var i = 0; i < sheetRefs.length; i += 1) {
      var ref = sheetRefs[i];
      if (!zip.has(ref.path)) {
        continue;
      }
      sheetNames.push(ref.name);
      sheets[ref.name] = parseSheetRows(await zip.text(ref.path), sharedStrings);
    }

    return {
      sheetNames: sheetNames,
      sheets: sheets
    };
  }

  async function parseZip(buffer) {
    var bytes = new Uint8Array(buffer);
    var eocdOffset = findEndOfCentralDirectory(bytes);
    if (eocdOffset < 0) {
      throw new Error("This does not look like a valid .xlsx workbook.");
    }

    var entryCount = getUint16(bytes, eocdOffset + 10);
    var centralOffset = getUint32(bytes, eocdOffset + 16);
    var offset = centralOffset;
    var entries = new Map();
    var decoder = new TextDecoder("utf-8");

    for (var i = 0; i < entryCount; i += 1) {
      if (getUint32(bytes, offset) !== 0x02014b50) {
        throw new Error("The workbook zip directory is not readable.");
      }

      var method = getUint16(bytes, offset + 10);
      var compressedSize = getUint32(bytes, offset + 20);
      var fileNameLength = getUint16(bytes, offset + 28);
      var extraLength = getUint16(bytes, offset + 30);
      var commentLength = getUint16(bytes, offset + 32);
      var localOffset = getUint32(bytes, offset + 42);
      var nameBytes = bytes.slice(offset + 46, offset + 46 + fileNameLength);
      var name = decoder.decode(nameBytes);

      entries.set(name, {
        name: name,
        method: method,
        compressedSize: compressedSize,
        localOffset: localOffset
      });

      offset += 46 + fileNameLength + extraLength + commentLength;
    }

    async function entryBytes(name) {
      var entry = entries.get(name);
      if (!entry) {
        throw new Error("Missing workbook part: " + name);
      }

      var local = entry.localOffset;
      if (getUint32(bytes, local) !== 0x04034b50) {
        throw new Error("Workbook part is not readable: " + name);
      }

      var nameLength = getUint16(bytes, local + 26);
      var extraLength = getUint16(bytes, local + 28);
      var dataStart = local + 30 + nameLength + extraLength;
      var compressed = bytes.slice(dataStart, dataStart + entry.compressedSize);

      if (entry.method === 0) {
        return compressed;
      }

      if (entry.method === 8) {
        return new Uint8Array(await inflateDeflateRaw(compressed));
      }

      throw new Error("Unsupported workbook compression method: " + entry.method);
    }

    return {
      has: function (name) {
        return entries.has(name);
      },
      text: async function (name) {
        return decoder.decode(await entryBytes(name));
      }
    };
  }

  function findEndOfCentralDirectory(bytes) {
    var min = Math.max(0, bytes.length - 66000);
    for (var i = bytes.length - 22; i >= min; i -= 1) {
      if (getUint32(bytes, i) === 0x06054b50) {
        return i;
      }
    }
    return -1;
  }

  async function inflateDeflateRaw(bytes) {
    var formats = ["deflate-raw", "deflate"];
    var lastError = null;

    for (var i = 0; i < formats.length; i += 1) {
      try {
        var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(formats[i]));
        return await new Response(stream).arrayBuffer();
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError || new Error("The workbook could not be decompressed.");
  }

  function getUint16(bytes, offset) {
    return bytes[offset] | (bytes[offset + 1] << 8);
  }

  function getUint32(bytes, offset) {
    return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
  }

  function parseSharedStrings(xmlText) {
    var doc = xml(xmlText);
    return byLocalName(doc, "si").map(function (node) {
      return node.textContent.trim();
    });
  }

  function parseWorkbookSheets(workbookXml, relsXml) {
    var workbookDoc = xml(workbookXml);
    var relsDoc = xml(relsXml);
    var rels = new Map();

    byLocalName(relsDoc, "Relationship").forEach(function (rel) {
      rels.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
    });

    return byLocalName(workbookDoc, "sheet").map(function (sheet) {
      var relId = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheet.getAttribute("r:id");
      var target = rels.get(relId);
      return {
        name: sheet.getAttribute("name"),
        path: normalizeWorkbookPath(target)
      };
    });
  }

  function normalizeWorkbookPath(target) {
    if (!target) {
      return "";
    }
    var path = target.replace(/^\/+/, "");
    if (path.indexOf("xl/") === 0) {
      return path;
    }
    return "xl/" + path.replace(/^\.\//, "");
  }

  function parseSheetRows(xmlText, sharedStrings) {
    var doc = xml(xmlText);
    return byLocalName(doc, "row").map(function (row) {
      var cells = rowChildren(row, "c");
      var values = [];
      cells.forEach(function (cell) {
        var ref = cell.getAttribute("r") || "A1";
        var index = columnIndex(ref);
        values[index] = cellText(cell, sharedStrings);
      });
      for (var i = 0; i < values.length; i += 1) {
        if (typeof values[i] === "undefined") {
          values[i] = "";
        }
      }
      return {
        rowNumber: Number(row.getAttribute("r")) || 0,
        values: values
      };
    });
  }

  function cellText(cell, sharedStrings) {
    var type = cell.getAttribute("t");
    if (type === "inlineStr") {
      return cell.textContent.trim();
    }

    var valueNode = firstChildByLocalName(cell, "v");
    if (!valueNode) {
      return "";
    }

    var raw = valueNode.textContent || "";
    if (type === "s") {
      return sharedStrings[Number(raw)] || "";
    }

    return raw;
  }

  function columnIndex(ref) {
    var letters = (ref.match(/^[A-Z]+/) || ["A"])[0];
    var number = 0;
    for (var i = 0; i < letters.length; i += 1) {
      number = number * 26 + letters.charCodeAt(i) - 64;
    }
    return number - 1;
  }

  function xml(text) {
    var doc = new DOMParser().parseFromString(text, "application/xml");
    var parserError = byLocalName(doc, "parsererror")[0];
    if (parserError) {
      throw new Error("Workbook XML could not be parsed.");
    }
    return doc;
  }

  function byLocalName(node, localName) {
    return Array.prototype.slice.call(node.getElementsByTagName("*")).filter(function (child) {
      return child.localName === localName;
    });
  }

  function rowChildren(node, localName) {
    return Array.prototype.slice.call(node.children).filter(function (child) {
      return child.localName === localName;
    });
  }

  function firstChildByLocalName(node, localName) {
    return rowChildren(node, localName)[0] || null;
  }

  function buildPricingModel(workbook, fileName) {
    var sheets = workbook.sheets;
    var productModel = parseProductMap(sheets["Product Map"] || []);
    var rollupLookup = parseTodayRollup(sheets["Today's Rollup"] || []);
    var priceRows = [];
    var ppcRows = [];
    var dailySheets = workbook.sheetNames.filter(function (name) {
      return /Daily Prices$/i.test(name);
    });
    var ppcSheets = workbook.sheetNames.filter(function (name) {
      return /Daily Price per Count$/i.test(name) || /Daily Price Per$/i.test(name);
    });

    dailySheets.forEach(function (sheetName) {
      priceRows = priceRows.concat(parseDailySheet(sheetName, sheets[sheetName], productModel, rollupLookup, "price"));
    });

    ppcSheets.forEach(function (sheetName) {
      ppcRows = ppcRows.concat(parseDailySheet(sheetName, sheets[sheetName], productModel, rollupLookup, "ppc"));
    });

    attachLowestRetailers(priceRows);
    var ppcGaps = buildPpcGaps(ppcRows);
    var retailerScores = buildRetailerScores(priceRows);
    var recommendations = buildRecommendations(priceRows, ppcGaps, retailerScores, fileName);
    var dates = unique(priceRows.flatMap(function (row) {
      return row.series.map(function (point) {
        return point.date;
      });
    })).sort();

    return {
      sheets: workbook.sheetNames,
      productMap: productModel.products,
      productLookup: productModel.lookup,
      rollupLookup: rollupLookup,
      priceRows: priceRows,
      ppcRows: ppcRows,
      ppcGaps: ppcGaps,
      retailerScores: retailerScores,
      recommendations: recommendations,
      dateRange: dates.length ? formatDateHeader(dates[0]) + " to " + formatDateHeader(dates[dates.length - 1]) : ""
    };
  }

  function parseProductMap(rows) {
    if (!rows.length) {
      return { products: [], lookup: new Map(), retailerNames: new Map() };
    }

    var headers = rows[0].values;
    var products = [];
    var lookup = new Map();
    var retailerNames = new Map();

    rows.slice(1).forEach(function (row) {
      var get = headerGetter(headers, row.values);
      var retailerId = clean(get("RetailerID"));
      var retailerName = clean(get("Retailername")) || RETAILER_BY_ID[retailerId] || "";
      var gtin = clean(get("GTIN Map")) || clean(get("Common UPC (from client)"));
      var commonUpc = clean(get("Common UPC (from client)"));
      var gtinRid = clean(get("GTIN_RID")) || (gtin && retailerId ? gtin + "_" + retailerId : "");
      var product = {
        retailerId: retailerId,
        retailerName: retailerName,
        channel: channelFor(retailerId, retailerName),
        brand: clean(get("Brand")),
        title: clean(get("Beacon Title")) || clean(get("Title (from client)")) || clean(get("Short(er) Title")),
        shortTitle: clean(get("Short(er) Title")),
        productCount: toNumber(get("Product Count")),
        retailerSku: clean(get("Retailer SKU (Stackline)")) || clean(get("Retailer SKU (provided by client)")),
        gtin: gtin,
        commonUpc: commonUpc,
        gtinRid: gtinRid,
        guardrail: toNumber(get("NonPromo MAP")),
        url: clean(get("URL"))
      };

      if (!product.gtin || !product.retailerId) {
        return;
      }

      products.push(product);
      retailerNames.set(product.retailerId, product.retailerName);
      addProductKeys(lookup, product);
    });

    return { products: products, lookup: lookup, retailerNames: retailerNames };
  }

  function addProductKeys(lookup, product) {
    var keys = [];
    if (product.gtin && product.retailerId) {
      keys.push(product.gtin + "|" + product.retailerId);
    }
    if (product.commonUpc && product.retailerId) {
      keys.push(product.commonUpc + "|" + product.retailerId);
    }
    if (product.gtinRid) {
      keys.push(product.gtinRid);
    }
    if (product.retailerSku && product.retailerId) {
      keys.push(product.retailerSku + "|" + product.retailerId);
    }
    keys.forEach(function (key) {
      lookup.set(key, product);
    });
  }

  function parseTodayRollup(rows) {
    var lookup = new Map();
    var headerIndex = rows.findIndex(function (row) {
      return row.values.some(function (value) {
        return clean(value) === "GTIN (UPC)";
      });
    });

    if (headerIndex < 0) {
      return lookup;
    }

    var headers = rows[headerIndex].values.map(clean);
    var gtinIndex = headers.indexOf("GTIN (UPC)");
    var firstMoverIndex = headers.indexOf("First Mover of Violation Period");
    var lowestRetailerIndex = headers.indexOf("Lowest Price Retailer");

    rows.slice(headerIndex + 1).forEach(function (row) {
      var gtin = clean(row.values[gtinIndex]);
      if (!gtin) {
        return;
      }
      lookup.set(gtin, {
        firstMover: clean(row.values[firstMoverIndex]),
        lowestPriceRetailer: clean(row.values[lowestRetailerIndex])
      });
    });

    return lookup;
  }

  function parseDailySheet(sheetName, rows, productModel, rollupLookup, mode) {
    if (!rows || !rows.length) {
      return [];
    }

    var headerIndex = rows.findIndex(function (row) {
      return row.values.some(function (value) {
        return clean(value) === "GTIN (UPC)";
      });
    });

    if (headerIndex < 0) {
      return [];
    }

    var headers = rows[headerIndex].values.map(clean);
    var dateColumns = headers
      .map(function (header, index) {
        return { header: header, index: index };
      })
      .filter(function (item) {
        return /^20\d{6}$/.test(item.header);
      });

    var brandFallback = brandFromSheet(sheetName);
    var parsed = [];

    rows.slice(headerIndex + 1).forEach(function (row, rowOffset) {
      var get = headerGetter(headers, row.values);
      var gtin = clean(get("GTIN (UPC)"));
      var retailerId = clean(get("Retailer ID"));
      if (!gtin || !retailerId) {
        return;
      }

      var product = findProduct(productModel.lookup, gtin, retailerId, clean(get("GTIN_RID")), clean(get("Retailer SKU")));
      var retailerName = clean(get("Retailer Name")) || (product && product.retailerName) || RETAILER_BY_ID[retailerId] || "Retailer " + retailerId;
      var count = toNumber(get("Product Count")) || (product && product.productCount) || null;
      var guardrail = (mode === "ppc")
        ? firstNumber(toNumber(get("MAP / Count")), ((product && product.guardrail && count) ? product.guardrail / count : null))
        : firstNumber(toNumber(get("MAP")), product ? product.guardrail : null);

      var series = dateColumns.map(function (dateColumn) {
        var raw = clean(row.values[dateColumn.index]);
        return {
          date: dateColumn.header,
          raw: raw,
          value: toNumber(raw),
          oos: /^OOS$/i.test(raw)
        };
      });

      var latestPoint = latestSeriesPoint(series);
      var latestValue = firstNumber(toNumber(get("Lowest UPC Price Latest Day")), latestPoint ? latestPoint.value : null);
      var latestRaw = (latestPoint && latestPoint.raw) || "";
      var daysViolation = parseDays(get("Consecutive Days in Violation"));
      var daysAtPrice = parseDays(get("Consecutive Days at Today's Price"));
      var rollup = rollupLookup.get(gtin) || {};
      var title = clean(get("Title")) || (product && product.title) || gtin;
      var brand = clean(get("Brand")) || (product && product.brand) || brandFallback;
      var channel = channelFor(retailerId, retailerName);
      var belowGuardrail = Boolean(latestValue !== null && guardrail !== null && latestValue < guardrail - 0.004);
      var gapValue = belowGuardrail ? guardrail - latestValue : 0;
      var gapPct = belowGuardrail && guardrail ? gapValue / guardrail : 0;

      parsed.push({
        id: mode + "-" + sheetName + "-" + gtin + "-" + retailerId + "-" + rowOffset,
        mode: mode,
        sheetName: sheetName,
        gtin: gtin,
        retailerId: retailerId,
        retailerName: retailerName,
        channel: channel,
        brand: brand,
        title: title,
        shortTitle: (product && product.shortTitle) || title,
        productCount: count,
        sizeGroup: clean(get("Size Group")) || "TBD",
        url: clean(get("PDP URL")) || (product && product.url) || "",
        retailerSku: clean(get("Retailer SKU")) || (product && product.retailerSku) || "",
        guardrail: guardrail,
        latestValue: latestValue,
        latestRaw: latestRaw,
        latestDate: latestPoint ? latestPoint.date : "",
        belowGuardrail: belowGuardrail,
        gapValue: gapValue,
        gapPct: gapPct,
        daysViolation: daysViolation,
        daysAtPrice: daysAtPrice,
        isOos: /^OOS$/i.test(latestRaw) || /^OOS$/i.test(clean(get("Consecutive Days in Violation"))),
        firstMover: rollup.firstMover || "",
        rollupLowestRetailer: rollup.lowestPriceRetailer || "",
        latestDayLowest: truthy(get("Latest Day at Lowest price?")),
        latestDayViolation: truthy(get("Latest Day in MAP Violation?")),
        series: series
      });
    });

    return parsed;
  }

  function findProduct(lookup, gtin, retailerId, gtinRid, retailerSku) {
    return lookup.get(gtinRid) ||
      lookup.get(gtin + "|" + retailerId) ||
      lookup.get(retailerSku + "|" + retailerId) ||
      null;
  }

  function attachLowestRetailers(rows) {
    var groups = groupBy(rows.filter(function (row) {
      return row.latestValue !== null;
    }), function (row) {
      return row.gtin;
    });

    Object.keys(groups).forEach(function (gtin) {
      var sorted = groups[gtin].slice().sort(function (a, b) {
        return a.latestValue - b.latestValue;
      });
      var lowest = sorted[0];
      sorted.forEach(function (row) {
        row.currentLowestRetailer = lowest ? lowest.retailerName : "";
        row.currentLowestPrice = lowest ? lowest.latestValue : null;
        row.isCurrentLowest = lowest ? row.id === lowest.id : false;
      });
    });
  }

  function buildPpcGaps(ppcRows) {
    var groups = groupBy(ppcRows.filter(function (row) {
      return row.latestValue !== null;
    }), function (row) {
      return row.gtin;
    });

    return Object.keys(groups).map(function (gtin) {
      var rows = groups[gtin];
      if (rows.length < 2) {
        return null;
      }
      var sorted = rows.slice().sort(function (a, b) {
        return a.latestValue - b.latestValue;
      });
      var low = sorted[0];
      var high = sorted[sorted.length - 1];
      var gapPct = low.latestValue ? (high.latestValue - low.latestValue) / low.latestValue : 0;
      return {
        id: "ppc-gap-" + gtin,
        gtin: gtin,
        brand: low.brand,
        title: low.title,
        lowRetailer: low.retailerName,
        highRetailer: high.retailerName,
        lowValue: low.latestValue,
        highValue: high.latestValue,
        gapPct: gapPct,
        rows: rows
      };
    }).filter(Boolean).sort(function (a, b) {
      return b.gapPct - a.gapPct;
    });
  }

  function buildRetailerScores(rows) {
    var groups = groupBy(rows, function (row) {
      return row.retailerName;
    });

    return Object.keys(groups).map(function (retailer) {
      var retailerRows = groups[retailer];
      var total = retailerRows.length || 1;
      var below = retailerRows.filter(function (row) { return row.belowGuardrail; }).length;
      var extended = retailerRows.filter(function (row) { return row.daysViolation >= 7; }).length;
      var oos = retailerRows.filter(function (row) { return row.isOos; }).length;
      var score = Math.max(0, Math.round(100 - (below / total) * 68 - (extended / total) * 22 - (oos / total) * 10));
      return {
        retailer: retailer,
        channel: retailerRows[0] ? retailerRows[0].channel : "",
        total: total,
        below: below,
        extended: extended,
        oos: oos,
        score: score
      };
    }).sort(function (a, b) {
      return a.score - b.score;
    });
  }

  function buildRecommendations(priceRows, ppcGaps, retailerScores, fileName) {
    var saved = loadSavedApprovalStatuses();
    var recommendations = [];
    var belowRows = priceRows
      .filter(function (row) { return row.belowGuardrail; })
      .sort(function (a, b) {
        return b.daysViolation - a.daysViolation || b.gapPct - a.gapPct;
      })
      .slice(0, 24);

    belowRows.forEach(function (row) {
      var severity = severityForRow(row);
      var id = "below-" + row.gtin + "-" + row.retailerId;
      recommendations.push(withSavedStatus(saved, {
        id: id,
        type: "Guardrail flag",
        severity: severity,
        brand: row.brand,
        retailer: row.retailerName,
        channel: row.channel,
        title: "Review below-guardrail pricing",
        scope: row.brand + " | " + row.retailerName + " | " + row.gtin,
        evidence: formatMoney(row.latestValue) + " latest price vs " + formatMoney(row.guardrail) + " guardrail; " + row.daysViolation + " consecutive days" + (row.firstMover ? "; first mover: " + row.firstMover : ""),
        businessProblem: "The retailer/SKU combination is priced below the NonPromo MAP guardrail in the workbook and may create pricing discipline risk.",
        dataRequired: "Product Map, daily price history, NonPromo MAP guardrail, retailer ID, price-per-count sheet, and product URL evidence.",
        analyticalApproach: "Compare the latest crawled price against the guardrail, review consecutive days below guardrail, identify lowest-price position, and check whether the pattern is isolated or repeated.",
        expectedImpact: "Sharper RGM visibility into corridor pressure and a cleaner queue for commercial review.",
        risks: "Crawl error, temporary promotion, OOS effects, retailer context not captured in the workbook, or legal sensitivity around guardrail interpretation.",
        volumeImpact: "Hypothesis only: lower visible pricing may support short-term volume, but the workbook does not contain units to measure lift.",
        marginImpact: "Hypothesis only: below-guardrail pricing may pressure margin and channel discipline; margin data is not present in this workbook.",
        competitiveResponse: "Monitor whether other retailers follow the low price and whether the first mover changes over the violation period.",
        retailerImplications: "Review with RGM before any retailer-facing action. The tool should flag the issue, not draft outreach language.",
        consumerImplications: "Consumers may anchor to the lower price if it persists, raising future price resistance.",
        approval: "RGM approval required. Legal review required if the recommendation leads to external action or corridor/guardrail language.",
        legalReview: true,
        source: fileName,
        raw: row
      }));
    });

    ppcGaps.filter(function (gap) {
      return gap.gapPct >= 0.2;
    }).slice(0, 8).forEach(function (gap) {
      var id = "ppc-" + gap.gtin;
      recommendations.push(withSavedStatus(saved, {
        id: id,
        type: "Pack-price architecture",
        severity: gap.gapPct >= 0.5 ? "High" : "Medium",
        brand: gap.brand,
        retailer: "Multiple retailers",
        channel: "Cross-channel",
        title: "Review price-per-count gap",
        scope: gap.brand + " | " + gap.gtin,
        evidence: formatMoney(gap.lowValue) + " per count at " + gap.lowRetailer + " vs " + formatMoney(gap.highValue) + " at " + gap.highRetailer,
        businessProblem: "Price-per-count spread is wide enough to create pack-price architecture tension across retailers or channels.",
        dataRequired: "Daily price-per-count sheet, product count, retailer mapping, and guardrail reference.",
        analyticalApproach: "Compare latest price per count across retailers for the same GTIN and flag large relative spreads.",
        expectedImpact: "Improves visibility into consumer value signals and potential pack architecture inconsistencies.",
        risks: "Different pack formats, marketplace listings, OOS, and retailer-specific assortment rules may explain the spread.",
        volumeImpact: "Hypothesis only: shoppers may shift to the retailer or pack with the most favorable price-per-count signal.",
        marginImpact: "Hypothesis only: large price-per-count gaps may indicate margin leakage or value perception risk.",
        competitiveResponse: "Track whether low price-per-count positions become the reference point for other retailers.",
        retailerImplications: "RGM should confirm whether the spread reflects intended architecture before any commercial action.",
        consumerImplications: "Consumers comparing unit value may see inconsistent value across retailers.",
        approval: "RGM approval required before action. Legal review required if action references guardrails or retailer-specific conduct.",
        legalReview: true,
        source: fileName,
        raw: gap
      }));
    });

    retailerScores.filter(function (score) {
      return score.score < 72;
    }).slice(0, 6).forEach(function (score) {
      var id = "retailer-score-" + normalizeId(score.retailer);
      recommendations.push(withSavedStatus(saved, {
        id: id,
        type: "Retailer discipline",
        severity: score.score < 55 ? "High" : "Medium",
        brand: "Portfolio",
        retailer: score.retailer,
        channel: score.channel,
        title: "Prioritize retailer discipline review",
        scope: score.retailer + " | " + score.channel,
        evidence: score.below + " below-guardrail flags, " + score.extended + " extended flags, discipline score " + score.score,
        businessProblem: "A retailer has a concentration of guardrail flags or extended below-guardrail periods across the loaded workbook.",
        dataRequired: "Daily prices, guardrail values, retailer mapping, and consecutive days in violation.",
        analyticalApproach: "Score retailer discipline using below-guardrail rate, extended violation rate, and OOS flags.",
        expectedImpact: "Helps RGM prioritize where to spend review time across retailers and channels.",
        risks: "The score is directional and should not be treated as a contractual or legal finding.",
        volumeImpact: "Hypothesis only: repeated low prices can influence cross-retailer demand patterns.",
        marginImpact: "Hypothesis only: repeated low prices may indicate margin pressure or channel conflict.",
        competitiveResponse: "Other retailers may react if low pricing becomes visible and persistent.",
        retailerImplications: "RGM should review the retailer pattern internally before any follow-up.",
        consumerImplications: "Persistent lower prices may reset consumer expectations.",
        approval: "RGM approval required. Legal review required for any recommendation that moves toward external action.",
        legalReview: true,
        source: fileName,
        raw: score
      }));
    });

    priceRows.filter(function (row) {
      return row.isOos;
    }).slice(0, 5).forEach(function (row) {
      var id = "oos-" + row.gtin + "-" + row.retailerId;
      recommendations.push(withSavedStatus(saved, {
        id: id,
        type: "Data and availability",
        severity: "Watch",
        brand: row.brand,
        retailer: row.retailerName,
        channel: row.channel,
        title: "Confirm OOS before pricing readout",
        scope: row.brand + " | " + row.retailerName + " | " + row.gtin,
        evidence: "Latest workbook status shows OOS or non-price value.",
        businessProblem: "OOS status can distort lowest-price, price-per-count, and guardrail interpretation.",
        dataRequired: "Daily price sheet, PDP availability evidence, and retailer/product mapping.",
        analyticalApproach: "Separate availability gaps from price discipline flags before using the row in an RGM recommendation.",
        expectedImpact: "Reduces false positives and keeps the review queue commercially credible.",
        risks: "Availability status may change quickly or vary by geography.",
        volumeImpact: "Hypothesis only: OOS can suppress observed demand and hide pricing effects.",
        marginImpact: "Hypothesis only: no margin conclusion should be drawn from OOS rows alone.",
        competitiveResponse: "Competitors may gain share if availability gaps persist.",
        retailerImplications: "Review as an availability/data issue before any price-discipline conclusion.",
        consumerImplications: "Consumers may switch if the product remains unavailable.",
        approval: "RGM approval required before any action. Legal review is not triggered by OOS alone unless tied to an external recommendation.",
        legalReview: false,
        source: fileName,
        raw: row
      }));
    });

    return recommendations.sort(function (a, b) {
      return severityWeight(b.severity) - severityWeight(a.severity);
    });
  }

  function withSavedStatus(saved, recommendation) {
    recommendation.status = saved[recommendation.id] || "Draft";
    return recommendation;
  }

  function severityForRow(row) {
    if (row.daysViolation >= 14 || row.gapPct >= 0.15) {
      return "High";
    }
    if (row.daysViolation >= 3 || row.gapPct >= 0.05) {
      return "Medium";
    }
    return "Watch";
  }

  function severityWeight(severity) {
    return { High: 3, Medium: 2, Watch: 1, Low: 0 }[severity] || 0;
  }

  function populateFilters() {
    fillSelect(els.brandFilter, ["All"].concat(unique(state.priceRows.map(function (row) { return row.brand; })).sort()));
    fillSelect(els.retailerFilter, ["All"].concat(unique(state.priceRows.map(function (row) { return row.retailerName; })).sort()));
    fillSelect(els.channelFilter, ["All"].concat(unique(state.priceRows.map(function (row) { return row.channel; })).sort()));
    fillSelect(els.severityFilter, ["All", "High", "Medium", "Watch", "Low"]);
    els.audienceFilter.value = state.filters.audience;
  }

  function fillSelect(select, values) {
    select.innerHTML = values.map(function (value) {
      return '<option value="' + escapeHtml(value) + '">' + escapeHtml(value) + "</option>";
    }).join("");
    if (values.indexOf(state.filters[select.id.replace("-filter", "")]) >= 0) {
      select.value = state.filters[select.id.replace("-filter", "")];
    }
  }

  function render() {
    if (!state.workbook) {
      return;
    }

    els.cockpitView.classList.toggle("hidden", state.activeTab !== "cockpit");
    els.recommendationsView.classList.toggle("hidden", state.activeTab !== "recommendations");
    els.skuView.classList.toggle("hidden", state.activeTab !== "sku");
    els.strategistView.classList.toggle("hidden", state.activeTab !== "strategist");
    els.approvalsView.classList.toggle("hidden", state.activeTab !== "approvals");
    els.dataView.classList.toggle("hidden", state.activeTab !== "data");

    if (state.activeTab === "cockpit") {
      renderCockpit();
    } else if (state.activeTab === "recommendations") {
      renderRecommendations();
    } else if (state.activeTab === "sku") {
      renderSkuExplorer();
    } else if (state.activeTab === "strategist") {
      renderStrategist();
    } else if (state.activeTab === "approvals") {
      renderApprovals();
    } else if (state.activeTab === "data") {
      renderDataMap();
    }
  }

  function filteredPriceRows() {
    return state.priceRows.filter(function (row) {
      return matchesBaseFilters(row);
    });
  }

  function filteredRecommendations() {
    return state.recommendations.filter(function (rec) {
      if (state.filters.brand !== "All" && rec.brand !== state.filters.brand && rec.brand !== "Portfolio") {
        return false;
      }
      if (state.filters.retailer !== "All" && rec.retailer !== state.filters.retailer && rec.retailer !== "Multiple retailers") {
        return false;
      }
      if (state.filters.channel !== "All" && rec.channel !== state.filters.channel && rec.channel !== "Cross-channel") {
        return false;
      }
      if (state.filters.severity !== "All" && rec.severity !== state.filters.severity) {
        return false;
      }
      return true;
    });
  }

  function matchesBaseFilters(row) {
    if (state.filters.brand !== "All" && row.brand !== state.filters.brand) {
      return false;
    }
    if (state.filters.retailer !== "All" && row.retailerName !== state.filters.retailer) {
      return false;
    }
    if (state.filters.channel !== "All" && row.channel !== state.filters.channel) {
      return false;
    }
    return true;
  }

  function renderCockpit() {
    var rows = filteredPriceRows();
    var below = rows.filter(function (row) { return row.belowGuardrail; });
    var extended = rows.filter(function (row) { return row.daysViolation >= 7; });
    var oos = rows.filter(function (row) { return row.isOos; });
    var ppcGaps = state.ppcGaps.filter(function (gap) {
      return gap.gapPct >= 0.2 && (state.filters.brand === "All" || gap.brand === state.filters.brand);
    });
    var scores = filteredRetailerScores();
    var avgScore = scores.length ? Math.round(scores.reduce(function (sum, score) { return sum + score.score; }, 0) / scores.length) : 0;
    var selectedRow = selectedPriceRow(rows);
    var selectedRec = selectedRow ? recommendationForRow(selectedRow) : null;
    var highRecs = filteredRecommendations().filter(function (rec) { return rec.severity === "High"; }).length;

    els.cockpitView.innerHTML = [
      '<section class="command-strip">',
      '<div>',
      '<span class="kicker">Loaded workbook</span>',
      '<strong>' + escapeHtml(state.fileName || "Workbook") + "</strong>",
      '<span>' + escapeHtml(state.dateRange || "Latest available period") + "</span>",
      "</div>",
      '<div class="strip-actions">',
      '<span class="status-chip">Guardrail = review signal</span>',
      '<span class="status-chip">No auto execution</span>',
      '<button class="button ghost" id="export-brief">Export brief</button>',
      "</div>",
      "</section>",
      '<section class="summary-grid">',
      metric("Needs RGM review", highRecs, "High-severity recommendations in the current filter", highRecs ? "danger" : "good"),
      metric("Guardrail flags", below.length, "Latest price below NonPromo MAP guardrail", below.length ? "danger" : "good"),
      metric("Extended flags", extended.length, "7+ consecutive days below guardrail", extended.length ? "warning" : "good"),
      metric("Retailer score", avgScore || "n/a", "Average discipline score for selected filters", avgScore >= 80 ? "good" : "info"),
      metric("PPC gaps", ppcGaps.length, "20%+ same-GTIN price-per-count spread", ppcGaps.length ? "warning" : "good"),
      "</section>",
      '<section class="cockpit-board">',
      '<aside class="section-panel queue-panel">',
      sectionHeader("Priority Queue", below.length + " guardrail flags after filters", ""),
      renderIssueQueue(below, oos),
      "</aside>",
      '<section class="section-panel evidence-panel">',
      sectionHeader("Evidence", selectedRow ? compactTitle(selectedRow) : "Select an issue", selectedRow ? selectedRow.retailerName + " | " + selectedRow.channel : "Choose a priority card to inspect", ""),
      selectedRow ? renderIssueEvidence(selectedRow) : '<div class="empty-state">Select a card in the queue to inspect evidence.</div>',
      "</section>",
      '<aside class="section-panel decision-panel">',
      sectionHeader("Decision", selectedRec ? selectedRec.title : "RGM review", selectedRec ? selectedRec.scope : "No issue selected", ""),
      selectedRec ? renderDecisionPanel(selectedRec, selectedRow) : '<div class="empty-state">A recommendation appears here after selecting a flagged issue.</div>',
      "</aside>",
      "</section>",
      '<section class="content-grid">',
      '<div class="section-panel">',
      sectionHeader("Retailer Discipline", "Directional scoring for review prioritization", ""),
      renderRetailerBars(scores),
      "</div>",
      '<div class="section-panel">',
      sectionHeader("Price-Per-Count Architecture", "Largest same-GTIN spreads across retailers", ""),
      renderPpcTable(ppcGaps),
      "</div>",
      "</section>"
    ].join("");

    document.querySelectorAll("[data-price-row-id]").forEach(function (card) {
      card.addEventListener("click", function () {
        state.selectedPriceRowId = card.getAttribute("data-price-row-id");
        renderCockpit();
      });
    });

    document.querySelectorAll("[data-set-rec-status]").forEach(function (button) {
      button.addEventListener("click", function () {
        var rec = state.recommendations.find(function (item) {
          return item.id === button.getAttribute("data-set-rec-status");
        });
        if (rec) {
          rec.status = button.getAttribute("data-status-value");
          saveApprovalStatus(rec.id, rec.status);
          renderCockpit();
        }
      });
    });

    document.querySelectorAll("[data-open-rec]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.selectedRecId = button.getAttribute("data-open-rec");
        state.activeTab = "recommendations";
        setActiveTabButton("recommendations");
        render();
      });
    });

    var exportBrief = document.getElementById("export-brief");
    if (exportBrief) {
      exportBrief.addEventListener("click", exportExecutiveBrief);
    }
  }

  function selectedPriceRow(rows) {
    return rows.find(function (item) {
      return item.id === state.selectedPriceRowId;
    }) || rows.find(function (item) {
      return item.belowGuardrail;
    }) || rows[0] || null;
  }

  function renderIssueQueue(belowRows, oosRows) {
    var priorityRows = belowRows.slice().sort(function (a, b) {
      return severityWeight(severityForRow(b)) - severityWeight(severityForRow(a)) ||
        b.daysViolation - a.daysViolation ||
        b.gapPct - a.gapPct;
    }).slice(0, 18);

    if (!priorityRows.length) {
      return [
        '<div class="empty-state compact">',
        "<strong>No active guardrail flags.</strong>",
        "<span>" + escapeHtml(oosRows.length ? oosRows.length + " OOS rows remain available in SKU Explorer." : "Selected filters look clean.") + "</span>",
        "</div>"
      ].join("");
    }

    return [
      '<div class="issue-list">',
      priorityRows.map(renderIssueCard).join(""),
      "</div>"
    ].join("");
  }

  function renderIssueCard(row) {
    var selected = row.id === state.selectedPriceRowId ? " selected" : "";
    var severity = severityForRow(row);
    return [
      '<button class="issue-card' + selected + '" type="button" data-price-row-id="' + escapeHtml(row.id) + '">',
      '<span class="issue-topline">',
      severityPill(severity),
      '<span>' + escapeHtml(row.retailerName) + "</span>",
      "</span>",
      '<strong>' + escapeHtml(compactTitle(row)) + "</strong>",
      '<span class="issue-meta">' + escapeHtml(row.brand) + " | " + escapeHtml(row.channel) + "</span>",
      '<span class="issue-metrics">',
      '<span><b>' + escapeHtml(String(row.daysViolation || 0)) + "</b> days</span>",
      '<span><b>' + formatMoney(row.latestValue) + "</b> latest</span>",
      '<span><b>' + formatPercent(row.gapPct) + "</b> gap</span>",
      "</span>",
      "</button>"
    ].join("");
  }

  function renderIssueEvidence(row) {
    var lowest = row.currentLowestRetailer || row.rollupLowestRetailer || "n/a";
    return [
      '<div class="evidence-stack">',
      '<div class="evidence-hero">',
      '<div>' + severityPill(severityForRow(row)) + '<h3>' + escapeHtml(compactTitle(row)) + "</h3><p>" + escapeHtml(row.title) + "</p></div>",
      '<div class="guardrail-delta"><span>Below guardrail</span><strong>' + formatMoney(row.gapValue) + "</strong></div>",
      "</div>",
      '<div class="evidence-grid">',
      evidenceTile("Latest", formatMoney(row.latestValue), formatDateHeader(row.latestDate)),
      evidenceTile("Guardrail", formatMoney(row.guardrail), "NonPromo MAP guardrail"),
      evidenceTile("Days", String(row.daysViolation || 0), "Consecutive below flag"),
      evidenceTile("First mover", row.firstMover || "n/a", "From Today's Rollup"),
      evidenceTile("Lowest", lowest, row.isCurrentLowest ? "Current lowest row" : "Current or rollup signal"),
      evidenceTile("Status", row.isOos ? "OOS" : "In price feed", row.latestRaw || "Numeric price"),
      "</div>",
      '<div class="chart-card">' + buildLineChart(row) + "</div>",
      '<div class="evidence-actions">',
      row.url ? '<a class="button ghost" target="_blank" rel="noopener" href="' + escapeAttribute(row.url) + '">Open PDP</a>' : "",
      '<button class="button ghost" type="button" data-open-rec="' + escapeHtml(recommendationIdForRow(row)) + '">View recommendation</button>',
      "</div>",
      "</div>"
    ].join("");
  }

  function evidenceTile(label, value, note) {
    return [
      '<div class="evidence-tile">',
      '<span>' + escapeHtml(label) + "</span>",
      '<strong>' + escapeHtml(value || "n/a") + "</strong>",
      '<em>' + escapeHtml(note || "") + "</em>",
      "</div>"
    ].join("");
  }

  function renderDecisionPanel(rec, row) {
    var nextStage = rec.legalReview ? "Legal Review" : "RGM Review";
    var rowNote = row ? formatMoney(row.latestValue) + " latest vs " + formatMoney(row.guardrail) + " guardrail" : rec.evidence;
    return [
      '<div class="decision-stack">',
      '<div class="decision-summary">',
      '<span class="approval-pill">' + escapeHtml(rec.type) + "</span>",
      '<span class="approval-pill">' + escapeHtml(rec.status) + "</span>",
      rec.legalReview ? '<span class="approval-pill legal">Legal trigger</span>' : "",
      "</div>",
      '<p class="decision-copy">' + escapeHtml(rec.businessProblem) + "</p>",
      '<div class="mini-checklist">',
      miniCheck("Evidence", rowNote),
      miniCheck("Impact", rec.expectedImpact),
      miniCheck("Risk", rec.risks),
      miniCheck("Approval", "RGM approval required. No automatic price execution."),
      "</div>",
      '<div class="decision-actions">',
      '<button class="button" type="button" data-set-rec-status="' + escapeHtml(rec.id) + '" data-status-value="' + escapeHtml(nextStage) + '">Move to ' + escapeHtml(nextStage) + "</button>",
      '<button class="button ghost" type="button" data-set-rec-status="' + escapeHtml(rec.id) + '" data-status-value="Revised">Needs revision</button>',
      "</div>",
      "</div>"
    ].join("");
  }

  function miniCheck(label, text) {
    return [
      '<div class="mini-check">',
      '<span>' + escapeHtml(label) + "</span>",
      '<p>' + escapeHtml(text) + "</p>",
      "</div>"
    ].join("");
  }

  function recommendationIdForRow(row) {
    if (!row) {
      return "";
    }
    if (row.belowGuardrail) {
      return "below-" + row.gtin + "-" + row.retailerId;
    }
    if (row.isOos) {
      return "oos-" + row.gtin + "-" + row.retailerId;
    }
    return "";
  }

  function recommendationForRow(row) {
    var id = recommendationIdForRow(row);
    return state.recommendations.find(function (rec) {
      return rec.id === id;
    }) || filteredRecommendations()[0] || null;
  }

  function compactTitle(row) {
    return row.shortTitle || row.title || row.gtin || "SKU";
  }

  function setActiveTabButton(tabName) {
    document.querySelectorAll(".tab").forEach(function (tab) {
      tab.classList.toggle("active", tab.getAttribute("data-tab") === tabName);
    });
  }

  function filteredRetailerScores() {
    return state.retailerScores.filter(function (score) {
      if (state.filters.retailer !== "All" && score.retailer !== state.filters.retailer) {
        return false;
      }
      if (state.filters.channel !== "All" && score.channel !== state.filters.channel) {
        return false;
      }
      return true;
    });
  }

  function metric(label, value, detail, tone) {
    return [
      '<article class="metric ' + tone + '">',
      '<span>' + escapeHtml(label) + "</span>",
      "<strong>" + escapeHtml(String(value)) + "</strong>",
      '<span>' + escapeHtml(detail) + "</span>",
      "</article>"
    ].join("");
  }

  function sectionHeader(title, subtitle, actionHtml) {
    return [
      '<div class="section-header">',
      "<div><h2>" + escapeHtml(title) + "</h2><p>" + escapeHtml(subtitle) + "</p></div>",
      actionHtml ? '<div class="actions-row">' + actionHtml + "</div>" : "",
      "</div>"
    ].join("");
  }

  function renderGuardrailTable(rows) {
    var sorted = rows.slice().sort(function (a, b) {
      return b.daysViolation - a.daysViolation || b.gapPct - a.gapPct;
    }).slice(0, 16);

    if (!sorted.length) {
      return '<div class="empty-state">No below-guardrail rows for the selected filters.</div>';
    }

    return [
      '<div class="table-wrap"><table><thead><tr>',
      "<th>Severity</th><th>Brand</th><th>Retailer</th><th>SKU / GTIN</th><th>Latest</th><th>Guardrail</th><th>Days</th><th>First mover</th><th>Lowest</th>",
      "</tr></thead><tbody>",
      sorted.map(function (row) {
        return [
          '<tr class="row-action" data-price-row-id="' + escapeHtml(row.id) + '">',
          "<td>" + severityPill(severityForRow(row)) + "</td>",
          "<td>" + escapeHtml(row.brand) + "</td>",
          "<td>" + escapeHtml(row.retailerName) + '<br><span class="fine-print">' + escapeHtml(row.channel) + "</span></td>",
          "<td>" + escapeHtml(row.shortTitle || row.title) + '<br><span class="fine-print">' + escapeHtml(row.gtin) + "</span></td>",
          "<td>" + formatMoney(row.latestValue) + "</td>",
          "<td>" + formatMoney(row.guardrail) + "</td>",
          "<td>" + escapeHtml(String(row.daysViolation || 0)) + "</td>",
          "<td>" + escapeHtml(row.firstMover || "n/a") + "</td>",
          "<td>" + escapeHtml(row.currentLowestRetailer || row.rollupLowestRetailer || "") + "</td>",
          "</tr>"
        ].join("");
      }).join(""),
      "</tbody></table></div>"
    ].join("");
  }

  function renderTrend() {
    var rows = filteredPriceRows();
    var row = rows.find(function (item) {
      return item.id === state.selectedPriceRowId;
    }) || rows.find(function (item) {
      return item.belowGuardrail;
    }) || rows[0];

    if (!row) {
      return '<div class="empty-state">Select a flagged row to view a price trend.</div>';
    }

    var chart = buildLineChart(row);
    return [
      chart,
      '<div class="trend-meta">',
      '<div><span>Brand</span><br><strong>' + escapeHtml(row.brand) + "</strong></div>",
      '<div><span>Retailer</span><br><strong>' + escapeHtml(row.retailerName) + "</strong></div>",
      '<div><span>Latest price</span><br><strong>' + formatMoney(row.latestValue) + "</strong></div>",
      '<div><span>Guardrail</span><br><strong>' + formatMoney(row.guardrail) + "</strong></div>",
      "</div>",
      '<p class="fine-print">' + escapeHtml(row.title) + "</p>"
    ].join("");
  }

  function buildLineChart(row) {
    var points = row.series.filter(function (point) {
      return point.value !== null;
    });

    if (points.length < 2) {
      return '<div class="empty-state">Not enough numeric daily prices for a line chart.</div>';
    }

    var width = 660;
    var height = 220;
    var pad = 26;
    var values = points.map(function (point) { return point.value; });
    if (row.guardrail !== null) {
      values.push(row.guardrail);
    }
    var min = Math.min.apply(Math, values);
    var max = Math.max.apply(Math, values);
    if (min === max) {
      min = min * 0.94;
      max = max * 1.06;
    }
    var innerWidth = width - pad * 2;
    var innerHeight = height - pad * 2;

    function x(index) {
      return pad + (points.length === 1 ? 0 : index / (points.length - 1)) * innerWidth;
    }

    function y(value) {
      return pad + ((max - value) / (max - min)) * innerHeight;
    }

    var path = points.map(function (point, index) {
      return (index === 0 ? "M" : "L") + x(index).toFixed(1) + " " + y(point.value).toFixed(1);
    }).join(" ");
    var guardrailY = row.guardrail !== null ? y(row.guardrail).toFixed(1) : null;
    var labels = [points[0], points[points.length - 1]];

    return [
      '<svg class="line-chart" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Daily price chart">',
      '<line class="chart-axis" x1="' + pad + '" y1="' + (height - pad) + '" x2="' + (width - pad) + '" y2="' + (height - pad) + '"></line>',
      guardrailY ? '<line class="chart-guardrail" x1="' + pad + '" y1="' + guardrailY + '" x2="' + (width - pad) + '" y2="' + guardrailY + '"></line>' : "",
      '<path class="chart-line" d="' + path + '"></path>',
      labels.map(function (point, index) {
        return '<text x="' + (index === 0 ? pad : width - pad - 66) + '" y="' + (height - 6) + '" font-size="12" fill="#61717d">' + escapeHtml(formatDateHeader(point.date)) + "</text>";
      }).join(""),
      guardrailY ? '<text x="' + (width - pad - 88) + '" y="' + (Number(guardrailY) - 7) + '" font-size="12" fill="#b42318">Guardrail</text>' : "",
      "</svg>"
    ].join("");
  }

  function renderRetailerBars(scores) {
    if (!scores.length) {
      return '<div class="empty-state">No retailer scores for the selected filters.</div>';
    }

    return [
      '<div class="bar-list">',
      scores.slice(0, 12).map(function (score) {
        var tone = score.score < 60 ? "bad" : score.score < 80 ? "warn" : "";
        return [
          '<div class="bar-row">',
          '<div class="bar-label" title="' + escapeHtml(score.retailer) + '">' + escapeHtml(score.retailer) + "</div>",
          '<div class="bar-track"><div class="bar-fill ' + tone + '" style="width:' + Math.max(4, score.score) + '%"></div></div>',
          '<div class="bar-score">' + score.score + "</div>",
          "</div>"
        ].join("");
      }).join(""),
      "</div>"
    ].join("");
  }

  function renderPpcTable(gaps) {
    var top = gaps.slice(0, 10);
    if (!top.length) {
      return '<div class="empty-state">No 20%+ price-per-count spreads for the selected filters.</div>';
    }

    return [
      '<div class="table-wrap"><table><thead><tr>',
      "<th>Brand</th><th>GTIN</th><th>Low PPC</th><th>High PPC</th><th>Gap</th>",
      "</tr></thead><tbody>",
      top.map(function (gap) {
        return [
          "<tr>",
          "<td>" + escapeHtml(gap.brand) + "</td>",
          "<td>" + escapeHtml(gap.gtin) + '<br><span class="fine-print">' + escapeHtml(gap.title) + "</span></td>",
          "<td>" + formatMoney(gap.lowValue) + '<br><span class="fine-print">' + escapeHtml(gap.lowRetailer) + "</span></td>",
          "<td>" + formatMoney(gap.highValue) + '<br><span class="fine-print">' + escapeHtml(gap.highRetailer) + "</span></td>",
          "<td>" + formatPercent(gap.gapPct) + "</td>",
          "</tr>"
        ].join("");
      }).join(""),
      "</tbody></table></div>"
    ].join("");
  }

  function renderRecommendations() {
    var recs = filteredRecommendations();
    var selected = ensureSelectedRecommendation(recs);
    var actionHtml = [
      '<button class="button ghost" id="export-recs">Export CSV</button>',
      '<button class="button ghost" id="export-brief-rec">Export brief</button>'
    ].join("");

    els.recommendationsView.innerHTML = [
      '<section class="command-strip">',
      '<div><span class="kicker">Recommendation workspace</span><strong>' + recs.length + " draft actions</strong><span>No automatic price execution. RGM approval required.</span></div>",
      '<div class="strip-actions">' + actionHtml + "</div>",
      "</section>",
      '<section class="recommendation-workspace">',
      '<aside class="section-panel rec-list-panel">',
      sectionHeader("Queue", "Select one recommendation to review", ""),
      recs.length ? '<div class="rec-list">' + recs.map(renderRecommendationListCard).join("") + "</div>" : '<div class="empty-state">No recommendations for the selected filters.</div>',
      "</aside>",
      '<section class="section-panel rec-detail-panel">',
      selected ? renderRecommendationDetail(selected) : '<div class="empty-state">Select a recommendation to inspect.</div>',
      "</section>",
      "</section>"
    ].join("");

    document.querySelectorAll("[data-select-rec]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.selectedRecId = button.getAttribute("data-select-rec");
        renderRecommendations();
      });
    });

    document.querySelectorAll("[data-rec-tab]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.activeRecTab = button.getAttribute("data-rec-tab");
        renderRecommendations();
      });
    });

    bindApprovalSelects(renderRecommendations);

    var exportRecs = document.getElementById("export-recs");
    var exportBrief = document.getElementById("export-brief-rec");
    if (exportRecs) {
      exportRecs.addEventListener("click", exportRecommendationsCsv);
    }
    if (exportBrief) {
      exportBrief.addEventListener("click", exportExecutiveBrief);
    }
  }

  function ensureSelectedRecommendation(recs) {
    var selected = recs.find(function (rec) {
      return rec.id === state.selectedRecId;
    });
    if (!selected && recs.length) {
      selected = recs[0];
      state.selectedRecId = selected.id;
    }
    return selected || null;
  }

  function renderRecommendationListCard(rec) {
    var selected = rec.id === state.selectedRecId ? " selected" : "";
    return [
      '<button class="rec-list-card' + selected + '" type="button" data-select-rec="' + escapeHtml(rec.id) + '">',
      '<span class="issue-topline">' + severityPill(rec.severity) + '<span>' + escapeHtml(rec.type) + "</span></span>",
      '<strong>' + escapeHtml(rec.title) + "</strong>",
      '<span class="issue-meta">' + escapeHtml(rec.scope) + "</span>",
      '<span class="rec-card-footer"><span>' + escapeHtml(rec.status) + "</span>" + (rec.legalReview ? "<span>Legal</span>" : "<span>RGM</span>") + "</span>",
      "</button>"
    ].join("");
  }

  function renderRecommendationDetail(rec) {
    var tabs = ["Summary", "Evidence", "Impact", "Risks", "Approval"];
    return [
      '<div class="rec-detail-head">',
      '<div>',
      '<span class="kicker">' + escapeHtml(rec.type) + "</span>",
      '<h2>' + escapeHtml(rec.title) + "</h2>",
      '<p>' + escapeHtml(rec.scope) + "</p>",
      "</div>",
      '<div class="rec-head-actions">' + severityPill(rec.severity) + '<select data-status-rec="' + escapeHtml(rec.id) + '">' + approvalOptions(rec.status) + "</select></div>",
      "</div>",
      '<div class="detail-tabs">',
      tabs.map(function (tab) {
        return '<button type="button" class="' + (state.activeRecTab === tab ? "active" : "") + '" data-rec-tab="' + escapeHtml(tab) + '">' + escapeHtml(tab) + "</button>";
      }).join(""),
      "</div>",
      '<div class="rec-tab-panel">',
      renderRecTabContent(rec, state.activeRecTab),
      "</div>"
    ].join("");
  }

  function renderRecTabContent(rec, tab) {
    if (tab === "Evidence") {
      return [
        '<div class="evidence-grid single-row">',
        evidenceTile("Signal", rec.evidence, "Workbook evidence"),
        evidenceTile("Brand", rec.brand, "Filterable scope"),
        evidenceTile("Retailer", rec.retailer, rec.channel),
        "</div>",
        '<div class="rec-field wide"><span>Analytical approach</span><p>' + escapeHtml(rec.analyticalApproach) + "</p></div>",
        '<div class="rec-field wide"><span>Data required</span><p>' + escapeHtml(rec.dataRequired) + "</p></div>"
      ].join("");
    }
    if (tab === "Impact") {
      return [
        '<div class="impact-grid">',
        impactCard("Expected impact", rec.expectedImpact),
        impactCard("Volume hypothesis", rec.volumeImpact),
        impactCard("Margin hypothesis", rec.marginImpact),
        impactCard("Competitive response", rec.competitiveResponse),
        impactCard("Retailer implications", rec.retailerImplications),
        impactCard("Consumer implications", rec.consumerImplications),
        "</div>"
      ].join("");
    }
    if (tab === "Risks") {
      return [
        '<div class="risk-panel">',
        '<h3>Risk Review</h3>',
        '<p>' + escapeHtml(rec.risks) + "</p>",
        '<div class="mini-checklist">',
        miniCheck("Legal", rec.legalReview ? "Legal review is triggered if this moves toward external action or corridor language." : "Legal review is not triggered by the flag alone."),
        miniCheck("Data quality", "Confirm crawl accuracy, product mapping, OOS status, and temporary promo context before escalation."),
        miniCheck("Decision boundary", "The agent flags the issue. Humans decide whether to act."),
        "</div>",
        "</div>"
      ].join("");
    }
    if (tab === "Approval") {
      return [
        '<div class="approval-path">',
        APPROVAL_STAGES.map(function (stage) {
          var active = stage === rec.status ? " active" : "";
          return '<button type="button" class="approval-step' + active + '" data-set-rec-status="' + escapeHtml(rec.id) + '" data-status-value="' + escapeHtml(stage) + '">' + escapeHtml(stage) + "</button>";
        }).join(""),
        "</div>",
        '<div class="approval-note">',
        '<strong>Human approval requirements</strong>',
        '<p>' + escapeHtml(rec.approval) + "</p>",
        '<p>No automatic price execution. RGM approval required before action.</p>',
        "</div>"
      ].join("");
    }
    return [
      '<div class="summary-cards">',
      impactCard("Business problem", rec.businessProblem),
      impactCard("Recommendation", "Flag for RGM review. Do not generate retailer outreach language from this prototype."),
      impactCard("Expected impact", rec.expectedImpact),
      impactCard("Approval", rec.approval),
      "</div>"
    ].join("");
  }

  function impactCard(label, text) {
    return '<div class="impact-card"><span>' + escapeHtml(label) + "</span><p>" + escapeHtml(text) + "</p></div>";
  }

  function approvalOptions(status) {
    return APPROVAL_STAGES.map(function (stage) {
      return '<option value="' + escapeHtml(stage) + '"' + (stage === status ? " selected" : "") + ">" + escapeHtml(stage) + "</option>";
    }).join("");
  }

  function bindApprovalSelects(afterChange) {
    document.querySelectorAll("[data-status-rec]").forEach(function (select) {
      select.addEventListener("change", function () {
        var rec = state.recommendations.find(function (item) {
          return item.id === select.getAttribute("data-status-rec");
        });
        if (rec) {
          rec.status = select.value;
          saveApprovalStatus(rec.id, rec.status);
          if (afterChange) {
            afterChange();
          }
        }
      });
    });

    document.querySelectorAll("[data-set-rec-status]").forEach(function (button) {
      button.addEventListener("click", function () {
        var rec = state.recommendations.find(function (item) {
          return item.id === button.getAttribute("data-set-rec-status");
        });
        if (rec) {
          rec.status = button.getAttribute("data-status-value");
          saveApprovalStatus(rec.id, rec.status);
          if (afterChange) {
            afterChange();
          }
        }
      });
    });
  }

  function renderSkuExplorer() {
    var rows = filteredPriceRows().slice().sort(function (a, b) {
      return severityWeight(severityForRow(b)) - severityWeight(severityForRow(a)) ||
        b.daysViolation - a.daysViolation ||
        Number(b.belowGuardrail) - Number(a.belowGuardrail);
    });
    var selected = selectedPriceRow(rows);

    els.skuView.innerHTML = [
      '<section class="command-strip">',
      '<div><span class="kicker">SKU Explorer</span><strong>' + rows.length + " rows</strong><span>Inspect brand, retailer, channel, guardrail, OOS, and price history.</span></div>",
      '<div class="strip-actions"><span class="status-chip">Workbook-only view</span><span class="status-chip">Impact = hypothesis</span></div>',
      "</section>",
      '<section class="sku-layout">',
      '<aside class="section-panel sku-list-panel">',
      sectionHeader("Rows", "Sorted by risk and duration", ""),
      renderSkuList(rows),
      "</aside>",
      '<section class="section-panel sku-detail-panel">',
      selected ? renderSkuDetail(selected) : '<div class="empty-state">Select a SKU row to inspect.</div>',
      "</section>",
      "</section>"
    ].join("");

    document.querySelectorAll("[data-price-row-id]").forEach(function (card) {
      card.addEventListener("click", function () {
        state.selectedPriceRowId = card.getAttribute("data-price-row-id");
        renderSkuExplorer();
      });
    });

    document.querySelectorAll("[data-open-rec]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.selectedRecId = button.getAttribute("data-open-rec");
        state.activeTab = "recommendations";
        setActiveTabButton("recommendations");
        render();
      });
    });
  }

  function renderSkuList(rows) {
    if (!rows.length) {
      return '<div class="empty-state">No SKU rows match the selected filters.</div>';
    }
    return [
      '<div class="sku-list">',
      rows.slice(0, 80).map(function (row) {
        var selected = row.id === state.selectedPriceRowId ? " selected" : "";
        var status = row.isOos ? "OOS" : row.belowGuardrail ? "Below guardrail" : "In range";
        return [
          '<button class="sku-card' + selected + '" type="button" data-price-row-id="' + escapeHtml(row.id) + '">',
          '<span class="issue-topline">' + severityPill(severityForRow(row)) + '<span>' + escapeHtml(status) + "</span></span>",
          '<strong>' + escapeHtml(compactTitle(row)) + "</strong>",
          '<span class="issue-meta">' + escapeHtml(row.brand) + " | " + escapeHtml(row.retailerName) + " | " + escapeHtml(row.channel) + "</span>",
          '<span class="rec-card-footer"><span>' + formatMoney(row.latestValue) + "</span><span>" + escapeHtml(String(row.daysViolation || 0)) + " days</span></span>",
          "</button>"
        ].join("");
      }).join(""),
      "</div>"
    ].join("");
  }

  function renderSkuDetail(row) {
    var ppcRows = state.ppcRows.filter(function (item) {
      return item.gtin === row.gtin && matchesBaseFilters(item);
    }).sort(function (a, b) {
      return (a.latestValue || 0) - (b.latestValue || 0);
    });
    return [
      '<div class="sku-detail-head">',
      '<div><span class="kicker">' + escapeHtml(row.brand) + "</span><h2>" + escapeHtml(compactTitle(row)) + "</h2><p>" + escapeHtml(row.title) + "</p></div>",
      '<div>' + severityPill(severityForRow(row)) + "</div>",
      "</div>",
      renderIssueEvidence(row),
      '<div class="section-divider"></div>',
      sectionHeader("Price Per Count", "Same-GTIN comparison in loaded workbook", ""),
      renderSkuPpcRows(ppcRows),
      '<div class="section-divider"></div>',
      sectionHeader("Governance", "How this row should be used", ""),
      '<div class="summary-cards">',
      impactCard("Decision boundary", "Flag for RGM review only. Do not execute prices automatically."),
      impactCard("Legal", row.belowGuardrail ? "Legal review is triggered only if a recommendation proposes external action or guardrail/corridor language." : "No legal trigger from this row alone."),
      impactCard("Data quality", "Confirm crawl accuracy, product mapping, OOS status, and promo context before escalation."),
      "</div>"
    ].join("");
  }

  function renderSkuPpcRows(rows) {
    if (!rows.length) {
      return '<div class="empty-state">No price-per-count rows found for this SKU under current filters.</div>';
    }
    return [
      '<div class="ppc-chip-grid">',
      rows.map(function (row) {
        return [
          '<div class="ppc-chip">',
          '<span>' + escapeHtml(row.retailerName) + "</span>",
          '<strong>' + formatMoney(row.latestValue) + "</strong>",
          '<em>' + escapeHtml(row.channel) + "</em>",
          "</div>"
        ].join("");
      }).join(""),
      "</div>"
    ].join("");
  }

  function renderApprovals() {
    var recs = filteredRecommendations();
    var high = recs.filter(function (rec) { return rec.severity === "High"; }).length;
    var legal = recs.filter(function (rec) { return rec.legalReview; }).length;
    var approved = recs.filter(function (rec) { return rec.status === "Approved"; }).length;

    els.approvalsView.innerHTML = [
      '<section class="summary-grid compact-metrics">',
      metric("In queue", recs.length, "Recommendations matching filters", "info"),
      metric("High severity", high, "Needs faster review", high ? "danger" : "good"),
      metric("Legal triggers", legal, "Only if recommendation moves to action", legal ? "warning" : "good"),
      metric("Approved", approved, "Human approved recommendations", approved ? "good" : "info"),
      "</section>",
      '<section class="approval-board">',
      APPROVAL_STAGES.map(function (stage) {
        var stageRecs = recs.filter(function (rec) {
          return rec.status === stage;
        });
        return [
          '<section class="approval-column">',
          '<div class="approval-column-head"><strong>' + escapeHtml(stage) + "</strong><span>" + stageRecs.length + "</span></div>",
          stageRecs.length ? stageRecs.map(renderApprovalCard).join("") : '<div class="empty-state small">No items</div>',
          "</section>"
        ].join("");
      }).join(""),
      "</section>"
    ].join("");

    bindApprovalSelects(renderApprovals);
    document.querySelectorAll("[data-select-rec]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.selectedRecId = button.getAttribute("data-select-rec");
        state.activeTab = "recommendations";
        setActiveTabButton("recommendations");
        render();
      });
    });
  }

  function renderApprovalCard(rec) {
    return [
      '<article class="approval-card">',
      '<button type="button" data-select-rec="' + escapeHtml(rec.id) + '">',
      severityPill(rec.severity),
      '<strong>' + escapeHtml(rec.title) + "</strong>",
      '<span>' + escapeHtml(rec.scope) + "</span>",
      "</button>",
      '<select data-status-rec="' + escapeHtml(rec.id) + '">' + approvalOptions(rec.status) + "</select>",
      "</article>"
    ].join("");
  }

  function renderStrategist() {
    els.strategistView.innerHTML = [
      '<section class="chat-layout">',
      '<div class="chat-panel">',
      '<div id="chat-log" class="chat-log">' + state.chat.map(renderChatMessage).join("") + "</div>",
      '<form id="chat-form" class="chat-input">',
      '<input id="chat-question" type="text" placeholder="Ask about guardrails, retailer discipline, or price-per-count gaps">',
      '<button class="button" id="chat-submit" type="submit">Ask</button>',
      "</form>",
      "</div>",
      '<aside class="section-panel">',
      sectionHeader("Starter Questions", "Local answers from the loaded workbook", ""),
      '<div class="prompt-list">',
      promptButton("Create an executive summary."),
      promptButton("Where are we below guardrail?"),
      promptButton("Which retailers need RGM review?"),
      promptButton("Where are price-per-count gaps largest?"),
      promptButton("Design a pricing experiment."),
      promptButton("Which promos should we stop?"),
      "</div>",
      "</aside>",
      "</section>"
    ].join("");

    var form = document.getElementById("chat-form");
    var input = document.getElementById("chat-question");
    var submitButton = document.getElementById("chat-submit");
    var submitCurrentQuestion = function () {
      if (!input) {
        return;
      }
      var question = input.value;
      input.value = "";
      submitQuestion(question);
      setTimeout(function () {
        var nextInput = document.getElementById("chat-question");
        if (nextInput) {
          nextInput.focus();
        }
      }, 0);
    };

    if (form && input) {
      form.addEventListener("submit", function (event) {
        event.preventDefault();
        submitCurrentQuestion();
      });
      input.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
          event.preventDefault();
          submitCurrentQuestion();
        }
      });
    }

    if (submitButton) {
      submitButton.addEventListener("click", function (event) {
        event.preventDefault();
        submitCurrentQuestion();
      });
    }

    document.querySelectorAll("[data-prompt]").forEach(function (button) {
      button.addEventListener("click", function () {
        submitQuestion(button.getAttribute("data-prompt"));
      });
    });

    var log = document.getElementById("chat-log");
    log.scrollTop = log.scrollHeight;
  }

  function renderEmptyChat() {
    state.chat = [
      {
        role: "agent",
        text: "Load a workbook to start the pricing discipline review."
      }
    ];
  }

  function promptButton(text) {
    return '<button type="button" data-prompt="' + escapeHtml(text) + '">' + escapeHtml(text) + "</button>";
  }

  function submitQuestion(question) {
    var cleaned = clean(question);
    if (!cleaned) {
      return;
    }
    state.chat.push({ role: "user", text: escapeHtml(cleaned) });
    state.chat.push({ role: "agent", text: answerQuestion(cleaned), html: true });
    renderStrategist();
  }

  function renderChatMessage(message) {
    var body = message.html ? message.text : escapeHtml(message.text);
    return '<div class="chat-message ' + message.role + '">' + body + "</div>";
  }

  function answerQuestion(question) {
    var lower = question.toLowerCase();
    var rows = filteredPriceRows();
    var below = rows.filter(function (row) { return row.belowGuardrail; }).sort(function (a, b) {
      return b.daysViolation - a.daysViolation || b.gapPct - a.gapPct;
    });
    var recs = filteredRecommendations();

    if (lower.indexOf("promo") >= 0 || lower.indexOf("promotion") >= 0) {
      return "<p>This workbook does not include units, revenue, trade spend, repeat rate, or promo mechanics, so I would not recommend stopping promos from this file alone.</p><p>For this MVP, I can flag pricing discipline and price-per-count risks. Promo optimization should wait for the sales/margin feed.</p>";
    }

    if (lower.indexOf("experiment") >= 0 || lower.indexOf("test") >= 0) {
      var row = below[0];
      if (!row) {
        return "<p>No below-guardrail row is available under the current filters. A test design should start after RGM selects a SKU, retailer, and target guardrail question.</p>";
      }
      return [
        "<p><strong>Experiment design hypothesis:</strong> " + escapeHtml(row.brand) + " at " + escapeHtml(row.retailerName) + " should be reviewed for persistent below-guardrail pricing.</p>",
        "<ul>",
        "<li><strong>Business problem:</strong> Latest crawled price is " + formatMoney(row.latestValue) + " vs " + formatMoney(row.guardrail) + " guardrail.</li>",
        "<li><strong>Approach:</strong> Use a test/control retailer or region only after RGM confirms legal and sales guardrails.</li>",
        "<li><strong>Success metrics:</strong> Guardrail compliance, price-per-count position, retailer response, consumer price signal, and later sales/margin impact if a second feed is added.</li>",
        "<li><strong>Approval:</strong> RGM approval required. Legal review required before external action.</li>",
        "</ul>"
      ].join("");
    }

    if (lower.indexOf("price-per-count") >= 0 || lower.indexOf("pack") >= 0 || lower.indexOf("architecture") >= 0) {
      var gaps = state.ppcGaps.filter(function (gap) {
        return gap.gapPct >= 0.2 && (state.filters.brand === "All" || gap.brand === state.filters.brand);
      }).slice(0, 5);
      if (!gaps.length) {
        return "<p>No 20%+ price-per-count gaps are visible under the current filters.</p>";
      }
      return "<p>Largest price-per-count gaps:</p>" + list(gaps.map(function (gap) {
        return gap.brand + " " + gap.gtin + ": " + formatPercent(gap.gapPct) + " spread from " + gap.lowRetailer + " to " + gap.highRetailer;
      }));
    }

    if (lower.indexOf("retailer") >= 0 || lower.indexOf("discipline") >= 0) {
      var scores = filteredRetailerScores().slice(0, 6);
      return "<p>Retailers to prioritize for RGM review:</p>" + list(scores.map(function (score) {
        return score.retailer + ": score " + score.score + ", " + score.below + " below-guardrail flags, " + score.extended + " extended flags";
      }));
    }

    if (lower.indexOf("summary") >= 0 || lower.indexOf("executive") >= 0) {
      return executiveSummaryHtml(rows, recs);
    }

    if (lower.indexOf("below") >= 0 || lower.indexOf("guardrail") >= 0 || lower.indexOf("leak") >= 0) {
      if (!below.length) {
        return "<p>No below-guardrail rows are visible under the current filters.</p>";
      }
      return "<p>Top below-guardrail rows:</p>" + list(below.slice(0, 6).map(function (row) {
        return row.brand + " at " + row.retailerName + ": " + formatMoney(row.latestValue) + " vs " + formatMoney(row.guardrail) + ", " + row.daysViolation + " days";
      }));
    }

    return "<p>I can answer from the loaded workbook about guardrail flags, retailer discipline, first-mover indicators, OOS flags, and price-per-count architecture. I will label volume and margin as hypotheses because this workbook does not contain sales or margin fields.</p>";
  }

  function executiveSummaryHtml(rows, recs) {
    var below = rows.filter(function (row) { return row.belowGuardrail; });
    var extended = rows.filter(function (row) { return row.daysViolation >= 7; });
    var oos = rows.filter(function (row) { return row.isOos; });
    var high = recs.filter(function (rec) { return rec.severity === "High"; });

    return [
      "<p><strong>Executive summary:</strong></p>",
      "<ul>",
      "<li>" + below.length + " current below-guardrail flags in the selected view.</li>",
      "<li>" + extended.length + " rows have 7 or more consecutive flagged days.</li>",
      "<li>" + oos.length + " rows show OOS or non-price status.</li>",
      "<li>" + high.length + " high-severity recommendations require RGM review.</li>",
      "<li>No automatic price execution. RGM approval required before action.</li>",
      "</ul>"
    ].join("");
  }

  function renderDataMap() {
    var brandCounts = countBy(state.productMap, function (product) { return product.brand || "Unknown"; });
    var retailerCounts = countBy(state.productMap, function (product) { return product.retailerName || "Unknown"; });

    els.dataView.innerHTML = [
      '<section class="data-grid">',
      '<div class="section-panel">',
      sectionHeader("Workbook Sheets", state.sheets.length + " sheets loaded", ""),
      '<div class="table-wrap"><table><thead><tr><th>Sheet</th></tr></thead><tbody>',
      state.sheets.map(function (name) {
        return "<tr><td>" + escapeHtml(name) + "</td></tr>";
      }).join(""),
      "</tbody></table></div>",
      "</div>",
      '<div class="section-panel">',
      sectionHeader("Normalized Tables", "Purpose-built for the first-mover workbook", ""),
      '<div class="table-wrap"><table><thead><tr><th>Table</th><th>Rows</th></tr></thead><tbody>',
      tableRow("Product map", state.productMap.length),
      tableRow("Daily price rows", state.priceRows.length),
      tableRow("Daily price-per-count rows", state.ppcRows.length),
      tableRow("Recommendations", state.recommendations.length),
      "</tbody></table></div>",
      "</div>",
      '<div class="section-panel">',
      sectionHeader("Brand Coverage", "Rows from Product Map", ""),
      keyValueTable(brandCounts),
      "</div>",
      '<div class="section-panel">',
      sectionHeader("Retailer Coverage", "Rows from Product Map", ""),
      keyValueTable(retailerCounts),
      "</div>",
      "</section>",
      '<section class="section-panel">',
      sectionHeader("Field Coverage", "Current MVP boundaries", ""),
      '<div class="table-wrap"><table><thead><tr><th>Available</th><th>Not in this workbook</th></tr></thead><tbody>',
      "<tr><td>Guardrail, retailer, brand, SKU, product count, daily price, price per count, OOS, consecutive days, first mover, lowest retailer</td><td>Units, revenue, gross margin, COGS, trade spend, promo ROI, repeat rate</td></tr>",
      "</tbody></table></div>",
      '<p class="fine-print">NonPromo MAP is treated as a commercial guardrail. The app does not label rows as legal violations.</p>',
      "</section>"
    ].join("");
  }

  function tableRow(label, value) {
    return "<tr><td>" + escapeHtml(label) + "</td><td>" + escapeHtml(String(value)) + "</td></tr>";
  }

  function keyValueTable(values) {
    var rows = Object.keys(values).sort().map(function (key) {
      return tableRow(key, values[key]);
    }).join("");
    return '<div class="table-wrap"><table><thead><tr><th>Name</th><th>Rows</th></tr></thead><tbody>' + rows + "</tbody></table></div>";
  }

  function exportRecommendationsCsv() {
    var recs = filteredRecommendations();
    var headers = [
      "severity",
      "type",
      "brand",
      "retailer",
      "channel",
      "title",
      "scope",
      "evidence",
      "business_problem",
      "data_required",
      "analytical_approach",
      "expected_impact",
      "risks",
      "volume_impact",
      "margin_impact",
      "competitive_response",
      "retailer_implications",
      "consumer_implications",
      "approval",
      "legal_review",
      "status"
    ];

    var csv = [
      headers.join(","),
      recs.map(function (rec) {
        return [
          rec.severity,
          rec.type,
          rec.brand,
          rec.retailer,
          rec.channel,
          rec.title,
          rec.scope,
          rec.evidence,
          rec.businessProblem,
          rec.dataRequired,
          rec.analyticalApproach,
          rec.expectedImpact,
          rec.risks,
          rec.volumeImpact,
          rec.marginImpact,
          rec.competitiveResponse,
          rec.retailerImplications,
          rec.consumerImplications,
          rec.approval,
          rec.legalReview ? "Yes" : "No",
          rec.status
        ].map(csvCell).join(",");
      }).join("\n")
    ].join("\n");

    downloadFile("pricing-ai-recommendations.csv", csv, "text/csv");
  }

  function exportExecutiveBrief() {
    var rows = filteredPriceRows();
    var recs = filteredRecommendations();
    var high = recs.filter(function (rec) { return rec.severity === "High"; });
    var below = rows.filter(function (row) { return row.belowGuardrail; });
    var extended = rows.filter(function (row) { return row.daysViolation >= 7; });
    var ppcGaps = state.ppcGaps.filter(function (gap) { return gap.gapPct >= 0.2; });

    var text = [
      "Pricing AI Strategist - Executive Brief",
      "Workbook: " + state.fileName,
      "Period: " + (state.dateRange || "n/a"),
      "",
      "Summary",
      "- Below-guardrail flags: " + below.length,
      "- Extended flags (7+ days): " + extended.length,
      "- Price-per-count gaps (20%+): " + ppcGaps.length,
      "- Recommendations: " + recs.length,
      "- High-severity recommendations: " + high.length,
      "",
      "Top Recommendations",
      recs.slice(0, 8).map(function (rec, index) {
        return (index + 1) + ". [" + rec.severity + "] " + rec.title + " - " + rec.scope + " - " + rec.evidence;
      }).join("\n"),
      "",
      "Governance",
      "No automatic price execution. RGM approval required before action. Legal review required when recommendations reference guardrails, corridor language, channel conflict, or external-facing retailer action."
    ].join("\n");

    downloadFile("pricing-ai-executive-brief.txt", text, "text/plain");
  }

  function downloadFile(fileName, text, mimeType) {
    var blob = new Blob([text], { type: mimeType });
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function csvCell(value) {
    var text = value === null || typeof value === "undefined" ? "" : String(value);
    return '"' + text.replace(/"/g, '""') + '"';
  }

  function loadSavedApprovalStatuses() {
    try {
      return JSON.parse(localStorage.getItem("pricing-ai-approval-statuses") || "{}");
    } catch (error) {
      return {};
    }
  }

  function saveApprovalStatus(id, status) {
    var saved = loadSavedApprovalStatuses();
    saved[id] = status;
    localStorage.setItem("pricing-ai-approval-statuses", JSON.stringify(saved));
  }

  function pickDefaultSelectedRow(rows) {
    var row = rows.filter(function (item) {
      return item.belowGuardrail;
    }).sort(function (a, b) {
      return b.daysViolation - a.daysViolation || b.gapPct - a.gapPct;
    })[0] || rows[0];
    return row ? row.id : "";
  }

  function headerGetter(headers, values) {
    var normalized = headers.map(normalizeHeader);
    return function (header) {
      var index = normalized.indexOf(normalizeHeader(header));
      return index >= 0 ? values[index] : "";
    };
  }

  function normalizeHeader(header) {
    return clean(header).toLowerCase().replace(/\s+/g, " ");
  }

  function brandFromSheet(sheetName) {
    return clean(sheetName.split("-")[0]).replace(/\s+$/, "");
  }

  function channelFor(retailerId, retailerName) {
    if (CHANNEL_BY_RETAILER_ID[retailerId]) {
      return CHANNEL_BY_RETAILER_ID[retailerId];
    }
    var name = (retailerName || "").toLowerCase();
    if (name.indexOf("amazon") >= 0) {
      return "Amazon / Marketplace";
    }
    if (name.indexOf("costco") >= 0 || name.indexOf("sam") >= 0) {
      return "Club";
    }
    if (name.indexOf("cvs") >= 0 || name.indexOf("walgreens") >= 0) {
      return "Drug";
    }
    if (name.indexOf("kroger") >= 0) {
      return "Grocery";
    }
    if (name.indexOf("walmart") >= 0 || name.indexOf("target") >= 0) {
      return "Mass / Digital";
    }
    return "Other";
  }

  function latestSeriesPoint(series) {
    for (var i = series.length - 1; i >= 0; i -= 1) {
      if (series[i].raw !== "") {
        return series[i];
      }
    }
    return series[series.length - 1] || null;
  }

  function parseDays(value) {
    var text = clean(value);
    if (!text || /^OOS$/i.test(text) || /^n\/a$/i.test(text)) {
      return 0;
    }
    var number = Number(text);
    return Number.isFinite(number) ? number : 0;
  }

  function truthy(value) {
    var text = clean(value).toLowerCase();
    return text === "true" || text === "yes" || text === "1" || text === "y";
  }

  function toNumber(value) {
    if (value === null || typeof value === "undefined") {
      return null;
    }
    var text = String(value).trim();
    if (!text || /^n\/a$/i.test(text) || /^OOS$/i.test(text) || /^TBD$/i.test(text)) {
      return null;
    }
    var number = Number(text.replace(/[$,%]/g, ""));
    return Number.isFinite(number) ? number : null;
  }

  function firstNumber() {
    for (var i = 0; i < arguments.length; i += 1) {
      if (arguments[i] !== null && typeof arguments[i] !== "undefined" && Number.isFinite(arguments[i])) {
        return arguments[i];
      }
    }
    return null;
  }

  function clean(value) {
    return value === null || typeof value === "undefined" ? "" : String(value).trim();
  }

  function unique(values) {
    return Array.from(new Set(values.filter(function (value) {
      return clean(value);
    })));
  }

  function groupBy(items, keyFn) {
    return items.reduce(function (acc, item) {
      var key = keyFn(item) || "Unknown";
      if (!acc[key]) {
        acc[key] = [];
      }
      acc[key].push(item);
      return acc;
    }, {});
  }

  function countBy(items, keyFn) {
    return items.reduce(function (acc, item) {
      var key = keyFn(item) || "Unknown";
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
  }

  function formatMoney(value) {
    if (value === null || typeof value === "undefined" || Number.isNaN(value)) {
      return "n/a";
    }
    return "$" + Number(value).toFixed(value < 1 ? 3 : 2);
  }

  function formatPercent(value) {
    if (value === null || typeof value === "undefined" || Number.isNaN(value)) {
      return "n/a";
    }
    return Math.round(value * 100) + "%";
  }

  function formatDateHeader(value) {
    var text = clean(value);
    if (!/^20\d{6}$/.test(text)) {
      return text || "n/a";
    }
    var year = text.slice(0, 4);
    var month = Number(text.slice(4, 6));
    var day = Number(text.slice(6, 8));
    var names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return names[month - 1] + " " + day + ", " + year;
  }

  function severityPill(severity) {
    var css = {
      High: "severity-high",
      Medium: "severity-medium",
      Watch: "severity-watch",
      Low: "severity-low"
    }[severity] || "severity-low";
    return '<span class="severity-pill ' + css + '">' + escapeHtml(severity) + "</span>";
  }

  function list(items) {
    return "<ul>" + items.map(function (item) {
      return "<li>" + escapeHtml(item) + "</li>";
    }).join("") + "</ul>";
  }

  function normalizeId(value) {
    return clean(value).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  }

  function escapeHtml(value) {
    return clean(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeAttribute(value) {
    return escapeHtml(value);
  }
}());
