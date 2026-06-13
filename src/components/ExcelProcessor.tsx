/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from "react";
import { 
  FileSpreadsheet, 
  Upload, 
  ArrowRight, 
  Settings2, 
  Download, 
  RefreshCcw, 
  FileCheck, 
  Info, 
  Check, 
  AlertTriangle,
  Flame,
  HelpCircle
} from "lucide-react";
import * as XLSX from "xlsx";
import { ProductCatalogRow, InvoiceDetailRow, OutputRow, ThanhTienFormula, ColumnMapping } from "../types";

interface ExcelProcessorProps {
  products: ProductCatalogRow[];
  invoices: InvoiceDetailRow[];
  onDataChange: (products: ProductCatalogRow[], invoices: InvoiceDetailRow[]) => void;
  formula: ThanhTienFormula;
  setFormula: (f: ThanhTienFormula) => void;
  isNegativeCKTM: boolean;
  setIsNegativeCKTM: (val: boolean) => void;
}

// Helper to safely parse localized Vietnamese metrics/money values (handling variations of thousands separators)
function parseVietnameseNumber(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return val;
  let str = String(val).trim();
  if (!str) return 0;
  
  // Remove currency units and extra spaces
  str = str.replace(/[đ|d|VND|VNĐ|Vnđ|vnđ]/g, "").trim();
  
  // Clean custom punctuation
  const dotCount = (str.match(/\./g) || []).length;
  const commaCount = (str.match(/,/g) || []).length;
  
  if (dotCount > 1) {
    str = str.replace(/\./g, "");
    if (commaCount === 1) {
      str = str.replace(/,/g, ".");
    }
  } else if (commaCount > 1) {
    str = str.replace(/,/g, "");
  } else if (dotCount === 1 && commaCount === 1) {
    const dotIdx = str.indexOf(".");
    const commaIdx = str.indexOf(",");
    if (dotIdx < commaIdx) {
      str = str.replace(/\./g, "").replace(/,/g, ".");
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (dotCount === 1) {
    const parts = str.split(".");
    if (parts[1].length === 3) {
      str = str.replace(/\./g, "");
    }
  } else if (commaCount === 1) {
    const parts = str.split(",");
    if (parts[1].length === 3) {
      str = str.replace(/,/g, "");
    } else {
      str = str.replace(/,/g, ".");
    }
  }
  
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
}

// Helper to convert index 0, 1, 2... to Excel-style alphabetical column label (A, B, C... Z, AA... AZ, BA, BB...)
function getExcelColumnLabel(index: number): string {
  let label = "";
  let temp = index;
  while (temp >= 0) {
    label = String.fromCharCode((temp % 26) + 65) + label;
    temp = Math.floor(temp / 26) - 1;
  }
  return label;
}

// Helper to safely fetch value with flexible/insensitive keys (handling spaces, casing, and Vietnamese NFC/NFD Unicode normalization)
function getValueByFlexibleKey(row: any, targetKey: string): any {
  if (!row || !targetKey) return undefined;
  
  // Try direct match first (highest performance)
  if (row[targetKey] !== undefined) return row[targetKey];
  
  const normTarget = String(targetKey).trim().normalize("NFC").toLowerCase();
  if (!normTarget) return undefined;

  // Let's iterate over keys and match both trimmed and normalized
  const keys = Object.keys(row);
  
  // 1. Direct trimmed match
  const trimmedTarget = targetKey.trim();
  if (row[trimmedTarget] !== undefined) return row[trimmedTarget];

  // 2. Exact match using normalized forms
  for (const k of keys) {
    const normK = k.trim().normalize("NFC").toLowerCase();
    if (normK === normTarget) {
      return row[k];
    }
  }
  
  // 3. Substring match on normalized forms (fallback - only for meaningful keys to prevent false positives)
  for (const k of keys) {
    const normK = k.trim().normalize("NFC").toLowerCase();
    if (normK.length >= 3 && normTarget.length >= 3) {
      if (normK.includes(normTarget) || normTarget.includes(normK)) {
        return row[k];
      }
    }
  }
  
  return undefined;
}

// Robust helper to automatically match the best column name based on prioritized patterns
function findBestColumnMatch(
  cols: string[],
  patterns: { exact: string[]; contains: string[] },
  excludeContains?: string[]
): string | undefined {
  const normalizedCols = cols.map(c => ({
    original: c,
    clean: c.trim().normalize("NFC").toLowerCase()
  }));

  // Filter out columns matching exclusion patterns
  const filteredCols = normalizedCols.filter(item => {
    if (!excludeContains) return true;
    return !excludeContains.some(exclude => item.clean.includes(exclude.trim().normalize("NFC").toLowerCase()));
  });

  // 1. Search for exact matches in priority order
  for (const pattern of patterns.exact) {
    const pClean = pattern.trim().normalize("NFC").toLowerCase();
    const matched = filteredCols.find(item => item.clean === pClean);
    if (matched) return matched.original;
  }

  // 2. Search for substring "contains" matches in priority order
  for (const pattern of patterns.contains) {
    const pClean = pattern.trim().normalize("NFC").toLowerCase();
    const matched = filteredCols.find(item => item.clean.includes(pClean));
    if (matched) return matched.original;
  }

  return undefined;
}

export default function ExcelProcessor({
  products,
  invoices,
  onDataChange,
  formula,
  setFormula,
  isNegativeCKTM,
  setIsNegativeCKTM,
}: ExcelProcessorProps) {
  const [outputData, setOutputData] = useState<OutputRow[]>([]);
  const [activeTab, setActiveTab] = useState<"upload" | "mapping" | "preview">("upload");
  const [processingLogs, setProcessingLogs] = useState<string[]>([]);

  // NEW: Step-by-step state management
  const [activeStepTab, setActiveStepTab] = useState<"step1" | "step2" | "step3">("step1");
  const [activeStep2SubTab, setActiveStep2SubTab] = useState<"step2_1" | "step2_2">("step2_1");
  const [step2Invoices, setStep2Invoices] = useState<any[]>([]);
  const [step2Processed, setStep2Processed] = useState<boolean>(false);

  const [step2Stats, setStep2Stats] = useState<{
    taxRates: number[];
    hasInvoiceDiscount: boolean;
    totalInvoicesCount: number;
    invoiceTaxSummaries: { invoiceId: string; itemCount: number; taxRates: number[]; taxCount: number }[];
  } | null>(null);
  const [step3Processed, setStep3Processed] = useState<boolean>(false);
  const [step2SearchQuery, setStep2SearchQuery] = useState("");
  const [step2Page, setStep2Page] = useState(1);
  
  // Custom manual mappings if headers differ from standard
  const [colMapping, setColMapping] = useState<ColumnMapping>({
    productCodeCol: "Mã hàng",
    productTaxCol: "Thuế bán hàng",
    invoiceIdCol: "Mã hóa đơn",
    invoiceProductCodeCol: "Mã hàng",
    invoiceProductNameCol: "Tên hàng",
    invoicePriceCol: "Giá bán",
    invoiceQtyCol: "Số lượng",
    invoiceDiscountPercentCol: "Giảm giá %",
    invoiceDiscountAmountCol: "Giảm giá hóa đơn",
  });

  const [availableCatalogCols, setAvailableCatalogCols] = useState<string[]>([]);
  const [availableInvoiceCols, setAvailableInvoiceCols] = useState<string[]>([]);

  // Automatically reset subsequent steps if source data changes
  useEffect(() => {
    setStep2Processed(false);
    setStep2Invoices([]);
    setStep2Stats(null);
    setStep3Processed(false);
    setOutputData([]);
  }, [products, invoices]);

  // Automatically adjust column matching list once files load
  useEffect(() => {
    if (products.length > 0) {
      const firstRow = products[0];
      const cols = Object.keys(firstRow);
      setAvailableCatalogCols(cols);
      
      // Auto mapping heuristics using prioritized findBestColumnMatch
      const code = findBestColumnMatch(cols, {
        exact: ["mã món", "ma mon", "mã hàng", "ma hang", "mã sản phẩm", "mã sp", "ma sp", "mã món ăn"],
        contains: ["mã món", "ma mon", "mã hàng", "ma hang", "mã sp", "ma sp"]
      }, ["loại", "nhóm", "thêm"]) || cols[0];
      
      const tax = findBestColumnMatch(cols, {
        exact: ["thuế bán hàng", "thue ban hang", "thuế suất", "thue suat", "thuế %", "mức thuế", "muc thue"],
        contains: ["thuế", "thue", "vat", "tax", "%"]
      }) || cols[1] || cols[0];
      
      setColMapping(prev => ({
        ...prev,
        productCodeCol: code || prev.productCodeCol,
        productTaxCol: tax || prev.productTaxCol,
      }));

      addLog(`[Hệ thống] Tự động khớp cột Catalog: Mã hàng hóa sản phẩm là '${code}', Cột thuế suất là '${tax}'`);
    }
  }, [products]);

  useEffect(() => {
    if (invoices.length > 0) {
      const firstRow = invoices[0];
      const cols = Object.keys(firstRow);
      setAvailableInvoiceCols(cols);

      // Auto mapping heuristics using prioritized findBestColumnMatch
      const invoiceId = findBestColumnMatch(cols, {
        exact: ["mã hóa đơn", "ma hoa don", "mã nhóm hóa đơn", "số hóa đơn", "so hoa don", "ma_hd", "so_hd", "mã nhóm", "mã số thuế"],
        contains: ["mã nhóm", "hóa đơn", "hoa don", "số hd", "so hd", "ma hd"]
      }) || cols[0];

      const code = findBestColumnMatch(cols, {
        exact: ["mã hàng", "ma hang", "ma_hang", "mã sản phẩm", "mã sp", "ma sp", "mã món", "ma mon"],
        contains: ["mã hàng", "ma hang", "mã sp", "ma sp", "mã món", "ma mon"]
      }, ["loại", "nhóm", "thêm"]) || cols.find(c => c.toLowerCase().includes("mã")) || cols[1] || cols[0];

      const name = findBestColumnMatch(cols, {
        exact: ["tên hàng", "ten hang", "tên sản phẩm", "tên hàng hóa", "tên món", "ten mon"],
        contains: ["tên hàng", "ten hang", "tên món", "ten mon", "tên sp", "tên sản phẩm", "món ăn"]
      }) || cols[2] || cols[0];

      const price = findBestColumnMatch(cols, {
        exact: ["giá bán", "gia ban", "đơn giá", "don gia", "đơn giá bán", "giá", "gia"],
        contains: ["giá bán", "gia ban", "đơn giá", "don gia", "giá", "gia"]
      }, ["trạng thái", "status", "giao", "nhận", "ck", "chiết khấu", "giảm giá", "giảm"]) || cols[3] || cols[0];

      const qty = findBestColumnMatch(cols, {
        exact: ["số lượng", "so luong", "sl", "quantity", "qty"],
        contains: ["số lượng", "so luong", "sl", "qty"]
      }) || cols[4] || cols[0];

      const lineDisc = findBestColumnMatch(cols, {
        exact: ["giảm giá %", "giam gia %", "chiết khấu %", "chiet khau %", "chiết khấu", "chiet khau"],
        contains: ["giảm giá %", "giam gia %", "chiết khấu", "chiet khau", "ck %", "ck(%)", "giảm %"]
      }) || cols[5] || cols[0];

      const hdBdisc = findBestColumnMatch(cols, {
        exact: ["giảm giá hóa đơn", "giam gia hoa don", "chiết khấu hóa đơn", "chiet khau hoa don", "giam_gia_hd", "ck hd"],
        contains: ["giảm giá hóa đơn", "giam gia hoa don", "chiết khấu hóa đơn", "giam_gia_hd", "chiết khấu hd", "ck hd"]
      }) || cols[6] || cols[0];

      setColMapping(prev => ({
        ...prev,
        invoiceIdCol: invoiceId || prev.invoiceIdCol,
        invoiceProductCodeCol: code || prev.invoiceProductCodeCol,
        invoiceProductNameCol: name || prev.invoiceProductNameCol,
        invoicePriceCol: price || prev.invoicePriceCol,
        invoiceQtyCol: qty || prev.invoiceQtyCol,
        invoiceDiscountPercentCol: lineDisc || prev.invoiceDiscountPercentCol,
        invoiceDiscountAmountCol: hdBdisc || prev.invoiceDiscountAmountCol,
      }));

      addLog(`[Hệ thống] Tự động khớp cột Hóa đơn: Mã HĐ là '${invoiceId}', Mã hàng là '${code}', Giá bán là '${price}', S.Lượng là '${qty}'`);
    }
  }, [invoices]);

  // Read upload excel or CSV
  const handleFileUpload = (type: "products" | "invoices", file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;
        const workbook = XLSX.read(data, { type: "array" });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          alert("File excel không có trang dữ liệu!");
          return;
        }
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<any>(sheet);

        if (jsonData.length === 0) {
          alert("Tải lên thất bại: File trống hoặc không có dòng dữ liệu hợp lệ!");
          return;
        }

        if (type === "products") {
          const firstRow = jsonData[0] || {};
          const cols = Object.keys(firstRow);
          
          // Identify columns representing Code and Tax rate accurately
          const codeCol = cols.find(c => {
            const lower = c.toLowerCase().trim();
            if (lower.includes("loại") || lower.includes("nhóm") || lower.includes("thêm")) return false;
            return (
              lower === "mã món" ||
              lower === "ma mon" ||
              lower.includes("mã món") ||
              lower.includes("ma mon") ||
              lower === "mã hàng" ||
              lower === "ma hang" ||
              lower.includes("mã hàng") ||
              lower.includes("ma hang") ||
              lower.includes("mã sản phẩm") ||
              lower.includes("code") ||
              lower === "mã"
            );
          }) || cols.find(c => c.toLowerCase().includes("mã")) || cols[0];

          const taxCol = cols.find(c => {
            const lower = c.toLowerCase().trim();
            return (
              lower === "thuế bán hàng" ||
              lower === "thue ban hang" ||
              lower.includes("thuế bán hàng") ||
              lower.includes("thue ban hang") ||
              lower.includes("thuế suất") ||
              lower.includes("thue suat") ||
              lower.includes("thuế") ||
              lower.includes("thue") ||
              lower.includes("vat") ||
              lower.includes("tax") ||
              lower.includes("%")
            );
          }) || cols[1] || cols[0];

          // Map every row strictly to retain only Mã món and Thuế bán hàng
          const formattedProducts = jsonData.map((row: any) => {
            const rawCode = String(row[codeCol] || "").trim();
            const rawTax = row[taxCol];
            
            let parsedTax = 0.08; // Default fallback is 8%
            if (rawTax !== undefined && rawTax !== null) {
              if (typeof rawTax === "number") {
                parsedTax = rawTax;
              } else {
                const replaced = String(rawTax).replace("%", "").trim();
                const parsed = parseFloat(replaced);
                if (!isNaN(parsed)) {
                  parsedTax = parsed > 1 ? parsed / 100 : parsed;
                }
              }
            }
            return {
              "Mã món": rawCode,
              "Thuế bán hàng": parsedTax
            };
          }).filter(p => p["Mã món"] !== ""); // Filter out empty lines if any

          // Update column mappings to our normalized properties
          setColMapping(prev => ({
            ...prev,
            productCodeCol: "Mã món",
            productTaxCol: "Thuế bán hàng"
          }));

          onDataChange(formattedProducts, invoices);
          addLog(`Đã tải Danh sách hàng hóa từ: ${file.name} (${formattedProducts.length} dòng). Chỉ giữ lại 2 cột: 'Mã món' (từ '${codeCol}') và 'Thuế bán hàng' (từ '${taxCol}').`);
        } else {
          // Keep all rows in raw state so they can be viewed in Step 1
          onDataChange(products, jsonData);
          addLog(`Đã tải danh sách chi tiết hóa đơn từ: ${file.name} (${jsonData.length} dòng thành công).`);
        }
      } catch (err: any) {
        console.error("Error parsing file:", err);
        alert(`Có lỗi xảy ra khi đọc file: ${err.message || err}`);
        addLog(`[Lỗi] Không thể phân tích file ${file.name}. Lỗi: ${err.message || err}`);
      }
    };
    reader.onerror = (err) => {
      alert("Lỗi đọc file từ máy tính!");
      addLog(`[Lỗi] FileReader gặp sự cố với tệp: ${file.name}`);
    };
    reader.readAsArrayBuffer(file);
  };

  const addLog = (msg: string) => {
    const time = new Date().toLocaleTimeString("vi-VN");
    setProcessingLogs(p => [`[${time}] ${msg}`, ...p]);
  };

  // ----------------------------------------------------
  // BƯỚC 2: CHUẨN HÓA & LÀM GIÀU DỮ LIỆU
  // ----------------------------------------------------
  const runStep2Processing = () => {
    if (products.length === 0 || invoices.length === 0) {
      alert("Vui lòng nạp đầy đủ dữ liệu gốc tại Bước 1!");
      return;
    }

    addLog("[Bước 2] Đang tiến hành làm giàu dữ liệu nguồn...");
    
    let toProcess = invoices;

    // Build tax dictionary lookup from Catalog (using clean normalized keys)
    const taxDict: Record<string, number> = {};
    products.forEach((p) => {
      const rawCode = String(p["Mã món"] || "").trim().toLowerCase();
      const rawTax = p["Thuế bán hàng"];
      
      let parsedTax = 0.08; // Default fallback: 8%
      if (rawTax !== undefined && rawTax !== null) {
        if (typeof rawTax === "number") {
          parsedTax = rawTax;
        } else {
          const replaced = String(rawTax).replace("%", "").trim();
          const parsed = parseFloat(replaced);
          if (!isNaN(parsed)) {
            parsedTax = parsed > 1 ? parsed / 100 : parsed;
          }
        }
      }
      if (rawCode) {
        taxDict[rawCode] = parsedTax;
      }
    });

    // Enrich invoice details with 3 columns: "Thuế bán hàng", "Giá bán trước thuế" and "Tiền thuế"
    const enriched = toProcess.map((row) => {
      const rawItemCode = getValueByFlexibleKey(row, colMapping.invoiceProductCodeCol);
      const itemCode = String(rawItemCode || "").trim().toLowerCase();
      const taxRate = taxDict[itemCode] !== undefined ? taxDict[itemCode] : 0.08;
      
      const rawPrice = getValueByFlexibleKey(row, colMapping.invoicePriceCol);
      const priceVal = parseVietnameseNumber(rawPrice);
      
      // Formula: Cột BA ( Giá bán trước thuế ) = cột AU (Giá bán) / ( 1 + cột AZ ( Thuế bán hàng ))
      const priceBeforeTax = Math.round((priceVal / (1 + taxRate)) * 100) / 100;
      // Formula: Cột BB (Tiền thuế) = cột BA (Giá bán trước thuế) * cột AZ ( Thuế bán hàng )
      const taxAmount = Math.round((priceBeforeTax * taxRate) * 100) / 100;

      return {
        ...row,
        "Thuế bán hàng": taxRate,
        "Giá bán trước thuế": priceBeforeTax,
        "Tiền thuế": taxAmount
      };
    });

    const uniqueTaxes = Array.from(new Set(enriched.map((item) => item["Thuế bán hàng"]))) as number[];
    const hasDiscount = enriched.some((item) => {
      const rawAmt = getValueByFlexibleKey(item, colMapping.invoiceDiscountAmountCol);
      const parsedAmt = Math.max(0, parseVietnameseNumber(rawAmt));
      return parsedAmt > 0;
    });

    // Calculate unique invoice id counts
    const uniqueInvoiceIds = Array.from(
      new Set(enriched.map((item) => String(getValueByFlexibleKey(item, colMapping.invoiceIdCol) || "").trim()).filter(Boolean))
    );
    const totalInvoicesCount = uniqueInvoiceIds.length;

    // Calculate the number of tax levels for each unique invoice
    const invoiceTaxDetails: Record<string, { itemCount: number; taxRates: Set<number> }> = {};
    enriched.forEach((item) => {
      const invId = String(getValueByFlexibleKey(item, colMapping.invoiceIdCol) || "").trim();
      if (!invId) return;
      const taxRate = item["Thuế bán hàng"] !== undefined ? item["Thuế bán hàng"] : 0.08;

      if (!invoiceTaxDetails[invId]) {
        invoiceTaxDetails[invId] = { itemCount: 0, taxRates: new Set<number>() };
      }
      invoiceTaxDetails[invId].itemCount += 1;
      invoiceTaxDetails[invId].taxRates.add(taxRate);
    });

    const invoiceTaxSummaries = Object.entries(invoiceTaxDetails).map(([invId, data]) => ({
      invoiceId: invId,
      itemCount: data.itemCount,
      taxRates: Array.from(data.taxRates).sort((a, b) => a - b),
      taxCount: data.taxRates.size,
    }));

    setStep2Invoices(enriched);
    setStep2Stats({
      taxRates: uniqueTaxes,
      hasInvoiceDiscount: hasDiscount,
      totalInvoicesCount,
      invoiceTaxSummaries,
    });
    setStep2Processed(true);
    setStep2Page(1);

    addLog(`[Bước 2] Đã bổ sung thành công cột 'Thuế bán hàng', 'Giá bán trước thuế' và 'Tiền thuế' cho ${enriched.length} dòng dữ liệu.`);
    addLog(`[Bước 2] Thống kê: Tổng số có ${totalInvoicesCount} hóa đơn. Phát hiện có ${uniqueTaxes.length} mức thuế suất khác nhau.`);
    setActiveStepTab("step2");
  };

  // Export intermediate Step 2 processor file
  const handleExportStep2 = () => {
    if (step2Invoices.length === 0) {
      alert("Không có dữ liệu Bước 2 để tải!");
      return;
    }

    addLog("[Bước 2] Đang tạo tệp Excel đã bổ sung cột...");

    // Create custom output keeping 100% of original keys and values in exact order, plus supplementary columns
    const formattedRows = step2Invoices.map((row) => {
      const outputRow: Record<string, any> = {};
      
      // Extract original keys from first loaded invoice row
      const origKeys = invoices.length > 0 ? Object.keys(invoices[0]) : Object.keys(row).filter(k => k !== "Thuế bán hàng" && k !== "Giá bán trước thuế" && k !== "Tiền thuế");
      
      // Map exact original data 100%
      origKeys.forEach((key) => {
        outputRow[key] = row[key];
      });
      
      // Append calculated tax, base priced, and tax amount
      const limitTax = row["Thuế bán hàng"];
      outputRow["Thuế bán hàng"] = limitTax !== undefined ? `${Math.round(limitTax * 100)}%` : "8%";
      outputRow["Giá bán trước thuế"] = row["Giá bán trước thuế"] !== undefined ? row["Giá bán trước thuế"] : 0;
      outputRow["Tiền thuế"] = row["Tiền thuế"] !== undefined ? row["Tiền thuế"] : 0;
      
      return outputRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(formattedRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DanhSachChiTietHoaDon_Daxuly");
    XLSX.writeFile(workbook, "DanhSachChiTietHoaDon_Daxuly.xlsx");
    addLog("[Bước 2] Đã tải xuống thành công tệp 'DanhSachChiTietHoaDon_Daxuly.xlsx' về máy tính.");
  };

  // ----------------------------------------------------
  // BƯỚC 3: GỘP NHÓM & ÁNH XẠ CHUYỂN ĐỔI GTGT
  // ----------------------------------------------------
  const runStep3Processing = () => {
    if (!step2Processed || step2Invoices.length === 0) {
      alert("Vui lòng thực hiện Bước 2 (Cập nhật dữ liệu bổ sung) trước khi chạy Bước 3!");
      return;
    }

    addLog("[Bước 3] Bắt đầu gộp nhóm và ánh xạ sang Mau_import_hoa_don_GTGT...");
    
    // Group invoices by 'Mã hóa đơn'
    const groupedInvoices: Record<string, any[]> = {};
    step2Invoices.forEach((row) => {
      const invId = String(getValueByFlexibleKey(row, colMapping.invoiceIdCol) || "").trim();
      if (!invId) return;
      if (!groupedInvoices[invId]) {
        groupedInvoices[invId] = [];
      }
      groupedInvoices[invId].push(row);
    });

    const finalRows: OutputRow[] = [];
    const dObj = new Date();
    const dayStr = String(dObj.getDate()).padStart(2, "0");
    const monthStr = String(dObj.getMonth() + 1).padStart(2, "0");
    const yearStr = dObj.getFullYear();
    const currentDateStr = `${dayStr}/${monthStr}/${yearStr}`;

    Object.entries(groupedInvoices).forEach(([invId, groupItems]) => {
      // First line of each group contains customer info, other row templates leave them blank
      let isFirstRowOfGroup = true;

      // Extract general group values (e.g. invoice discount if exists)
      const rawInvDiscount = getValueByFlexibleKey(groupItems[0], colMapping.invoiceDiscountAmountCol);
      const invoiceDiscount = Math.max(0, parseVietnameseNumber(rawInvDiscount));

      // We'll track total item amount (inclusive of taxes) to divide trade discounts proportionally
      let tempGroupOutputRows: OutputRow[] = [];
      let totalItemValueForAllocation = 0;

      groupItems.forEach((item) => {
        const itemCode = String(getValueByFlexibleKey(item, colMapping.invoiceProductCodeCol) || "").trim();
        const itemName = String(getValueByFlexibleKey(item, colMapping.invoiceProductNameCol) || "").trim();
        
        // Match tax rate enriched in step 2
        const taxRate = item["Thuế bán hàng"] !== undefined ? item["Thuế bán hàng"] : 0.08;

        const qty = parseVietnameseNumber(getValueByFlexibleKey(item, colMapping.invoiceQtyCol));
        const giaTruocThue = item["Giá bán trước thuế"] !== undefined ? item["Giá bán trước thuế"] : 0;
        const discountPercent = parseVietnameseNumber(getValueByFlexibleKey(item, colMapping.invoiceDiscountPercentCol));

        // Apply selected 'Thành tiền' calculation formulas
        let thanhTien = 0;
        if (formula === "standard_accounting") {
          const discAmount = giaTruocThue * qty * (discountPercent / 100);
          thanhTien = Math.round(((giaTruocThue * qty) - discAmount) * 100) / 100;
        } else if (formula === "standard_discount_rate") {
          thanhTien = Math.round((giaTruocThue * qty * (1 - discountPercent / 100)) * 100) / 100;
        } else if (formula === "multiplier") {
          thanhTien = Math.round((giaTruocThue * discountPercent) * 100) / 100;
        } else {
          thanhTien = Math.round((giaTruocThue * qty) * 100) / 100;
        }

        const tChietKhau = Math.round(giaTruocThue * qty * (discountPercent / 100));
        const tVat = Math.round(thanhTien * taxRate);
        const totalRowAmount = Math.round(thanhTien + tVat);
        totalItemValueForAllocation += totalRowAmount;

        const mappedTaxStr = `${Math.round(taxRate * 100)}%`;

        const outRow: OutputRow = {
          "Mã nhóm hóa đơn": invId,
          "Ngày hóa đơn": isFirstRowOfGroup ? currentDateStr : "",
          "Mã khách hàng": isFirstRowOfGroup ? "KL" : "",
          "Mã số thuế": "",
          "Tên khách hàng": isFirstRowOfGroup ? "Khách lẻ không lấy hóa đơn" : "",
          "Họ tên người mua": "",
          "Số CCCD": "",
          "Địa chỉ": "",
          "Điện thoại": "",
          "Email nhận HĐ": "",
          "Hình thức thanh toán": isFirstRowOfGroup ? "TM/CK" : "",
          "Tên ngân hàng": "",
          "Mã hàng": itemCode,
          "Tên hàng hóa, dịch vụ *": itemName,
          "Tính chất hàng hóa, dịch vụ": "Hàng hóa, dịch vụ",
          "Đơn vị tính": "",
          "Số lượng": qty,
          "Đơn giá": giaTruocThue,
          "Chiết khấu (%)": discountPercent,
          "Tiền chiết khấu": tChietKhau > 0 ? tChietKhau : "",
          "Thành tiền": thanhTien,
          "% VAT": mappedTaxStr,
          "Tiền VAT": tVat > 0 ? tVat : "",
          "Tổng tiền *": totalRowAmount,
        };

        tempGroupOutputRows.push(outRow);
        isFirstRowOfGroup = false;
      });

      // Insert separate CKTM rows for each unique tax rate group if trade discount exists
      if (invoiceDiscount > 0 && totalItemValueForAllocation > 0) {
        const vatGroups: Record<number, OutputRow[]> = {};
        tempGroupOutputRows.forEach((row) => {
          // retrieve actual numeric rate of the code
          const origMatch = groupItems.find(g => {
            const cleanG = String(getValueByFlexibleKey(g, colMapping.invoiceProductCodeCol) || "").trim().replace(/\s+/g, "").toLowerCase();
            const cleanRow = String(row["Mã hàng"] || "").trim().replace(/\s+/g, "").toLowerCase();
            return cleanG === cleanRow;
          });
          const taxRate = origMatch && origMatch["Thuế bán hàng"] !== undefined ? origMatch["Thuế bán hàng"] : 0.08;
          
          if (vatGroups[taxRate] === undefined) {
            vatGroups[taxRate] = [];
          }
          vatGroups[taxRate].push(row);
        });

        Object.entries(vatGroups).forEach(([taxRateStr, rowsOfTax]) => {
          const rate = parseFloat(taxRateStr);
          const totalOfTaxRateGroup = rowsOfTax.reduce((sum, r) => sum + r["Tổng tiền *"], 0);

          if (totalOfTaxRateGroup <= 0) return;

          // Proportional allocation: Giảm giá hóa đơn * Tổng tiền hành cùng thuế suất / Tổng hóa đơn
          const allocatedTotalWithTax = invoiceDiscount * (totalOfTaxRateGroup / totalItemValueForAllocation);
          const allocatedTotalRounded = Math.round(allocatedTotalWithTax);

          // Apply sign negative / positive
          const totalVal = isNegativeCKTM ? -allocatedTotalRounded : allocatedTotalRounded;
          
          // Back-calculate: Thành tiền = Tổng tiền / (1 + % VAT)
          const allocatedThanhTien = Math.round((totalVal / (1 + rate)) * 100) / 100;

          // Unit Price is allocatedThanhTien since Quantity defaults to 1 as requested
          const allocatedDonGia = allocatedThanhTien;
          const allocatedTienVat = Math.round(totalVal - allocatedThanhTien);

          const vatPctStr = `${Math.round(rate * 100)}%`;

          const cktmRow: OutputRow = {
            "Mã nhóm hóa đơn": invId,
            "Ngày hóa đơn": "",
            "Mã khách hàng": "",
            "Mã số thuế": "",
            "Tên khách hàng": "",
            "Họ tên người mua": "",
            "Số CCCD": "",
            "Địa chỉ": "",
            "Điện thoại": "",
            "Email nhận HĐ": "",
            "Hình thức thanh toán": "",
            "Tên ngân hàng": "",
            "Mã hàng": `CKTM${vatPctStr}`,
            "Tên hàng hóa, dịch vụ *": "Chiết khấu thương mại",
            "Tính chất hàng hóa, dịch vụ": "Hàng hóa chiết khấu thương mại",
            "Đơn vị tính": "",
            "Số lượng": 1, // Mặc định SL là 1 cho dòng Chiết khấu thương mại như yêu cầu!
            "Đơn giá": allocatedDonGia,
            "Chiết khấu (%)": 0,
            "Tiền chiết khấu": "",
            "Thành tiền": allocatedThanhTien,
            "% VAT": vatPctStr,
            "Tiền VAT": allocatedTienVat,
            "Tổng tiền *": totalVal,
          };

          tempGroupOutputRows.push(cktmRow);
        });
      }

      finalRows.push(...tempGroupOutputRows);
    });

    setOutputData(finalRows);
    setStep3Processed(true);
    addLog(`[Bước 3] Hoàn thành ánh xạ! Đã gộp và tạo thành công ${finalRows.length} dòng dữ liệu.`);
    setActiveStepTab("step3");
  };

  // Export processed data to standard Excel file
  const handleExport = () => {
    if (outputData.length === 0) {
      alert("Không có dữ liệu đầu ra để xuất bản! Vui lòng hoàn thành các bước trước.");
      return;
    }

    addLog("[Bước 3] Đang ghi xuất tệp tin Import Mau_import_hoa_don_GTGT...");

    // Prepare sheet columns ensuring exact 24 columns in strict requested sequence
    const cleanedRows = outputData.map((row) => ({
      "Mã nhóm hóa đơn *": row["Mã nhóm hóa đơn"],
      "Ngày hóa đơn": row["Ngày hóa đơn"],
      "Mã khách hàng": row["Mã khách hàng"],
      "Mã số thuế": row["Mã số thuế"],
      "Tên khách hàng": row["Tên khách hàng"],
      "Họ tên người mua": row["Họ tên người mua"],
      "Số CCCD": row["Số CCCD"],
      "Địa chỉ": row["Địa chỉ"],
      "Điện thoại": row["Điện thoại"],
      "Email nhận HĐ": row["Email nhận HĐ"],
      "Hình thức thanh toán": row["Hình thức thanh toán"],
      "Tên ngân hàng": row["Tên ngân hàng"],
      "Mã hàng": row["Mã hàng"],
      "Tên hàng hóa, dịch vụ *": row["Tên hàng hóa, dịch vụ *"],
      "Tính chất hàng hóa, dịch vụ": row["Tính chất hàng hóa, dịch vụ"],
      "Đơn vị tính": row["Đơn vị tính"],
      "Số lượng": row["Số lượng"],
      "Đơn giá": row["Đơn giá"],
      "Chiết khấu (%)": row["Chiết khấu (%)"],
      "Tiền chiết khấu": row["Tiền chiết khấu"],
      "Thành tiền": row["Thành tiền"],
      "% VAT": row["% VAT"],
      "Tiền VAT": row["Tiền VAT"],
      "Tổng tiền *": row["Tổng tiền *"],
    }));

    const worksheet = XLSX.utils.json_to_sheet(cleanedRows);
    const workbook = XLSX.utils.book_new();

    // Auto fit column widths for premium visual experience (24 columns total)
    const colWidths = [
      { wch: 18 }, // Mã nhóm hóa đơn *
      { wch: 15 }, // Ngày hóa đơn
      { wch: 15 }, // Mã khách hàng
      { wch: 15 }, // Mã số thuế
      { wch: 30 }, // Tên khách hàng
      { wch: 20 }, // Họ tên người mua
      { wch: 15 }, // Số CCCD
      { wch: 25 }, // Địa chỉ
      { wch: 15 }, // Điện thoại
      { wch: 25 }, // Email nhận HĐ
      { wch: 20 }, // Hình thức thanh toán
      { wch: 15 }, // Tên ngân hàng
      { wch: 15 }, // Mã hàng
      { wch: 35 }, // Tên hàng hóa, dịch vụ *
      { wch: 30 }, // Tính chất hàng hóa, dịch vụ
      { wch: 15 }, // Đơn vị tính
      { wch: 10 }, // Số lượng
      { wch: 15 }, // Đơn giá
      { wch: 15 }, // Chiết khấu (%)
      { wch: 15 }, // Tiền chiết khấu
      { wch: 15 }, // Thành tiền
      { wch: 10 }, // % VAT
      { wch: 15 }, // Tiền VAT
      { wch: 18 }, // Tổng tiền *
    ];
    worksheet["!cols"] = colWidths;

    XLSX.utils.book_append_sheet(workbook, worksheet, "Mau_import_hoa_don_GTGT");
    XLSX.writeFile(workbook, "Mau_import_hoa_don_GTGT.xlsx");
    addLog("Xuất thành công tập tin Mau_import_hoa_don_GTGT.xlsx về máy tính của bạn!");
  };

  const filteredSummaries = step2Stats?.invoiceTaxSummaries
    ? step2Stats.invoiceTaxSummaries.filter(s =>
        s.invoiceId.toLowerCase().includes(step2SearchQuery.toLowerCase())
      )
    : [];

  return (
    <div className="space-y-6">
      {/* Visual Stepper Progress Bar */}
      <div className="bg-white rounded-xl border border-slate-200/80 p-4 shadow-sm">
        <div className="flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center space-x-3">
            <span className="p-2 bg-blue-50 text-blue-600 rounded-lg">
              <Settings2 className="h-5 w-5" />
            </span>
            <div>
              <h3 className="font-bold text-slate-800 text-sm">Phân Tách Quy Trình 3 Bước Để Chuyển Đổi Hóa Đơn</h3>
              <p className="text-xs text-slate-500 mt-0.5">Vận hành theo trình tự kiểm soát kế toán để tránh sai lệch dữ liệu thuế.</p>
            </div>
          </div>
          
          {/* Circular/Line visual step indicator */}
          <div className="flex items-center space-x-2 md:space-x-4">
            <button
              onClick={() => setActiveStepTab("step1")}
              className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
                activeStepTab === "step1"
                  ? "bg-blue-600 text-white border-blue-600"
                  : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
              }`}
            >
              <span className="w-5 h-5 flex items-center justify-center bg-white/20 rounded-full text-[10px]">1</span>
              <span>Bước 1: Nạp File</span>
            </button>
            <div className="text-slate-300">➔</div>
            <button
              onClick={() => {
                if (products.length === 0 || invoices.length === 0) {
                  alert("Vui lòng nạp đủ dữ liệu nguồn tại Bước 1 trước!");
                  return;
                }
                setActiveStepTab("step2");
              }}
              className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
                activeStepTab === "step2"
                  ? "bg-blue-600 text-white border-blue-600"
                  : step2Processed
                  ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                  : "bg-slate-50 text-slate-400 border-slate-200 hover:bg-slate-100"
              }`}
            >
              <span className={`w-5 h-5 flex items-center justify-center rounded-full text-[10px] ${step2Processed ? "bg-emerald-600 text-white" : "bg-white/20"}`}>
                {step2Processed ? "✓" : "2"}
              </span>
              <span>Bước 2: Bổ Sung Cột</span>
            </button>
            <div className="text-slate-300">➔</div>
            <button
              onClick={() => {
                if (!step2Processed) {
                  alert("Vui lòng thực hiện Bước 2 (Cập nhật dữ liệu bổ sung) trước!");
                  return;
                }
                setActiveStepTab("step3");
              }}
              className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
                activeStepTab === "step3"
                  ? "bg-blue-600 text-white border-blue-600"
                  : step3Processed
                  ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                  : "bg-slate-50 text-slate-400 border-slate-200 hover:bg-slate-100"
              }`}
            >
              <span className={`w-5 h-5 flex items-center justify-center rounded-full text-[10px] ${step3Processed ? "bg-emerald-600 text-white" : "bg-white/20"}`}>
                {step3Processed ? "✓" : "3"}
              </span>
              <span>Bước 3: Tạo File Import</span>
            </button>
          </div>
        </div>
      </div>

      {/* STEP 1 CONTAINER: NẠP TỆP SOURCE PATH DATA */}
      {activeStepTab === "step1" && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Upload Box 1: Catalog */}
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between h-60 group relative overflow-hidden">
              <div>
                <div className="flex items-center space-x-2 text-blue-600 mb-3">
                  <FileSpreadsheet className="h-5 w-5" />
                  <h4 className="text-xs font-extrabold text-slate-400 uppercase tracking-wider">1. Tệp Danh Mục Hàng Hóa</h4>
                </div>
                <h3 className="font-bold text-slate-800 text-sm mt-1">DanhSachSanPham.xlsx</h3>
                <p className="text-xs text-slate-500 leading-relaxed mt-1">
                  Chứa bảng tra cứu chuẩn của kế toán gồm [Mã hàng] và [Thuế bán hàng] để tự động đối chiếu thuế suất bán ra khi bán hàng.
                </p>
              </div>
              
              <div className="mt-4">
                <label className="flex flex-col items-center justify-center border border-dashed border-slate-300 group-hover:border-blue-400 rounded-lg py-3 px-3 cursor-pointer bg-slate-50 hover:bg-blue-50/20 transition duration-200">
                  <Upload className="h-5 w-5 text-slate-400 group-hover:text-blue-500 mb-1 transition duration-200" />
                  <span className="text-[11px] text-slate-600 group-hover:text-blue-700 font-bold">Tải lên file Danh mục hàng hóa</span>
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    className="hidden" 
                    onChange={(e) => {
                      if (e.target.files?.[0]) {
                        handleFileUpload("products", e.target.files[0]);
                      }
                      e.target.value = "";
                    }}
                  />
                </label>
                {products.length > 0 ? (
                  <div className="flex items-center space-x-1.5 mt-2.5 text-[11px] text-blue-600 font-bold bg-blue-50 p-1.5 rounded border border-blue-100">
                    <span className="inline-block w-2 h-2 rounded-full bg-blue-500 animate-pulse"></span>
                    <span>Đã nạp {products.length} dòng hàng hóa thành công</span>
                  </div>
                ) : (
                  <div className="text-[10px] text-slate-400 mt-2 italic">Chưa chọn tập tin</div>
                )}
              </div>
            </div>

            {/* Upload Box 2: Invoices details */}
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between h-60 group relative overflow-hidden">
              <div>
                <div className="flex items-center space-x-2 text-teal-600 mb-3">
                  <FileSpreadsheet className="h-5 w-5" />
                  <h4 className="text-xs font-extrabold text-slate-400 uppercase tracking-wider">2. Tệp Danh Sách Hóa Đơn Gốc</h4>
                </div>
                <h3 className="font-bold text-slate-800 text-sm mt-1">DanhSachChiTietHoaDon.xlsx</h3>
                <p className="text-xs text-slate-500 leading-relaxed mt-1">
                  Chứa chi tiết các dòng nguyên bản của hóa đơn bán ra (chưa bổ sung thông tin thuế bán hàng hóa và giá trước thuế).
                </p>
              </div>

              <div className="mt-4">
                <label className="flex flex-col items-center justify-center border border-dashed border-slate-300 group-hover:border-teal-400 rounded-lg py-3 px-3 cursor-pointer bg-slate-50 hover:bg-teal-50/20 transition duration-200">
                  <Upload className="h-5 w-5 text-slate-400 group-hover:text-teal-500 mb-1 transition duration-200" />
                  <span className="text-[11px] text-slate-600 group-hover:text-teal-700 font-bold">Tải lên file Chi tiết hóa đơn cần chuyển đổi</span>
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    className="hidden" 
                    onChange={(e) => {
                      if (e.target.files?.[0]) {
                        handleFileUpload("invoices", e.target.files[0]);
                      }
                      e.target.value = "";
                    }}
                  />
                </label>
                {invoices.length > 0 ? (
                  <div className="flex items-center space-x-1.5 mt-2.5 text-[11px] text-teal-600 font-bold bg-teal-50 p-1.5 rounded border border-teal-100">
                    <span className="inline-block w-2 h-2 rounded-full bg-teal-500 animate-pulse"></span>
                    <span>Đã nạp {invoices.length} dòng hóa đơn thành công</span>
                  </div>
                ) : (
                  <div className="text-[10px] text-slate-400 mt-2 italic">Chưa chọn tập tin</div>
                )}
              </div>
            </div>
          </div>

          {/* Displays full loaded source contents below */}
          {(products.length > 0 || invoices.length > 0) && (
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-5 space-y-4">
              <div className="border-b border-slate-100 pb-3 flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider">Hiển thị thông tin đã nạp được đầy đủ</h4>
                  <p className="text-[11px] text-slate-500 mt-0.5">Xem dữ liệu thô của các tệp nguồn để xác tín quy trình Bước 1.</p>
                </div>

                <div className="flex flex-wrap items-center gap-3">


                  {products.length > 0 && invoices.length > 0 && (
                    <button
                      onClick={runStep2Processing}
                      id="goto-step2-btn-from-preview"
                      className="flex items-center space-x-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-bold shadow-sm transition active:scale-95 cursor-pointer"
                    >
                      <span>Tiến hành Bước 2: Bổ xung cột</span>
                      <ArrowRight className="h-4 w-4" />
                    </button>
                  )}
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                {/* Product Catalog display block */}
                <div className="lg:col-span-5 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700">Tra cứu danh mục ({products.length} dòng)</span>
                    <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-bold text-slate-500">Catalog</span>
                  </div>
                  {products.length === 0 ? (
                    <div className="h-32 flex items-center justify-center border rounded bg-slate-50 text-slate-400 text-xs italic">Chưa có dữ liệu danh mục sản phẩm</div>
                  ) : (
                    <div className="border rounded-lg overflow-hidden max-h-64 overflow-y-auto">
                      <table className="w-full text-[11px] text-left text-slate-500">
                        <thead className="bg-slate-50 sticky top-0 text-slate-700 border-b font-semibold z-10">
                          <tr>
                            <th className="p-2">Mã món</th>
                            <th className="p-2">Thuế bán hàng</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y font-mono">
                          {products.slice(0, 10).map((row, idx) => {
                            const codeVal = String(row["Mã món"] || "");
                            const taxVal = row["Thuế bán hàng"];
                            return (
                              <tr key={idx} className="hover:bg-slate-50/50">
                                <td className="p-2 text-blue-600 font-bold">{codeVal}</td>
                                <td className="p-2 text-amber-800 font-semibold">
                                  {typeof taxVal === "number" ? `${Math.round(taxVal * 100)}%` : String(taxVal || "")}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      {products.length > 10 && (
                        <div className="bg-slate-50 p-1.5 text-center text-[10px] text-slate-400 font-sans border-t">Hiển thị 10 dòng đầu...</div>
                      )}
                    </div>
                  )}
                </div>

                {/* Raw Invoice rows display block */}
                <div className="lg:col-span-7 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700">Danh sách hóa đơn thô ({invoices.length} dòng)</span>
                    <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded font-bold text-slate-500">Original Invoices</span>
                  </div>
                  {invoices.length === 0 ? (
                    <div className="h-32 flex items-center justify-center border rounded bg-slate-50 text-slate-400 text-xs italic">Chưa có dữ liệu hóa đơn nhập vào</div>
                  ) : (
                    <div className="border rounded-lg overflow-hidden max-h-64 overflow-y-auto overflow-x-auto">
                      <table className="w-full text-[11px] text-left text-slate-500 whitespace-nowrap">
                        <thead className="bg-slate-50 sticky top-0 text-slate-700 border-b font-semibold z-10">
                          <tr>
                            {Object.keys(invoices[0] || {}).map((col, idx) => (
                              <th key={idx} className="p-2 border-r last:border-r-0 border-slate-200">{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y font-mono">
                          {invoices.slice(0, 15).map((row, idx) => {
                            const cells = Object.keys(invoices[0] || {});
                            return (
                              <tr key={idx} className="hover:bg-slate-50/50">
                                {cells.map((col, cIdx) => {
                                  const val = row[col];
                                  const formattedVal = typeof val === "number" ? val.toLocaleString("vi-VN") : String(val !== undefined && val !== null ? val : "");
                                  return (
                                    <td key={cIdx} className="p-2 border-r last:border-r-0 border-slate-150 text-slate-700 font-sans">
                                      {formattedVal}
                                    </td>
                                  );
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      {invoices.length > 15 && (
                        <div className="bg-slate-50 p-1.5 text-center text-[10px] text-slate-400 font-sans border-t">Hiển thị 15 dòng đầu trong tổng số {invoices.length} dòng...</div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* STEP 2 CONTAINER: CHUẨN HÓA & BỔ SUNG CỘT */}
      {activeStepTab === "step2" && (
        <div className="space-y-6">
          <div className="bg-white rounded-xl border border-slate-200/80 p-5 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
              <div className="space-y-1">
                <span className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded text-[10px] font-bold uppercase tracking-wider">Tính toán nạp thuế & đơn giá trước thuế</span>
                <h3 className="font-bold text-slate-800 text-sm">Bước 2: Cập Nhật Dữ Liệu Bổ Sung</h3>
                <p className="text-xs text-slate-500 leading-relaxed max-w-2xl">
                  Hành động này sẽ thực hiện truy vấn với danh mục hàng hóa dựa vào [Mã hàng] để xác định và đưa cột <strong>Thuế bán hàng</strong> vào danh sách hóa đơn. 
                  Đồng thời dựa vào cột <strong>Giá bán</strong> dòng đầu để tạo cột <strong>Giá bán trước thuế</strong> = Giá bán / (1 + Thuế bán hàng), làm tròn 2 chữ số thập phân.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 shrink-0">
                <button
                  onClick={runStep2Processing}
                  id="trigger-step2-enrich-btn"
                  className="flex items-center space-x-1.5 px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-bold shadow-md transition duration-200 active:scale-95 cursor-pointer"
                >
                  <RefreshCcw className="h-4 w-4 shrink-0 animate-spin-slow" />
                  <span>Cập nhật dữ liệu bổ sung</span>
                </button>
                
                {step2Processed && (
                  <button
                    onClick={handleExportStep2}
                    id="download-step2-xlsx"
                    className="flex items-center space-x-1.5 px-3 py-2.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-lg text-xs font-bold shadow-sm transition cursor-pointer"
                  >
                    <Download className="h-4 w-4 shrink-0" />
                    <span>Tải Excel chi tiết đã bổ sung (.xlsx)</span>
                  </button>
                )}
              </div>
            </div>

            {/* Step 2 statistical outputs */}
            {step2Processed && step2Stats ? (
              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="bg-blue-50/70 border border-blue-150 rounded-xl p-4 flex items-center space-x-4">
                    <div className="p-3 bg-blue-500 text-white rounded-lg shrink-0">
                      <Check className="h-5 w-5" />
                    </div>
                    <div className="space-y-0.5">
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Tổng số hóa đơn phát hiện</div>
                      <div className="text-sm font-bold text-slate-800">
                        Có <span className="text-blue-600">{step2Stats.totalInvoicesCount}</span> hóa đơn
                      </div>
                      <p className="text-[10px] text-slate-500 font-sans">Được đếm từ nguồn dữ liệu sau khi gộp trùng lặp Mã hóa đơn.</p>
                    </div>
                  </div>

                  <div className="bg-emerald-50/70 border border-emerald-150 rounded-xl p-4 flex items-center space-x-4">
                    <div className="p-3 bg-emerald-500 text-white rounded-lg shrink-0">
                      <Check className="h-5 w-5" />
                    </div>
                    <div className="space-y-0.5">
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mức thuế suất phát sinh</div>
                      <div className="text-sm font-bold text-slate-800">
                        Có {step2Stats.taxRates.length} mức: <span className="text-emerald-700 font-mono">{step2Stats.taxRates.map(r => `${Math.round(r * 100)}%`).join(", ")}</span>
                      </div>
                      <p className="text-[10px] text-slate-500 font-sans">Xác thực tự động từ look-up mã kết nối giữa Mã hàng & Mã món.</p>
                    </div>
                  </div>

                  <div className="bg-teal-50/70 border border-teal-150 rounded-xl p-4 flex items-center space-x-4">
                    <div className="p-3 bg-teal-500 text-white rounded-lg shrink-0">
                      <Check className="h-5 w-5" />
                    </div>
                    <div className="space-y-0.5">
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Trạng thái giảm giá hóa đơn</div>
                      <div className="text-sm font-bold text-slate-800 flex items-center space-x-1.5">
                        {step2Stats.hasInvoiceDiscount ? (
                          <span className="px-2 py-0.5 bg-red-100 text-red-700 text-xs rounded-full font-sans font-bold">CÓ giảm giá</span>
                        ) : (
                          <span className="px-2 py-0.5 bg-slate-100 text-slate-500 text-xs rounded-full font-sans font-medium">KHÔNG giảm giá</span>
                        )}
                      </div>
                      <p className="text-[10px] text-slate-500 font-sans">Chiết khấu cấp hóa đơn phát hiện được phân bổ thông minh ở Bước 3.</p>
                    </div>
                  </div>
                </div>

                {/* Chẩn đoán khớp cột & kiểm tra công thức tính toán */}
                {step2Invoices.length > 0 && invoices.length > 0 && (
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                    <div className="border-b border-slate-200 pb-2">
                      <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider flex items-center space-x-1.5 text-indigo-700">
                        <span className="inline-block w-2 h-2 rounded-full bg-indigo-500 animate-pulse"></span>
                        <span>Bản đồ chẩn đoán giá trị & công thức dòng đầu tiên</span>
                      </h4>
                      <p className="text-[10px] text-slate-500 mt-0.5 font-sans">
                        Kiểm thử tự động giúp bạn xác minh hệ thống có lấy đúng nguồn cột [Giá bán] hay không, và xem trực quan kết quả từng bước làm tròn.
                      </p>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 text-xs font-sans">
                      <div className="bg-white p-2.5 rounded-lg border border-slate-150 shadow-sm flex flex-col justify-between">
                        <div>
                          <span className="block text-[9px] text-slate-400 font-bold uppercase tracking-wider">Cột bóc tách đơn giá (AU)</span>
                          <span className="block text-slate-700 font-bold font-mono mt-0.5 border-b pb-1 truncate" title={colMapping.invoicePriceCol}>
                            {colMapping.invoicePriceCol}
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px] text-slate-500 mt-1">
                          <span>Dòng 1 trong file:</span>
                          <span className="font-bold font-mono text-indigo-600">
                            {String(getValueByFlexibleKey(invoices[0], colMapping.invoicePriceCol) !== undefined ? getValueByFlexibleKey(invoices[0], colMapping.invoicePriceCol) : "Trống")}
                          </span>
                        </div>
                      </div>

                      <div className="bg-white p-2.5 rounded-lg border border-slate-150 shadow-sm flex flex-col justify-between">
                        <div>
                          <span className="block text-[9px] text-slate-400 font-bold uppercase tracking-wider">Mã hàng & Mức thuế (AZ)</span>
                          <span className="block text-slate-700 font-bold font-mono mt-0.5 border-b pb-1 truncate" title={String(getValueByFlexibleKey(invoices[0], colMapping.invoiceProductCodeCol) || "N/A")}>
                            {String(getValueByFlexibleKey(invoices[0], colMapping.invoiceProductCodeCol) || "N/A")}
                          </span>
                        </div>
                        <div className="flex justify-between text-[10px] text-slate-500 mt-1">
                          <span>Thuế suất lookup:</span>
                          <span className="font-bold font-mono text-emerald-600">
                            {Math.round((step2Invoices[0]["Thuế bán hàng"] || 0) * 100)}%
                          </span>
                        </div>
                      </div>

                      <div className="bg-white p-2.5 rounded-lg border border-slate-150 shadow-sm flex flex-col justify-between">
                        <div>
                          <span className="block text-[9px] text-slate-400 font-bold uppercase tracking-wider">Giá bán trước thuế (BA)</span>
                          <span className="block text-slate-700 font-bold font-mono mt-0.5 border-b pb-1 text-blue-700">
                            {(step2Invoices[0]["Giá bán trước thuế"] ?? 0).toLocaleString("vi-VN")} đ
                          </span>
                        </div>
                        <div className="text-[9px] text-slate-400 font-mono italic mt-1">
                          Hiệu chỉnh: AU / (1 + AZ)
                        </div>
                      </div>

                      <div className="bg-white p-2.5 rounded-lg border border-slate-150 shadow-sm flex flex-col justify-between">
                        <div>
                          <span className="block text-[9px] text-slate-400 font-bold uppercase tracking-wider">Tiền thuế bổ sung (BB)</span>
                          <span className="block text-slate-700 font-bold font-mono mt-0.5 border-b pb-1 text-amber-700">
                            {(step2Invoices[0]["Tiền thuế"] ?? 0).toLocaleString("vi-VN")} đ
                          </span>
                        </div>
                        <div className="text-[9px] text-slate-400 font-mono italic mt-1">
                          Hiệu chỉnh: BA * AZ
                        </div>
                      </div>
                    </div>

                    <div className="text-[10px] text-indigo-800 bg-indigo-50/50 p-3 rounded-lg border border-indigo-100 flex items-start space-x-1.5 font-sans leading-relaxed">
                      <Info className="h-4 w-4 mt-0.5 text-indigo-500 shrink-0" />
                      <span>
                        <strong>Giải trình dòng thứ nhất:</strong> Đơn giá/Giá bán gốc (Cột AU) đọc được từ Excel là <span className="font-bold font-mono">{getValueByFlexibleKey(invoices[0], colMapping.invoicePriceCol)}</span> (bóc tách dạng số = <span className="font-bold text-slate-900">{parseVietnameseNumber(getValueByFlexibleKey(invoices[0], colMapping.invoicePriceCol)).toLocaleString("vi-VN")}</span>). 
                        Thuế suất xác định được từ danh mục (Cột AZ) là <span className="font-bold font-mono text-emerald-700">{Math.round((step2Invoices[0]["Thuế bán hàng"] ?? 0.08) * 100)}%</span>.
                        Do đó, Giá bán trước thuế (BA) = <span className="font-bold font-mono text-slate-900">{parseVietnameseNumber(getValueByFlexibleKey(invoices[0], colMapping.invoicePriceCol)).toLocaleString("vi-VN")} / (1 + {step2Invoices[0]["Thuế bán hàng"] ?? 0.08})</span> = <span className="text-blue-700 font-bold">{(step2Invoices[0]["Giá bán trước thuế"] ?? 0).toLocaleString("vi-VN")} đ</span>.
                        Tiền thuế (BB) = <span className="font-bold font-mono text-slate-900">{(step2Invoices[0]["Giá bán trước thuế"] ?? 0).toLocaleString("vi-VN")} * {Math.round((step2Invoices[0]["Thuế bán hàng"] ?? 0.08) * 100)}%</span> = <span className="text-amber-700 font-bold">{(step2Invoices[0]["Tiền thuế"] ?? 0).toLocaleString("vi-VN")} đ</span>.
                      </span>
                    </div>
                  </div>
                )}

                {/* Phân tích số mức thuế suất theo từng hóa đơn cụ thể */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200/60 pb-2">
                    <div>
                      <h4 className="font-bold text-slate-800 text-xs">Chi tiết thống kê sản phẩm & mức thuế suất theo từng hóa đơn</h4>
                      <p className="text-[10px] text-slate-500">Tìm thấy {filteredSummaries.length} hóa đơn. Mỗi dòng tương ứng hiển thị số lượng mặt hàng và danh sách mức thuế suất.</p>
                    </div>
                    <div className="flex items-center space-x-2 shrink-0">
                      <input
                        type="text"
                        placeholder="Tìm hóa đơn..."
                        value={step2SearchQuery}
                        onChange={(e) => {
                          setStep2SearchQuery(e.target.value);
                          setStep2Page(1);
                        }}
                        className="px-2.5 py-1 text-[11px] border rounded-lg bg-white text-slate-700 outline-none focus:ring-1 focus:ring-blue-500 w-40 font-semibold"
                      />
                    </div>
                  </div>

                  <div className="bg-white border rounded-lg overflow-hidden max-h-56 overflow-y-auto">
                    <table className="w-full text-left text-[11px] text-slate-600 divide-y divide-slate-200 font-mono">
                      <thead className="bg-[#f8fafc] sticky top-0 font-semibold text-slate-700 border-b z-10">
                        <tr>
                          <th className="py-2 px-3 font-sans">Mã hóa đơn</th>
                          <th className="py-2 px-3 text-center font-sans">Số sản phẩm (số dòng)</th>
                          <th className="py-2 px-3 text-center font-sans">Số lượng mức thuế</th>
                          <th className="py-2 px-3 text-right font-sans">Các mức thuế suất phát sinh</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredSummaries.length === 0 ? (
                          <tr>
                            <td colSpan={4} className="p-4 text-center text-slate-400 italic font-sans text-xs">Không tìm thấy mã hóa đơn phù hợp</td>
                          </tr>
                        ) : (
                          filteredSummaries.slice((step2Page - 1) * 8, step2Page * 8).map((summary, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/70 transition">
                              <td className="py-1.5 px-3 font-semibold text-blue-600">{summary.invoiceId}</td>
                              <td className="py-1.5 px-3 text-center font-bold text-slate-700">{summary.itemCount} sp</td>
                              <td className="py-1.5 px-3 text-center">
                                <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${summary.taxCount > 1 ? 'bg-amber-50 text-amber-700 border border-amber-100' : 'bg-emerald-50 text-emerald-700 border border-emerald-100'}`}>
                                  {summary.taxCount} mức thuế
                                </span>
                              </td>
                              <td className="py-1.5 px-3 text-right text-slate-600 font-sans font-semibold">
                                {summary.taxRates.map(r => `${Math.round(r * 100)}%`).join(", ")}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>

                  {filteredSummaries.length > 8 && (
                    <div className="flex items-center justify-between text-[10px] text-slate-500 font-sans pt-1">
                      <span>Hiển thị từ {(step2Page - 1) * 8 + 1} đến {Math.min(step2Page * 8, filteredSummaries.length)} trong tổng số {filteredSummaries.length} hóa đơn</span>
                      <div className="flex items-center space-x-1">
                        <button
                          onClick={() => setStep2Page(p => Math.max(1, p - 1))}
                          disabled={step2Page === 1}
                          className="px-2 py-0.5 border border-slate-200 rounded bg-white hover:bg-slate-50 disabled:opacity-50 text-slate-600 font-bold"
                        >
                          Trước
                        </button>
                        <span className="px-2 font-bold font-mono text-slate-700">Trang {step2Page} / {Math.ceil(filteredSummaries.length / 8)}</span>
                        <button
                          onClick={() => setStep2Page(p => Math.min(Math.ceil(filteredSummaries.length / 8), p + 1))}
                          disabled={step2Page === Math.ceil(filteredSummaries.length / 8)}
                          className="px-2 py-0.5 border border-slate-200 rounded bg-white hover:bg-slate-50 disabled:opacity-50 text-slate-600 font-bold"
                        >
                          Sau
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center p-12 bg-slate-50/50 rounded-xl border border-dashed text-slate-400 space-y-2">
                <AlertTriangle className="h-8 w-8 text-amber-500" />
                <span className="text-xs font-semibold text-slate-600">Dữ liệu Bước 2 chưa được nạp làm giàu.</span>
                <p className="text-[10px] text-slate-500">Bấm nút "Cập nhật dữ liệu bổ sung" ở góc phải để tạo và hiển thị file chi tiết có bổ sung 2 cột.</p>
              </div>
            )}
          </div>

          {/* Step 2 Excel visual preview TABLE: DanhSachChiTietHoaDon_Daxuly */}
          {step2Processed && step2Invoices.length > 0 && (
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden space-y-4">
              {/* Nested Sub-Tabs header selection */}
              <div className="bg-slate-50/50 p-2.5 border-b border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
                <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 shadow-inner">
                  <button
                    onClick={() => setActiveStep2SubTab("step2_1")}
                    className={`px-4 py-1.5 rounded-md text-xs font-bold transition duration-200 flex items-center space-x-1.5 ${
                      activeStep2SubTab === "step2_1"
                        ? "bg-white text-blue-600 shadow-sm"
                        : "text-slate-500 hover:text-slate-800"
                    }`}
                  >
                    <span className="px-1.5 py-0.5 rounded bg-blue-100 text-blue-800 text-[10px] uppercase font-mono">Bản 2.1</span>
                    <span>Bổ sung Thuế suất</span>
                  </button>
                  <button
                    onClick={() => setActiveStep2SubTab("step2_2")}
                    className={`px-4 py-1.5 rounded-md text-xs font-bold transition duration-200 flex items-center space-x-1.5 ${
                      activeStep2SubTab === "step2_2"
                        ? "bg-white text-indigo-600 shadow-sm"
                        : "text-slate-500 hover:text-slate-800"
                    }`}
                  >
                    <span className="px-1.5 py-0.5 rounded bg-indigo-100 text-indigo-800 text-[10px] uppercase font-mono">Bản 2.2</span>
                    <span>Bổ sung Giá trước thuế & Tiền thuế</span>
                  </button>
                </div>
                <div className="flex items-center space-x-2 font-sans text-[11px]">
                  <span className="font-bold text-slate-500">Đánh dấu Excel Alphabet:</span>
                  <span className="bg-amber-100 text-amber-800 font-bold border border-amber-200 px-2 py-0.5 rounded">Đã bật (A, B, C...)</span>
                </div>
              </div>

              {activeStep2SubTab === "step2_1" ? (
                <div className="space-y-2">
                  <div className="px-5 text-xs text-slate-500">
                    <span className="font-bold text-slate-800">Bước 2.1:</span> Chỉ nạp thêm cột <strong className="text-blue-600 font-bold">Thuế bán hàng (bổ sung)</strong> tra cứu tương ứng từ Danh mục hàng hóa (A, B, C... để xác định vị trí).
                  </div>
                  <div className="overflow-x-auto max-h-[300px]">
                    <table className="w-full text-left text-[11px] text-slate-600 divide-y divide-slate-200 whitespace-nowrap">
                      <thead className="bg-[#f8fafc] sticky top-0 font-semibold text-slate-700 border-b z-10">
                        {/* Alphabetical row label */}
                        <tr className="bg-slate-100/90 text-slate-500 font-mono font-bold text-center text-[10px] border-b border-slate-200">
                          {Object.keys(invoices[0] || {}).map((col, idx) => (
                            <th key={`alpha-21-${idx}`} className="py-1 px-3 border-r border-slate-200 bg-slate-100 font-bold">
                              {getExcelColumnLabel(idx)}
                            </th>
                          ))}
                          <th className="py-1 px-3 text-center text-blue-800 bg-blue-100 border-r border-slate-200 font-bold font-mono">
                            {getExcelColumnLabel(Object.keys(invoices[0] || {}).length)}
                          </th>
                        </tr>
                        {/* Title Headers */}
                        <tr>
                          {Object.keys(invoices[0] || {}).map((col, idx) => (
                            <th key={`head-21-${idx}`} className="py-2.5 px-3 border-r border-slate-250 text-slate-700 font-sans">{col}</th>
                          ))}
                          <th className="py-2.5 px-3 text-center text-blue-800 bg-blue-50/40 border-r border-slate-250 font-sans">Thuế bán hàng (bổ sung)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y font-mono">
                        {step2Invoices.slice(0, 15).map((row, idx) => {
                          const cells = Object.keys(invoices[0] || {});
                          const taxRateVal = row["Thuế bán hàng"];
                          const taxRatePercent = taxRateVal !== undefined ? `${Math.round(taxRateVal * 100)}%` : "8%";
                          return (
                            <tr key={`row-21-${idx}`} className="hover:bg-slate-50/70 transition">
                              {cells.map((col, cIdx) => {
                                const val = row[col];
                                let formattedVal = "";
                                if (typeof val === "number") {
                                  formattedVal = val.toLocaleString("vi-VN");
                                } else {
                                  const parsedVal = parseFloat(String(val || "").replace(/,/g, ""));
                                  if (!isNaN(parsedVal) && String(val).trim().length > 3 && (String(val).includes(".") || String(val).includes(","))) {
                                    formattedVal = parsedVal.toLocaleString("vi-VN");
                                  } else {
                                    formattedVal = String(val !== undefined && val !== null ? val : "");
                                  }
                                }
                                return (
                                  <td key={`cell-21-${idx}-${cIdx}`} className="py-2 px-3 border-r border-slate-150 text-slate-700 font-sans">
                                    {formattedVal}
                                  </td>
                                );
                              })}
                              <td className="py-2 px-3 text-center font-bold text-blue-600 bg-blue-50/10 border-r border-slate-150">
                                {taxRatePercent}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : (
                <div className="space-y-2">
                  <div className="px-5 text-xs text-slate-500">
                    <span className="font-bold text-slate-800">Bước 2.2:</span> Bổ sung cột <strong className="text-teal-600 font-bold">Giá bán trước thuế (bổ sung)</strong> và cột <strong className="text-purple-600 font-bold">Tiền thuế (bổ sung)</strong>. Công thức: <strong className="text-teal-700">Giá bán trước thuế = Giá bán / (1 + Thuế bán hàng)</strong> và <strong className="text-purple-700">Tiền thuế = Giá bán trước thuế * Thuế bán hàng</strong> để kiểm tra tương quan.
                  </div>
                  <div className="overflow-x-auto max-h-[300px]">
                    <table className="w-full text-left text-[11px] text-slate-600 divide-y divide-slate-200 whitespace-nowrap">
                      <thead className="bg-[#f8fafc] sticky top-0 font-semibold text-slate-700 border-b z-10">
                        {/* Alphabetical row label */}
                        <tr className="bg-slate-100/90 text-slate-500 font-mono font-bold text-center text-[10px] border-b border-slate-200">
                          {Object.keys(invoices[0] || {}).map((col, idx) => (
                            <th key={`alpha-22-${idx}`} className="py-1 px-3 border-r border-slate-200 bg-slate-100 font-bold">
                              {getExcelColumnLabel(idx)}
                            </th>
                          ))}
                          <th className="py-1 px-3 text-center text-blue-800 bg-blue-100 border-r border-slate-200 font-bold font-mono">
                            {getExcelColumnLabel(Object.keys(invoices[0] || {}).length)}
                          </th>
                          <th className="py-1 px-3 text-right text-teal-800 bg-teal-100 border-r border-slate-200 font-bold font-mono">
                            {getExcelColumnLabel(Object.keys(invoices[0] || {}).length + 1)}
                          </th>
                          <th className="py-1 px-3 text-right text-purple-800 bg-purple-100 font-bold font-mono">
                            {getExcelColumnLabel(Object.keys(invoices[0] || {}).length + 2)}
                          </th>
                        </tr>
                        {/* Title Headers */}
                        <tr>
                          {Object.keys(invoices[0] || {}).map((col, idx) => (
                            <th key={`head-22-${idx}`} className="py-2.5 px-3 border-r border-slate-250 text-slate-700 font-sans">{col}</th>
                          ))}
                          <th className="py-2.5 px-3 text-center text-blue-800 bg-blue-50/40 border-r border-slate-250 font-sans">Thuế bán hàng (bổ sung)</th>
                          <th className="py-2.5 px-3 text-right text-teal-800 bg-teal-50/40 border-r border-slate-250 font-sans">Giá bán trước thuế (bổ sung)</th>
                          <th className="py-2.5 px-3 text-right text-purple-800 bg-purple-50/40 font-sans">Tiền thuế (bổ sung)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y font-mono">
                        {step2Invoices.slice(0, 15).map((row, idx) => {
                          const cells = Object.keys(invoices[0] || {});
                          const taxRateVal = row["Thuế bán hàng"];
                          const taxRatePercent = taxRateVal !== undefined ? `${Math.round(taxRateVal * 100)}%` : "8%";
                          const priceBeforeTax = row["Giá bán trước thuế"] || 0;
                          const taxAmount = row["Tiền thuế"] || 0;
                          return (
                            <tr key={`row-22-${idx}`} className="hover:bg-slate-50/70 transition">
                              {cells.map((col, cIdx) => {
                                const val = row[col];
                                let formattedVal = "";
                                if (typeof val === "number") {
                                  formattedVal = val.toLocaleString("vi-VN");
                                } else {
                                  const parsedVal = parseFloat(String(val || "").replace(/,/g, ""));
                                  if (!isNaN(parsedVal) && String(val).trim().length > 3 && (String(val).includes(".") || String(val).includes(","))) {
                                    formattedVal = parsedVal.toLocaleString("vi-VN");
                                  } else {
                                    formattedVal = String(val !== undefined && val !== null ? val : "");
                                  }
                                }
                                return (
                                  <td key={`cell-22-${idx}-${cIdx}`} className="py-2 px-3 border-r border-slate-150 text-slate-700 font-sans">
                                    {formattedVal}
                                  </td>
                                );
                              })}
                              <td className="py-2 px-3 text-center font-bold text-blue-600 bg-blue-50/10 border-r border-slate-150">
                                {taxRatePercent}
                              </td>
                              <td className="py-2 px-3 text-right font-bold text-teal-700 bg-teal-50/10 border-r border-slate-150">
                                {priceBeforeTax.toLocaleString("vi-VN")} đ
                              </td>
                              <td className="py-2 px-3 text-right font-bold text-purple-700 bg-purple-50/10">
                                {taxAmount.toLocaleString("vi-VN")} đ
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              <div className="p-4 bg-slate-50 border-t flex flex-col sm:flex-row items-center justify-between gap-4">
                <span className="text-[10px] text-slate-500 font-sans">
                  Hiển thị 15 dòng đầu trong tổng số {step2Invoices.length} dòng dữ liệu. 100% cột dữ liệu gốc của file được giữ nguyên, các cột bổ sung có dán nhãn chữ cái Alphabet Excel tương đối phía trên giúp kiểm chứng công thức chính xác.
                </span>
                <button
                  onClick={runStep3Processing}
                  id="goto-step3-from-step2"
                  className="flex items-center space-x-1 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-bold shadow-md cursor-pointer whitespace-nowrap"
                >
                  <span>Chuyển sang Bước 3</span>
                  <ArrowRight className="h-4.5 w-4.5" />
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* STEP 3 CONTAINER: ÁNH XẠ GỘP NHÓM & XUẤT BẢN FILE ĐẤU NỐI GTGT */}
      {activeStepTab === "step3" && (
        <div className="space-y-6">
          <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6 animate-fade-in">
            
            {/* Left box: Dynamic auto-detected column matching setups */}
            <div className="lg:col-span-8 flex flex-col justify-between space-y-4">
              <div>
                <div className="flex items-center space-x-2 pb-2 border-b border-slate-100 mb-4">
                  <FileCheck className="h-4.5 w-4.5 text-blue-600" />
                  <span className="text-xs font-bold text-slate-700 uppercase tracking-widest">Đầu Bố Trí Khớp Cột Excel</span>
                </div>
                
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mã sản phẩm (Catalog)</span>
                    <select
                      value={colMapping.productCodeCol}
                      onChange={(e) => setColMapping({ ...colMapping, productCodeCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableCatalogCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  
                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mức thuế (Catalog)</span>
                    <select
                      value={colMapping.productTaxCol}
                      onChange={(e) => setColMapping({ ...colMapping, productTaxCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableCatalogCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mã hóa đơn (Dòng HD)</span>
                    <select
                      value={colMapping.invoiceIdCol}
                      onChange={(e) => setColMapping({ ...colMapping, invoiceIdCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Giá bán (Dòng HD)</span>
                    <select
                      value={colMapping.invoicePriceCol}
                      onChange={(e) => setColMapping({ ...colMapping, invoicePriceCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Số lượng (Dòng HD)</span>
                    <select
                      value={colMapping.invoiceQtyCol}
                      onChange={(e) => setColMapping({ ...colMapping, invoiceQtyCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-150 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  <div>
                    <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Giảm giá % (Dòng HD)</span>
                    <select
                      value={colMapping.invoiceDiscountPercentCol}
                      onChange={(e) => setColMapping({ ...colMapping, invoiceDiscountPercentCol: e.target.value })}
                      className="w-full text-[11px] bg-slate-50 border border-slate-150 rounded p-1.5 text-slate-700 font-mono mt-0.5 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                    >
                      {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              <div className="text-[10px] text-blue-800 bg-blue-50/70 p-3 rounded-lg border border-blue-100 flex items-start space-x-1.5 leading-relaxed font-sans mt-2">
                <Info className="h-4 w-4 shrink-0 text-blue-500 mt-0.5" />
                <span>Hệ thống tự động phát hiện chính xác tên cột (Auto heuristics) của file hóa đơn và danh mục sản phẩm của bạn. Bạn có thể thay đổi các tùy chọn trên để đồng bộ hóa nếu tệp nhập của bạn có cấu trúc cột khác biệt.</span>
              </div>
            </div>

            {/* Right box: Execution actions */}
            <div className="lg:col-span-4 bg-slate-50/50 p-4 rounded-xl border border-slate-150 flex flex-col justify-between space-y-4">
              <div className="space-y-2">
                <div className="flex items-center space-x-2 pb-1 border-b border-slate-200">
                  <Settings2 className="h-4.5 w-4.5 text-slate-600" />
                  <span className="text-xs font-bold text-slate-700 uppercase tracking-widest">Tiến Trình Chuyển Đổi</span>
                </div>
                <div className="text-[11px] text-slate-600 space-y-2 leading-relaxed font-sans">
                  <p>
                    • Giá trị <strong>Tổng tiền</strong> của dòng Chiết khấu thương mại (CKTM) mặc định luôn luôn mang <strong>giá trị dương (+)</strong>.
                  </p>
                  <p>
                    • Công thức Thành tiền dòng hóa đơn được tinh chỉnh tự động theo quy chuẩn: <code className="bg-slate-200/80 px-1 py-0.5 rounded text-indigo-700 font-mono text-[10px]">Thành tiền = (Giá trước thuế * Số lượng) - Tiền chiết khấu</code>.
                  </p>
                </div>
              </div>

              <button
                onClick={runStep3Processing}
                id="run-step3-mapping-btn"
                className="w-full flex items-center justify-center space-x-2 py-3.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-extrabold text-xs uppercase tracking-wider transition shadow-md hover:shadow-lg active:scale-[0.98] cursor-pointer"
              >
                <span>Thực Hiện Chuyển Đổi & Tạo Import</span>
                <ArrowRight className="h-4 w-4 animate-pulse" />
              </button>
            </div>

          </div>

          {/* Step 3 XML output Preview visual table */}
          {step3Processed && outputData.length > 0 && (
            <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
              <div className="border-b border-slate-200 bg-slate-50/50 px-5 py-3.5 flex items-center justify-between">
                <div>
                  <span className="text-xs font-bold text-slate-700">Preview Layout File Import GTGT (Mau_import_hoa_don_GTGT.xlsx - {outputData.length} dòng)</span>
                  <p className="text-[10px] text-slate-400 mt-0.5">Bảng hiển thị đầy đủ 24 cột chuẩn của tệp ánh xạ đấu nối cùng ký tự Alphabet nhận dạng từ A-X.</p>
                </div>
                
                <button
                  onClick={handleExport}
                  id="xlsx-download-final-step3"
                  className="flex items-center space-x-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-extrabold cursor-pointer shadow-md transition duration-200 active:scale-95"
                >
                  <Download className="h-4 w-4" />
                  <span>Tải Excel Import GTGT (.xlsx)</span>
                </button>
              </div>

              <div className="overflow-x-auto max-h-[420px]">
                <table className="min-w-[2400px] text-left text-[11px] text-slate-600 divide-y divide-slate-200">
                  <thead className="bg-[#f8fafc] sticky top-0 font-semibold text-slate-700 border-b border-slate-200 z-10">
                    <tr className="border-b border-slate-100 bg-slate-50">
                      {[
                        { name: "Mã nhóm hóa đơn *", letter: "A", key: "Mã nhóm hóa đơn" },
                        { name: "Ngày hóa đơn", letter: "B", key: "Ngày hóa đơn" },
                        { name: "Mã khách hàng", letter: "C", key: "Mã khách hàng" },
                        { name: "Mã số thuế", letter: "D", key: "Mã số thuế" },
                        { name: "Tên khách hàng", letter: "E", key: "Tên khách hàng" },
                        { name: "Họ tên người mua", letter: "F", key: "Họ tên người mua" },
                        { name: "Số CCCD", letter: "G", key: "Số CCCD" },
                        { name: "Địa chỉ", letter: "H", key: "Địa chỉ" },
                        { name: "Điện thoại", letter: "I", key: "Điện thoại" },
                        { name: "Email nhận HĐ", letter: "J", key: "Email nhận HĐ" },
                        { name: "Hình thức thanh toán", letter: "K", key: "Hình thức thanh toán" },
                        { name: "Tên ngân hàng", letter: "L", key: "Tên ngân hàng" },
                        { name: "Mã hàng", letter: "M", key: "Mã hàng" },
                        { name: "Tên hàng hóa, dịch vụ *", letter: "N", key: "Tên hàng hóa, dịch vụ *" },
                        { name: "Tính chất hàng hóa, dịch vụ", letter: "O", key: "Tính chất hàng hóa, dịch vụ" },
                        { name: "Đơn vị tính", letter: "P", key: "Đơn vị tính" },
                        { name: "Số lượng", letter: "Q", key: "Số lượng" },
                        { name: "Đơn giá", letter: "R", key: "Đơn giá" },
                        { name: "Chiết khấu (%)", letter: "S", key: "Chiết khấu (%)" },
                        { name: "Tiền chiết khấu", letter: "T", key: "Tiền chiết khấu" },
                        { name: "Thành tiền", letter: "U", key: "Thành tiền" },
                        { name: "% VAT", letter: "V", key: "% VAT" },
                        { name: "Tiền VAT", letter: "W", key: "Tiền VAT" },
                        { name: "Tổng tiền *", letter: "X", key: "Tổng tiền *" }
                      ].map((h, i) => (
                        <th key={i} className="py-2 px-3 border-r border-slate-200/60 last:border-0 truncate" style={{ minWidth: i === 4 || i === 13 ? '220px' : i === 0 || i === 7 ? '180px' : '110px' }}>
                          <span className="inline-block px-1.5 py-0.5 bg-blue-50 border border-blue-200 rounded text-blue-700 font-mono font-extrabold text-[10px] mb-1">
                            {h.letter}
                          </span>
                          <span className="block text-slate-800 font-bold text-[10px] tracking-tight">{h.name}</span>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y font-mono">
                    {outputData.map((row, idx) => {
                      const isCKTM = row["Mã hàng"].startsWith("CKTM");
                      
                      const rowKeys: (keyof OutputRow)[] = [
                        "Mã nhóm hóa đơn",
                        "Ngày hóa đơn",
                        "Mã khách hàng",
                        "Mã số thuế",
                        "Tên khách hàng",
                        "Họ tên người mua",
                        "Số CCCD",
                        "Địa chỉ",
                        "Điện thoại",
                        "Email nhận HĐ",
                        "Hình thức thanh toán",
                        "Tên ngân hàng",
                        "Mã hàng",
                        "Tên hàng hóa, dịch vụ *",
                        "Tính chất hàng hóa, dịch vụ",
                        "Đơn vị tính",
                        "Số lượng",
                        "Đơn giá",
                        "Chiết khấu (%)",
                        "Tiền chiết khấu",
                        "Thành tiền",
                        "% VAT",
                        "Tiền VAT",
                        "Tổng tiền *"
                      ];

                      return (
                        <tr 
                          key={idx} 
                          className={`hover:bg-slate-50/80 transition ${
                            isCKTM ? "bg-indigo-50/50 text-indigo-900 border-l-4 border-l-indigo-500" : ""
                          } ${
                            row["Ngày hóa đơn"] ? "border-t-2 border-slate-250 font-bold bg-slate-50/25" : ""
                          }`}
                        >
                          {rowKeys.map((k, colIdx) => {
                            const val = row[k];
                            let content: React.ReactNode = "-";
                            let cellClass = "py-1.5 px-3 text-slate-700 border-r border-slate-100/60 last:border-r-0 truncate";

                            if (val !== undefined && val !== null && val !== "") {
                              if (["Đơn giá", "Tiền chiết khấu", "Thành tiền", "Tiền VAT", "Tổng tiền *"].includes(k)) {
                                const num = Number(val);
                                content = !isNaN(num) ? `${num.toLocaleString("vi-VN")} đ` : String(val);
                                cellClass += " text-right font-semibold text-slate-900";
                              } else if (k === "Mã nhóm hóa đơn") {
                                content = String(val);
                                cellClass += " font-extrabold text-slate-800";
                              } else if (k === "Mã hàng") {
                                content = String(val);
                                cellClass += isCKTM ? " text-indigo-650 font-bold bg-indigo-50/20" : " text-blue-650 font-bold";
                              } else if (k === "% VAT") {
                                content = String(val);
                                cellClass += " text-center font-extrabold text-amber-800 bg-amber-50/10";
                              } else if (k === "Số lượng") {
                                content = String(val);
                                cellClass += " text-center font-extrabold text-slate-900";
                              } else {
                                content = String(val);
                                if (k === "Tên khách hàng" || k === "Tên hàng hóa, dịch vụ *") {
                                  cellClass += " font-sans text-slate-800";
                                } else {
                                  cellClass += " font-sans";
                                }
                              }
                            }

                            return (
                              <td key={colIdx} className={cellClass} title={String(val || "")}>
                                {content}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Unified processing details log trails */}
      <div className="bg-slate-900 rounded-xl border border-slate-950 p-4 shadow-inner text-[11px] font-mono text-slate-350 space-y-2">
        <div className="flex items-center justify-between border-b border-slate-800 pb-2 text-slate-400">
          <span className="flex items-center space-x-1.5 font-bold uppercase tracking-wider text-[10px]">
            <RefreshCcw className="h-3.5 w-3.5 animate-spin-slow text-blue-500 shrink-0" />
            <span>Hệ Thống Phân Tích & Ghi Nhật Ký (Audit Trail Logs)</span>
          </span>
          <button 
            onClick={() => setProcessingLogs([])}
            className="hover:text-white transition text-[10px] underline cursor-pointer"
          >
            Xóa lịch sử log
          </button>
        </div>
        
        <div className="max-h-24 overflow-y-auto space-y-1">
          {processingLogs.length === 0 ? (
            <div className="text-slate-500 italic py-2">Hệ thống đang hoạt động ổn định. Đang chờ tác vụ nạp tệp ...</div>
          ) : (
            processingLogs.map((log, i) => (
              <div key={i} className="py-0.5 border-b border-slate-800/20 last:border-none">
                {log}
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}
