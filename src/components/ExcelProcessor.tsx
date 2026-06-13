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
import ManualMockUploader from "./ManualMockUploader";

interface ExcelProcessorProps {
  products: ProductCatalogRow[];
  invoices: InvoiceDetailRow[];
  onDataChange: (products: ProductCatalogRow[], invoices: InvoiceDetailRow[]) => void;
  formula: ThanhTienFormula;
  setFormula: (f: ThanhTienFormula) => void;
  isNegativeCKTM: boolean;
  setIsNegativeCKTM: (val: boolean) => void;
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

  // Automatically adjust column matching list once files load
  useEffect(() => {
    if (products.length > 0) {
      const firstRow = products[0];
      const cols = Object.keys(firstRow);
      setAvailableCatalogCols(cols);
      
      // Auto mapping heuristics
      const code = cols.find(c => ["mã hàng", "ma hang", "ma_hang", "mã sản phẩm", "mã sp", "ma sp"].slice().includes(c.toLowerCase().trim())) || cols[0];
      const tax = cols.find(c => ["thuế bán hàng", "thue ban hang", "thuế", "thue", "vat", "thuế suất", "thue suat"].slice().includes(c.toLowerCase().trim())) || cols[1] || cols[0];
      
      setColMapping(prev => ({
        ...prev,
        productCodeCol: code || prev.productCodeCol,
        productTaxCol: tax || prev.productTaxCol,
      }));
    }
  }, [products]);

  useEffect(() => {
    if (invoices.length > 0) {
      const firstRow = invoices[0];
      const cols = Object.keys(firstRow);
      setAvailableInvoiceCols(cols);

      // Auto mapping heuristics
      const invoiceId = cols.find(c => ["mã hóa đơn", "ma hoa don", "mã nhóm hóa đơn", "số hóa đơn", "so hoa don", "ma_hd", "so_hd"].slice().includes(c.toLowerCase().trim())) || cols[0];
      const code = cols.find(c => ["mã hàng", "ma hang", "ma_hang", "mã sản phẩm", "mã sp", "ma sp"].slice().includes(c.toLowerCase().trim())) || cols[1] || cols[0];
      const name = cols.find(c => ["tên hàng", "ten hang", "tên sản phẩm", "tên hàng hóa", "tên hàng hóa, dịch vụ"].slice().includes(c.toLowerCase().trim())) || cols[2] || cols[0];
      const price = cols.find(c => ["giá bán", "gia ban", "đơn giá", "don gia", "giá", "gia", "đơn giá bán"].slice().includes(c.toLowerCase().trim())) || cols[3] || cols[0];
      const qty = cols.find(c => ["số lượng", "so luong", "sl", "qty"].slice().includes(c.toLowerCase().trim())) || cols[4] || cols[0];
      const lineDisc = cols.find(c => ["giảm giá %", "giam gia %", "chiết khấu %", "chiet khau %", "chiết khấu"].slice().includes(c.toLowerCase().trim())) || cols[5] || cols[0];
      const hdBdisc = cols.find(c => ["giảm giá hóa đơn", "giam gia hoa don", "chiết khấu hóa đơn", "giam_gia_hd"].slice().includes(c.toLowerCase().trim())) || cols[6] || cols[0];

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
    }
  }, [invoices]);

  // Read upload excel or CSV
  const handleFileUpload = (type: "products" | "invoices", file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) return;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json<any>(sheet);

      if (jsonData.length === 0) {
        alert("File này rỗng. Vui lòng kiểm tra lại thiết lập!");
        return;
      }

      if (type === "products") {
        onDataChange(jsonData, invoices);
        addLog(`Đã tải danh mục sản phẩm từ: ${file.name} (${jsonData.length} dòng)`);
      } else {
        onDataChange(products, jsonData);
        addLog(`Đã tải danh sách chi tiết hóa đơn từ: ${file.name} (${jsonData.length} dòng)`);
      }
    };
    reader.readAsBinaryString(file);
  };

  const addLog = (msg: string) => {
    const time = new Date().toLocaleTimeString("vi-VN");
    setProcessingLogs(p => [`[${time}] ${msg}`, ...p]);
  };

  // Logic calculation for STEP 1, 2, and 3
  const processTransformation = () => {
    if (products.length === 0 || invoices.length === 0) {
      alert("Vui lòng tải hoặc nạp cả hai file dữ liệu trước!");
      return;
    }

    addLog("Bắt đầu xử lý ánh xạ kế toán bán hàng...");
    const logs: string[] = [];
    
    // ----------------------------------------------------
    // BƯỚC 1: LÀM GIÀU DỮ LIỆU & TRUY XUẤT THUẾ SUẤT
    // ----------------------------------------------------
    // Build tax rate mapping dictionary from DanhSachSanPham
    const taxDict: Record<string, number> = {};
    products.forEach((p, idx) => {
      const rawCode = String(p[colMapping.productCodeCol] || "").trim();
      const rawTax = p[colMapping.productTaxCol];
      
      let parsedTax = 0.10; // Default fallback to 10% VAT
      if (rawTax !== undefined && rawTax !== null) {
        if (typeof rawTax === "number") {
          // If already decimal like 0.08 or 0.10
          parsedTax = rawTax;
        } else {
          // If string like "8%" or "10%"
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

    logs.push(`Đã tải bảng tra cứu gồm ${Object.keys(taxDict).length} sản phẩm chứa dữ liệu thuế.`);

    // ----------------------------------------------------
    // BƯỚC 2: QUY TRÌNH MAPPING SANG TEMPLATE & GROUPING
    // ----------------------------------------------------
    // Group invoices by 'Mã hóa đơn'
    const groupedInvoices: Record<string, any[]> = {};
    invoices.forEach((row, i) => {
      const invId = String(row[colMapping.invoiceIdCol] || "").trim();
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
      const rawInvDiscount = groupItems[0][colMapping.invoiceDiscountAmountCol];
      const invoiceDiscount = Math.max(0, parseFloat(String(rawInvDiscount || 0).replace(/,/g, "")) || 0);

      // We'll track total item amount (inclusive of taxes) to divide trade discounts proportionally
      let tempGroupOutputRows: OutputRow[] = [];
      let totalItemValueForAllocation = 0;

      groupItems.forEach((item) => {
        const itemCode = String(item[colMapping.invoiceProductCodeCol] || "").trim();
        const itemName = String(item[colMapping.invoiceProductNameCol] || "").trim();
        
        // Match tax from catalog using our parsed helper
        const taxRate = taxDict[itemCode] !== undefined ? taxDict[itemCode] : 0.10;

        const qty = parseFloat(String(item[colMapping.invoiceQtyCol] || 0).replace(/,/g, "")) || 0;
        const priceWithTax = parseFloat(String(item[colMapping.invoicePriceCol] || 0).replace(/,/g, "")) || 0;
        const discountPercent = parseFloat(String(item[colMapping.invoiceDiscountPercentCol] || 0).replace(/,/g, "")) || 0;

        // Formula 1: Giá trước thuế = [Giá bán] / (1 + [Thuế bán hàng]) -> round top 2 decimal places
        const giaTruocThue = Math.round((priceWithTax / (1 + taxRate)) * 100) / 100;

        // Apply selected 'Thành tiền' calculation formulas
        let thanhTien = 0;
        if (formula === "standard_accounting") {
          // (Giá trước thuế * Số lượng) - Tiền chiết khấu
          // Tiền chiết khấu = Đơn giá trước thuế * Số lượng * (Giảm giá % / 100)
          const discAmount = giaTruocThue * qty * (discountPercent / 100);
          thanhTien = Math.round(((giaTruocThue * qty) - discAmount) * 100) / 100;
        } else if (formula === "standard_discount_rate") {
          // (Giá trước thuế * Số lượng) * (1 - Giảm giá% / 100)
          thanhTien = Math.round((giaTruocThue * qty * (1 - discountPercent / 100)) * 100) / 100;
        } else if (formula === "multiplier") {
          // Đơn giá * Giảm giá (from prompt "Tính bằng Đơn giá nhân với Chiết khấu")
          thanhTien = Math.round((giaTruocThue * discountPercent) * 100) / 100;
        } else {
          // No discount on item
          thanhTien = Math.round((giaTruocThue * qty) * 100) / 100;
        }

        const totalRowAmount = Math.round((thanhTien * (1 + taxRate)) * 100) / 100;
        totalItemValueForAllocation += totalRowAmount;

        const mappedTaxStr = `${Math.round(taxRate * 100)}%`;

        const outRow: OutputRow = {
          "Ngày hóa đơn": isFirstRowOfGroup ? currentDateStr : "",
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
          "Số lượng": qty,
          "Đơn giá": giaTruocThue,
          "Chiết khấu (%)": discountPercent,
          "Tiền chiết khấu": "",
          "Thành tiền": thanhTien,
          "% VAT": mappedTaxStr,
          "Tiền VAT": "",
          "Tổng tiền *": totalRowAmount,
          "Mã nhóm hóa đơn": invId,
        };

        tempGroupOutputRows.push(outRow);
        isFirstRowOfGroup = false;
      });

      // ----------------------------------------------------
      // BƯỚC 3: PHÂN BỔ CHIẾT KHẤU THƯƠNG MẠI (CKTM)
      // ----------------------------------------------------
      // Check invoice-level discount (Giảm giá hóa đơn)
      if (invoiceDiscount > 0 && totalItemValueForAllocation > 0) {
        // Group the group's current rows by unique tax rates
        const vatGroups: Record<number, OutputRow[]> = {};
        tempGroupOutputRows.forEach((row) => {
          // retrieve numeric key for tax from product lookup
          const taxRate = taxDict[row["Mã hàng"]] !== undefined ? taxDict[row["Mã hàng"]] : 0.10;
          if (vatGroups[taxRate] === undefined) {
            vatGroups[taxRate] = [];
          }
          vatGroups[taxRate].push(row);
        });

        // Insert separate CKTM rows for each unique tax rate group
        Object.entries(vatGroups).forEach(([taxRateStr, rowsOfTax]) => {
          const rate = parseFloat(taxRateStr);
          const totalOfTaxRateGroup = rowsOfTax.reduce((sum, r) => sum + r["Tổng tiền *"], 0);

          if (totalOfTaxRateGroup <= 0) return;

          // Proportional allocation: [Giảm giá hóa đơn] * [Tổng tiền hàng cùng thuế suất] / [Tổng giá trị hóa đơn]
          const allocatedTotalWithTax = invoiceDiscount * (totalOfTaxRateGroup / totalItemValueForAllocation);
          const allocatedTotalRounded = Math.round(allocatedTotalWithTax);

          // Apply sign depending on if negative CKTM is chosen
          const totalVal = isNegativeCKTM ? -allocatedTotalRounded : allocatedTotalRounded;
          
          // Back-calculate: Thành tiền = [Tổng tiền *] / (1 + [% VAT])
          const allocatedThanhTien = Math.round((totalVal / (1 + rate)) * 100) / 100;
          const allocatedDonGia = allocatedThanhTien;

          const vatPctStr = `${Math.round(rate * 100)}%`;

          const cktmRow: OutputRow = {
            "Ngày hóa đơn": "",
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
            "Số lượng": "",
            "Đơn giá": allocatedDonGia,
            "Chiết khấu (%)": 0,
            "Tiền chiết khấu": "",
            "Thành tiền": allocatedThanhTien,
            "% VAT": vatPctStr,
            "Tiền VAT": "",
            "Tổng tiền *": totalVal,
            "Mã nhóm hóa đơn": invId,
          };

          tempGroupOutputRows.push(cktmRow);
        });
      }

      // Add temporary rows to final records
      finalRows.push(...tempGroupOutputRows);
    });

    logs.push(`Hoàn thành ánh xạ! Đã xuất ${finalRows.length} dòng dữ liệu sang cấu trúc template hóa đơn.`);
    setOutputData(finalRows);
    setProcessingLogs(p => [...logs.map(l => `[${new Date().toLocaleTimeString("vi-VN")}] ${l}`), ...p]);
    setActiveTab("preview");
  };

  // Export processed data to standard Excel file
  const handleExport = () => {
    if (outputData.length === 0) {
      alert("Không có dữ liệu đầu ra để xuất bản! Vui lòng bấm 'Thực hiện chuyển đổi' trước.");
      return;
    }

    // Prepare sheet columns ensuring "Mã nhóm hóa đơn" is the first column matching user's template
    const cleanedRows = outputData.map((row) => ({
      "Mã nhóm hóa đơn": row["Mã nhóm hóa đơn"],
      "Ngày hóa đơn": row["Ngày hóa đơn"],
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

    // Auto fit column widths for premium visual experience
    const colWidths = [
      { wch: 18 }, // Mã nhóm hóa đơn
      { wch: 15 }, // Ngày hóa đơn
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

  const handleLoadMock = (pData: ProductCatalogRow[], iData: InvoiceDetailRow[]) => {
    onDataChange(pData, iData);
    addLog(`Đã tả liên kết mẫu: ${pData.length} sản phẩm & ${iData.length} hóa đơn.`);
  };

  return (
    <div className="space-y-6">
      {/* File Upload / Mock Trigger block */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Upload Column 1: Product catalog */}
        <div className="bg-white rounded-lg p-5 border border-slate-200 shadow-sm flex flex-col justify-between h-56 group overflow-hidden">
          <div>
            <div className="flex items-center space-x-2 text-blue-600 mb-3">
              <FileSpreadsheet className="h-5 w-5" />
              <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest">1. Danh Mục Sản Phẩm</h4>
            </div>
            <p className="text-xs text-slate-500 leading-relaxed">
              Tải tập tin chứa danh mục mã hàng hóa kế toán, cột mã hàng và mức thuế suất tương ứng bán ra.
            </p>
          </div>
          
          <div className="mt-4">
            <label className="flex flex-col items-center justify-center border border-dashed border-slate-300 group-hover:border-blue-400 rounded-md py-3.5 px-3 cursor-pointer bg-slate-50 hover:bg-blue-50/20 transition duration-200">
              <Upload className="h-5 w-5 text-slate-400 group-hover:text-blue-500 mb-1 transition duration-200" />
              <span className="text-[11px] text-slate-600 group-hover:text-blue-700 font-medium">Tải lên DanhSachSanPham</span>
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv" 
                className="hidden" 
                onChange={(e) => e.target.files?.[0] && handleFileUpload("products", e.target.files[0])}
              />
            </label>
            {products.length > 0 && (
              <div className="flex items-center space-x-1 mt-2 text-[11px] text-blue-600 font-medium">
                <Check className="h-3 w-3" />
                <span>Đã nạp {products.length} dòng sản phẩm (Indexed)</span>
              </div>
            )}
          </div>
        </div>

        {/* Upload Column 2: Invoice Details */}
        <div className="bg-white rounded-lg p-5 border border-slate-200 shadow-sm flex flex-col justify-between h-56 group overflow-hidden">
          <div>
            <div className="flex items-center space-x-2 text-blue-600 mb-3">
              <FileSpreadsheet className="h-5 w-5" />
              <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest">2. Dữ Liệu Hóa Đơn</h4>
            </div>
            <p className="text-xs text-slate-500 leading-relaxed">
              Tải tập tin chi tiết dòng bán hàng cần chuyển đổi định dạng. Dòng 1 phải là dòng tiêu đề cột.
            </p>
          </div>

          <div className="mt-4">
            <label className="flex flex-col items-center justify-center border border-dashed border-slate-300 group-hover:border-blue-400 rounded-md py-3.5 px-3 cursor-pointer bg-slate-50 hover:bg-blue-50/20 transition duration-200">
              <Upload className="h-5 w-5 text-slate-400 group-hover:text-blue-500 mb-1 transition duration-200" />
              <span className="text-[11px] text-slate-600 group-hover:text-blue-700 font-medium">Tải lên DanhSachChiTietHoaDon</span>
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv" 
                className="hidden" 
                onChange={(e) => e.target.files?.[0] && handleFileUpload("invoices", e.target.files[0])}
              />
            </label>
            {invoices.length > 0 && (
              <div className="flex items-center space-x-1 mt-2 text-[11px] text-teal-600 font-medium">
                <Check className="h-3 w-3" />
                <span>Đã nạp {invoices.length} dòng hóa đơn</span>
              </div>
            )}
          </div>
        </div>

        {/* Column 3: Processing Action console */}
        <div className="bg-white rounded-lg p-5 border border-slate-200 shadow-sm flex flex-col justify-between h-56">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">3. Automation Panel</span>
              {products.length > 0 && invoices.length > 0 && (
                <span className="px-2 py-0.5 bg-green-550/10 text-green-600 rounded text-[10px] font-bold border border-green-500/20">
                  Ready
                </span>
              )}
            </div>
            <h3 className="font-semibold text-slate-800 text-sm">Cấu Hình Hoạt Động</h3>
            <p className="text-xs text-slate-500 leading-relaxed mt-1">
              Hệ thống tự động liên kết mã hàng, xử lý thuế VAT và phân bổ giảm giá hóa đơn tương ứng.
            </p>
          </div>

          <button
            onClick={processTransformation}
            disabled={products.length === 0 || invoices.length === 0}
            id="run-data-mapping-btn"
            className={`w-full flex items-center justify-center space-x-2 py-3 rounded font-bold text-xs uppercase tracking-wider transition duration-200 cursor-pointer ${
              products.length > 0 && invoices.length > 0
                ? "bg-blue-600 hover:bg-blue-500 text-white shadow-sm active:scale-[0.98]"
                : "bg-slate-150 text-slate-400 border border-slate-200 cursor-not-allowed"
            }`}
          >
            <span>Thực Hiện Chuyển Đổi</span>
            <ArrowRight className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* Mock Data Injector */}
      <ManualMockUploader 
        onLoadMock={handleLoadMock} 
        isLoaded={products.length > 0 && invoices.length > 0} 
      />

      {/* Advanced formula parameters & QA */}
      {/* Advanced formula parameters & QA */}
      <div className="bg-white rounded-lg p-5 border border-slate-200 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6">
        <div className="lg:col-span-4 space-y-4">
          <div className="flex items-center space-x-2 pb-1 border-b border-slate-100">
            <Settings2 className="h-4 w-4 text-slate-500" />
            <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">Mapping Schema Options</span>
          </div>

          {/* Formula configuration */}
          <div className="space-y-2">
            <label className="block text-xs font-semibold text-slate-600">Định Nghĩa Thành Tiền dòng:</label>
            <select
              value={formula}
              onChange={(e) => setFormula(e.target.value as ThanhTienFormula)}
              className="w-full text-xs bg-slate-50 border border-slate-200 rounded p-2 text-slate-700 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-medium"
            >
              <option value="standard_accounting">Mẫu Chuẩn Kế Toán (Đơn giá * SL) - CK dòng</option>
              <option value="standard_discount_rate">Chuẩn Kế Toán %: Đơn giá * SL * (1 - Giảm %)</option>
              <option value="no_discount">Nhân thuần túy: Đơn giá * Số lượng (Không chiết khấu dòng)</option>
              <option value="multiplier">Đáp ứng cơ bản Prompt: Đơn giá * Giảm %</option>
            </select>
          </div>

          {/* Trade Discount behavior */}
          <div className="space-y-2 pt-2">
            <label className="block text-xs font-semibold text-slate-600">Dấu của Chiết khấu Thương mại (CKTM):</label>
            <div className="flex items-center space-x-3 bg-slate-50 p-2 rounded border border-slate-200">
              <label className="flex items-center space-x-1.5 cursor-pointer text-xs text-slate-600">
                <input 
                  type="radio" 
                  checked={isNegativeCKTM} 
                  onChange={() => setIsNegativeCKTM(true)}
                  className="accent-blue-600 h-3.5 w-3.5"
                />
                <span className="font-medium">Số âm (-)</span>
              </label>
              <label className="flex items-center space-x-1.5 cursor-pointer text-xs text-slate-600">
                <input 
                  type="radio" 
                  checked={!isNegativeCKTM} 
                  onChange={() => setIsNegativeCKTM(false)}
                  className="accent-blue-600 h-3.5 w-3.5"
                />
                <span className="font-medium">Số dương (+)</span>
              </label>
            </div>
          </div>
        </div>

        {/* Dynamic Column Matchings config in current spreadsheets */}
        <div className="lg:col-span-8 flex flex-col justify-between">
          <div className="flex items-center space-x-2 pb-1 border-b border-slate-100 mb-3">
            <FileCheck className="h-4 w-4 text-slate-500" />
            <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">Kiểm soát Khớp Cột Excel (Auto-detected)</span>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mã hàng (Catalog)</span>
              <select
                value={colMapping.productCodeCol}
                onChange={(e) => setColMapping({ ...colMapping, productCodeCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-600 font-mono mt-0.5"
              >
                {availableCatalogCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            
            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Thuế VAT (Catalog)</span>
              <select
                value={colMapping.productTaxCol}
                onChange={(e) => setColMapping({ ...colMapping, productTaxCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-600 font-mono mt-0.5"
              >
                {availableCatalogCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Mã hóa đơn (Dòng HD)</span>
              <select
                value={colMapping.invoiceIdCol}
                onChange={(e) => setColMapping({ ...colMapping, invoiceIdCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-650 font-mono mt-0.5"
              >
                {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Giá bán d.m (Dòng HD)</span>
              <select
                value={colMapping.invoicePriceCol}
                onChange={(e) => setColMapping({ ...colMapping, invoicePriceCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-650 font-mono mt-0.5"
              >
                {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Số lượng (Dòng HD)</span>
              <select
                value={colMapping.invoiceQtyCol}
                onChange={(e) => setColMapping({ ...colMapping, invoiceQtyCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-650 font-mono mt-0.5"
              >
                {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div>
              <span className="block text-[10px] text-slate-400 font-bold uppercase tracking-wider">Giảm giá % (Dòng HD)</span>
              <select
                value={colMapping.invoiceDiscountPercentCol}
                onChange={(e) => setColMapping({ ...colMapping, invoiceDiscountPercentCol: e.target.value })}
                className="w-full text-[11px] bg-slate-50 border border-slate-200 rounded p-1.5 text-slate-650 font-mono mt-0.5"
              >
                {availableInvoiceCols.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
          </div>
          
          <div className="mt-3 text-[11px] text-slate-550 bg-blue-50/40 p-2 rounded border border-blue-100 flex items-center space-x-1.5">
            <Info className="h-3.5 w-3.5 text-blue-500 shrink-0" />
            <span>Mẹo: Tên các cột của tập tin được cấu hình tự động ngay khi bạn tải tệp Excel lên hoặc bấm nạp dữ liệu mẫu.</span>
          </div>
        </div>
      </div>

      {/* Tabs list for logs and Output Table Preview */}
      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden shadow-sm">
        <div className="border-b border-slate-200 bg-slate-50/50 px-5 flex items-center justify-between">
          <div className="flex space-x-4">
            <button
              onClick={() => setActiveTab("upload")}
              className={`py-3 text-xs font-semibold border-b-2 px-1 cursor-pointer transition ${
                activeTab === "upload" 
                  ? "border-blue-600 text-blue-750" 
                  : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              Lịch Sử Log
            </button>
            <button
              onClick={() => setActiveTab("preview")}
              className={`py-3 text-xs font-semibold border-b-2 px-1 cursor-pointer transition ${
                activeTab === "preview" 
                  ? "border-blue-600 text-blue-750" 
                  : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              Bảng Preview Đầu Ra ({outputData.length} dòng)
            </button>
          </div>

          {outputData.length > 0 && (
            <button
              onClick={handleExport}
              id="download-xlsx-processed-btn"
              className="flex items-center space-x-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs font-semibold cursor-pointer shadow-sm transition active:scale-[0.98]"
            >
              <Download className="h-3.5 w-3.5" />
              <span>Tải Excel Import (.xlsx)</span>
            </button>
          )}
        </div>

        <div className="p-5 max-h-[360px] overflow-auto">
          {activeTab === "upload" && (
            <div className="space-y-1.5 font-mono text-[11px] text-slate-600">
              {processingLogs.length === 0 ? (
                <div className="text-slate-400 py-6 text-center">Chưa có hành động nào được ghi lại. Vui lòng bấm "Nạp dữ liệu mẫu" để bắt đầu thử nghiệm!</div>
              ) : (
                processingLogs.map((log, i) => (
                  <div key={i} className="py-0.5 border-b border-slate-100 last:border-none">
                    {log}
                  </div>
                ))
              )}
            </div>
          )}

          {activeTab === "preview" && (
            <div>
              {outputData.length === 0 ? (
                <div className="text-slate-400 py-10 text-center flex flex-col items-center justify-center space-y-2">
                  <AlertTriangle className="h-8 w-8 text-amber-500" />
                  <span className="text-slate-650 font-medium">Chưa có dữ liệu kết quả. Bấm nút "Thực Hiện Chuyển Đổi" phía trên để nạp kết quả!</span>
                </div>
              ) : (
                <div className="overflow-x-auto text-[11px] select-all leading-normal">
                  <table className="w-full text-left text-slate-600 divide-y divide-slate-200">
                    <thead className="bg-[#f8fafc] sticky top-0 font-semibold text-slate-700 border-b border-slate-200">
                      <tr>
                        <th className="py-2 px-2.5">Mã Nhóm</th>
                        <th className="py-2 px-1.5">Ngày HĐ</th>
                        <th className="py-2 px-2">Khách Hàng</th>
                        <th className="py-2 px-1.5">Hình Thức</th>
                        <th className="py-2 px-1.5">Mã Hàng</th>
                        <th className="py-2 px-2">Tên Hàng Hóa Dịch Vụ</th>
                        <th className="py-2 px-1.5">Tính Chất</th>
                        <th className="py-2 px-1.5 text-center">SL</th>
                        <th className="py-2 px-2 text-right">Đơn Giá trước thuế</th>
                        <th className="py-2 px-1 text-center">Giảm %</th>
                        <th className="py-2 px-2 text-right">Thành Tiền</th>
                        <th className="py-2 px-1.5 text-center">% VAT</th>
                        <th className="py-2 px-2 text-right">Tổng Tiền *</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {outputData.map((row, idx) => {
                        const isCKTM = row["Mã hàng"].startsWith("CKTM");
                        return (
                          <tr 
                            key={idx} 
                            className={`hover:bg-slate-55/63 transition ${
                              isCKTM ? "bg-indigo-50/40 text-indigo-900 border-l-2 border-l-indigo-500" : ""
                            } ${
                              row["Ngày hóa đơn"] ? "border-t border-slate-300 font-medium bg-slate-50/30" : ""
                            }`}
                          >
                            <td className="py-1.5 px-2.5 font-semibold text-slate-500">{row["Mã nhóm hóa đơn"]}</td>
                            <td className="py-1.5 px-1.5 text-slate-800">{row["Ngày hóa đơn"]}</td>
                            <td className="py-1.5 px-2 truncate max-w-[130px]" title={row["Tên khách hàng"]}>{row["Tên khách hàng"]}</td>
                            <td className="py-1.5 px-1.5">{row["Hình thức thanh toán"]}</td>
                            <td className={`py-1.5 px-1.5 font-bold ${isCKTM ? "text-indigo-650" : "text-blue-650"}`}>{row["Mã hàng"]}</td>
                            <td className="py-1.5 px-2 truncate max-w-[164px] text-slate-750 font-sans" title={row["Tên hàng hóa, dịch vụ *"]}>{row["Tên hàng hóa, dịch vụ *"]}</td>
                            <td className="py-1.5 px-1.5 opacity-70 font-sans">{row["Tính chất hàng hóa, dịch vụ"]}</td>
                            <td className="py-1.5 px-1.5 text-center font-bold">{row["Số lượng"]}</td>
                            <td className="py-1.5 px-2 text-right">{row["Đơn giá"] !== "" && typeof row["Đơn giá"] === "number" ? `${row["Đơn giá"].toLocaleString("vi-VN")} đ` : "-"}</td>
                            <td className="py-1.5 px-1 text-center">{row["Chiết khấu (%)"] ? `${row["Chiết khấu (%)"]}%` : "-"}</td>
                            <td className="py-1.5 px-2 text-right font-semibold text-slate-850">{row["Thành tiền"] !== "" && typeof row["Thành tiền"] === "number" ? `${row["Thành tiền"].toLocaleString("vi-VN")} đ` : "-"}</td>
                            <td className="py-1.5 px-1.5 text-center text-amber-800 font-semibold">{row["% VAT"]}</td>
                            <td className={`py-1.5 px-2 text-right font-bold ${isCKTM ? "text-indigo-800" : "text-slate-900"}`}>{row["Tổng tiền *"].toLocaleString("vi-VN")} đ</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
