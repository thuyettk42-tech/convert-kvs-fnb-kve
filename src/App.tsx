/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import { ProductCatalogRow, InvoiceDetailRow, ThanhTienFormula } from "./types";
import Header from "./components/Header";
import ExcelProcessor from "./components/ExcelProcessor";
import PythonScriptView from "./components/PythonScriptView";
import MathQADecrypter from "./components/MathQADecrypter";
import { Calendar, HelpCircle, AlertCircle, FileText, Check } from "lucide-react";

export default function App() {
  const [products, setProducts] = useState<ProductCatalogRow[]>([]);
  const [invoices, setInvoices] = useState<InvoiceDetailRow[]>([]);
  const [activeTab, setActiveTab] = useState<string>("processor");
  const [formula, setFormula] = useState<ThanhTienFormula>("standard_accounting");
  const [isNegativeCKTM, setIsNegativeCKTM] = useState<boolean>(false);

  const handleDataChange = (prodList: ProductCatalogRow[], invList: InvoiceDetailRow[]) => {
    setProducts(prodList);
    setInvoices(invList);
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] text-slate-800 font-sans antialiased flex flex-col">
      {/* Visual top Header */}
      <Header
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        productsCount={products.length}
        invoicesCount={invoices.length}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
        {/* Welcome information banner */}
        <div className="bg-[#1e293b] text-white border border-slate-700/80 rounded-lg p-5 shrink-0 shadow-sm relative overflow-hidden flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
          <div className="flex items-start space-x-3.5 z-10">
            <div className="p-2.5 bg-slate-800 text-blue-400 rounded-md shrink-0 border border-slate-700">
              <AlertCircle className="h-5.5 w-5.5" />
            </div>
            <div>
              <h3 className="font-semibold text-white text-sm tracking-tight">Cổng Kiểm Soát ETL & Đồng Bộ Kế Toán Thuế</h3>
              <p className="text-xs text-slate-300 mt-1 leading-relaxed max-w-3xl">
                Hệ thống hỗ trợ nạp tệp danh mục, tự khớp mã hàng hóa, tự động phân tách nhóm khách hàng trên dòng đầu tiên của hóa đơn, và tự phân bổ Chiết khấu Thương mại (CKTM) theo tỷ trọng hóa đơn dòng một cách tối ưu.
              </p>
            </div>
          </div>
          <div className="text-[10px] text-slate-300 font-mono bg-slate-800/80 px-3 py-1.5 rounded border border-slate-700 shrink-0 self-start sm:self-center">
            📅 UTC: 2026-06-12 22:38:31
          </div>
        </div>

        {/* Dynamic Display area of selected Tab */}
        <div className="transition-all duration-300">
          {activeTab === "processor" && (
            <ExcelProcessor
              products={products}
              invoices={invoices}
              onDataChange={handleDataChange}
              formula={formula}
              setFormula={setFormula}
              isNegativeCKTM={isNegativeCKTM}
              setIsNegativeCKTM={setIsNegativeCKTM}
            />
          )}

          {activeTab === "python" && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 min-h-[500px]">
              <div className="lg:col-span-8 flex flex-col">
                <PythonScriptView formula={formula} isNegativeCKTM={isNegativeCKTM} />
              </div>
              
              <div className="lg:col-span-4 bg-white rounded-xl p-5 border border-gray-150 shadow-sm flex flex-col justify-between">
                <div>
                  <h4 className="font-bold text-gray-800 text-xs uppercase tracking-wider mb-3">Tính Năng Của Code Python</h4>
                  <ul className="space-y-3.5 text-xs text-gray-600">
                    <li className="flex items-start space-x-2">
                      <Check className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
                      <span><strong>Xử lý hàng triệu dòng siêu tốc:</strong> Thư viện Pandas tối ưu hóa cấu trúc bộ nhớ dạng C-arrays, giúp xử lý vượt trội so với đọc lặp thủ công.</span>
                    </li>
                    <li className="flex items-start space-x-2">
                      <Check className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
                      <span><strong>Đồng bộ động với cấu hình giao diện:</strong> Tập mã Python liên tục cập nhật theo công thức Thành Tiền và định hướng Dấu Chiết khấu Thương mại mà bạn thay đổi ở Tab Đấu nối.</span>
                    </li>
                    <li className="flex items-start space-x-2">
                      <Check className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
                      <span><strong>Tự động đọc ghi đa định dạng:</strong> Chấp nhận cả tệp Excel (.xlsx) cũng như tệp dẹt CSV trong các nghiệp vụ lớn.</span>
                    </li>
                  </ul>
                </div>

                <div className="mt-6 bg-amber-50 rounded-lg p-3.5 border border-amber-100 text-[11px] text-amber-800 flex items-start space-x-2.5">
                  <HelpCircle className="h-4.5 w-4.5 text-amber-600 mt-0.5 shrink-0" />
                  <div>
                    Bạn có thể sao chép tệp lệnh trên về máy tính và chạy định kỳ hàng ngày dưới dạng cron-job hoặc tích hợp vào hệ thống ERP để ánh xạ hóa đơn tự động sang Mau_import_hoa_don_GTGT.
                  </div>
                </div>
              </div>
            </div>
          )}

          {activeTab === "qa" && (
            <div className="max-w-4xl mx-auto">
              <MathQADecrypter />
            </div>
          )}
        </div>
      </main>

      {/* Humble craft footer */}
      <footer className="bg-white border-t border-gray-150 py-5 mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between text-xs text-gray-400">
          <div>Ứng dụng ánh xạ kế toán bán hàng & phân bổ điện tử</div>
          <div className="mt-1 sm:mt-0 font-mono text-[10px]">Tối ưu hóa cấu trúc tệp • Microsoft Excel Compatible • Pandas v2.0 Py-Script</div>
        </div>
      </footer>
    </div>
  );
}
