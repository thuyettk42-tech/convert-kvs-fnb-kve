/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import { ProductCatalogRow, InvoiceDetailRow, ThanhTienFormula } from "./types";
import Header from "./components/Header";
import ExcelProcessor from "./components/ExcelProcessor";

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
