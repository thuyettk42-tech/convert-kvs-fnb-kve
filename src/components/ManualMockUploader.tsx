/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from "react";
import { Database, FileSpreadsheet, Sparkles, CheckCircle } from "lucide-react";
import { ProductCatalogRow, InvoiceDetailRow } from "../types";

interface ManualMockUploaderProps {
  onLoadMock: (products: ProductCatalogRow[], invoices: InvoiceDetailRow[]) => void;
  isLoaded: boolean;
}

export default function ManualMockUploader({ onLoadMock, isLoaded }: ManualMockUploaderProps) {
  // Generate highly realistic mock data for Vietnamese sales accounting context
  const mockProducts: ProductCatalogRow[] = [
    { "Mã hàng": "SP001", "Thuế bán hàng": 0.08, "Tên gốc": "Máy lọc nước Kangaroo KG10A3" },
    { "Mã hàng": "SP002", "Thuế bán hàng": 0.10, "Tên gốc": "Ấm siêu tốc Philips HD9306" },
    { "Mã hàng": "SP003", "Thuế bán hàng": 0.08, "Tên gốc": "Nồi cơm điện Toshiba RC-18NTFV" },
    { "Mã hàng": "SP004", "Thuế bán hàng": 0.00, "Tên gốc": "Rau sạch Đà Lạt organic (Mặt hàng không chịu thuế)" },
    { "Mã hàng": "SP005", "Thuế bán hàng": 0.05, "Tên gốc": "Sữa đậu nành đóng chai (Thuế suất ưu đãi)" },
  ];

  const mockInvoices: InvoiceDetailRow[] = [
    // Invoice HD001 - single item, no discount
    { 
      "Mã hóa đơn": "HD001", 
      "Mã hàng": "SP001", 
      "Tên hàng": "Máy lọc nước Kangaroo KG10A3 - Phiên bản Luxury", 
      "Giá bán": 8640000, 
      "Số lượng": 1, 
      "Giảm giá %": 0, 
      "Giảm giá hóa đơn": 0 
    },
    // Invoice HD002 - multiple items with different tax rates (8% and 10%) and an invoice-level discount (Chiết khấu thương mại)
    { 
      "Mã hóa đơn": "HD002", 
      "Mã hàng": "SP001", 
      "Tên hàng": "Máy lọc nước Kangaroo KG10A3 - Màng RO Dupont US", 
      "Giá bán": 8640000, 
      "Số lượng": 2, 
      "Giảm giá %": 5, // 5% line discount
      "Giảm giá hóa đơn": 1500000 // 1.5M absolute invoice discount to allocate across tax rates!
    },
    { 
      "Mã hóa đơn": "HD002", 
      "Mã hàng": "SP002", 
      "Tên hàng": "Ấm siêu tốc Philips HD9306 - Inox 1.5L", 
      "Giá bán": 660000, 
      "Số lượng": 3, 
      "Giảm giá %": 10, // 10% line discount
      "Giảm giá hóa đơn": 1500000 // same invoice-level discount
    },
    { 
      "Mã hóa đơn": "HD002", 
      "Mã hàng": "SP003", 
      "Tên hàng": "Nồi cơm điện Toshiba RC-18NTFV - Điện tử", 
      "Giá bán": 2700000, 
      "Số lượng": 1, 
      "Giảm giá %": 0, 
      "Giảm giá hóa đơn": 1500000 // same invoice-level discount
    },
    // Invoice HD003 - zero tax item and 5% items with discount
    { 
      "Mã hóa đơn": "HD003", 
      "Mã hàng": "SP004", 
      "Tên hàng": "Rau sạch Đà Lạt organic - Bắp cải & Xà lách", 
      "Giá bán": 120000, 
      "Số lượng": 10, 
      "Giảm giá %": 0, 
      "Giảm giá hóa đơn": 200000 
    },
    { 
      "Mã hóa đơn": "HD003", 
      "Mã hàng": "SP005", 
      "Tên hàng": "Sữa đậu nành đóng chai Fami - Thùng 40 hộp", 
      "Giá bán": 157500, 
      "Số lượng": 4, 
      "Giảm giá %": 2, 
      "Giảm giá hóa đơn": 200000 
    }
  ];

  const handleApply = () => {
    onLoadMock(mockProducts, mockInvoices);
  };

  return (
    <div id="manual-mock-uploader" className="bg-white rounded-lg p-5 border border-slate-200 shadow-sm duration-300">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex items-start space-x-3">
          <div className="p-2 bg-slate-150 rounded text-slate-700 border border-slate-200 flex-shrink-0">
            <Sparkles className="h-5 w-5 text-blue-500" />
          </div>
          <div>
            <h4 className="font-semibold text-slate-800 text-sm">Bộ dữ liệu mô phỏng chuẩn xác</h4>
            <p className="text-xs text-slate-500 mt-0.5">Dữ liệu kế toán đa dạng thuế suất (0%, 5%, 8%, 10%) để kiểm nghiệm phân bổ chiết khấu</p>
          </div>
        </div>
        <button
          onClick={handleApply}
          id="load-mock-datasets-btn"
          className={`flex items-center space-x-1.5 px-4 py-2 rounded font-semibold text-xs transition duration-200 cursor-pointer ${
            isLoaded 
              ? "bg-slate-100 text-slate-600 hover:bg-slate-200/80 border border-slate-200" 
              : "bg-blue-600 text-white hover:bg-blue-500 shadow-sm"
          }`}
        >
          {isLoaded ? (
            <>
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span>Dữ liệu đã sẵn sàng</span>
            </>
          ) : (
            <>
              <Database className="h-4 w-4" />
              <span>Nạp dữ liệu mẫu</span>
            </>
          )}
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-5">
        {/* Products Table Preview */}
        <div className="bg-slate-50/60 rounded-lg p-3.5 border border-slate-200/70">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-semibold text-slate-700 flex items-center space-x-1">
              <FileSpreadsheet className="h-3.5 w-3.5 text-slate-450" />
              <span>DanhSachSanPham ({mockProducts.length} dòng)</span>
            </span>
          </div>
          <div className="overflow-x-auto text-[11px] max-h-[140px]">
            <table className="w-full text-left text-slate-600 divide-y divide-slate-200">
              <thead>
                <tr className="text-slate-400 font-medium">
                  <th className="pb-1.5">Mã hàng</th>
                  <th className="pb-1.5">Mô tả sản phẩm (Tên gốc)</th>
                  <th className="pb-1.5 text-right">Thuế bán hàng</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-mono">
                {mockProducts.map((p, i) => (
                  <tr key={i} className="hover:bg-white/80 transition duration-150">
                    <td className="py-1 px-0.5 text-blue-600 font-semibold">{p["Mã hàng"]}</td>
                    <td className="py-1 truncate max-w-[150px]" title={p["Tên gốc"]}>{p["Tên gốc"]}</td>
                    <td className="py-1 text-right font-semibold text-amber-700">{(p["Thuế bán hàng"] * 100)}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Invoice Details preview */}
        <div className="bg-slate-50/60 rounded-lg p-3.5 border border-slate-200/70">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-semibold text-slate-700 flex items-center space-x-3">
              <FileSpreadsheet className="h-3.5 w-3.5 text-slate-450" />
              <span>DanhSachChiTietHoaDon ({mockInvoices.length} dòng)</span>
            </span>
          </div>
          <div className="overflow-x-auto text-[11px] max-h-[140px]">
            <table className="w-full text-left text-slate-600 divide-y divide-slate-200">
              <thead>
                <tr className="text-slate-400 font-medium">
                  <th className="pb-1.5">Hóa đơn</th>
                  <th className="pb-1.5">Mã hàng</th>
                  <th className="pb-1.5 text-right">Giá bán (gồm thuế)</th>
                  <th className="pb-1.5 text-center">SL</th>
                  <th className="pb-1.5 text-right">Giảm hóa đơn</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-mono">
                {mockInvoices.map((inv, i) => (
                  <tr key={i} className="hover:bg-white/80 transition duration-150">
                    <td className="py-1 px-0.5 font-bold text-slate-700">{inv["Mã hóa đơn"]}</td>
                    <td className="py-1 px-0.5 text-blue-600 font-semibold">{inv["Mã hàng"]}</td>
                    <td className="py-1 text-right text-slate-900">{inv["Giá bán"].toLocaleString("vi-VN")} đ</td>
                    <td className="py-1 text-center font-bold">{inv["Số lượng"]}</td>
                    <td className="py-1 text-right text-orange-600 font-semibold">{inv["Giảm giá hóa đơn"] > 0 ? `${inv["Giảm giá hóa đơn"].toLocaleString("vi-VN")} đ` : "-"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
