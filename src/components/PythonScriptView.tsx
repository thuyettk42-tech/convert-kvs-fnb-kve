/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import { Copy, Check, FileCode, Play, AlertCircle } from "lucide-react";

interface PythonScriptViewProps {
  formula: string;
  isNegativeCKTM: boolean;
}

export default function PythonScriptView({ formula, isNegativeCKTM }: PythonScriptViewProps) {
  const [copied, setCopied] = useState(false);

  // Generate python python code based on state
  const getPythonCode = () => {
    let formulaExpr = "";
    if (formula === "standard_accounting") {
      formulaExpr = `df_merged["Thành tiền"] = (df_merged["Giá trước thuế"] * df_merged["Số lượng"]) * (1 - df_merged["Giảm giá %"] / 100)`;
    } else if (formula === "standard_discount_rate") {
      formulaExpr = `df_merged["Thành tiền"] = (df_merged["Giá trước thuế"] * df_merged["Số lượng"]) * (1 - df_merged["Giảm giá %"] / 100)`;
    } else if (formula === "multiplier") {
      formulaExpr = `df_merged["Thành tiền"] = df_merged["Giá trước thuế"] * df_merged["Giảm giá %"]`;
    } else {
      formulaExpr = `df_merged["Thành tiền"] = df_merged["Giá trước thuế"] * df_merged["Số lượng"]`;
    }

    const cktmSign = isNegativeCKTM ? "-" : "";

    return `import pandas as pd
import numpy as np
from datetime import datetime

r"""
========================================================================================
SCRIPT ĐỒNG BỘ & CHUYỂN ĐỔI DỮ LIỆU KẾ TOÁN BÁN HÀNG SANG TEMPLATE HÓA ĐƠN ĐIỆN TỬ GTGT
========================================================================================
- Vai trò: Data Engineer & Python Developer
- Thư viện yêu cầu: pandas, openpyxl, numpy
- Cài đặt thư viện: pip install pandas openpyxl numpy
========================================================================================
"""

def execute_accounting_mapping(prod_catalog_path, invoice_details_path, output_path="Mau_import_hoa_don_GTGT.xlsx"):
    # ----------------------------------------------------------------------------------
    # BƯỚC 1: Đọc tệp đầu vào
    # ----------------------------------------------------------------------------------
    print("-> Đang tải dữ liệu đầu vào...")
    # Đọc danh mục hàng hóa (Mã hàng, Thuế bán hàng)
    # Hỗ trợ cả file xlsx và csv
    if prod_catalog_path.endswith('.csv'):
        df_prod = pd.read_csv(prod_catalog_path)
    else:
        df_prod = pd.read_excel(prod_catalog_path)
        
    if invoice_details_path.endswith('.csv'):
        df_inv = pd.read_csv(invoice_details_path)
    else:
        df_inv = pd.read_excel(invoice_details_path)

    # Chuẩn hóa tên cột (loại bỏ khoảng trắng dư thừa)
    df_prod.columns = df_prod.columns.str.strip()
    df_inv.columns = df_inv.columns.str.strip()

    # ----------------------------------------------------------------------------------
    # BƯỚC 2: QUY TRÌNH XỬ LÝ 1 - LÀM GIÀU DỮ LIỆU
    # ----------------------------------------------------------------------------------
    print("-> Đang thực hiện làm giàu dữ liệu...")
    
    # 1. Truy xuất cột 'Thuế bán hàng' từ danh mục sản phẩm dựa theo 'Mã hàng'
    # Đảm bảo kiểu dữ liệu chuỗi cho mã hàng để so khớp chính xác
    df_prod["Mã hàng"] = df_prod["Mã hàng"].astype(str).str.strip()
    df_inv["Mã hàng"] = df_inv["Mã hàng"].astype(str).str.strip()
    
    # Lấy bảng ánh xạ giữa Mã hàng và Thuế suất
    tax_mapping = df_prod.set_index("Mã hàng")["Thuế bán hàng"].to_dict()
    
    # Bổ sung cột Thuế bán hàng vào chi tiết hóa đơn
    df_inv["Thuế bán hàng"] = df_inv["Mã hàng"].map(tax_mapping)
    # Điền giá trị mặc định nếu không tìm thấy thuế suất (mặc định 10% = 0.10)
    df_inv["Thuế bán hàng"] = df_inv["Thuế bán hàng"].fillna(0.10)
    
    # 2. Tạo cột 'Giá trước thuế' theo công thức: Giá bán / (1 + Thuế bán hàng)
    # Tự động phát hiện cột Giá bán (bằng cách so khớp mờ nếu cần)
    price_col = "Giá bán" if "Giá bán" in df_inv.columns else [c for c in df_inv.columns if "giá" in c.lower() or "price" in c.lower()][0]
    
    # Chuyển đổi định dạng số
    df_inv[price_col] = pd.to_numeric(df_inv[price_col], errors='coerce').fillna(0)
    df_inv["Thuế bán hàng"] = pd.to_numeric(df_inv["Thuế bán hàng"], errors='coerce').fillna(0)
    
    df_inv["Giá trước thuế"] = df_inv[price_col] / (1 + df_inv["Thuế bán hàng"])
    # Làm tròn giá trước thuế đến 2 chữ số thập phân
    df_inv["Giá trước thuế"] = df_inv["Giá trước thuế"].round(2)

    # ----------------------------------------------------------------------------------
    # BƯỚC 3: QUY TRÌNH XỬ LÝ 2 - MAPPING DỮ LIỆU ĐẦU RA SANG TEMPLATE
    # ----------------------------------------------------------------------------------
    print("-> Ánh xạ dữ liệu sang cấu trúc template...")
    df_merged = df_inv.copy()
    
    # Chuyển đổi các cột số cần thiết
    df_merged["Số lượng"] = pd.to_numeric(df_merged["Số lượng"], errors='coerce').fillna(0)
    df_merged["Giảm giá %"] = pd.to_numeric(df_merged["Giảm giá %"], errors='coerce').fillna(0)
    df_merged["Giảm giá hóa đơn"] = pd.to_numeric(df_merged["Giảm giá hóa đơn"], errors='coerce').fillna(0)

    # Áp dụng công thức Thành tiền theo tùy chọn được cấu hình:
    # ${formulaExpr}
    ${formulaExpr}
    df_merged["Thành tiền"] = df_merged["Thành tiền"].round(2)
    
    # Tính Tổng tiền = Thành tiền * (1 + % VAT)
    df_merged["Tổng tiền *"] = df_merged["Thành tiền"] * (1 + df_merged["Thuế bán hàng"])
    df_merged["Tổng tiền *"] = df_merged["Tổng tiền *"].round(2)

    # Danh sách các nhóm hóa đơn duy nhất
    invoice_groups = df_merged["Mã hóa đơn"].unique()
    final_rows = []

    # Định dạng ngày hiện tại (dd/mm/yyyy)
    current_date_str = datetime.now().strftime("%d/%m/%Y")

    for inv_id in invoice_groups:
        # Lấy tất cả các dòng của hóa đơn này
        df_group = df_merged[df_merged["Mã hóa đơn"] == inv_id].copy()
        
        # Lấy giá trị Giảm giá hóa đơn đầu tiên của nhóm
        invoice_discount = df_group["Giảm giá hóa đơn"].iloc[0] if len(df_group) > 0 else 0
        total_inv_group_amount = df_group["Tổng tiền *"].sum()
        
        # Duyệt qua các mặt hàng thực tế trong nhóm để thêm dòng hóa đơn
        first_row_processed = False
        
        for idx, row in df_group.iterrows():
            # Quy tắc gộp nhóm: Dòng đầu tiên chứa thông tin khách hàng, các dòng sau để trống
            out_row = {
                "Ngày hóa đơn": current_date_str if not first_row_processed else "",
                "Mã số thuế": "",
                "Tên khách hàng": "Khách lẻ không lấy hóa đơn" if not first_row_processed else "",
                "Họ tên người mua": "",
                "Số CCCD": "",
                "Địa chỉ": "",
                "Điện thoại": "",
                "Email nhận HĐ": "",
                "Hình thức thanh toán": "TM/CK" if not first_row_processed else "",
                "Tên ngân hàng": "",
                "Mã hàng": row["Mã hàng"],
                "Tên hàng hóa, dịch vụ *": row["Tên hàng"],
                "Tính chất hàng hóa, dịch vụ": "Hàng hóa, dịch vụ",
                "Số lượng": row["Số lượng"],
                "Đơn giá": row["Giá trước thuế"],
                "Chiết khấu (%)": row["Giảm giá %"],
                "Tiền chiết khấu": "",
                "Thành tiền": row["Thành tiền"],
                "% VAT": f"{int(row['Thuế bán hàng'] * 100)}%" if row["Thuế bán hàng"] > 0 else "0%",
                "Tiền VAT": "",
                "Tổng tiền *": row["Tổng tiền *"],
                "Mã nhóm hóa đơn": inv_id
            }
            final_rows.append(out_row)
            first_row_processed = True

        # ------------------------------------------------------------------------------
        # BƯỚC 4: QUY TRÌNH XỬ LÝ 3 - BỔ SUNG DÒNG CHIẾT KHẤU THƯƠNG MẠI
        # ------------------------------------------------------------------------------
        if invoice_discount > 0 and total_inv_group_amount > 0:
            # Nhóm các mặt hàng theo mức thuế suất để phân bổ CKTM
            vat_rates = df_group["Thuế bán hàng"].unique()
            print(f"   -> Phát hiện CKTM hóa đơn {inv_id}, thực hiện phân bổ cho {len(vat_rates)} nhóm thuế...")

            for rate in vat_rates:
                # Lọc hàng hóa có nhóm thuế VAT này
                df_rate_group = df_group[df_group["Thuế bán hàng"] == rate]
                sum_total_rate = df_rate_group["Tổng tiền *"].sum()
                
                if sum_total_rate == 0:
                    continue
                
                # Tính tổng tiền CKTM phân bổ: [Giảm giá hóa đơn] * [Tổng tiền nhóm thuế] / [Tổng giá trị đơn hàng]
                allocated_total = (invoice_discount * (sum_total_rate / total_inv_group_amount))
                allocated_total_rounded = round(allocated_total) # Làm tròn tới hàng đơn vị
                
                # Chiết khấu thương mại làm giảm doanh thu, gán dấu tương ứng
                total_val = ${cktmSign}allocated_total_rounded
                
                # Tính toán ngược lại các chỉ tiêu còn lại
                vat_factor = 1 + rate
                allocated_thanh_tien = round(total_val / vat_factor, 2)
                allocated_don_gia = allocated_thanh_tien
                
                vat_pct_str = f"{int(rate * 100)}%" if rate > 0 else "0%"
                
                # Tạo dòng CKTM (các trường thông tin khách hàng đều để trống vì đứng sau trong nhóm)
                cktm_row = {
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
                    "Mã hàng": f"CKTM{vat_pct_str}",
                    "Tên hàng hóa, dịch vụ *": "Chiết khấu thương mại",
                    "Tính chất hàng hóa, dịch vụ": "Hàng hóa chiết khấu thương mại",
                    "Số lượng": "",
                    "Đơn giá": allocated_don_gia,
                    "Chiết khấu (%)": "",
                    "Tiền chiết khấu": "",
                    "Thành tiền": allocated_thanh_tien,
                    "% VAT": vat_pct_str,
                    "Tiền VAT": "",
                    "Tổng tiền *": total_val,
                    "Mã nhóm hóa đơn": inv_id
                }
                final_rows.append(cktm_row)

    # ----------------------------------------------------------------------------------
    # BƯỚC 5: Ghi file đầu ra theo Mẫu import hóa đơn GTGT
    # ----------------------------------------------------------------------------------
    df_out = pd.DataFrame(final_rows)
    
    # Loại bỏ cột bổ trợ nhóm nội bộ trước khi xuất file
    df_out_clean = df_out.drop(columns=["Mã nhóm hóa đơn"])
    
    # Xuất file Excel
    df_out_clean.to_excel(output_path, index=False)
    print(f"✔ Đã xử lý thành công! File kết quả lưu tại: {output_path}")
    return df_out_clean

# --------------------------------------------------------------------------------------
# CHẠY THỬ NGHIỆM SCRIPT
# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    # Thay thế đường dẫn thực tế của bạn tại đây
    # execute_accounting_mapping("DanhSachSanPham.xlsx", "DanhSachChiTietHoaDon.xlsx")
    pass
`;
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(getPythonCode());
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div id="python-script-view" className="bg-[#0f172a] text-slate-100 rounded-lg overflow-hidden border border-slate-700 shadow-xl flex flex-col h-full">
      {/* Script Header */}
      <div className="flex items-center justify-between px-5 py-4 bg-[#0f172a] border-b border-slate-800">
        <div className="flex items-center space-x-3">
          <div className="p-2 bg-slate-800 rounded text-blue-400 border border-slate-700">
            <FileCode className="h-5 w-5" />
          </div>
          <div>
            <h3 className="font-semibold text-sm text-slate-200">Python Script Tự Động Hóa</h3>
            <p className="text-xs text-slate-500">Sử dụng Pandas và Openpyxl chuẩn hóa kế toán</p>
          </div>
        </div>
        <button
          onClick={copyToClipboard}
          id="copy-python-code-btn"
          className="flex items-center space-x-1.5 px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 active:bg-slate-700 text-xs text-slate-300 hover:text-white rounded transition duration-250 cursor-pointer shadow-sm border border-slate-700 font-medium"
        >
          {copied ? (
            <>
              <Check className="h-3.5 w-3.5 text-blue-400" />
              <span className="text-blue-400">Đã Sao Chép!</span>
            </>
          ) : (
            <>
              <Copy className="h-3.5 w-3.5" />
              <span>Sao chép mã</span>
            </>
          )}
        </button>
      </div>

      {/* Warning/Guidance bar */}
      <div className="bg-slate-800/40 text-blue-300 px-5 py-3 border-b border-slate-800 text-xs flex items-start space-x-2.5">
        <AlertCircle className="h-4.5 w-4.5 text-blue-400 mt-0.5 shrink-0" />
        <div>
          <span className="font-semibold">Mẹo chạy cục bộ:</span> Chạy lệnh <code className="bg-slate-900 px-1 py-0.5 rounded text-blue-300 font-mono text-[11px] border border-slate-800">pip install pandas openpyxl</code> trên máy tính cá nhân của bạn, lưu đoạn code ở dưới thành file <code className="font-semibold text-slate-100 font-mono">mapping_data.py</code> và thiết lập đường dẫn tệp để xử lý tự động hàng triệu dòng siêu tốc!
        </div>
      </div>

      {/* Code Area */}
      <div className="flex-1 overflow-auto p-5 font-mono text-xs leading-relaxed text-slate-300 bg-[#0f172a]">
        <pre className="whitespace-pre-wrap select-all selection:bg-slate-800 select-text">
          {getPythonCode()}
        </pre>
      </div>
    </div>
  );
}
