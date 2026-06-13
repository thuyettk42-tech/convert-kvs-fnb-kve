/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from "react";
import { CheckCircle2, ShieldAlert, Scale } from "lucide-react";

export default function MathQADecrypter() {
  return (
    <div id="math-qa-decrypter" className="bg-white rounded-lg p-6 border border-slate-200 shadow-sm space-y-6">
      {/* Title */}
      <div className="flex items-center space-x-3 pb-3 border-b border-slate-150">
        <div className="p-2 bg-slate-100 rounded text-slate-700 border border-slate-200 shadow-inner">
          <Scale className="h-5 w-5" />
        </div>
        <div>
          <h3 className="font-semibold text-slate-800 text-sm">Phân Tích Toán Học Kế Toán &amp; Bản Chất File Mẫu</h3>
          <p className="text-xs text-slate-500">Đối chiếu công thức theo quy định của Tổng Cục Thuế (Nghị định 123/2020/NĐ-CP)</p>
        </div>
      </div>

      {/* Answer the specific user question */}
      <div className="bg-slate-50/80 rounded p-5 border border-slate-200 space-y-3.5">
        <h4 className="font-semibold text-slate-800 text-xs uppercase tracking-wider flex items-center space-x-1.5">
          <CheckCircle2 className="h-4 w-4 text-blue-500" />
          <span>Làm rõ câu hỏi: Công thức chuẩn xác cho cột "Thành tiền"</span>
        </h4>
        
        <p className="text-xs text-slate-700 leading-relaxed font-medium">
          Trong hệ thống kế toán và tờ khai hóa đơn GTGT tiêu chuẩn, công thức tự chọn luôn bám sát khung pháp lý:
        </p>
        
        <div className="bg-white px-4 py-3 rounded border border-slate-200 font-mono text-center shadow-inner my-2">
          <span className="text-xs font-bold text-blue-600 block uppercase tracking-wider">THÀNH TIỀN CHUẨN KẾ TOÁN</span>
          <span className="text-sm font-bold text-slate-900 block mt-1">
            Thành tiền = (Giá trước thuế * Số lượng) - Tiền chiết khấu dòng
          </span>
          <span className="text-[10px] text-slate-500 block mt-1">
            = (Giá trước thuế * Số lượng) * (1 - Giảm giá % / 100)
          </span>
        </div>

        <p className="text-xs text-slate-600 leading-relaxed">
          Sở dĩ có điều này vì trong Nghị định 123/2020/NĐ-CP về hóa đơn điện tử, <strong>Chiết khấu thương mại của từng dòng hàng (chiết khấu dòng)</strong> được khấu trừ trực tiếp vào giá trị chưa thuế của mặt hàng trước khi tính VAT.
        </p>

        <p className="text-xs text-slate-500 leading-relaxed">
          💡 Hệ thống này được lập trình vượt trội bằng cách tích hợp đồng thời cả 4 chế độ công thức để anh tự chọn và chuyển đổi linh hoạt bằng dropdown ngay trên giao diện của công cụ, bảo đảm tệp nhập liệu (import) luôn hoàn thiện đúng ý anh!
        </p>
      </div>

      {/* Advanced calculation table */}
      <div className="space-y-3">
        <h4 className="font-semibold text-slate-800 text-xs uppercase tracking-wider">So Sánh 4 Tùy Chọn Tính Toán Trên Hệ Thống:</h4>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="p-4 rounded border border-slate-200 bg-white hover:bg-slate-50/50 transition">
            <span className="px-2 py-0.5 bg-slate-100 text-slate-700 border border-slate-200 rounded text-[10px] font-bold">Lựa chọn 1 (Khuyên dùng)</span>
            <h5 className="font-medium text-slate-800 text-xs mt-2">Mẫu Chuẩn Kế Toán (Đã trừ Chiết khấu dòng)</h5>
            <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
              Trừ chiết khấu dòng trước khi tính thuế. Tiền VAT dòng được tính dựa trên số Thành tiền đã giảm trừ này. Đáp ứng hoàn hảo cấu trúc hóa đơn GTGT.
            </p>
          </div>

          <div className="p-4 rounded border border-slate-200 bg-white hover:bg-slate-50/50 transition">
            <span className="px-2 py-0.5 bg-blue-50/50 text-blue-700 border border-blue-100/40 rounded text-[10px] font-bold">Lựa chọn 2</span>
            <h5 className="font-medium text-slate-800 text-xs mt-2">Nhân thuần túy (Không chiết khấu dòng)</h5>
            <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
              Thành tiền = Giá trước thuế * Số lượng. Không tự động trừ Chiết khấu dòng (thích hợp cho hóa đơn mà chiết khấu chỉ thể hiện dạng phân bổ tổng).
            </p>
          </div>

          <div className="p-4 rounded border border-slate-200 bg-white hover:bg-slate-50/50 transition">
            <span className="px-2 py-0.5 bg-purple-50/60 text-purple-700 border border-purple-150 rounded text-[10px] font-bold">Lựa chọn 3</span>
            <h5 className="font-medium text-slate-800 text-xs mt-2">Đơn giá * Giảm % (Theo prompt gốc)</h5>
            <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
              Thành tiền = Đơn giá * Giảm giá %. Đây là công thức toán học thô, hữu ích nếu trường 'Giảm giá %' của anh thực tế lưu trữ hệ số số lượng đặc biệt.
            </p>
          </div>

          <div className="p-4 rounded border border-slate-200 bg-white hover:bg-slate-50/50 transition">
            <span className="px-2 py-0.5 bg-orange-50 text-orange-700 border border-orange-100 rounded text-[10px] font-bold">Ủy thác VAT</span>
            <h5 className="font-medium text-slate-800 text-xs mt-2">Phân bổ Chiết khấu Thương mại (CKTM)</h5>
            <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
              Khi nhóm hóa đơn có giá trị 'Giảm giá hóa đơn' &gt; 0, hệ thống tự động sinh ra dòng CKTM tương ứng kèm mã <code className="font-mono text-purple-700 font-bold">CKTM8%</code>, <code className="font-mono text-purple-700 font-bold">CKTM10%</code>... để khấu trừ tiền trước thuế một cách hoàn hảo nhất.
            </p>
          </div>
        </div>
      </div>

      {/* Technical schema check list for accountants */}
      <div className="border-t border-slate-150 pt-5 space-y-3">
        <h4 className="font-semibold text-slate-750 text-xs uppercase tracking-wider flex items-center space-x-2">
          <ShieldAlert className="h-4.5 w-4.5 text-orange-550" />
          <span>Lưu ý kỹ thuật đặc biệt cho nhập liệu</span>
        </h4>
        <ul className="text-xs text-slate-600 space-y-2.5 list-disc pl-5">
          <li className="leading-relaxed">
            <strong>Gộp nhóm hóa đơn:</strong> Dòng dữ liệu đầu tiên của nhóm hóa đơn giữ thông tin người mua (khách lẻ không lấy hóa đơn, TM/CK), các dòng phụ của nhóm hóa đơn bắt buộc phải để trống thông tin người mua để hệ thống hóa đơn điện tử nhận dạng gộp tệp.
          </li>
          <li className="leading-relaxed">
            <strong>Tạo dòng chiết khấu thương mại động:</strong> Hóa đơn có thể có nhiều sản phẩm với mức thuế suất suất khác nhau. Phân bổ chiết khấu được làm tròn chuẩn xác tới hàng đơn vị đóng vai trò tối thiết để tổng toán dòng thành hợp lệ tuyệt đối.
          </li>
        </ul>
      </div>
    </div>
  );
}
