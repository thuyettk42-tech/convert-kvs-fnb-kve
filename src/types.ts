/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ProductCatalogRow {
  "Mã hàng": string;
  "Thuế bán hàng": number; // decimal e.g., 0.08 or 0.10
  [key: string]: any;
}

export interface InvoiceDetailRow {
  "Mã hóa đơn": string;
  "Mã hàng": string;
  "Tên hàng": string;
  "Giá bán": number;
  "Số lượng": number;
  "Giảm giá %": number;
  "Giảm giá hóa đơn": number;
  [key: string]: any;
}

export interface OutputRow {
  "Ngày hóa đơn": string;
  "Mã số thuế": string;
  "Tên khách hàng": string;
  "Họ tên người mua": string;
  "Số CCCD": string;
  "Địa chỉ": string;
  "Điện thoại": string;
  "Email nhận HĐ": string;
  "Hình thức thanh toán": string;
  "Tên ngân hàng": string;
  "Mã hàng": string;
  "Tên hàng hóa, dịch vụ *": string;
  "Tính chất hàng hóa, dịch vụ": string;
  "Số lượng": string | number;
  "Đơn giá": number;
  "Chiết khấu (%)": number;
  "Tiền chiết khấu": string;
  "Thành tiền": number;
  "% VAT": string | number;
  "Tiền VAT": string;
  "Tổng tiền *": number;
  "Mã nhóm hóa đơn": string; // Internal grouping helper
}

export type ThanhTienFormula = 
  | "standard_accounting" // (Đơn giá * Số lượng) - Tiền chiết khấu
  | "standard_discount_rate" // (Giá trước thuế * Số lượng) * (1 - Giảm giá % / 100)
  | "multiplier"          // Đơn giá * Chiết khấu % (from literal prompt)
  | "no_discount";        // Đơn giá * Số lượng

export interface ColumnMapping {
  // Mapping for Product Catalogue (DanhSachSanPham)
  productCodeCol: string;
  productTaxCol: string;

  // Mapping for Invoice Details (DanhSachChiTietHoaDon)
  invoiceIdCol: string;
  invoiceProductCodeCol: string;
  invoiceProductNameCol: string;
  invoicePriceCol: string;
  invoiceQtyCol: string;
  invoiceDiscountPercentCol: string;
  invoiceDiscountAmountCol: string;
}
