/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from "react";
import { Scale, FileSpreadsheet, Layers, BookOpen, Settings } from "lucide-react";

interface HeaderProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  productsCount: number;
  invoicesCount: number;
}

export default function Header({ activeTab, setActiveTab, productsCount, invoicesCount }: HeaderProps) {
  return (
    <header className="bg-[#1e293b] text-white sticky top-0 z-50 shadow-md">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex justify-between h-16 items-center">
          {/* Logo & Name */}
          <div className="flex items-center space-x-3.5">
            <div className="w-10 h-10 bg-blue-500 rounded flex items-center justify-center font-bold text-[11px] text-white shadow-md">
              KVE
            </div>
            <div>
              <h1 className="font-semibold text-white text-base flex items-center tracking-tight">
                <span>Convert Tool KVE</span>
              </h1>
              <p className="text-[11px] text-slate-400 leading-none mt-1">
                Tự động hóa chuẩn hóa &amp; phân bổ chiết khấu thương mại hóa đơn GTGT v2.1
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-3">
            {/* Status indicator badge */}
            <div className="hidden md:flex items-center gap-1.5 px-2.5 py-0.5 bg-green-500/10 text-green-400 rounded-full text-[10px] font-medium border border-green-500/20">
              <span className="w-1.5 h-1.5 bg-green-400 rounded-full animate-pulse"></span>
              <span>Environment Ready</span>
            </div>

            {/* Nav Links */}
            <div className="flex items-center space-x-1 bg-slate-800 p-1 rounded-md border border-slate-700/50">
              <button
                onClick={() => setActiveTab("processor")}
                className={`flex items-center space-x-1.5 px-3 py-1.5 rounded text-xs transition duration-200 cursor-pointer ${
                  activeTab === "processor"
                    ? "bg-blue-600 text-white shadow-sm font-semibold"
                    : "text-slate-300 hover:text-white"
                }`}
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                <span>Chuyển Đổi Excel</span>
                {(productsCount > 0 || invoicesCount > 0) && (
                  <span className="h-2 w-2 rounded-full bg-green-400 animate-pulse" />
                )}
              </button>
            </div>
          </div>
        </div>
      </div>
    </header>
  );
}
