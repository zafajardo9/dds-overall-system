"use client";

import React from "react";
import { FileText, FileDown } from "lucide-react";
import { Button } from "@/components/ui/button";

interface SidebarFooterProps {
  onSaveTransmittal: () => void;
  onExportPdf: () => void;
  onExportDocx: () => void;
  isGeneratingPdf: boolean;
  isGeneratingDocx: boolean;
}

export const SidebarFooter: React.FC<SidebarFooterProps> = ({
  onSaveTransmittal,
  onExportPdf,
  onExportDocx,
  isGeneratingPdf,
  isGeneratingDocx,
}) => {
  return (
    <div className="p-8 border-t border-slate-200 bg-white shadow-[0_-10px_30px_rgba(0,0,0,0.03)]">
      <div className="mb-3">
        <Button
          onClick={onSaveTransmittal}
          className="w-full py-3 rounded-[24px] bg-brand-600 text-white text-[10px] font-black uppercase tracking-[0.2em] shadow-lg hover:bg-brand-500 active:scale-95"
        >
          Save to History
        </Button>
      </div>
      <div className="grid grid-cols-2 gap-3 mb-3">
        <Button
          onClick={onExportPdf}
          disabled={isGeneratingPdf}
          className="flex flex-col items-center justify-center py-4 h-auto bg-slate-900 text-white rounded-[24px] shadow-2xl hover:bg-slate-800 active:scale-95 group"
        >
          <FileText className="w-5 h-5 mb-1 group-hover:scale-110 transition-transform" />
          <span className="text-[9px] font-black uppercase tracking-[0.2em]">
            Export PDF
          </span>
        </Button>
        <Button
          onClick={onExportDocx}
          disabled={isGeneratingDocx}
          variant="outline"
          className="flex flex-col items-center justify-center py-4 h-auto rounded-[24px] hover:border-brand-500 hover:text-brand-600 active:scale-95 group"
        >
          <FileDown className="w-5 h-5 mb-1 group-hover:scale-110 transition-transform" />
          <span className="text-[9px] font-black uppercase tracking-[0.2em]">
            Export Word
          </span>
        </Button>
      </div>
    </div>
  );
};
