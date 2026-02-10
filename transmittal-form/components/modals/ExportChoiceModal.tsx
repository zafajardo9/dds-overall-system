"use client";

import React from "react";
import { Download, Cloud, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";

export type ExportFormat = "pdf" | "docx" | "csv";

interface ExportChoiceModalProps {
  isOpen: boolean;
  format: ExportFormat;
  fileName: string;
  isUploading: boolean;
  onDownloadLocal: () => void;
  onUploadToDrive: () => void;
  onClose: () => void;
}

const formatLabels: Record<ExportFormat, string> = {
  pdf: "PDF",
  docx: "Word Document",
  csv: "CSV Spreadsheet",
};

export const ExportChoiceModal: React.FC<ExportChoiceModalProps> = ({
  isOpen,
  format,
  fileName,
  isUploading,
  onDownloadLocal,
  onUploadToDrive,
  onClose,
}) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-sm rounded-[28px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white">
        <div className="p-8 pb-4 text-center">
          <h3 className="text-xl font-black text-slate-800 font-display">
            Export {formatLabels[format]}
          </h3>
          <p className="text-xs text-slate-400 mt-2 truncate max-w-[280px] mx-auto">
            {fileName}
          </p>
        </div>

        <div className="p-6 pt-2 space-y-3">
          <button
            onClick={onDownloadLocal}
            disabled={isUploading}
            className="w-full flex items-center gap-4 p-4 rounded-2xl border border-slate-100 hover:border-slate-300 hover:bg-slate-50 transition-all text-left group"
          >
            <div className="flex-shrink-0 w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center group-hover:bg-blue-50">
              <Download className="w-4 h-4 text-slate-500 group-hover:text-blue-600" />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-semibold text-slate-800">Download to device</p>
              <p className="text-[10px] text-slate-400">Save to your local computer</p>
            </div>
          </button>

          <button
            onClick={onUploadToDrive}
            disabled={isUploading}
            className="w-full flex items-center gap-4 p-4 rounded-2xl border border-slate-100 hover:border-slate-300 hover:bg-slate-50 transition-all text-left group"
          >
            <div className="flex-shrink-0 w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center group-hover:bg-green-50">
              {isUploading ? (
                <Loader2 className="w-4 h-4 text-green-600 animate-spin" />
              ) : (
                <Cloud className="w-4 h-4 text-slate-500 group-hover:text-green-600" />
              )}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-semibold text-slate-800">
                {isUploading ? "Uploading..." : "Save to Google Drive"}
              </p>
              <p className="text-[10px] text-slate-400">Choose a folder in your Drive</p>
            </div>
          </button>
        </div>

        <div className="px-6 pb-6">
          <Button
            onClick={onClose}
            variant="ghost"
            className="w-full rounded-xl text-xs text-slate-400"
            disabled={isUploading}
          >
            Cancel
          </Button>
        </div>
      </div>
    </div>
  );
};
