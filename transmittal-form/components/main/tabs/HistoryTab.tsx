"use client";

import React from "react";
import { Clock, X, CheckCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { HistoryItem } from "@/types";

interface HistoryTabProps {
  history: HistoryItem[];
  onLoadFromHistory: (item: HistoryItem) => void;
  onDeleteHistoryItem: (id: string, e: React.MouseEvent) => void;
  linkedSheetTitle: string | null;
  sheetUrlInput: string;
  onSheetUrlInputChange: (value: string) => void;
  sheetLinkError: string;
  isLinkingSheet: boolean;
  onLinkSheet: () => void;
  onUnlinkSheet: () => void;
}

export const HistoryTab: React.FC<HistoryTabProps> = ({
  history,
  onLoadFromHistory,
  onDeleteHistoryItem,
  linkedSheetTitle,
  sheetUrlInput,
  onSheetUrlInputChange,
  sheetLinkError,
  isLinkingSheet,
  onLinkSheet,
  onUnlinkSheet,
}) => {
  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
      {/* Google Sheets Link */}
      <div className="p-5 bg-emerald-50/60 border border-emerald-200 rounded-[32px]">
        <h3 className="text-[9px] font-black text-emerald-700 uppercase tracking-widest mb-3">
          Google Sheet Sync
        </h3>
        {linkedSheetTitle ? (
          <div className="flex items-center justify-between gap-3">
            <div className="flex items-center gap-2 min-w-0">
              <span className="shrink-0 w-6 h-6 rounded-full bg-emerald-200 flex items-center justify-center text-emerald-700">
                <CheckCircle className="w-3.5 h-3.5" />
              </span>
              <span className="text-[10px] font-bold text-emerald-800 truncate">
                {linkedSheetTitle}
              </span>
            </div>
            <button
              onClick={onUnlinkSheet}
              className="text-[9px] font-black text-red-400 hover:text-red-600 uppercase tracking-widest whitespace-nowrap"
            >
              Unlink
            </button>
          </div>
        ) : (
          <div className="space-y-2">
            <div className="flex gap-2">
              <Input
                className="text-[10px] flex-1"
                placeholder="Paste Google Sheets URL..."
                value={sheetUrlInput}
                onChange={(e) => onSheetUrlInputChange(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && onLinkSheet()}
              />
              <Button
                onClick={onLinkSheet}
                disabled={isLinkingSheet || !sheetUrlInput.trim()}
                className="rounded-xl text-[9px] font-black uppercase tracking-widest px-4"
                size="sm"
              >
                {isLinkingSheet ? "..." : "Link"}
              </Button>
            </div>
            {sheetLinkError && (
              <p className="text-[9px] font-bold text-red-500">{sheetLinkError}</p>
            )}
            <p className="text-[9px] text-emerald-600">
              New transmittals will be logged as rows in the linked sheet.
            </p>
          </div>
        )}
      </div>

      <div className="flex justify-between items-center mb-4">
        <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
          Workspace Snapshots
        </h2>
        <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
          {history.length}/20
        </span>
      </div>
      {history.length === 0 ? (
        <div className="text-center py-20 bg-slate-50/50 rounded-[40px] border border-dashed border-slate-200">
          <p className="text-[10px] font-black text-slate-300 uppercase tracking-widest">
            No Recent Snapshots
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {history.map((item) => (
            <div
              key={item.id}
              onClick={() => onLoadFromHistory(item)}
              className="group p-5 bg-white border border-slate-200 rounded-[32px] cursor-pointer hover:border-brand-500 hover:shadow-xl transition-all relative overflow-hidden"
            >
              <div className="flex justify-between items-start mb-2">
                <div className="shrink-0 w-8 h-8 rounded-full bg-slate-50 flex items-center justify-center text-slate-400 group-hover:bg-brand-50 group-hover:text-brand-500 transition-all">
                  <Clock className="w-4 h-4" />
                </div>
                <button
                  onClick={(e) => onDeleteHistoryItem(item.id, e)}
                  className="p-2 text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="space-y-1">
                <p className="text-[10px] font-black text-slate-800 line-clamp-1">
                  {item.projectName}
                </p>
                <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest">
                  {item.transmittalNumber}
                </p>
                <div className="flex flex-wrap gap-x-4 gap-y-1 text-[9px] text-slate-500">
                  <span className="font-semibold text-slate-600">
                    {item.createdBy}
                  </span>
                  <span>Prepared: {item.preparedBy || "—"}</span>
                  <span>Noted: {item.notedBy || "—"}</span>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
