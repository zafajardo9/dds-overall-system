"use client";

import React, { useRef } from "react";
import { Zap, Upload, HardDrive } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { AppData } from "@/types";

interface ContentTabProps {
  smartInput: string;
  onSmartInputChange: (value: string) => void;
  isAnalyzingText: boolean;
  onSmartAnalysis: () => void;
  isParsing: boolean;
  parseProgress: { current: number; total: number };
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  onBatchUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
  isDriveReady: boolean;
  onOpenDriveModal: () => void;
  statusMsg: string;
  statusType: "info" | "error";
  transmissionMethod: AppData["transmissionMethod"];
  onUpdateTransmission: (method: string, checked: boolean) => void;
  notes: string;
  onUpdateNotes: (value: string) => void;
}

export const ContentTab: React.FC<ContentTabProps> = ({
  smartInput,
  onSmartInputChange,
  isAnalyzingText,
  onSmartAnalysis,
  isParsing,
  parseProgress,
  fileInputRef,
  onBatchUpload,
  isDriveReady,
  onOpenDriveModal,
  statusMsg,
  statusType,
  transmissionMethod,
  onUpdateTransmission,
  notes,
  onUpdateNotes,
}) => {
  const transmissionOptions = [
    { key: "personalDelivery", label: "Hand Delivered" },
    { key: "pickUp", label: "Pick-up" },
    { key: "grabLalamove", label: "Courier" },
    { key: "registeredMail", label: "Registered Mail" },
  ] as const;

  return (
    <div className="space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <section>
        <div className="flex justify-between items-end mb-4">
          <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
            Intelligent Import
          </h2>
          {isParsing && (
            <span className="text-[10px] font-bold text-brand-600 animate-pulse uppercase">
              Parsing {parseProgress.current}/{parseProgress.total}...
            </span>
          )}
        </div>
        <div className="space-y-4">
          <div className="relative group">
            <Textarea
              className="w-full p-5 bg-slate-100/50 border border-slate-200 rounded-3xl text-xs font-medium focus:bg-white focus:ring-4 focus:ring-brand-500/5 focus:border-brand-500 outline-none transition-all placeholder-slate-300 min-h-[120px]"
              placeholder="Paste text, file names, Google Drive file links, or folder links here..."
              value={smartInput}
              onChange={(e) => onSmartInputChange(e.target.value)}
            />
            <div className="absolute right-4 bottom-4 flex gap-2">
              <Button
                onClick={onSmartAnalysis}
                disabled={!smartInput.trim() || isAnalyzingText}
                size="icon"
                className="rounded-xl shadow-lg"
              >
                <Zap className={`w-4 h-4 ${isAnalyzingText ? "animate-spin" : ""}`} />
              </Button>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Button
              onClick={() => fileInputRef.current?.click()}
              variant="outline"
              className="flex items-center justify-center gap-2 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-sm"
            >
              <Upload className="w-4 h-4 text-slate-400" />
              Upload Files
            </Button>
            <Button
              onClick={onOpenDriveModal}
              disabled={!isDriveReady}
              variant="outline"
              className="flex items-center justify-center gap-2 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-sm"
            >
              <HardDrive className="w-4 h-4 text-slate-400" />
              Browse Drive
            </Button>
          </div>
          <input
            type="file"
            multiple
            ref={fileInputRef}
            className="hidden"
            accept=".pdf,image/*"
            onChange={onBatchUpload}
          />
          {statusMsg && (
            <div
              className={`p-4 rounded-2xl text-[10px] font-bold border animate-in slide-in-from-right duration-300 ${statusType === "error" ? "bg-red-50 border-red-100 text-red-600" : "bg-brand-50 border-brand-100 text-brand-600"}`}
            >
              {statusMsg}
            </div>
          )}
        </div>
      </section>

      <section className="space-y-4">
        <div className="flex justify-between items-center">
          <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em]">
            Transmission
          </h2>
        </div>
        <div className="p-6 bg-slate-50/50 rounded-[32px] border border-slate-100/80 space-y-4">
          <div className="grid grid-cols-2 gap-4">
            {transmissionOptions.map(({ key, label }) => (
              <label
                key={key}
                className="flex items-center gap-3 p-3 bg-white rounded-2xl border border-slate-100 cursor-pointer hover:border-brand-500 transition-all"
              >
                <input
                  type="checkbox"
                  checked={transmissionMethod[key]}
                  onChange={(e) => onUpdateTransmission(key, e.target.checked)}
                  className="w-4 h-4 rounded-md text-brand-600 border-slate-300"
                />
                <span className="text-[10px] font-black text-slate-600 uppercase tracking-widest">
                  {label}
                </span>
              </label>
            ))}
          </div>
          <div className="space-y-1">
            <Label className="text-[10px] font-black text-muted-foreground uppercase tracking-widest ml-1">
              Internal Notes
            </Label>
            <Textarea
              rows={2}
              value={notes}
              onChange={(e) => onUpdateNotes(e.target.value)}
            />
          </div>
        </div>
      </section>
    </div>
  );
};
