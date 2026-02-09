"use client";

import React, { useState } from "react";
import { ChevronDown } from "lucide-react";
import {
  AlertDialog,
  AlertDialogContent,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogAction,
  AlertDialogCancel,
} from "@/components/ui/alert-dialog";

interface SidebarHeaderProps {
  isDriveReady: boolean;
  onSignOut: () => void;
  onResetWorkspace: () => void;
  showPreview: boolean;
  onTogglePreview: (show: boolean) => void;
}

export const SidebarHeader: React.FC<SidebarHeaderProps> = ({
  isDriveReady,
  onSignOut,
  onResetWorkspace,
  showPreview,
  onTogglePreview,
}) => {
  const [showResetDialog, setShowResetDialog] = useState(false);

  return (
    <>
      <div className="p-8 border-b border-slate-100 bg-white/40">
        <div className="flex justify-between items-center mb-1">
          <h1 className="font-display font-black text-2xl text-slate-900 tracking-tighter">
            Smart Transmittal
          </h1>
          <div className="flex items-center gap-3">
            <button
              onClick={() => onTogglePreview(true)}
              className="lg:hidden text-slate-400 hover:text-brand-600 transition-colors"
              title="Minimize Sidebar"
            >
              <ChevronDown className="w-5 h-5" />
            </button>
            <button
              onClick={onSignOut}
              className="text-[9px] font-black text-slate-400 hover:text-slate-700 uppercase tracking-widest"
            >
              Sign out
            </button>
          </div>
        </div>
        <div className="flex justify-between items-center mt-1">
          <p className="text-[10px] text-slate-400 font-black uppercase tracking-widest">
            {isDriveReady
              ? "✓ Drive Connected"
              : "Enterprise Controller v2.1"}
          </p>
          <button
            onClick={() => setShowResetDialog(true)}
            className="text-[9px] font-black text-slate-400 hover:text-red-500 uppercase tracking-widest transition-colors"
          >
            Clear Workspace
          </button>
        </div>
      </div>

      <AlertDialog open={showResetDialog} onOpenChange={setShowResetDialog}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Clear Workspace?</AlertDialogTitle>
            <AlertDialogDescription>
              This will reset the current form. Your saved History snapshots will not be affected.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel onClick={() => setShowResetDialog(false)}>
              Cancel
            </AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                setShowResetDialog(false);
                onResetWorkspace();
              }}
            >
              Reset Form
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
};
