"use client";

import React, { useState, useEffect, useCallback } from "react";
import {
  Folder,
  FolderOpen,
  ChevronRight,
  ArrowLeft,
  Loader2,
  Search,
  Home,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { listDriveFolders } from "@/services/googleDriveService";

interface BreadcrumbItem {
  id: string;
  name: string;
}

interface FolderPickerModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelect: (folderId: string, folderName: string) => void;
}

export const FolderPickerModal: React.FC<FolderPickerModalProps> = ({
  isOpen,
  onClose,
  onSelect,
}) => {
  const [folders, setFolders] = useState<Array<{ id: string; name: string }>>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [search, setSearch] = useState("");
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([
    { id: "root", name: "My Drive" },
  ]);

  const currentFolderId = breadcrumbs[breadcrumbs.length - 1].id;
  const currentFolderName = breadcrumbs[breadcrumbs.length - 1].name;

  const loadFolders = useCallback(async (parentId: string, searchQuery?: string) => {
    setIsLoading(true);
    setError("");
    try {
      const result = await listDriveFolders(
        parentId === "root" ? undefined : parentId,
        searchQuery || undefined,
      );
      setFolders(result);
    } catch (err: any) {
      setError(err.message || "Failed to load folders");
      setFolders([]);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    if (isOpen) {
      setBreadcrumbs([{ id: "root", name: "My Drive" }]);
      setSearch("");
      setError("");
      loadFolders("root");
    }
  }, [isOpen, loadFolders]);

  useEffect(() => {
    if (!isOpen) return;
    const timer = setTimeout(() => {
      loadFolders(currentFolderId, search);
    }, 300);
    return () => clearTimeout(timer);
  }, [search, currentFolderId, isOpen, loadFolders]);

  const navigateToFolder = (folderId: string, folderName: string) => {
    setBreadcrumbs((prev) => [...prev, { id: folderId, name: folderName }]);
    setSearch("");
  };

  const navigateBack = () => {
    if (breadcrumbs.length <= 1) return;
    setBreadcrumbs((prev) => prev.slice(0, -1));
    setSearch("");
  };

  const navigateToBreadcrumb = (index: number) => {
    setBreadcrumbs((prev) => prev.slice(0, index + 1));
    setSearch("");
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
      <div className="bg-white w-full max-w-md rounded-[28px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white flex flex-col max-h-[70vh]">
        {/* Header */}
        <div className="p-6 pb-3 border-b border-slate-100 bg-slate-50/50">
          <div className="flex justify-between items-center mb-3">
            <h3 className="text-lg font-black text-slate-800 font-display">
              Choose Folder
            </h3>
            <Button
              onClick={onClose}
              variant="ghost"
              size="icon"
              className="rounded-full"
            >
              ✕
            </Button>
          </div>

          {/* Breadcrumbs */}
          <div className="flex items-center gap-1 text-xs text-slate-500 overflow-x-auto mb-3 scrollbar-hide">
            {breadcrumbs.map((crumb, i) => (
              <React.Fragment key={crumb.id}>
                {i > 0 && <ChevronRight className="w-3 h-3 flex-shrink-0 text-slate-300" />}
                <button
                  onClick={() => navigateToBreadcrumb(i)}
                  className={`flex-shrink-0 px-1.5 py-0.5 rounded hover:bg-slate-100 transition-colors ${
                    i === breadcrumbs.length - 1
                      ? "font-bold text-slate-700"
                      : "text-slate-400 hover:text-slate-600"
                  }`}
                >
                  {i === 0 ? <Home className="w-3 h-3 inline" /> : crumb.name}
                </button>
              </React.Fragment>
            ))}
          </div>

          {/* Search */}
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400 pointer-events-none" />
            <Input
              placeholder="Search folders..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="rounded-xl pl-9 text-xs"
            />
          </div>
        </div>

        {/* Folder list */}
        <div className="flex-1 overflow-y-auto p-3 space-y-1 min-h-[200px]">
          {breadcrumbs.length > 1 && !search && (
            <button
              onClick={navigateBack}
              className="w-full flex items-center gap-3 p-3 rounded-xl hover:bg-slate-50 transition-all text-left text-xs text-slate-400"
            >
              <ArrowLeft className="w-4 h-4" />
              Back
            </button>
          )}

          {isLoading ? (
            <div className="flex items-center justify-center py-10 text-slate-400">
              <Loader2 className="w-5 h-5 animate-spin mr-2" />
              <span className="text-sm">Loading folders...</span>
            </div>
          ) : error ? (
            <div className="flex flex-col items-center justify-center py-10 text-red-400">
              <p className="text-xs">{error}</p>
            </div>
          ) : folders.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-10 text-slate-400">
              <FolderOpen className="w-7 h-7 mb-2" />
              <p className="text-xs font-medium">
                {search ? "No matching folders" : "No subfolders here"}
              </p>
            </div>
          ) : (
            folders.map((folder) => (
              <button
                key={folder.id}
                onDoubleClick={() => navigateToFolder(folder.id, folder.name)}
                onClick={() => navigateToFolder(folder.id, folder.name)}
                className="w-full flex items-center gap-3 p-3 rounded-xl border border-transparent hover:border-slate-200 hover:bg-slate-50 transition-all text-left group"
              >
                <Folder className="w-4 h-4 text-amber-400 flex-shrink-0" />
                <span className="text-xs text-slate-700 truncate flex-1">
                  {folder.name}
                </span>
                <ChevronRight className="w-3 h-3 text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity flex-shrink-0" />
              </button>
            ))
          )}
        </div>

        {/* Footer: select current folder */}
        <div className="p-4 border-t border-slate-100 bg-slate-50/50 flex gap-2">
          <Button
            onClick={() => onSelect(currentFolderId, currentFolderName)}
            className="flex-1 rounded-xl text-xs font-bold"
            disabled={isLoading}
          >
            Save here
            {currentFolderName !== "My Drive" && (
              <span className="ml-1 opacity-70 truncate max-w-[120px]">
                ({currentFolderName})
              </span>
            )}
          </Button>
          <Button
            onClick={onClose}
            variant="outline"
            className="rounded-xl text-xs"
          >
            Cancel
          </Button>
        </div>
      </div>
    </div>
  );
};
