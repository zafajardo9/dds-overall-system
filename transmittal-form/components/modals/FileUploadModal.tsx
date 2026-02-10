"use client";

import React, { useState, useRef, useCallback } from "react";
import {
  Upload,
  X,
  FileText,
  Image as ImageIcon,
  Cloud,
  Loader2,
  Trash2,
  CheckCircle2,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { uploadFileToDrive } from "@/services/googleDriveService";
import { FolderPickerModal } from "./FolderPickerModal";

interface FileUploadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUploadFiles: (files: File[]) => void;
  isDriveReady: boolean;
  isParsing: boolean;
  parseProgress: { current: number; total: number };
}

export const FileUploadModal: React.FC<FileUploadModalProps> = ({
  isOpen,
  onClose,
  onUploadFiles,
  isDriveReady,
  isParsing,
  parseProgress,
}) => {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [saveToDrive, setSaveToDrive] = useState(false);
  const [folderPickerOpen, setFolderPickerOpen] = useState(false);
  const [driveUploadStatus, setDriveUploadStatus] = useState<
    "idle" | "uploading" | "done" | "error"
  >("idle");
  const [driveUploadMsg, setDriveUploadMsg] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  const acceptedTypes = ["application/pdf", "image/"];

  const isAccepted = (file: File) =>
    acceptedTypes.some((t) =>
      t.endsWith("/") ? file.type.startsWith(t) : file.type === t,
    );

  const addFiles = useCallback(
    (newFiles: File[]) => {
      const valid = newFiles.filter(isAccepted);
      if (valid.length === 0) return;
      setSelectedFiles((prev) => {
        const names = new Set(prev.map((f) => f.name));
        const deduped = valid.filter((f) => !names.has(f.name));
        return [...prev, ...deduped];
      });
    },
    [],
  );

  const removeFile = (index: number) => {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    addFiles(files);
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.from(e.target.files) : [];
    addFiles(files);
    e.target.value = "";
  };

  const handleProcess = () => {
    if (selectedFiles.length === 0) return;
    if (saveToDrive && isDriveReady) {
      setFolderPickerOpen(true);
    } else {
      onUploadFiles(selectedFiles);
      setSelectedFiles([]);
      onClose();
    }
  };

  const handleFolderSelected = async (folderId: string, folderName: string) => {
    setFolderPickerOpen(false);
    setDriveUploadStatus("uploading");
    setDriveUploadMsg(`Saving to ${folderName}...`);

    try {
      for (const file of selectedFiles) {
        const blob = new Blob([file], { type: file.type });
        await uploadFileToDrive(blob, file.name, folderId);
      }
      setDriveUploadStatus("done");
      setDriveUploadMsg(`${selectedFiles.length} file(s) saved to ${folderName}`);
    } catch (err: any) {
      setDriveUploadStatus("error");
      setDriveUploadMsg(`Drive upload failed: ${err.message}`);
    }

    // Also process them for data extraction
    onUploadFiles(selectedFiles);
    setTimeout(() => {
      setSelectedFiles([]);
      setDriveUploadStatus("idle");
      setDriveUploadMsg("");
      setSaveToDrive(false);
      onClose();
    }, 1500);
  };

  const handleClose = () => {
    if (isParsing || driveUploadStatus === "uploading") return;
    setSelectedFiles([]);
    setSaveToDrive(false);
    setDriveUploadStatus("idle");
    setDriveUploadMsg("");
    onClose();
  };

  const getFileIcon = (file: File) => {
    if (file.type.startsWith("image/"))
      return <ImageIcon className="w-4 h-4 text-purple-400" />;
    return <FileText className="w-4 h-4 text-blue-400" />;
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1048576).toFixed(1)} MB`;
  };

  if (!isOpen) return null;

  return (
    <>
      <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-md animate-in fade-in duration-300">
        <div className="bg-white w-full max-w-lg rounded-[28px] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-300 border border-white flex flex-col max-h-[80vh]">
          {/* Header */}
          <div className="p-6 pb-3 flex justify-between items-center">
            <div>
              <h3 className="text-xl font-black text-slate-800 font-display">
                Upload Files
              </h3>
              <p className="text-[10px] text-slate-400 mt-1">
                PDF documents and images supported
              </p>
            </div>
            <Button
              onClick={handleClose}
              variant="ghost"
              size="icon"
              className="rounded-full"
              disabled={isParsing || driveUploadStatus === "uploading"}
            >
              <X className="w-4 h-4" />
            </Button>
          </div>

          {/* Drop zone */}
          <div className="px-6">
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`relative border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all ${
                isDragging
                  ? "border-blue-400 bg-blue-50/50 scale-[1.01]"
                  : "border-slate-200 hover:border-slate-300 hover:bg-slate-50/50"
              }`}
            >
              <div className="flex flex-col items-center gap-3">
                <div
                  className={`w-12 h-12 rounded-2xl flex items-center justify-center transition-colors ${
                    isDragging ? "bg-blue-100" : "bg-slate-100"
                  }`}
                >
                  <Upload
                    className={`w-5 h-5 ${isDragging ? "text-blue-500" : "text-slate-400"}`}
                  />
                </div>
                <div>
                  <p className="text-sm font-bold text-slate-700">
                    {isDragging ? "Drop files here" : "Drag & drop files here"}
                  </p>
                  <p className="text-[10px] text-slate-400 mt-1">
                    or click to browse your device
                  </p>
                </div>
              </div>
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept=".pdf,image/*"
                className="hidden"
                onChange={handleFileInput}
              />
            </div>
          </div>

          {/* File list */}
          {selectedFiles.length > 0 && (
            <div className="flex-1 overflow-y-auto px-6 mt-4 space-y-2 min-h-0 max-h-[200px]">
              <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                {selectedFiles.length} file{selectedFiles.length !== 1 ? "s" : ""} selected
              </p>
              {selectedFiles.map((file, index) => (
                <div
                  key={`${file.name}-${index}`}
                  className="flex items-center gap-3 p-3 rounded-xl bg-slate-50/70 border border-slate-100 group"
                >
                  {getFileIcon(file)}
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-semibold text-slate-700 truncate">
                      {file.name}
                    </p>
                    <p className="text-[10px] text-slate-400">
                      {formatSize(file.size)}
                    </p>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      removeFile(index);
                    }}
                    className="opacity-0 group-hover:opacity-100 transition-opacity p-1 hover:bg-red-50 rounded-lg"
                  >
                    <Trash2 className="w-3 h-3 text-red-400" />
                  </button>
                </div>
              ))}
            </div>
          )}

          {/* Drive toggle + status */}
          <div className="px-6 mt-4 space-y-3">
            {isDriveReady && selectedFiles.length > 0 && (
              <label className="flex items-center gap-3 p-3 rounded-xl border border-slate-100 cursor-pointer hover:bg-slate-50 transition-all">
                <input
                  type="checkbox"
                  checked={saveToDrive}
                  onChange={(e) => setSaveToDrive(e.target.checked)}
                  className="w-4 h-4 rounded text-blue-600 border-slate-300"
                />
                <Cloud className="w-4 h-4 text-slate-400" />
                <span className="text-xs font-medium text-slate-600">
                  Also save original files to Google Drive
                </span>
              </label>
            )}

            {driveUploadMsg && (
              <div
                className={`p-3 rounded-xl text-[10px] font-bold border flex items-center gap-2 ${
                  driveUploadStatus === "error"
                    ? "bg-red-50 border-red-100 text-red-600"
                    : driveUploadStatus === "done"
                      ? "bg-green-50 border-green-100 text-green-600"
                      : "bg-blue-50 border-blue-100 text-blue-600"
                }`}
              >
                {driveUploadStatus === "uploading" && (
                  <Loader2 className="w-3 h-3 animate-spin" />
                )}
                {driveUploadStatus === "done" && (
                  <CheckCircle2 className="w-3 h-3" />
                )}
                {driveUploadMsg}
              </div>
            )}

            {isParsing && (
              <div className="p-3 rounded-xl text-[10px] font-bold border bg-blue-50 border-blue-100 text-blue-600 flex items-center gap-2">
                <Loader2 className="w-3 h-3 animate-spin" />
                Parsing {parseProgress.current}/{parseProgress.total} files...
              </div>
            )}
          </div>

          {/* Actions */}
          <div className="p-6 pt-4 flex gap-3">
            <Button
              onClick={handleProcess}
              disabled={
                selectedFiles.length === 0 ||
                isParsing ||
                driveUploadStatus === "uploading"
              }
              className="flex-1 rounded-xl text-xs font-bold"
            >
              {isParsing ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin mr-2" />
                  Processing...
                </>
              ) : saveToDrive ? (
                "Upload & Save to Drive"
              ) : (
                `Upload ${selectedFiles.length || ""} File${selectedFiles.length !== 1 ? "s" : ""}`
              )}
            </Button>
            <Button
              onClick={handleClose}
              variant="outline"
              className="rounded-xl text-xs"
              disabled={isParsing || driveUploadStatus === "uploading"}
            >
              Cancel
            </Button>
          </div>
        </div>
      </div>

      <FolderPickerModal
        isOpen={folderPickerOpen}
        onClose={() => {
          setFolderPickerOpen(false);
          // Still process files even if folder picker is cancelled
          onUploadFiles(selectedFiles);
          setSelectedFiles([]);
          setSaveToDrive(false);
          onClose();
        }}
        onSelect={handleFolderSelected}
      />
    </>
  );
};
