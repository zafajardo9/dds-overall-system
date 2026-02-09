"use client";

import { useState, useEffect } from "react";
import { WifiOff } from "lucide-react";

export function OfflineAlert() {
  const [isOffline, setIsOffline] = useState(false);

  useEffect(() => {
    const goOffline = () => setIsOffline(true);
    const goOnline = () => setIsOffline(false);

    // Check initial state
    if (typeof navigator !== "undefined" && !navigator.onLine) {
      setIsOffline(true);
    }

    window.addEventListener("offline", goOffline);
    window.addEventListener("online", goOnline);
    return () => {
      window.removeEventListener("offline", goOffline);
      window.removeEventListener("online", goOnline);
    };
  }, []);

  if (!isOffline) return null;

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/40 backdrop-blur-sm">
      <div className="mx-4 w-full max-w-sm rounded-2xl bg-white p-8 text-center shadow-2xl">
        <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-red-100">
          <WifiOff className="h-8 w-8 text-red-600" />
        </div>
        <h2 className="mb-2 text-lg font-bold text-slate-900">
          No Internet Connection
        </h2>
        <p className="text-sm text-slate-500">
          You appear to be offline. Please check your network connection.
          This dialog will dismiss automatically when you reconnect.
        </p>
      </div>
    </div>
  );
}
