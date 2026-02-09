"use client";

import React from "react";

export const LoadingScreen: React.FC = () => {
  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-100 font-sans">
      <div className="bg-white border border-slate-200 rounded-3xl shadow-2xl p-10 text-center">
        <div className="w-12 h-12 border-4 border-brand-500 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
        <p className="text-xs font-black uppercase tracking-widest text-slate-500">
          Loading session
        </p>
      </div>
    </div>
  );
};
