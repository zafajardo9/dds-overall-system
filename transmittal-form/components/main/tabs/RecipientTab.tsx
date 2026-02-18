"use client";

import React, { useRef, useEffect } from "react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { RecipientInfo } from "@/types";

const ExpandingTextarea = ({
  value,
  onChange,
  placeholder,
  className = "",
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  className?: string;
}) => {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
      textareaRef.current.style.height = `${textareaRef.current.scrollHeight}px`;
    }
  }, [value]);
  return (
    <Textarea
      ref={textareaRef}
      className={`overflow-hidden resize-none min-h-[44px] ${className}`}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      rows={1}
    />
  );
};

interface RecipientTabProps {
  recipient: RecipientInfo;
  onUpdateField: (section: "recipient", field: string, value: any) => void;
}

export const RecipientTab: React.FC<RecipientTabProps> = ({
  recipient,
  onUpdateField,
}) => {
  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <h2 className="text-xs font-black text-slate-800 uppercase tracking-[0.2em] mb-4">
        Recipient Dossier
      </h2>
      <div className="space-y-4">
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            To
          </Label>
          <Input
            value={recipient.to}
            onChange={(e) => onUpdateField("recipient", "to", e.target.value)}
          />
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Organization
          </Label>
          <Input
            value={recipient.company}
            onChange={(e) =>
              onUpdateField("recipient", "company", e.target.value)
            }
          />
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Attention
          </Label>
          <Input
            value={recipient.attention}
            onChange={(e) =>
              onUpdateField("recipient", "attention", e.target.value)
            }
          />
        </div>
        <div className="space-y-1">
          <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
            Full Address
          </Label>
          <ExpandingTextarea
            value={recipient.address}
            onChange={(v) => onUpdateField("recipient", "address", v)}
            placeholder="Full address..."
          />
        </div>
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Contact No.
            </Label>
            <Input
              value={recipient.contactNumber}
              onChange={(e) =>
                onUpdateField("recipient", "contactNumber", e.target.value)
              }
            />
          </div>
          <div className="space-y-1">
            <Label className="text-[9px] font-black text-muted-foreground uppercase tracking-widest ml-2">
              Email
            </Label>
            <Input
              value={recipient.email}
              onChange={(e) =>
                onUpdateField("recipient", "email", e.target.value)
              }
            />
          </div>
        </div>
      </div>
    </div>
  );
};
