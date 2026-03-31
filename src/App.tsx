/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";
import * as XLSX from 'xlsx';
import { 
  Camera, 
  FileSpreadsheet, 
  Trash2, 
  Loader2, 
  Plus, 
  Image as ImageIcon,
  CheckCircle2,
  AlertCircle,
  ChevronRight,
  Receipt,
  Fuel,
  Utensils,
  Wrench,
  UserCircle,
  CreditCard,
  Globe
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for Tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface BillEntry {
  id: string;
  srNo: string;
  type: string;
  invoiceNo: string;
  date: string;
  amount: number;
  rawResponse?: string;
}

interface IqamaEntry {
  id: string;
  name: string;
  iqamaNo: string;
  nationality: string;
  expiryDate?: string;
}

const BILL_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    srNo: { type: Type.STRING, description: "Serial number if present on the bill" },
    type: { type: Type.STRING, description: "Category of the bill (e.g., Fuel, Food, Tool, Medical, Rent, etc.)" },
    invoiceNo: { type: Type.STRING, description: "Invoice, receipt, or bill number" },
    date: { type: Type.STRING, description: "Date of the bill in YYYY-MM-DD format" },
    amount: { type: Type.NUMBER, description: "Total amount on the bill as a number" },
    isReadable: { type: Type.BOOLEAN, description: "True if the bill is clear and readable, false otherwise" },
    errorReason: { type: Type.STRING, description: "If not readable, provide a brief reason like 'Image blurry', 'Not a bill', 'Text too small', etc." },
  },
  required: ["isReadable"],
};

const IQAMA_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    name: { type: Type.STRING, description: "Full name of the person" },
    iqamaNo: { type: Type.STRING, description: "Iqama number (convert Arabic numerals to standard digits)" },
    nationality: { type: Type.STRING, description: "Nationality of the person" },
    isReadable: { type: Type.BOOLEAN, description: "True if the Iqama is clear and readable, false otherwise" },
    errorReason: { type: Type.STRING, description: "If not readable, provide a brief reason" },
  },
  required: ["isReadable"],
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'bills' | 'iqama'>('bills');
  const [bills, setBills] = useState<BillEntry[]>([]);
  const [iqamas, setIqamas] = useState<IqamaEntry[]>([]);
  const [activeScans, setActiveScans] = useState(0);
  const [queue, setQueue] = useState<{ file: File; type: 'bills' | 'iqama' }[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isRetrying, setIsRetrying] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const galleryInputRef = useRef<HTMLInputElement>(null);

  const MAX_CONCURRENT = 2; // Low concurrency to respect free tier limits

  // Load from local storage on mount
  useEffect(() => {
    const savedBills = localStorage.getItem('scanned_bills');
    const savedIqamas = localStorage.getItem('scanned_iqamas');
    
    if (savedBills) {
      try {
        setBills(JSON.parse(savedBills));
      } catch (e) {
        console.error("Failed to load bills", e);
      }
    }

    if (savedIqamas) {
      try {
        setIqamas(JSON.parse(savedIqamas));
      } catch (e) {
        console.error("Failed to load iqamas", e);
      }
    }
  }, []);

  // Save to local storage on change
  useEffect(() => {
    localStorage.setItem('scanned_bills', JSON.stringify(bills));
  }, [bills]);

  useEffect(() => {
    localStorage.setItem('scanned_iqamas', JSON.stringify(iqamas));
  }, [iqamas]);

  // Queue Processor
  useEffect(() => {
    if (queue.length > 0 && activeScans < MAX_CONCURRENT) {
      const nextItem = queue[0];
      setQueue(prev => prev.slice(1));
      processFile(nextItem.file, nextItem.type);
    }
  }, [queue, activeScans]);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setQueue(prev => [...prev, ...files.map(file => ({ file, type: activeTab }))]);
    if (fileInputRef.current) fileInputRef.current.value = '';
    if (galleryInputRef.current) galleryInputRef.current.value = '';
  };

  const processFile = async (file: File, scanType: 'bills' | 'iqama') => {
    setActiveScans(prev => prev + 1);
    setError(null);

    try {
      // Use the default free Gemini API key
      const apiKey = (window as any).ENV?.GEMINI_API_KEY || import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY;
      
      if (!apiKey) {
        setError("AI configuration missing. Please add 'GEMINI_API_KEY' to your Secrets.");
        setActiveScans(prev => prev - 1);
        return;
      }

      const ai = new GoogleGenAI({ apiKey });

      // Compress image for faster upload
      const compressedBase64 = await compressImage(file);
      
      const prompt = scanType === 'bills' 
        ? "FAST SCAN: Extract bill JSON {srNo, type, invoiceNo, date, amount, isReadable, errorReason}. Languages: Arabic, Hindi, English. Return JSON only."
        : "IQAMA SCAN: Extract Iqama JSON {name, iqamaNo, nationality, isReadable, errorReason}. IMPORTANT: Convert Arabic numerals (١٢٣) to standard digits (123) for iqamaNo. Return JSON only.";

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            parts: [
              { text: prompt },
              { inlineData: { data: compressedBase64.split(',')[1], mimeType: "image/jpeg" } }
            ]
          }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: scanType === 'bills' ? BILL_SCHEMA : IQAMA_SCHEMA,
          thinkingConfig: { thinkingLevel: ThinkingLevel.LOW }
        }
      });

      const result = JSON.parse(response.text || '{}');
      
      if (result.isReadable === false) {
        setError(`Skipped: ${result.errorReason || "Unreadable image"}`);
      } else {
        if (scanType === 'bills') {
          const newBill: BillEntry = {
            id: crypto.randomUUID(),
            srNo: result.srNo || '',
            type: result.type || 'Other',
            invoiceNo: result.invoiceNo || 'N/A',
            date: result.date || new Date().toISOString().split('T')[0],
            amount: result.amount || 0,
          };
          setBills(prev => [newBill, ...prev]);
        } else {
          const newIqama: IqamaEntry = {
            id: crypto.randomUUID(),
            name: result.name || 'N/A',
            iqamaNo: result.iqamaNo || 'N/A',
            nationality: result.nationality || 'N/A',
          };
          setIqamas(prev => [newIqama, ...prev]);
        }
      }
    } catch (err: any) {
      console.error("Extraction error:", err);
      const isRateLimit = err?.message?.includes('429') || err?.status === 429 || err?.message?.includes('RESOURCE_EXHAUSTED');
      
      if (isRateLimit) {
        setIsRetrying(true);
        setError(`Daily limit reached (Free AI). Retrying in 15s...`);
        // Wait 15 seconds before re-queuing
        await new Promise(resolve => setTimeout(resolve, 15000));
        setQueue(prev => [...prev, { file, type: scanType }]);
        setIsRetrying(false);
      } else {
        setError("Error processing file. Skipping.");
      }
    } finally {
      setActiveScans(prev => prev - 1);
    }
  };

  const compressImage = (file: File): Promise<string> => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (e) => {
        const img = new Image();
        img.src = e.target?.result as string;
        img.onload = () => {
          const canvas = document.createElement('canvas');
          let width = img.width;
          let height = img.height;
          
          // Max 1200px for fast processing but clear OCR
          const MAX_SIZE = 1200;
          if (width > height) {
            if (width > MAX_SIZE) {
              height *= MAX_SIZE / width;
              width = MAX_SIZE;
            }
          } else {
            if (height > MAX_SIZE) {
              width *= MAX_SIZE / height;
              height = MAX_SIZE;
            }
          }
          
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx?.drawImage(img, 0, 0, width, height);
          resolve(canvas.toDataURL('image/jpeg', 0.7)); // 70% quality is enough for OCR
        };
      };
    });
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = error => reject(error);
    });
  };

  const exportToExcel = () => {
    const workbook = XLSX.utils.book_new();
    const date = new Date().toISOString().split('T')[0];

    if (activeTab === 'bills') {
      if (bills.length === 0) return;
      const worksheet = XLSX.utils.json_to_sheet(bills.map(({ id, ...rest }) => ({
        'Sr No': rest.srNo,
        'Type': rest.type,
        'Invoice No': rest.invoiceNo,
        'Date': rest.date,
        'Amount': rest.amount
      })));
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bills");
      XLSX.writeFile(workbook, `Bills_Export_${date}.xlsx`);
    } else {
      if (iqamas.length === 0) return;
      const worksheet = XLSX.utils.json_to_sheet(iqamas.map(({ id, ...rest }) => ({
        'Name': rest.name,
        'Iqama No': rest.iqamaNo,
        'Nationality': rest.nationality
      })));
      XLSX.utils.book_append_sheet(workbook, worksheet, "Iqamas");
      XLSX.writeFile(workbook, `Iqamas_Export_${date}.xlsx`);
    }
  };

  const deleteBill = (id: string) => {
    setBills(prev => prev.filter(b => b.id !== id));
  };

  const deleteIqama = (id: string) => {
    setIqamas(prev => prev.filter(i => i.id !== id));
  };

  const clearAll = () => {
    if (window.confirm(`Are you sure you want to clear all ${activeTab === 'bills' ? 'bills' : 'iqamas'}?`)) {
      if (activeTab === 'bills') setBills([]);
      else setIqamas([]);
    }
  };

  const getIcon = (type: string) => {
    switch (type) {
      case 'Fuel': return <Fuel className="w-5 h-5 text-blue-500" />;
      case 'Food': return <Utensils className="w-5 h-5 text-orange-500" />;
      case 'Tool': return <Wrench className="w-5 h-5 text-gray-500" />;
      default: return <Receipt className="w-5 h-5 text-green-500" />;
    }
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1E] font-sans selection:bg-zinc-950 selection:text-white">
      {/* Background Gradient */}
      <div className="fixed inset-0 bg-[radial-gradient(circle_at_top_right,_var(--tw-gradient-stops))] from-zinc-100 via-transparent to-transparent pointer-events-none" />
      
      {/* Header */}
      <header className="bg-white/70 backdrop-blur-xl border-b border-zinc-200/50 sticky top-0 z-50 px-6 py-4">
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-zinc-950 p-2.5 rounded-2xl shadow-2xl shadow-zinc-200/50">
              <Camera className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-black tracking-tight text-zinc-900">Smart Scanner</h1>
              <div className="flex items-center gap-2">
                <span className="flex h-1.5 w-1.5 rounded-full bg-emerald-500 animate-pulse" />
                <p className="text-[9px] text-zinc-400 font-bold uppercase tracking-[0.25em]">Free AI Extraction Engine</p>
              </div>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            {(activeTab === 'bills' ? bills.length > 0 : iqamas.length > 0) && (
              <motion.button 
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                onClick={exportToExcel}
                className="flex items-center gap-2 bg-zinc-900 hover:bg-zinc-800 text-white px-5 py-2.5 rounded-xl text-sm font-bold transition-all shadow-md active:scale-95"
              >
                <FileSpreadsheet className="w-4 h-4" />
                Export Data
              </motion.button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto p-6 space-y-8">
        {/* Dashboard Summary */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white p-8 rounded-[2.5rem] border border-zinc-200/60 shadow-sm space-y-1 group hover:shadow-xl hover:shadow-zinc-200/30 transition-all duration-500">
            <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">Total Scanned</p>
            <div className="flex items-baseline gap-2">
              <span className="text-5xl font-black tracking-tighter group-hover:scale-105 transition-transform origin-left duration-500">{activeTab === 'bills' ? bills.length : iqamas.length}</span>
              <span className="text-zinc-400 font-bold text-sm">entries</span>
            </div>
          </div>
          
          {activeTab === 'bills' ? (
            <div className="bg-zinc-950 p-8 rounded-[2.5rem] text-white shadow-2xl shadow-zinc-300/50 space-y-1 md:col-span-2 relative overflow-hidden group">
              <div className="absolute top-0 right-0 -mr-12 -mt-12 w-48 h-48 bg-emerald-500/10 rounded-full blur-3xl group-hover:bg-emerald-500/20 transition-colors duration-700" />
              <p className="text-[10px] font-black text-zinc-500 uppercase tracking-widest relative z-10">Aggregate Value</p>
              <div className="flex items-baseline gap-2 relative z-10">
                <span className="text-5xl font-black tracking-tighter text-emerald-400">
                  SAR {bills.reduce((acc, b) => acc + b.amount, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                </span>
              </div>
            </div>
          ) : (
            <div className="bg-zinc-950 p-8 rounded-[2.5rem] text-white shadow-2xl shadow-zinc-300/50 space-y-1 md:col-span-2 relative overflow-hidden group">
              <div className="absolute top-0 right-0 -mr-12 -mt-12 w-48 h-48 bg-blue-500/10 rounded-full blur-3xl group-hover:bg-blue-500/20 transition-colors duration-700" />
              <p className="text-[10px] font-black text-zinc-500 uppercase tracking-widest relative z-10">Nationalities Tracked</p>
              <div className="flex items-baseline gap-2 relative z-10">
                <span className="text-5xl font-black tracking-tighter text-blue-400">
                  {new Set(iqamas.map(i => i.nationality)).size}
                </span>
                <span className="text-zinc-500 font-bold text-sm">unique origins</span>
              </div>
            </div>
          )}
        </div>

        {/* Tabs Switcher */}
        <div className="flex bg-zinc-200/50 p-1.5 rounded-[1.5rem] max-w-md mx-auto">
          <button 
            onClick={() => setActiveTab('bills')}
            className={cn(
              "flex-1 py-3 rounded-2xl font-bold text-sm flex items-center justify-center gap-2 transition-all duration-300",
              activeTab === 'bills' ? "bg-white text-zinc-950 shadow-md scale-[1.02]" : "text-zinc-500 hover:text-zinc-800"
            )}
          >
            <Receipt className="w-4 h-4" />
            Bill Ledger
          </button>
          <button 
            onClick={() => setActiveTab('iqama')}
            className={cn(
              "flex-1 py-3 rounded-2xl font-bold text-sm flex items-center justify-center gap-2 transition-all duration-300",
              activeTab === 'iqama' ? "bg-white text-zinc-950 shadow-md scale-[1.02]" : "text-zinc-500 hover:text-zinc-800"
            )}
          >
            <CreditCard className="w-4 h-4" />
            Iqama Registry
          </button>
        </div>

        {/* Dropzone Area */}
        <section 
          className={cn(
            "relative group overflow-hidden bg-white rounded-[2.5rem] p-12 border-2 border-dashed transition-all duration-500",
            activeScans > 0 ? "border-zinc-950 bg-zinc-50" : "border-zinc-200 hover:border-zinc-400 hover:bg-zinc-50/50"
          )}
          onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add('border-zinc-950', 'bg-zinc-50'); }}
          onDragLeave={(e) => { e.preventDefault(); e.currentTarget.classList.remove('border-zinc-950', 'bg-zinc-50'); }}
          onDrop={(e) => {
            e.preventDefault();
            e.currentTarget.classList.remove('border-zinc-950', 'bg-zinc-50');
            const files = Array.from(e.dataTransfer.files);
            if (files.length > 0) {
              setQueue(prev => [...prev, ...files.map(file => ({ file, type: activeTab }))]);
            }
          }}
        >
          <div className="relative z-10 flex flex-col items-center text-center space-y-6">
            <div className="w-24 h-24 bg-zinc-100 rounded-[2rem] flex items-center justify-center shadow-inner group-hover:scale-110 transition-transform duration-500">
              {activeScans > 0 ? (
                <Loader2 className="w-12 h-12 text-zinc-950 animate-spin" />
              ) : (
                <Plus className="w-12 h-12 text-zinc-400 group-hover:text-zinc-950 transition-colors" />
              )}
            </div>
            
            <div className="space-y-2">
              <h2 className="text-3xl font-black tracking-tight text-zinc-900">
                {activeTab === 'bills' ? 'Import Bills' : 'Import Iqamas'}
              </h2>
              <p className="text-zinc-500 font-medium max-w-sm mx-auto leading-relaxed">
                Drag and drop multiple images here or use the professional capture tools below.
              </p>
            </div>

            <div className="flex flex-wrap justify-center gap-4 w-full max-w-lg">
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="flex-1 min-w-[160px] bg-zinc-950 hover:bg-zinc-800 text-white py-4 px-6 rounded-2xl font-bold text-base flex items-center justify-center gap-3 transition-all shadow-xl shadow-zinc-200 active:scale-95"
              >
                <Camera className="w-5 h-5" />
                Live Capture
              </button>
              
              <button 
                onClick={() => galleryInputRef.current?.click()}
                className="flex-1 min-w-[160px] bg-white border-2 border-zinc-950 hover:bg-zinc-50 text-zinc-950 py-4 px-6 rounded-2xl font-bold text-base flex items-center justify-center gap-3 transition-all active:scale-95"
              >
                <ImageIcon className="w-5 h-5" />
                Media Library
              </button>
            </div>

            <input 
              type="file" 
              ref={fileInputRef}
              onChange={handleFileSelect}
              accept="image/*"
              capture="environment"
              multiple
              className="hidden"
            />
            <input 
              type="file" 
              ref={galleryInputRef}
              onChange={handleFileSelect}
              accept="image/*"
              multiple
              className="hidden"
            />
          </div>

          {/* Background Decorative Elements */}
          <div className="absolute top-0 right-0 -mr-16 -mt-16 w-64 h-64 bg-zinc-100 rounded-full blur-3xl opacity-50" />
          <div className="absolute bottom-0 left-0 -ml-16 -mb-16 w-64 h-64 bg-zinc-100 rounded-full blur-3xl opacity-50" />
        </section>

        {/* Batch Progress */}
        <AnimatePresence>
          {(queue.length > 0 || activeScans > 0) && (
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-zinc-950 text-white rounded-[2rem] p-6 shadow-2xl space-y-4"
            >
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="bg-white/10 p-2 rounded-lg">
                    <Loader2 className="w-5 h-5 text-emerald-400 animate-spin" />
                  </div>
                  <div>
                    <p className="font-bold text-sm">Batch Processing Active</p>
                    <p className="text-[10px] text-zinc-400 font-bold uppercase tracking-widest">
                      {activeScans} processing • {queue.length} in queue
                    </p>
                  </div>
                </div>
                <span className="text-xs font-mono bg-white/10 px-3 py-1 rounded-full">
                  {Math.round((activeScans / (activeScans + queue.length)) * 100) || 0}%
                </span>
              </div>
              
              <div className="w-full bg-white/10 h-2 rounded-full overflow-hidden">
                <motion.div 
                  className="h-full bg-emerald-500 rounded-full"
                  initial={{ width: "0%" }}
                  animate={{ width: `${(1 - (queue.length / (activeScans + queue.length || 1))) * 100}%` }}
                />
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {error && (
          <motion.div 
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className={cn(
              "p-4 rounded-2xl flex items-center justify-between gap-3 text-sm font-bold border",
              error.includes('limit') ? "bg-amber-50 border-amber-100 text-amber-700" : "bg-red-50 border-red-100 text-red-600"
            )}
          >
            <div className="flex items-center gap-3">
              <AlertCircle className="w-5 h-5 shrink-0" />
              {error}
            </div>
          </motion.div>
        )}

        {/* List Section */}
        <div className="space-y-6">
          <div className="flex items-center justify-between px-2">
            <div className="flex items-center gap-3">
              <h3 className="font-black text-2xl tracking-tight">
                {activeTab === 'bills' ? 'Recent Ledger' : 'Identity Registry'}
              </h3>
              <span className="bg-zinc-200 text-zinc-600 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-widest">
                {activeTab === 'bills' ? bills.length : iqamas.length}
              </span>
            </div>
            {(activeTab === 'bills' ? bills.length > 0 : iqamas.length > 0) && (
              <button 
                onClick={clearAll}
                className="text-[10px] font-black text-zinc-400 hover:text-red-500 uppercase tracking-[0.2em] transition-colors"
              >
                Purge All Data
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 gap-4">
            <AnimatePresence mode="popLayout">
              {activeTab === 'bills' ? (
                bills.length === 0 && activeScans === 0 ? (
                  <motion.div 
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="bg-white/50 border-2 border-dashed border-zinc-200 rounded-[2.5rem] py-20 text-center"
                  >
                    <Receipt className="w-16 h-16 text-zinc-200 mx-auto mb-4" />
                    <p className="text-zinc-400 font-bold text-lg">Your ledger is currently empty</p>
                  </motion.div>
                ) : (
                  bills.map((bill) => (
                    <motion.div
                      key={bill.id}
                      layout
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95 }}
                      className="bg-white p-6 rounded-[2rem] border border-zinc-200 shadow-sm flex items-center gap-6 hover:shadow-xl hover:shadow-zinc-200/50 transition-all duration-300 group"
                    >
                      <div className="bg-zinc-50 p-4 rounded-2xl group-hover:bg-zinc-100 transition-colors">
                        {getIcon(bill.type)}
                      </div>
                      
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-3">
                          <span className="font-black text-lg truncate text-zinc-900">{bill.invoiceNo}</span>
                          <span className="text-[9px] bg-zinc-100 text-zinc-500 px-2 py-0.5 rounded-full font-black uppercase tracking-widest">
                            {bill.type}
                          </span>
                        </div>
                        <div className="flex items-center gap-4 text-xs font-bold text-zinc-400 mt-1">
                          <span className="flex items-center gap-1.5"><ChevronRight className="w-3 h-3" /> {bill.date}</span>
                          {bill.srNo && <span className="flex items-center gap-1.5"><ChevronRight className="w-3 h-3" /> SR: {bill.srNo}</span>}
                        </div>
                      </div>

                      <div className="text-right space-y-1">
                        <p className="font-black text-2xl tracking-tighter text-zinc-900">
                          <span className="text-sm font-bold text-zinc-400 mr-1">SAR</span>
                          {bill.amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                        </p>
                        <button 
                          onClick={() => deleteBill(bill.id)}
                          className="text-zinc-300 hover:text-red-500 transition-colors p-1"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                      </div>
                    </motion.div>
                  ))
                )
              ) : (
                iqamas.length === 0 && activeScans === 0 ? (
                  <motion.div 
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="bg-white/50 border-2 border-dashed border-zinc-200 rounded-[2.5rem] py-20 text-center"
                  >
                    <UserCircle className="w-16 h-16 text-zinc-200 mx-auto mb-4" />
                    <p className="text-zinc-400 font-bold text-lg">No identities registered</p>
                  </motion.div>
                ) : (
                  iqamas.map((iqama) => (
                    <motion.div
                      key={iqama.id}
                      layout
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, scale: 0.95 }}
                      className="bg-white p-6 rounded-[2rem] border border-zinc-200 shadow-sm flex items-center gap-6 hover:shadow-xl hover:shadow-zinc-200/50 transition-all duration-300 group"
                    >
                      <div className="bg-zinc-50 p-4 rounded-2xl group-hover:bg-zinc-100 transition-colors">
                        <UserCircle className="w-6 h-6 text-zinc-950" />
                      </div>
                      
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-3">
                          <span className="font-black text-lg truncate text-zinc-900 uppercase">{iqama.name}</span>
                        </div>
                        <div className="flex items-center gap-6 text-xs font-bold text-zinc-400 mt-1">
                          <span className="flex items-center gap-2">
                            <CreditCard className="w-4 h-4 text-zinc-300" />
                            <span className="font-mono text-zinc-600">{iqama.iqamaNo}</span>
                          </span>
                          <span className="flex items-center gap-2">
                            <Globe className="w-4 h-4 text-zinc-300" />
                            {iqama.nationality}
                          </span>
                        </div>
                      </div>

                      <div className="text-right">
                        <button 
                          onClick={() => deleteIqama(iqama.id)}
                          className="text-zinc-300 hover:text-red-500 transition-colors p-2"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                      </div>
                    </motion.div>
                  ))
                )
              )}
            </AnimatePresence>
          </div>
        </div>
      </main>

      {/* Floating Action Bar */}
      <AnimatePresence>
        {activeTab === 'bills' && bills.length > 0 && (
          <motion.div 
            initial={{ y: 100, x: "-50%" }}
            animate={{ y: 0, x: "-50%" }}
            exit={{ y: 100, x: "-50%" }}
            className="fixed bottom-8 left-1/2 w-[calc(100%-3rem)] max-w-xl z-40"
          >
            <div className="bg-zinc-950/90 backdrop-blur-2xl text-white p-6 rounded-[2.5rem] shadow-[0_32px_64px_-15px_rgba(0,0,0,0.5)] flex items-center justify-between border border-white/10">
              <div className="space-y-0.5">
                <p className="text-[10px] font-black text-zinc-500 uppercase tracking-[0.25em]">Consolidated Total</p>
                <p className="text-3xl font-black tracking-tighter">
                  <span className="text-sm font-bold text-zinc-500 mr-2">SAR</span>
                  {bills.reduce((acc, b) => acc + b.amount, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div className="h-12 w-[1px] bg-white/10 mx-6" />
              <div className="text-right">
                <p className="text-[10px] font-black text-zinc-500 uppercase tracking-[0.25em]">Record Count</p>
                <p className="text-2xl font-black tracking-tighter">{bills.length}</p>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
