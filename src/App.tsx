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
  Wrench
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

export default function App() {
  const [bills, setBills] = useState<BillEntry[]>([]);
  const [activeScans, setActiveScans] = useState(0);
  const [queue, setQueue] = useState<File[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [previewImage, setPreviewImage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const galleryInputRef = useRef<HTMLInputElement>(null);

  const MAX_CONCURRENT = 3; // Process 3 bills at a time for speed

  // Load from local storage on mount
  useEffect(() => {
    const savedBills = localStorage.getItem('scanned_bills');
    
    if (savedBills) {
      try {
        setBills(JSON.parse(savedBills));
      } catch (e) {
        console.error("Failed to load bills", e);
      }
    }
  }, []);

  // Save to local storage on change
  useEffect(() => {
    localStorage.setItem('scanned_bills', JSON.stringify(bills));
  }, [bills]);

  // Queue Processor
  useEffect(() => {
    if (queue.length > 0 && activeScans < MAX_CONCURRENT) {
      const nextFile = queue[0];
      setQueue(prev => prev.slice(1));
      processFile(nextFile);
    }
  }, [queue, activeScans]);

  // Initialize Gemini
  const ai = new GoogleGenAI({ apiKey: import.meta.env.VITE_GEMINI_API_KEY || '' });

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setQueue(prev => [...prev, ...files]);
    if (fileInputRef.current) fileInputRef.current.value = '';
    if (galleryInputRef.current) galleryInputRef.current.value = '';
  };

  const processFile = async (file: File) => {
    setActiveScans(prev => prev + 1);
    setError(null);

    try {
      // Compress image for faster upload
      const compressedBase64 = await compressImage(file);
      
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            parts: [
              { text: "FAST SCAN: Extract bill JSON {srNo, type, invoiceNo, date, amount, isReadable, errorReason}. Languages: Arabic, Hindi, English. Return JSON only." },
              { inlineData: { data: compressedBase64.split(',')[1], mimeType: "image/jpeg" } }
            ]
          }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: BILL_SCHEMA,
          thinkingConfig: { thinkingLevel: ThinkingLevel.LOW }
        }
      });

      const result = JSON.parse(response.text || '{}');
      
      if (result.isReadable === false) {
        setError(`Skipped: ${result.errorReason || "Unreadable image"}`);
      } else {
        const newBill: BillEntry = {
          id: crypto.randomUUID(),
          srNo: result.srNo || '',
          type: result.type || 'Other',
          invoiceNo: result.invoiceNo || 'N/A',
          date: result.date || new Date().toISOString().split('T')[0],
          amount: result.amount || 0,
        };
        setBills(prev => [newBill, ...prev]);
      }
    } catch (err: any) {
      console.error("Extraction error:", err);
      if (err?.message?.includes('429') || err?.status === 429) {
        setError(`Rate limit hit. Re-queuing...`);
        setQueue(prev => [...prev, file]); // Put back in queue
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
    if (bills.length === 0) return;

    const worksheet = XLSX.utils.json_to_sheet(bills.map(({ id, ...rest }) => ({
      'Sr No': rest.srNo,
      'Type': rest.type,
      'Invoice No': rest.invoiceNo,
      'Date': rest.date,
      'Amount': rest.amount
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Bills");
    
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Bills_Export_${date}.xlsx`);
  };

  const deleteBill = (id: string) => {
    setBills(prev => prev.filter(b => b.id !== id));
  };

  const clearAll = () => {
    if (window.confirm("Are you sure you want to clear all entries?")) {
      setBills([]);
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
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans pb-20">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10 px-4 py-4">
        <div className="max-w-4xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-black p-2 rounded-xl">
              <Camera className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">Bill Scanner</h1>
              <p className="text-xs text-gray-500 font-medium uppercase tracking-wider">AI-Powered Extraction</p>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            {bills.length > 0 && (
              <button 
                onClick={exportToExcel}
                className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg text-sm font-semibold transition-all shadow-sm active:scale-95"
              >
                <FileSpreadsheet className="w-4 h-4" />
                Export Excel
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-4xl mx-auto p-4 space-y-6">
        {/* Action Card */}
        <section className="bg-white rounded-3xl p-8 border border-gray-100 shadow-sm text-center space-y-4">
          <div className="mx-auto w-20 h-20 bg-gray-50 rounded-full flex items-center justify-center border-2 border-dashed border-gray-200">
            {activeScans > 0 ? (
              <Loader2 className="w-10 h-10 text-black animate-spin" />
            ) : (
              <ImageIcon className="w-10 h-10 text-gray-400" />
            )}
          </div>
          
          <div className="space-y-2">
            <h2 className="text-2xl font-bold">Scan your bill</h2>
            <p className="text-gray-500 max-w-xs mx-auto">
              Take a photo of your fuel, food, or tool bill to automatically extract details.
            </p>
          </div>

          <div className="flex flex-col sm:flex-row gap-3 pt-4">
            <button 
              onClick={() => fileInputRef.current?.click()}
              disabled={activeScans > 0 || queue.length > 0}
              className="flex-1 bg-black hover:bg-gray-800 text-white py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              <Camera className="w-6 h-6" />
              Capture
            </button>
            
            <button 
              onClick={() => galleryInputRef.current?.click()}
              disabled={activeScans > 0 || queue.length > 0}
              className="flex-1 bg-white border-2 border-black hover:bg-gray-50 text-black py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              <ImageIcon className="w-6 h-6" />
              Gallery
            </button>

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

          {queue.length > 0 && (
            <div className="bg-blue-50 text-blue-700 p-3 rounded-xl text-sm font-bold flex items-center justify-between">
              <span>Queue: {queue.length} left • Active: {activeScans}</span>
              <Loader2 className="w-4 h-4 animate-spin" />
            </div>
          )}

          {error && (
            <div className="bg-red-50 text-red-600 p-3 rounded-xl flex items-center gap-2 text-sm font-medium">
              <AlertCircle className="w-4 h-4" />
              {error}
            </div>
          )}
        </section>

        {/* Processing Preview */}
        <AnimatePresence>
          {activeScans > 0 && (
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl p-4 border border-gray-100 shadow-sm flex items-center gap-4"
            >
              <div className="bg-black p-4 rounded-xl">
                <Loader2 className="w-6 h-6 text-white animate-spin" />
              </div>
              <div className="flex-1">
                <p className="font-bold">Turbo Scanning {activeScans} bills...</p>
                <div className="w-full bg-gray-100 h-1.5 rounded-full mt-2 overflow-hidden">
                  <motion.div 
                    initial={{ x: '-100%' }}
                    animate={{ x: '100%' }}
                    transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
                    className="w-1/2 h-full bg-black rounded-full"
                  />
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* List Section */}
        <div className="space-y-4">
          <div className="flex items-center justify-between px-2">
            <h3 className="font-bold text-lg">Scanned Entries ({bills.length})</h3>
            {bills.length > 0 && (
              <button 
                onClick={clearAll}
                className="text-xs font-bold text-red-500 hover:text-red-600 uppercase tracking-wider"
              >
                Clear All
              </button>
            )}
          </div>

          {bills.length === 0 && activeScans === 0 ? (
            <div className="bg-gray-50 border-2 border-dashed border-gray-200 rounded-3xl py-12 text-center">
              <Receipt className="w-12 h-12 text-gray-300 mx-auto mb-3" />
              <p className="text-gray-400 font-medium">No bills scanned yet</p>
            </div>
          ) : (
            <div className="space-y-3">
              <AnimatePresence mode="popLayout">
                {bills.map((bill) => (
                  <motion.div
                    key={bill.id}
                    layout
                    initial={{ opacity: 0, x: -20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm flex items-center gap-4 hover:border-gray-300 transition-colors group"
                  >
                    <div className="bg-gray-50 p-3 rounded-xl">
                      {getIcon(bill.type)}
                    </div>
                    
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="font-bold truncate">{bill.invoiceNo}</span>
                        <span className="text-[10px] bg-gray-100 px-1.5 py-0.5 rounded font-bold text-gray-500 uppercase tracking-wider">
                          {bill.type}
                        </span>
                      </div>
                      <div className="flex items-center gap-3 text-xs text-gray-500 mt-0.5">
                        <span>{bill.date}</span>
                        {bill.srNo && <span>• Sr: {bill.srNo}</span>}
                      </div>
                    </div>

                    <div className="text-right">
                      <p className="font-black text-lg">₹{bill.amount.toLocaleString()}</p>
                      <button 
                        onClick={() => deleteBill(bill.id)}
                        className="text-gray-300 hover:text-red-500 transition-colors p-1"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </motion.div>
                ))}
              </AnimatePresence>
            </div>
          )}
        </div>
      </main>

      {/* Floating Stats */}
      {bills.length > 0 && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 w-[calc(100%-2rem)] max-w-md">
          <div className="bg-black text-white p-4 rounded-2xl shadow-2xl flex items-center justify-between">
            <div>
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Total Amount</p>
              <p className="text-2xl font-black">₹{bills.reduce((acc, b) => acc + b.amount, 0).toLocaleString()}</p>
            </div>
            <div className="text-right">
              <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Items</p>
              <p className="text-xl font-bold">{bills.length}</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
