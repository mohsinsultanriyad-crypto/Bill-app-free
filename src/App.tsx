/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { GoogleGenAI, Type } from "@google/genai";
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
  Settings,
  X,
  Edit2,
  Tag
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
  const [categories, setCategories] = useState<string[]>(['Fuel', 'Food', 'Tool', 'Other']);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [newCategoryName, setNewCategoryName] = useState('');
  const [editingCategory, setEditingCategory] = useState<{ index: number, name: string } | null>(null);
  const [queue, setQueue] = useState<File[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [previewImage, setPreviewImage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const galleryInputRef = useRef<HTMLInputElement>(null);

  // Load from local storage on mount
  useEffect(() => {
    const savedBills = localStorage.getItem('scanned_bills');
    const savedCategories = localStorage.getItem('bill_categories');
    
    if (savedBills) {
      try {
        setBills(JSON.parse(savedBills));
      } catch (e) {
        console.error("Failed to load bills", e);
      }
    }
    
    if (savedCategories) {
      try {
        setCategories(JSON.parse(savedCategories));
      } catch (e) {
        console.error("Failed to load categories", e);
      }
    }
  }, []);

  // Save to local storage on change
  useEffect(() => {
    localStorage.setItem('scanned_bills', JSON.stringify(bills));
  }, [bills]);

  useEffect(() => {
    localStorage.setItem('bill_categories', JSON.stringify(categories));
  }, [categories]);

  // Queue Processor
  useEffect(() => {
    if (queue.length > 0 && !isProcessing) {
      processFile(queue[0]);
    }
  }, [queue, isProcessing]);

  // Initialize Gemini
  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || '' });

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setQueue(prev => [...prev, ...files]);
    if (fileInputRef.current) fileInputRef.current.value = '';
    if (galleryInputRef.current) galleryInputRef.current.value = '';
  };

  const processFile = async (file: File) => {
    setIsProcessing(true);
    setError(null);

    // Show preview
    const reader = new FileReader();
    reader.onload = (event) => setPreviewImage(event.target?.result as string);
    reader.readAsDataURL(file);

    try {
      const base64Data = await fileToBase64(file);
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            parts: [
              { text: "Extract bill information from this image. Support multiple languages including Arabic, Hindi, and English. Be precise with the amount and date. If the image is blurry, not a bill, or unreadable, set isReadable to false and provide a reason in errorReason. If Sr No is not found, leave it empty. Return the data in JSON format." },
              { inlineData: { data: base64Data.split(',')[1], mimeType: file.type } }
            ]
          }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: BILL_SCHEMA,
        }
      });

      const result = JSON.parse(response.text || '{}');
      
      if (result.isReadable === false) {
        setError(`Skipped: ${result.errorReason || "Unreadable image"}`);
        setQueue(prev => prev.slice(1));
        setPreviewImage(null);
        setIsProcessing(false);
        return;
      }

      const newBill: BillEntry = {
        id: crypto.randomUUID(),
        srNo: result.srNo || '',
        type: 'Other', // Default to Other as requested
        invoiceNo: result.invoiceNo || 'N/A',
        date: result.date || new Date().toISOString().split('T')[0],
        amount: result.amount || 0,
      };

      setBills(prev => [newBill, ...prev]);
      setQueue(prev => prev.slice(1));
      setPreviewImage(null);
    } catch (err: any) {
      console.error("Extraction error:", err);
      if (err?.message?.includes('429') || err?.status === 429) {
        setError(`Rate limit hit. Waiting 30s to retry... (${queue.length} left)`);
        setTimeout(() => {
          setIsProcessing(false);
        }, 30000);
        return;
      } else {
        setError("Error processing file. Skipping to next.");
        setQueue(prev => prev.slice(1));
        setPreviewImage(null);
      }
    } finally {
      setIsProcessing(false);
    }
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

  const updateBillCategory = (id: string, newType: string) => {
    setBills(prev => prev.map(b => b.id === id ? { ...b, type: newType } : b));
  };

  const addCategory = () => {
    if (newCategoryName && !categories.includes(newCategoryName)) {
      setCategories(prev => [...prev, newCategoryName]);
      setNewCategoryName('');
    }
  };

  const deleteCategory = (category: string) => {
    if (category === 'Other') return; // Keep Other
    setCategories(prev => prev.filter(c => c !== category));
    // Reset bills with this category to Other
    setBills(prev => prev.map(b => b.type === category ? { ...b, type: 'Other' } : b));
  };

  const startEditingCategory = (index: number, name: string) => {
    setEditingCategory({ index, name });
  };

  const saveEditedCategory = () => {
    if (editingCategory && editingCategory.name) {
      const oldName = categories[editingCategory.index];
      const newName = editingCategory.name;
      
      setCategories(prev => {
        const next = [...prev];
        next[editingCategory.index] = newName;
        return next;
      });
      
      // Update all bills with this category
      setBills(prev => prev.map(b => b.type === oldName ? { ...b, type: newName } : b));
      setEditingCategory(null);
    }
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
            <button 
              onClick={() => setIsSettingsOpen(true)}
              className="p-2 hover:bg-gray-100 rounded-full transition-colors"
            >
              <Settings className="w-5 h-5 text-gray-600" />
            </button>
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
            {isProcessing ? (
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
              disabled={isProcessing || queue.length > 0}
              className="flex-1 bg-black hover:bg-gray-800 text-white py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              <Camera className="w-6 h-6" />
              Capture
            </button>
            
            <button 
              onClick={() => galleryInputRef.current?.click()}
              disabled={isProcessing || queue.length > 0}
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
              <span>Queue: {queue.length} bills remaining...</span>
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
          {previewImage && (
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl p-4 border border-gray-100 shadow-sm flex items-center gap-4"
            >
              <img src={previewImage} alt="Preview" className="w-20 h-20 object-cover rounded-xl border border-gray-200" />
              <div className="flex-1">
                <p className="font-bold">Analyzing image...</p>
                <div className="w-full bg-gray-100 h-1.5 rounded-full mt-2 overflow-hidden">
                  <motion.div 
                    initial={{ x: '-100%' }}
                    animate={{ x: '100%' }}
                    transition={{ repeat: Infinity, duration: 1.5, ease: "linear" }}
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

          {bills.length === 0 && !isProcessing ? (
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
                        <select 
                          value={bill.type}
                          onChange={(e) => updateBillCategory(bill.id, e.target.value)}
                          className="text-[10px] bg-gray-100 px-1.5 py-0.5 rounded font-bold text-gray-500 uppercase border-none focus:ring-0 cursor-pointer hover:bg-gray-200 transition-colors"
                        >
                          {categories.map(cat => (
                            <option key={cat} value={cat}>{cat}</option>
                          ))}
                        </select>
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

      {/* Settings Modal */}
      <AnimatePresence>
        {isSettingsOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsSettingsOpen(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-md bg-white rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="bg-gray-100 p-2 rounded-xl">
                    <Settings className="w-5 h-5" />
                  </div>
                  <h2 className="text-xl font-bold">Settings</h2>
                </div>
                <button 
                  onClick={() => setIsSettingsOpen(false)}
                  className="p-2 hover:bg-gray-100 rounded-full transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-6 space-y-6">
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <h3 className="font-bold flex items-center gap-2">
                      <Tag className="w-4 h-4" />
                      Categories
                    </h3>
                  </div>

                  <div className="space-y-2 max-h-60 overflow-y-auto pr-2">
                    {categories.map((cat, idx) => (
                      <div key={cat} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl group">
                        {editingCategory?.index === idx ? (
                          <input 
                            autoFocus
                            value={editingCategory.name}
                            onChange={(e) => setEditingCategory({ ...editingCategory, name: e.target.value })}
                            onBlur={saveEditedCategory}
                            onKeyDown={(e) => e.key === 'Enter' && saveEditedCategory()}
                            className="flex-1 bg-white border-none focus:ring-2 focus:ring-black rounded-lg px-2 py-1 text-sm font-medium"
                          />
                        ) : (
                          <span className="text-sm font-medium">{cat}</span>
                        )}
                        
                        <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                          <button 
                            onClick={() => startEditingCategory(idx, cat)}
                            className="p-1.5 hover:bg-gray-200 rounded-lg text-gray-500"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          {cat !== 'Other' && (
                            <button 
                              onClick={() => deleteCategory(cat)}
                              className="p-1.5 hover:bg-red-100 rounded-lg text-red-500"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="flex gap-2">
                    <input 
                      placeholder="New category name..."
                      value={newCategoryName}
                      onChange={(e) => setNewCategoryName(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && addCategory()}
                      className="flex-1 bg-gray-50 border-none focus:ring-2 focus:ring-black rounded-xl px-4 py-3 text-sm"
                    />
                    <button 
                      onClick={addCategory}
                      className="bg-black text-white p-3 rounded-xl hover:bg-gray-800 transition-colors"
                    >
                      <Plus className="w-5 h-5" />
                    </button>
                  </div>
                </div>
              </div>

              <div className="p-6 bg-gray-50 border-t border-gray-100">
                <button 
                  onClick={() => setIsSettingsOpen(false)}
                  className="w-full bg-black text-white py-3 rounded-xl font-bold hover:bg-gray-800 transition-colors"
                >
                  Done
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
