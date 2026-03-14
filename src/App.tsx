/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileDown, FileSpreadsheet, Loader2, CheckCircle2, AlertCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

const MONTHS = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

interface PayoutRecord {
  ID: string;
  Currency: string;
  Amount: number;
  Status: string;
  Description: string;
  Created: string;
  'Estimate Arrival': string;
}

interface BTRecord {
  'Payout ID': string;
  Type: string;
  'Amount Payout': number;
  Currency: string;
  Net: number;
  Created: string;
}

export default function App() {
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const parseMonthOnly = (val: any): string | null => {
    if (!val) return null;
    
    let date: Date;
    if (val instanceof Date) {
      date = val;
    } else {
      // Try to parse string
      const dateStr = String(val);
      date = new Date(dateStr);
      
      // If native parsing fails, try manual extraction for formats like "Nov-28-2025"
      if (isNaN(date.getTime())) {
        const match = dateStr.match(/([A-Za-z]{3})/);
        if (match) {
          const m = match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
          return MONTHS.includes(m) ? m : null;
        }
        return null;
      }
    }

    return MONTHS[date.getMonth()];
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setError(null);
    setSuccess(false);

    try {
      const data = await file.arrayBuffer();
      // Use cellDates: true to handle Excel date objects correctly
      const workbook = XLSX.read(data, { cellDates: true });

      const payoutsSheet = workbook.Sheets['Payouts'];
      const btSheet = workbook.Sheets['Balance Transactions'] || workbook.Sheets['Balance-Transactions'];

      if (!payoutsSheet || !btSheet) {
        throw new Error('File must contain "Payouts" and "Balance Transactions" sheets.');
      }

      const payoutsData: PayoutRecord[] = XLSX.utils.sheet_to_json(payoutsSheet);
      const btData: BTRecord[] = XLSX.utils.sheet_to_json(btSheet);

      const newWorkbook = XLSX.utils.book_new();

      // 1. Add original Payouts
      XLSX.utils.book_append_sheet(newWorkbook, payoutsSheet, 'Payouts');

      // 2. Group Payouts by Month
      const payoutsByMonth: Record<string, PayoutRecord[]> = {};
      payoutsData.forEach(record => {
        const month = parseMonthOnly(record.Created);
        if (month) {
          if (!payoutsByMonth[month]) payoutsByMonth[month] = [];
          payoutsByMonth[month].push(record);
        }
      });

      // Add sheets in calendar order
      MONTHS.forEach(month => {
        if (payoutsByMonth[month]) {
          const sheetData = [...payoutsByMonth[month]];
          const totalAmount = sheetData.reduce((sum, r) => sum + (Number(r.Amount) || 0), 0);
          
          // @ts-ignore
          sheetData.push({
            ID: 'TOTAL',
            Amount: totalAmount
          });

          const ws = XLSX.utils.json_to_sheet(sheetData);
          XLSX.utils.book_append_sheet(newWorkbook, ws, `Payout-${month}`);
        }
      });

      // 3. Add original BT
      XLSX.utils.book_append_sheet(newWorkbook, btSheet, 'Balance-Transactions');

      // 4. Group BT by Month
      const btByMonth: Record<string, BTRecord[]> = {};
      btData.forEach(record => {
        const month = parseMonthOnly(record.Created);
        if (month) {
          if (!btByMonth[month]) btByMonth[month] = [];
          btByMonth[month].push(record);
        }
      });

      MONTHS.forEach(month => {
        if (btByMonth[month]) {
          const sheetData = [...btByMonth[month]];
          const totalNet = sheetData.reduce((sum, r) => sum + (Number(r.Net) || 0), 0);
          const totalAmountPayout = sheetData.reduce((sum, r) => sum + (Number(r['Amount Payout']) || 0), 0);

          // @ts-ignore
          sheetData.push({
            'Payout ID': 'TOTAL',
            'Amount Payout': totalAmountPayout,
            Net: totalNet
          });

          const ws = XLSX.utils.json_to_sheet(sheetData);
          XLSX.utils.book_append_sheet(newWorkbook, ws, `BT-${month}`);
        }
      });

      // Generate and download
      XLSX.writeFile(newWorkbook, `Processed_Payouts_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSuccess(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred while processing the file.');
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] text-[#1a1a1a] font-sans p-6 md:p-12">
      <div className="max-w-3xl mx-auto">
        {/* Header */}
        <header className="mb-12">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-black rounded-lg">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <h1 className="text-2xl font-semibold tracking-tight">Excel Payout Processor</h1>
          </div>
          <p className="text-[#9e9e9e] text-lg">
            Upload your Stripe payout report to generate monthly summaries with automated totals.
          </p>
        </header>

        {/* Main Card */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white rounded-[24px] shadow-sm border border-black/5 overflow-hidden"
        >
          <div className="p-8 md:p-12">
            <div 
              className={`
                relative border-2 border-dashed rounded-2xl p-12 transition-all duration-200
                flex flex-col items-center justify-center text-center
                ${isProcessing ? 'border-black/10 bg-black/5' : 'border-black/10 hover:border-black/30 hover:bg-black/[0.02] cursor-pointer'}
              `}
              onClick={() => !isProcessing && fileInputRef.current?.click()}
            >
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".xlsx, .xls"
                className="hidden"
              />

              <AnimatePresence mode="wait">
                {isProcessing ? (
                  <motion.div 
                    key="processing"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.9 }}
                    className="flex flex-col items-center"
                  >
                    <Loader2 className="w-12 h-12 text-black animate-spin mb-4" />
                    <p className="font-medium">Processing your file...</p>
                    <p className="text-sm text-[#9e9e9e] mt-1">This will only take a moment.</p>
                  </motion.div>
                ) : success ? (
                  <motion.div 
                    key="success"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.9 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-4">
                      <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <p className="font-medium text-emerald-600">Processing Complete!</p>
                    <p className="text-sm text-[#9e9e9e] mt-1">Your new file has been downloaded.</p>
                    <button 
                      onClick={(e) => {
                        e.stopPropagation();
                        setSuccess(false);
                      }}
                      className="mt-6 text-sm font-medium underline underline-offset-4 hover:text-black"
                    >
                      Upload another file
                    </button>
                  </motion.div>
                ) : (
                  <motion.div 
                    key="idle"
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.9 }}
                    className="flex flex-col items-center"
                  >
                    <div className="w-12 h-12 bg-black/5 rounded-full flex items-center justify-center mb-4">
                      <Upload className="w-6 h-6 text-black" />
                    </div>
                    <p className="font-medium text-lg">Click to upload or drag and drop</p>
                    <p className="text-sm text-[#9e9e9e] mt-1">Supports .xlsx and .xls files</p>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                className="mt-6 p-4 bg-red-50 border border-red-100 rounded-xl flex items-start gap-3 text-red-600"
              >
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <div className="text-sm">
                  <p className="font-semibold">Error</p>
                  <p>{error}</p>
                </div>
              </motion.div>
            )}
          </div>

          <div className="bg-[#fafafa] border-t border-black/5 p-6 px-8 md:px-12 flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="flex -space-x-2">
                {[1, 2, 3].map((i) => (
                  <div key={i} className="w-8 h-8 rounded-full border-2 border-white bg-[#e5e5e5] flex items-center justify-center text-[10px] font-bold">
                    {i}
                  </div>
                ))}
              </div>
              <p className="text-xs text-[#9e9e9e] font-medium uppercase tracking-wider">
                3-Step Automated Workflow
              </p>
            </div>
            <div className="flex items-center gap-6 text-xs text-[#9e9e9e]">
              <div className="flex items-center gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                <span>Monthly Grouping</span>
              </div>
              <div className="flex items-center gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                <span>Auto-Summation</span>
              </div>
            </div>
          </div>
        </motion.div>

        {/* Instructions */}
        <div className="mt-12 grid grid-cols-1 md:grid-cols-2 gap-8">
          <section>
            <h3 className="text-sm font-semibold uppercase tracking-wider text-[#9e9e9e] mb-4">Expected Format</h3>
            <ul className="space-y-3 text-sm text-[#666]">
              <li className="flex items-start gap-2">
                <span className="text-black font-bold">•</span>
                <span>Sheet <strong>"Payouts"</strong> with columns: Amount, Created</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-black font-bold">•</span>
                <span>Sheet <strong>"Balance Transactions"</strong> with columns: Net, Amount Payout, Created</span>
              </li>
            </ul>
          </section>
          <section>
            <h3 className="text-sm font-semibold uppercase tracking-wider text-[#9e9e9e] mb-4">Output Details</h3>
            <ul className="space-y-3 text-sm text-[#666]">
              <li className="flex items-start gap-2">
                <span className="text-black font-bold">•</span>
                <span>Monthly sheets (e.g., Payout-Nov, BT-Nov)</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-black font-bold">•</span>
                <span>Total rows automatically calculated at the bottom</span>
              </li>
            </ul>
          </section>
        </div>

        <footer className="mt-20 pt-8 border-t border-black/5 text-center">
          <p className="text-xs text-[#9e9e9e]">
            &copy; {new Date().getFullYear()} Excel Payout Processor. Built for precision.
          </p>
        </footer>
      </div>
    </div>
  );
}
