import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { Upload, FileSpreadsheet, Settings, Play, CheckCircle2, AlertCircle, X, Clock } from 'lucide-react';

export default function DeadstockProcessor() {
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [branchFiles, setBranchFiles] = useState<File[]>([]);
  const [targetFrekuensi, setTargetFrekuensi] = useState<number>(12);
  const [targetFMos, setTargetFMos] = useState<number>(3);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [previewData, setPreviewData] = useState<{ headers: string[], rows: any[][] } | null>(null);
  const [processTime, setProcessTime] = useState<string | null>(null);

  const handleMasterUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setMasterFile(e.target.files[0]);
      setError(null);
      setSuccess(null);
    }
  };

  const handleBranchUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files);
      setBranchFiles((prev) => [...prev, ...newFiles]);
      setError(null);
      setSuccess(null);
    }
  };

  const removeBranchFile = (index: number) => {
    setBranchFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const readFileAsText = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  };

  const parseFileToJSON = async (file: File) => {
    const isText = file.name.toLowerCase().endsWith('.txt') || file.name.toLowerCase().endsWith('.csv');
    let wb;
    if (isText) {
      const text = await readFileAsText(file);
      wb = XLSX.read(text, { type: 'string', raw: true });
    } else {
      const buffer = await readFileAsArrayBuffer(file);
      wb = XLSX.read(buffer, { 
        type: 'array',
        cellDates: false,
        cellStyles: false,
        cellNF: false,
        cellText: false
      });
    }
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
  };

  const processExcel = async () => {
    if (!masterFile) {
      setError("Silakan unggah file Master terlebih dahulu.");
      return;
    }
    if (branchFiles.length === 0) {
      setError("Silakan unggah minimal 1 file Cabang Tujuan.");
      return;
    }

    const startTime = performance.now();
    setIsProcessing(true);
    setProgress(0);
    setProgressText("Membaca Master File...");
    setError(null);
    setSuccess(null);
    setProcessTime(null);

    try {
      // 1. Baca Master File menggunakan XLSX (Sangat Cepat & Support TXT)
      const masterData = await parseFileToJSON(masterFile);

      if (masterData.length < 2) throw new Error("File Master kosong atau tidak valid.");

      // Otomatis tambahkan header FREKUENSI (AJ) dan AMOUNT (AK)
      const masterHeaders = [...masterData[0].slice(0, 35), 'FREKUENSI', 'AMOUNT'];
      const masterRows = masterData.slice(1);

      // 2. Baca Branch Files secara berurutan (Sekuensial) untuk mencegah Out of Memory
      const branchDataMap: Record<string, Map<string, { stock: number, avgDemand: number, frekuensi: number }>> = {};
      
      for (let i = 0; i < branchFiles.length; i++) {
        const file = branchFiles[i];
        
        // Update Progress & Beri jeda untuk Garbage Collection browser
        setProgress(10 + Math.floor((i / branchFiles.length) * 20)); // Progress 10% s/d 30%
        setProgressText(`Membaca File Cabang (${i + 1}/${branchFiles.length})...`);
        await new Promise(r => setTimeout(r, 50)); 

        const data = await parseFileToJSON(file);
        
        const branchMap = new Map();
        const branchName = file.name.replace(/\.[^/.]+$/, "");

        for (let j = 1; j < data.length; j++) {
          const row = data[j];
          const partCode = row[0]?.toString()?.trim(); // Kolom A
          if (partCode) {
            // Hitung frekuensi secara otomatis: Kolom U (20) s/d AF (31) >= 1
            let calculatedFreek = 0;
            for(let k = 20; k <= 31; k++) {
              if (Number(row[k]) >= 1) calculatedFreek++;
            }

            branchMap.set(partCode, {
              stock: Number(row[16]) || 0,      // Kolom Q (Index 16)
              frekuensi: calculatedFreek,       // Menggunakan hitungan otomatis
              avgDemand: Number(row[34]) || 0,  // Kolom AI (Index 34) - AVG DEMAND
            });
          }
        }
        branchDataMap[branchName] = branchMap;
        // Data file mentah akan otomatis dibersihkan oleh Garbage Collector setelah loop ini selesai
      }

      // 3. Siapkan Output Header
      const headerRow1 = [...masterHeaders.slice(0, 37)];
      const headerRow2 = [...masterHeaders.slice(0, 37)];

      branchFiles.forEach((file) => {
        const branchName = file.name.replace(/\.[^/.]+$/, "").toUpperCase();
        headerRow1.push(branchName, '', '', '', '', '', '');
        headerRow2.push('STOCK', 'AVG DEMAND', 'F MOS', 'FREEK', 'LAST COST', 'AMOUNT', 'AMBIL');
      });

      setProgress(30);
      setProgressText("Menghitung Alokasi...");
      await new Promise(r => setTimeout(r, 10));

      const escapeCSV = (val: any) => {
        if (val === null || val === undefined) return '';
        const str = String(val);
        if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
          return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
      };

      // 4. Proses Baris (Logika Inti) langsung ke format CSV Teks (Sangat Cepat & Anti-Crash)
      const csvLines: string[] = [];
      // Tambahkan BOM (Byte Order Mark) agar Excel membaca karakter spesial dengan benar
      csvLines.push('\uFEFF');
      
      // Masukkan Header
      csvLines.push(headerRow1.map(escapeCSV).join(',') + '\n');
      csvLines.push(headerRow2.map(escapeCSV).join(',') + '\n');

      const previewRows: any[][] = [];
      const totalRows = masterRows.length;
      let rowsProcessedCounter = 0;

      for (let i = 0; i < totalRows; i++) {
        const mRow = masterRows[i];
        
        // Update UI secara berkala (setiap 500 baris) agar tidak freeze
        if (i > 0 && i % 500 === 0) {
          setProgress(30 + Math.floor((i / totalRows) * 40)); // 30% sampai 70%
          setProgressText(`Memproses Baris ${i.toLocaleString()} dari ${totalRows.toLocaleString()}...`);
          await new Promise(r => setTimeout(r, 0));
        }

        const partCode = mRow[0]?.toString()?.trim();
        if (!partCode) continue;

        const category = mRow[6]?.toString()?.trim()?.toUpperCase(); // Kolom G (Index 6)
        const lastCost = Number(mRow[3]) || 0; // Kolom D (Index 3)
        const originalMasterStock = Number(mRow[16]) || 0; // Kolom Q (Index 16)
        let remainingMasterStock = originalMasterStock;

        // Hitung frekuensi Master secara otomatis: Kolom U (20) s/d AF (31) >= 1
        let masterFreek = 0;
        for(let j = 20; j <= 31; j++) {
          if (Number(mRow[j]) >= 1) masterFreek++;
        }
        const masterAmount = lastCost * originalMasterStock;

        // Ambil 35 kolom asli dari Master, lalu tambahkan hasil hitung FREKUENSI & AMOUNT
        const baseRow = Array.from({ length: 35 }, (_, i) => mRow[i] !== undefined ? mRow[i] : '');
        const rowData = [...baseRow, masterFreek, masterAmount];

        for (const file of branchFiles) {
          const branchName = file.name.replace(/\.[^/.]+$/, "");
          const bData = branchDataMap[branchName].get(partCode) || { stock: 0, avgDemand: 0, frekuensi: 0 };
          
          const { stock: bStock, avgDemand: bAvgDemand, frekuensi: bFreek } = bData;
          
          // 1. Hitung FMOS Cabang saat ini (Sebelum Alokasi)
          // Jika AVG Demand 0, set F MOS ke Infinity agar barang mati tidak dikirimi deadstock lagi
          let fMosOriginal = bAvgDemand > 0 ? bStock / bAvgDemand : Infinity;
          
          let ambil = 0;
          
          // 2. Filter: Ambil jika FMOS asli Cabang <= Target (Artinya cabang sedang butuh stok)
          if (category === 'D' && 
              bFreek >= targetFrekuensi && 
              fMosOriginal <= targetFMos && 
              remainingMasterStock > 0) {
            
            // Konsep "Winner Takes All": Cabang pertama yang butuh ambil SEMUA stock master
            ambil = remainingMasterStock;
            remainingMasterStock = 0;
          }

          // 3. Hitung FMOS Nyata Setelah Proses (Stok Lama + Ambil)
          let fMosAfter = bAvgDemand > 0 ? (ambil + bStock) / bAvgDemand : 0;

          // Push data dengan urutan sesuai prototype: STOCK, AVG DEMAND, F MOS, FREEK, LAST COST, AMOUNT, AMBIL
          rowData.push(
            bStock,
            bAvgDemand,
            Number(fMosAfter.toFixed(2)),
            bFreek,
            lastCost,
            lastCost * ambil, // Amount = Harga * Barang yang diambil
            ambil
          );
        }

        // Langsung konversi baris ke string CSV dan simpan ke array (Memory O(1) per baris)
        csvLines.push(rowData.map(escapeCSV).join(',') + '\n');
        
        if (previewRows.length < 10) previewRows.push(rowData);
        
        rowsProcessedCounter++;
      }

      setProgress(90);
      setProgressText("Membangun File CSV...");
      await new Promise(r => setTimeout(r, 10));

      // 5. Buat Blob CSV murni (Melewati batasan V8 JS Engine)
      const blob = new Blob(csvLines, { type: 'text/csv;charset=utf-8;' });
      
      setProgress(100);
      setProgressText("Selesai!");
      saveAs(blob, `Hasil_Alokasi_Deadstock_${new Date().getTime()}.csv`);
      
      const duration = ((performance.now() - startTime) / 1000).toFixed(2);
      setProcessTime(duration);
      setSuccess(`Berhasil! ${rowsProcessedCounter} baris diproses dalam ${duration} detik.`);
      setPreviewData({ headers: headerRow2, rows: previewRows });

    } catch (err: any) {
      console.error(err);
      setError(err.message || "Terjadi kesalahan saat memproses data.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-900">
      <div className="max-w-6xl mx-auto space-y-6">
        
        {/* Header */}
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-indigo-600 text-white rounded-xl shadow-indigo-200 shadow-lg">
              <FileSpreadsheet size={32} />
            </div>
            <div>
              <h1 className="text-2xl font-bold tracking-tight">Alokator Otomatis Deadstock</h1>
              <p className="text-slate-500 text-sm">Optimasi stok mati antar cabang dengan logika cerdas & cepat.</p>
            </div>
          </div>
          {processTime && (
            <div className="flex items-center gap-2 px-4 py-2 bg-emerald-50 text-emerald-700 rounded-full text-sm font-semibold border border-emerald-100">
              <Clock size={16} />
              Waktu Proses: {processTime}s
            </div>
          )}
        </div>

        {/* Configuration */}
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
          <h2 className="text-lg font-semibold flex items-center gap-2 mb-4 text-slate-700">
            <Settings size={20} className="text-indigo-500" />
            Parameter Syarat Alokasi
          </h2>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <label className="block text-sm font-semibold text-slate-600">Minimal Target Frekuensi</label>
              <input 
                type="number" 
                value={targetFrekuensi}
                onChange={(e) => setTargetFrekuensi(Number(e.target.value))}
                className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all bg-slate-50"
              />
              <p className="text-xs text-slate-400 italic">Cabang harus memiliki frekuensi penjualan &ge; angka ini.</p>
            </div>
            <div className="space-y-2">
              <label className="block text-sm font-semibold text-slate-600">Maksimal Target F MOS</label>
              <input 
                type="number" 
                value={targetFMos}
                onChange={(e) => setTargetFMos(Number(e.target.value))}
                className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all bg-slate-50"
              />
              <p className="text-xs text-slate-400 italic">Cabang harus memiliki F MOS &le; angka ini.</p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Master Upload */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col">
            <h2 className="text-lg font-semibold mb-4 text-slate-700">1. Unggah Data Master</h2>
            <div className="flex-1 border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center hover:border-indigo-400 hover:bg-indigo-50/30 transition-all relative group">
              <input 
                type="file" 
                accept=".xlsx, .xls, .txt, .csv" 
                onChange={handleMasterUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <Upload className="mx-auto text-slate-300 group-hover:text-indigo-400 mb-4 transition-colors" size={48} />
              {masterFile ? (
                <div className="space-y-2">
                  <div className="text-emerald-600 font-bold flex items-center justify-center gap-2 bg-emerald-50 py-2 px-4 rounded-lg inline-flex">
                    <CheckCircle2 size={18} />
                    {masterFile.name}
                  </div>
                  <p className="text-xs text-slate-400">Klik untuk mengganti file</p>
                </div>
              ) : (
                <div className="space-y-1">
                  <p className="font-bold text-slate-600">Klik atau seret file Master ke sini</p>
                  <p className="text-sm text-slate-400">Format yang didukung: .xlsx, .xls, .txt, .csv</p>
                </div>
              )}
            </div>
          </div>

          {/* Branch Upload */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <h2 className="text-lg font-semibold mb-4 text-slate-700">2. Unggah Data Cabang</h2>
            <div className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center hover:border-indigo-400 hover:bg-indigo-50/30 transition-all relative mb-4 group">
              <input 
                type="file" 
                accept=".xlsx, .xls, .txt, .csv" 
                multiple
                onChange={handleBranchUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <Upload className="mx-auto text-slate-300 group-hover:text-indigo-400 mb-4 transition-colors" size={48} />
              <p className="font-bold text-slate-600">Unggah File Cabang (Bisa banyak sekaligus)</p>
              <p className="text-sm text-slate-400">Urutan unggah menentukan prioritas alokasi</p>
            </div>

            {/* Branch List */}
            {branchFiles.length > 0 && (
              <div className="space-y-2 max-h-48 overflow-y-auto pr-2 custom-scrollbar">
                {branchFiles.map((file, index) => (
                  <div key={index} className="flex items-center justify-between bg-slate-50 p-3 rounded-xl border border-slate-100 group">
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="bg-indigo-600 text-white w-6 h-6 rounded-lg flex items-center justify-center text-xs font-bold shrink-0 shadow-sm">
                        {index + 1}
                      </div>
                      <span className="text-sm font-medium truncate text-slate-700">{file.name}</span>
                    </div>
                    <button 
                      onClick={() => removeBranchFile(index)}
                      className="text-slate-300 hover:text-red-500 transition-colors p-1"
                    >
                      <X size={18} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Action & Status */}
        <div className="flex flex-col items-center gap-4 py-4">
          {error && (
            <div className="flex items-center gap-3 text-red-600 bg-red-50 px-6 py-4 rounded-2xl w-full border border-red-100 shadow-sm animate-shake">
              <AlertCircle size={24} className="shrink-0" />
              <p className="text-sm font-bold">{error}</p>
            </div>
          )}
          
          {success && (
            <div className="flex items-center gap-3 text-emerald-700 bg-emerald-50 px-6 py-4 rounded-2xl w-full border border-emerald-100 shadow-sm">
              <CheckCircle2 size={24} className="shrink-0" />
              <p className="text-sm font-bold">{success}</p>
            </div>
          )}

          <button
            onClick={processExcel}
            disabled={isProcessing}
            className={`group relative flex items-center gap-3 px-12 py-5 rounded-2xl font-black text-lg text-white shadow-xl transition-all overflow-hidden ${
              isProcessing 
                ? 'bg-slate-400 cursor-not-allowed' 
                : 'bg-indigo-600 hover:bg-indigo-700 hover:shadow-indigo-200 hover:-translate-y-1 active:translate-y-0'
            }`}
          >
            {isProcessing ? (
              <div className="w-6 h-6 border-4 border-white/30 border-t-white rounded-full animate-spin shrink-0" />
            ) : (
              <Play size={28} fill="currentColor" className="shrink-0" />
            )}
            <div className="flex flex-col items-start text-left min-w-[200px]">
              <span>{isProcessing ? progressText : 'PROSES & UNDUH EXCEL'}</span>
              {isProcessing && (
                <div className="w-full mt-1 bg-white/20 rounded-full h-1.5 overflow-hidden flex relative">
                  <div className="bg-white h-full transition-all duration-300" style={{ width: `${progress}%` }} />
                </div>
              )}
            </div>
            {!isProcessing && (
              <div className="absolute inset-0 bg-white/10 translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-500 skew-x-12" />
            )}
          </button>
        </div>

        {/* Preview Table */}
        {previewData && (
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 overflow-hidden animate-slideUp">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <div className="w-2 h-8 bg-indigo-600 rounded-full" />
                <h2 className="text-xl font-black text-slate-800">Review Hasil Samping (10 Baris)</h2>
              </div>
              <button 
                onClick={() => setPreviewData(null)}
                className="text-slate-400 hover:text-slate-600 font-bold text-sm flex items-center gap-1"
              >
                <X size={16} /> Tutup
              </button>
            </div>
            <div className="overflow-x-auto border border-slate-100 rounded-2xl shadow-inner">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-800 text-white font-bold">
                  <tr>
                    {previewData.headers.map((header, i) => (
                      <th key={i} className="px-4 py-4 whitespace-nowrap border-r border-slate-700 last:border-0 uppercase tracking-wider">
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {previewData.rows.map((row, i) => (
                    <tr key={i} className="hover:bg-indigo-50/30 transition-colors">
                      {row.map((cell, j) => {
                        const isAmbil = previewData.headers[j]?.includes('AMBIL');
                        return (
                          <td 
                            key={j} 
                            className={`px-4 py-3 whitespace-nowrap border-r border-slate-100 last:border-0 font-medium ${
                              isAmbil ? 'bg-yellow-100 text-yellow-900 font-black border-yellow-200' : 'text-slate-600'
                            }`}
                          >
                            {typeof cell === 'object' ? JSON.stringify(cell) : String(cell ?? '')}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="mt-4 flex items-center gap-2 text-slate-400 text-xs italic">
              <AlertCircle size={14} />
              <span>Hanya menampilkan 10 baris pertama untuk pengecekan cepat.</span>
            </div>
          </div>
        )}

      </div>

      <style>{`
        @keyframes shake {
          0%, 100% { transform: translateX(0); }
          25% { transform: translateX(-4px); }
          75% { transform: translateX(4px); }
        }
        @keyframes slideUp {
          from { opacity: 0; transform: translateY(20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-shake { animation: shake 0.2s ease-in-out 0s 2; }
        .animate-slideUp { animation: slideUp 0.4s ease-out; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #e2e8f0; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #cbd5e1; }
      `}</style>
    </div>
  );
}
