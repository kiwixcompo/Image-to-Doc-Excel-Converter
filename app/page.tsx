'use client';

import { useState, useRef } from 'react';
import { GoogleGenAI, Type } from '@google/genai';
import { UploadCloud, FileText, Table as TableIcon, Loader2, Download, CheckCircle2, AlertCircle, X, Plus, FileArchive } from 'lucide-react';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, PageOrientation } from 'docx';
import * as xlsx from 'xlsx';
import JSZip from 'jszip';

interface FileWithPreview {
  file: File;
  preview: string;
}

export default function Home() {
  const [files, setFiles] = useState<FileWithPreview[]>([]);
  const [format, setFormat] = useState<'docx' | 'xlsx'>('docx');
  const [isConverting, setIsConverting] = useState(false);
  const [convertingIndex, setConvertingIndex] = useState(0);
  const [resultUrl, setResultUrl] = useState<string | null>(null);
  const [resultFileName, setResultFileName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFiles = (newFiles: File[]) => {
    const validFiles = newFiles.filter(f => f.type.startsWith('image/'));
    if (validFiles.length !== newFiles.length) {
      setError('Some files were ignored because they are not images.');
    } else {
      setError(null);
    }
    
    setFiles(prev => {
      const combined = [...prev, ...validFiles.map(f => ({ file: f, preview: URL.createObjectURL(f) }))];
      if (combined.length > 20) {
        setError('Maximum 20 images allowed. List has been truncated.');
        return combined.slice(0, 20);
      }
      return combined;
    });
    setResultUrl(null);
    setResultFileName(null);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      handleFiles(Array.from(e.target.files));
    }
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
      handleFiles(Array.from(e.dataTransfer.files));
    }
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
    setResultUrl(null);
    setResultFileName(null);
  };

  const handleConvert = async () => {
    if (files.length === 0) return;

    setIsConverting(true);
    setError(null);
    setResultUrl(null);
    setResultFileName(null);
    setConvertingIndex(0);

    try {
      const apiKey = process.env.NEXT_PUBLIC_GEMINI_API_KEY;
      if (!apiKey) {
        throw new Error('Gemini API key is not configured.');
      }

      const ai = new GoogleGenAI({ apiKey });
      const generatedFiles: { name: string, blob: Blob }[] = [];

      for (let i = 0; i < files.length; i++) {
        setConvertingIndex(i + 1);
        
        // Add a delay between requests for batch processing to respect free tier rate limits (15 RPM for Flash)
        if (i > 0) {
          await new Promise(resolve => setTimeout(resolve, 4000));
        }

        const fileObj = files[i].file;
        const base64Data = await fileToBase64(fileObj);

        const response = await ai.models.generateContent({
          model: 'gemini-3-flash-preview',
          contents: [
            {
              inlineData: {
                data: base64Data,
                mimeType: fileObj.type,
              },
            },
            {
              text: 'Extract the document heading and the tabular data from this image. ' +
                    '1. "filename": Create a short, safe filename (max 5 words, no extension) based on the main heading. ' +
                    '2. "headings": An array of strings representing the text headings/titles at the top of the document, line by line. ' +
                    '3. "table": A 2D array of strings representing the tabular data. The first array must be the column headers. ' +
                    'CRITICAL: Ensure table headers and cell contents are clean, single-line strings where possible. Do NOT add artificial line breaks or vertical stacking in the text. Capture all columns and rows accurately.',
            },
          ],
          config: {
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.OBJECT,
              properties: {
                filename: { type: Type.STRING },
                headings: { 
                  type: Type.ARRAY, 
                  items: { type: Type.STRING } 
                },
                table: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING }
                  }
                }
              },
              required: ["filename", "headings", "table"]
            }
          }
        });

        const jsonText = response.text || '{}';
        let data: { filename: string, headings: string[], table: string[][] };
        try {
          data = JSON.parse(jsonText);
        } catch (e) {
          throw new Error(`Failed to parse data from image ${i + 1}.`);
        }

        const safeFilename = (data.filename || `document_${i+1}`).replace(/[^a-z0-9]/gi, '_').toLowerCase();

        if (format === 'docx') {
          const headingParagraphs = (data.headings || []).map((h, idx) => new Paragraph({
            children: [new TextRun({ text: h, bold: true, size: idx === 0 ? 28 : 24 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 }
          }));

          const tableRows = (data.table || []).map((row, rowIndex) => new TableRow({
            children: row.map(cell => new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ 
                  // Force replacement of any newlines with spaces to prevent vertical stacking
                  text: cell ? cell.toString().replace(/\n/g, ' ') : "", 
                  bold: rowIndex === 0, 
                  size: 20 
                })],
                alignment: rowIndex === 0 ? AlignmentType.CENTER : AlignmentType.LEFT
              })],
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }))
          }));

          const docTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
              insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            },
            rows: tableRows,
          });

          const doc = new Document({
            sections: [{
              properties: {
                page: {
                  size: {
                    orientation: PageOrientation.LANDSCAPE,
                  },
                  margin: {
                    top: 720, // 0.5 inch margins to maximize horizontal space
                    right: 720,
                    bottom: 720,
                    left: 720,
                  }
                }
              },
              children: [...headingParagraphs, docTable],
            }],
          });

          const blob = await Packer.toBlob(doc);
          generatedFiles.push({ name: `${safeFilename}.docx`, blob });

        } else if (format === 'xlsx') {
          const aoa: any[][] = [];
          (data.headings || []).forEach(h => aoa.push([h]));
          if (data.headings && data.headings.length > 0) {
            aoa.push([]); // Spacer
          }
          
          (data.table || []).forEach(row => {
            // Remove newlines from excel cells too
            const cleanRow = row.map(cell => cell ? cell.toString().replace(/\n/g, ' ') : "");
            aoa.push(cleanRow);
          });

          const worksheet = xlsx.utils.aoa_to_sheet(aoa);
          
          // Auto-size columns roughly based on content length
          const colWidths = [];
          for (let c = 0; c < (data.table[0] || []).length; c++) {
            let maxLen = 10;
            for (let r = 0; r < data.table.length; r++) {
              const val = data.table[r][c];
              if (val && val.toString().length > maxLen) {
                maxLen = val.toString().length;
              }
            }
            colWidths.push({ wch: Math.min(maxLen + 2, 50) }); // Cap at 50 chars width
          }
          worksheet['!cols'] = colWidths;

          const workbook = xlsx.utils.book_new();
          xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
          
          const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'array' });
          const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          generatedFiles.push({ name: `${safeFilename}.xlsx`, blob });
        }
      }

      if (generatedFiles.length === 1) {
        // Single file: Download directly without zipping
        const url = URL.createObjectURL(generatedFiles[0].blob);
        setResultUrl(url);
        setResultFileName(generatedFiles[0].name);
      } else {
        // Multiple files: Zip them
        const zip = new JSZip();
        generatedFiles.forEach(f => zip.file(f.name, f.blob));
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        setResultUrl(url);
        setResultFileName('converted_documents.zip');
      }

    } catch (err: any) {
      console.error(err);
      setError(err.message || 'An error occurred during conversion.');
    } finally {
      setIsConverting(false);
      setConvertingIndex(0);
    }
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        if (typeof reader.result === 'string') {
          const base64 = reader.result.split(',')[1];
          resolve(base64);
        } else {
          reject(new Error('Failed to convert file to base64'));
        }
      };
      reader.onerror = error => reject(error);
    });
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] text-gray-900 font-sans p-6 md:p-12">
      <div className="max-w-5xl mx-auto space-y-8">
        <header className="space-y-2 text-center md:text-left">
          <h1 className="text-4xl font-light tracking-tight">Batch Document Converter</h1>
          <p className="text-gray-500 text-lg">Convert up to 20 images into perfectly formatted Word or Excel files.</p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Upload Section */}
          <div className="lg:col-span-2 space-y-6">
            <div
              className={`border-2 border-dashed rounded-2xl p-6 transition-colors
                ${files.length > 0 ? 'border-gray-300 bg-white' : 'border-gray-300 hover:border-gray-400 bg-white cursor-pointer text-center py-16'}`}
              onDragOver={(e) => e.preventDefault()}
              onDrop={handleDrop}
              onClick={() => files.length === 0 && fileInputRef.current?.click()}
            >
              <input
                type="file"
                multiple
                ref={fileInputRef}
                className="hidden"
                accept="image/*"
                onChange={handleFileChange}
              />
              
              {files.length > 0 ? (
                <div className="space-y-4">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="font-medium text-gray-700">Selected Images ({files.length}/20)</h3>
                    {files.length < 20 && (
                      <button 
                        onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}
                        className="text-sm text-blue-600 hover:text-blue-700 font-medium flex items-center gap-1"
                      >
                        <Plus className="w-4 h-4" /> Add More
                      </button>
                    )}
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4">
                    {files.map((f, i) => (
                      <div key={i} className="relative group rounded-xl overflow-hidden border border-gray-200 aspect-square bg-gray-50">
                        {/* eslint-disable-next-line @next/next/no-img-element */}
                        <img src={f.preview} alt="preview" className="object-cover w-full h-full" />
                        <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-start justify-end p-2">
                          <button 
                            onClick={(e) => { e.stopPropagation(); removeFile(i); }}
                            className="bg-white text-red-600 rounded-full p-1.5 hover:bg-red-50 transition-colors shadow-sm"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    ))}
                    {files.length < 20 && (
                      <div 
                        onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}
                        className="border-2 border-dashed border-gray-200 rounded-xl flex flex-col items-center justify-center cursor-pointer hover:bg-gray-50 hover:border-gray-300 aspect-square transition-colors text-gray-400 hover:text-gray-500"
                      >
                        <Plus className="w-8 h-8 mb-2" />
                        <span className="text-xs font-medium">Add Image</span>
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div className="space-y-4">
                  <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mx-auto">
                    <UploadCloud className="w-8 h-8 text-blue-500" />
                  </div>
                  <div>
                    <p className="font-medium text-lg">Click to upload or drag and drop</p>
                    <p className="text-sm text-gray-500 mt-1">Upload up to 20 images (SVG, PNG, JPG, GIF)</p>
                  </div>
                </div>
              )}
            </div>

            {error && (
              <div className="flex items-center gap-2 text-red-600 bg-red-50 p-4 rounded-xl text-sm border border-red-100">
                <AlertCircle className="w-5 h-5 flex-shrink-0" />
                <p>{error}</p>
              </div>
            )}
          </div>

          {/* Controls Section */}
          <div className="space-y-6">
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 space-y-6">
              <div className="space-y-3">
                <label className="text-sm font-semibold text-gray-700 uppercase tracking-wider">Output Format</label>
                <div className="grid grid-cols-1 gap-3">
                  <button
                    onClick={() => setFormat('docx')}
                    className={`flex items-center gap-3 p-4 rounded-xl border transition-all text-left
                      ${format === 'docx' 
                        ? 'border-blue-600 bg-blue-50 text-blue-800 ring-1 ring-blue-600 shadow-sm' 
                        : 'border-gray-200 hover:border-gray-300 text-gray-600 hover:bg-gray-50'}`}
                  >
                    <div className={`p-2 rounded-lg ${format === 'docx' ? 'bg-blue-100 text-blue-600' : 'bg-gray-100 text-gray-500'}`}>
                      <FileText className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="font-medium">Word Document</div>
                      <div className="text-xs opacity-80">.docx format</div>
                    </div>
                  </button>
                  <button
                    onClick={() => setFormat('xlsx')}
                    className={`flex items-center gap-3 p-4 rounded-xl border transition-all text-left
                      ${format === 'xlsx' 
                        ? 'border-green-600 bg-green-50 text-green-800 ring-1 ring-green-600 shadow-sm' 
                        : 'border-gray-200 hover:border-gray-300 text-gray-600 hover:bg-gray-50'}`}
                  >
                    <div className={`p-2 rounded-lg ${format === 'xlsx' ? 'bg-green-100 text-green-600' : 'bg-gray-100 text-gray-500'}`}>
                      <TableIcon className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="font-medium">Excel Spreadsheet</div>
                      <div className="text-xs opacity-80">.xlsx format</div>
                    </div>
                  </button>
                </div>
              </div>

              <button
                onClick={handleConvert}
                disabled={files.length === 0 || isConverting}
                className={`w-full py-4 rounded-xl font-medium flex items-center justify-center gap-2 transition-all
                  ${files.length === 0 
                    ? 'bg-gray-100 text-gray-400 cursor-not-allowed' 
                    : isConverting
                      ? 'bg-gray-900 text-white opacity-90 cursor-wait'
                      : 'bg-gray-900 text-white hover:bg-gray-800 shadow-md hover:shadow-lg'}`}
              >
                {isConverting ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    Converting {convertingIndex} of {files.length}...
                  </>
                ) : (
                  <>
                    {files.length === 1 ? <FileText className="w-5 h-5" /> : <FileArchive className="w-5 h-5" />}
                    {files.length === 1 ? 'Convert Document' : `Convert & Zip (${files.length})`}
                  </>
                )}
              </button>
            </div>

            {/* Result Section */}
            {resultUrl && (
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-green-200 bg-green-50/50 space-y-4 animate-in fade-in slide-in-from-bottom-4">
                <div className="flex items-center gap-3 text-green-700">
                  <CheckCircle2 className="w-6 h-6" />
                  <h3 className="font-medium text-lg">Ready to Download</h3>
                </div>
                <p className="text-sm text-gray-600">
                  Successfully converted {files.length} {files.length === 1 ? 'image' : 'images'}.
                </p>
                <a
                  href={resultUrl}
                  download={resultFileName || 'converted_document'}
                  className="w-full py-3 px-4 bg-white border border-gray-200 rounded-xl font-medium flex items-center justify-center gap-2 hover:bg-gray-50 transition-colors text-gray-900 shadow-sm"
                >
                  <Download className="w-5 h-5" />
                  {files.length === 1 ? `Download ${format.toUpperCase()}` : 'Download ZIP'}
                </a>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
