"use client";

import { useState, useRef } from "react";
import Link from "next/link";
import {
  Card,
  CardContent,
} from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import {
  ArrowLeft,
  UploadCloud,
  FileCheck2,
  FileSpreadsheet,
  Trash2,
  ShieldCheck,
  CheckCircle2,
  AlertCircle,
  Loader2,
} from "lucide-react";
import { convertPdfDocuments, type SupportedDocType } from "@/lib/api";

type DocType = SupportedDocType;

interface DocumentOption {
  id: DocType;
  title: string;
  code: string;
  description: string;
  details: string;
  badge: string;
}

const DOCUMENT_OPTIONS: DocumentOption[] = [
  {
    id: "sso",
    title: "ประกันสังคม",
    code: "SSO (สปส. 1-10)",
    description: "แบบยื่นรายการส่งเงินสมทบ กองทุนประกันสังคม",
    details: "รองรับเอกสาร สปส. 1-10 ส่วนที่ 1 และส่วนที่ 2",
    badge: "ประกันสังคม",
  },
  {
    id: "pnd1",
    title: "ภ.ง.ด. 1",
    code: "P.N.D. 1",
    description: "แบบยื่นรายการภาษีเงินได้หัก ณ ที่จ่าย (เงินเดือน/ค่าจ้าง)",
    details: "ภาษีเงินได้หัก ณ ที่จ่าย ตามมาตรา 40 (1) (2)",
    badge: "ภาษีเงินได้",
  },
  {
    id: "pnd3",
    title: "ภ.ง.ด. 3",
    code: "P.N.D. 3",
    description: "แบบยื่นรายการภาษีเงินได้หัก ณ ที่จ่าย (บุคคลธรรมดา)",
    details: "สำหรับผู้รับเงินที่มีหน้าที่เสียภาษีเงินได้บุคคลธรรมดา",
    badge: "บุคคลธรรมดา",
  },
  {
    id: "pnd53",
    title: "ภ.ง.ด. 53",
    code: "P.N.D. 53",
    description: "แบบยื่นรายการภาษีเงินได้หัก ณ ที่จ่าย (นิติบุคคล)",
    details: "สำหรับจ่ายเงินได้พึงประเมินให้แก่นิติบุคคล",
    badge: "นิติบุคคล",
  },
  {
    id: "statement",
    title: "SCB Bank Statement",
    code: "SCB Statement (เดินสะพัด)",
    description: "รายการเดินบัญชีธนาคารไทยพาณิชย์ (ประเภทบัญชีเดินสะพัด / Current Account)",
    details: "รองรับเอกสาร SCB Historical Statement สำหรับกระทบยอดรายการเดินบัญชี",
    badge: "SCB เดินสะพัด",
  },
];

export default function ConvertPage() {
  const [selectedDocType, setSelectedDocType] = useState<DocType>("sso");
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [dragActive, setDragActive] = useState<boolean>(false);
  const [uploadStatus, setUploadStatus] = useState<string | null>(null);
  const [isConverting, setIsConverting] = useState<boolean>(false);
  const [conversionError, setConversionError] = useState<string | null>(null);
  const [conversionSuccess, setConversionSuccess] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const currentDoc = DOCUMENT_OPTIONS.find((doc) => doc.id === selectedDocType);

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    const validPdfFiles = Array.from(files).filter(
      (file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")
    );

    if (validPdfFiles.length === 0 && files.length > 0) {
      setUploadStatus("กรุณาเลือกไฟล์ PDF เท่านั้น");
      return;
    }

    setSelectedFiles((prev) => [...prev, ...validPdfFiles]);
    setUploadStatus(null);
    setConversionError(null);
    setConversionSuccess(null);
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    handleFiles(e.dataTransfer.files);
  };

  const removeFile = (index: number) => {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const clearAllFiles = () => {
    setSelectedFiles([]);
    setConversionError(null);
    setConversionSuccess(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleConvert = async (format: "xlsx" | "csv") => {
    if (selectedFiles.length === 0) return;
    setConversionError(null);
    setConversionSuccess(null);

    setIsConverting(true);
    try {
      const result = await convertPdfDocuments(selectedFiles, selectedDocType, format);
      if (result.isZip) {
        setConversionSuccess(
          `แปลงเอกสารสำเร็จแล้ว ${selectedFiles.length} ไฟล์ รวมเป็นไฟล์ ZIP (${result.filename}) ดาวน์โหลดเรียบร้อย`
        );
      } else {
        setConversionSuccess(
          `แปลงเอกสารสำเร็จแล้ว (${result.filename}) ดาวน์โหลดไฟล์เรียบร้อย`
        );
      }
      setSelectedFiles([]);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : "เกิดข้อผิดพลาดในการแปลงไฟล์";
      setConversionError(message);
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <main className="min-h-screen bg-background text-foreground py-10 px-4 sm:px-6 lg:px-8">
      <div className="w-full max-w-4xl mx-auto space-y-8">
        {/* Header navigation & title */}
        <div>
          <Link
            href="/"
            className="inline-flex items-center gap-1.5 text-sm font-medium text-muted-foreground hover:text-foreground transition-colors mb-4"
          >
            <ArrowLeft className="size-4" />
            <span>กลับหน้าหลัก</span>
          </Link>
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
            <div>
              <h1 className="text-3xl font-extrabold tracking-tight">
                แปลงเอกสาร (PDF to CSV / XLSX)
              </h1>
              <p className="text-muted-foreground text-sm sm:text-base mt-1">
                เลือกลักษณะเอกสารที่ต้องการ และอัปโหลดไฟล์ PDF เพื่อทำการประมวลผล
              </p>
            </div>
            <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg bg-muted text-xs text-muted-foreground self-start sm:self-auto">
              <ShieldCheck className="size-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
              <span>ประมวลผลในเบราว์เซอร์ ไม่บันทึกข้อมูล</span>
            </div>
          </div>
        </div>

        {/* Step 1: Select Document Type */}
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <h2 className="text-base sm:text-lg font-semibold flex items-center gap-2">
              <span className="flex items-center justify-center size-6 rounded-full bg-primary text-primary-foreground text-xs font-bold">
                1
              </span>
              เลือกประเภทเอกสาร
            </h2>
            <span className="text-xs text-muted-foreground">
              เลือก 1 ประเภทเอกสาร
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {DOCUMENT_OPTIONS.map((doc) => {
              const isSelected = selectedDocType === doc.id;
              return (
                <button
                  key={doc.id}
                  type="button"
                  onClick={() => setSelectedDocType(doc.id)}
                  className={`text-left p-4 rounded-xl border transition-all duration-200 cursor-pointer relative ${
                    isSelected
                      ? "border-primary bg-primary/5 ring-2 ring-primary shadow-sm"
                      : "border-border bg-card hover:border-primary/50 hover:bg-muted/40"
                  }`}
                >
                  <div className="flex items-start justify-between mb-2">
                    <span
                      className={`text-xs font-medium px-2 py-0.5 rounded-full ${
                        isSelected
                          ? "bg-primary text-primary-foreground"
                          : "bg-muted text-muted-foreground"
                      }`}
                    >
                      {doc.badge}
                    </span>
                    {isSelected && (
                      <CheckCircle2 className="size-4 text-primary shrink-0" />
                    )}
                  </div>
                  <div className="font-bold text-lg text-card-foreground">
                    {doc.title}
                  </div>
                  <div className="text-xs font-semibold text-primary mb-1">
                    {doc.code}
                  </div>
                  <p className="text-xs text-muted-foreground line-clamp-2">
                    {doc.description}
                  </p>
                </button>
              );
            })}
          </div>
        </div>

        {/* Step 2: Upload Files */}
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <h2 className="text-base sm:text-lg font-semibold flex items-center gap-2">
              <span className="flex items-center justify-center size-6 rounded-full bg-primary text-primary-foreground text-xs font-bold">
                2
              </span>
              อัปโหลดไฟล์ PDF ({currentDoc?.title})
            </h2>
            {selectedFiles.length > 0 && (
              <Button
                variant="ghost"
                size="xs"
                onClick={clearAllFiles}
                className="text-xs text-destructive hover:text-destructive"
              >
                ล้างไฟล์ทั้งหมด ({selectedFiles.length})
              </Button>
            )}
          </div>

          <Card className="border-border bg-card">
            <CardContent className="p-6">
              <div
                onDragEnter={handleDrag}
                onDragOver={handleDrag}
                onDragLeave={handleDrag}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all duration-200 flex flex-col items-center justify-center gap-3 ${
                  dragActive
                    ? "border-primary bg-primary/10"
                    : "border-muted-foreground/30 hover:border-primary/60 hover:bg-muted/30"
                }`}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple
                  accept=".pdf,application/pdf"
                  className="hidden"
                  onChange={(e) => handleFiles(e.target.files)}
                />
                <div className="p-4 rounded-full bg-primary/10 text-primary">
                  <UploadCloud className="size-8" />
                </div>
                <div>
                  <p className="font-medium text-base">
                    ลากไฟล์มาวางที่นี่ หรือ{" "}
                    <span className="text-primary underline underline-offset-2">
                      คลิกเพื่อเลือกไฟล์
                    </span>
                  </p>
                  <p className="text-xs text-muted-foreground mt-1">
                    รองรับเฉพาะไฟล์ .PDF สำหรับ{" "}
                    <span className="font-semibold text-foreground">
                      {currentDoc?.title} ({currentDoc?.code})
                    </span>
                  </p>
                </div>
              </div>

              {uploadStatus && (
                <div className="mt-4 flex items-center gap-2 text-sm text-destructive bg-destructive/10 p-3 rounded-lg">
                  <AlertCircle className="size-4 shrink-0" />
                  <span>{uploadStatus}</span>
                </div>
              )}

              {/* Uploaded Files List */}
              {selectedFiles.length > 0 && (
                <div className="mt-6 space-y-3">
                  <h3 className="text-sm font-semibold text-card-foreground">
                    รายการไฟล์ที่เลือก ({selectedFiles.length} ไฟล์):
                  </h3>
                  <div className="divide-y divide-border border border-border rounded-lg max-h-60 overflow-y-auto">
                    {selectedFiles.map((file, idx) => (
                      <div
                        key={`${file.name}-${idx}`}
                        className="p-3 flex items-center justify-between text-sm hover:bg-muted/40 transition-colors"
                      >
                        <div className="flex items-center gap-3 min-w-0 pr-4">
                          <FileCheck2 className="size-5 text-primary shrink-0" />
                          <div className="truncate">
                            <p className="font-medium text-foreground truncate">
                              {file.name}
                            </p>
                            <p className="text-xs text-muted-foreground">
                              {(file.size / 1024).toFixed(1)} KB
                            </p>
                          </div>
                        </div>
                        <Button
                          variant="ghost"
                          size="icon-xs"
                          onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
                            e.stopPropagation();
                            removeFile(idx);
                          }}
                          className="text-muted-foreground hover:text-destructive shrink-0"
                          title="ลบไฟล์"
                        >
                          <Trash2 className="size-4" />
                        </Button>
                      </div>
                    ))}
                  </div>

                  {/* Feedback status messages */}
                  {conversionError && (
                    <div className="flex items-center gap-2 text-sm text-destructive bg-destructive/10 p-3 rounded-lg">
                      <AlertCircle className="size-4 shrink-0" />
                      <span>{conversionError}</span>
                    </div>
                  )}

                  {conversionSuccess && (
                    <div className="flex items-center gap-2 text-sm text-emerald-600 dark:text-emerald-400 bg-emerald-500/10 p-3 rounded-lg">
                      <CheckCircle2 className="size-4 shrink-0" />
                      <span>{conversionSuccess}</span>
                    </div>
                  )}

                  {/* Actions */}
                  <div className="pt-2 flex flex-col sm:flex-row items-center justify-end gap-3">
                    <Button
                      variant="outline"
                      size="default"
                      disabled={isConverting}
                      onClick={() => handleConvert("csv")}
                      className="w-full sm:w-auto"
                    >
                      {isConverting ? (
                        <Loader2 className="size-4 mr-2 animate-spin" />
                      ) : (
                        <FileSpreadsheet className="size-4 mr-2" />
                      )}
                      แปลงเป็น CSV
                    </Button>
                    <Button
                      size="default"
                      disabled={isConverting}
                      onClick={() => handleConvert("xlsx")}
                      className="w-full sm:w-auto"
                    >
                      {isConverting ? (
                        <Loader2 className="size-4 mr-2 animate-spin" />
                      ) : (
                        <FileSpreadsheet className="size-4 mr-2" />
                      )}
                      แปลงเป็น XLSX (Excel)
                    </Button>
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      </div>
    </main>
  );
}
