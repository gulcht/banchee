"use client";

import { useState, useRef } from "react";
import Link from "next/link";
import { Card, CardContent } from "@/components/ui/card";
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
  Scale,
  Info,
} from "lucide-react";
import { reconcilePayrollPdfs } from "@/lib/api";

export default function ReconcilePage() {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [dragActive, setDragActive] = useState<boolean>(false);
  const [uploadStatus, setUploadStatus] = useState<string | null>(null);
  const [isReconciling, setIsReconciling] = useState<boolean>(false);
  const [reconcileError, setReconcileError] = useState<string | null>(null);
  const [reconcileSuccess, setReconcileSuccess] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

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
    setReconcileError(null);
    setReconcileSuccess(null);
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
    setReconcileError(null);
    setReconcileSuccess(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleReconcile = async (format: "xlsx" | "csv") => {
    if (selectedFiles.length === 0) return;
    setReconcileError(null);
    setReconcileSuccess(null);

    setIsReconciling(true);
    try {
      const result = await reconcilePayrollPdfs(selectedFiles, format);
      setReconcileSuccess(
        `กระทบยอดสำเร็จแล้ว (${result.filename}) ดาวน์โหลดไฟล์เรียบร้อย`
      );
      setSelectedFiles([]);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (err: unknown) {
      if (err instanceof Error) {
        setReconcileError(err.message);
      } else {
        setReconcileError("เกิดข้อผิดพลาดในการกระทบยอดเอกสาร");
      }
    } finally {
      setIsReconciling(false);
    }
  };

  return (
    <main className="min-h-screen bg-background text-foreground py-8 px-4 sm:px-6 lg:px-8">
      <div className="max-w-5xl mx-auto space-y-8">
        {/* Navigation & Header */}
        <div className="space-y-4">
          <Link
            href="/"
            className="inline-flex items-center text-sm font-medium text-muted-foreground hover:text-primary transition-colors"
          >
            <ArrowLeft className="size-4 mr-1" />
            กลับหน้าหลัก
          </Link>

          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 border-b border-border pb-6">
            <div>
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-xl bg-primary/10 text-primary">
                  <Scale className="size-6" />
                </div>
                <div>
                  <h1 className="text-3xl font-extrabold tracking-tight">
                    กระทบยอดภาษีและประกันสังคม
                  </h1>
                  <p className="text-sm font-medium text-muted-foreground">
                    12-Month Payroll Reconciliation (สปส. 1-10 vs ภ.ง.ด. 1)
                  </p>
                </div>
              </div>
              <p className="mt-2 text-sm text-muted-foreground max-w-2xl">
                อัปโหลดไฟล์ PDF ใบแนบ ภ.ง.ด. 1 และ สปส. 1-10 ส่วนที่ 2 (รายเดือนทั้ง 12 เดือน) ระบบจะตรวจจับและกระทบยอดเงินเดือน ค่าจ้าง และภาษีหัก ณ ที่จ่ายให้อัตโนมัติ
              </p>
            </div>

            {/* Stateless & Privacy notice */}
            <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-muted/60 border border-border text-xs text-muted-foreground w-fit">
              <ShieldCheck className="size-4 text-emerald-500 shrink-0" />
              <span>ประมวลผลในหน่วยความจำ ไม่บันทึกข้อมูลส่วนบุคคล</span>
            </div>
          </div>
        </div>

        {/* Instructions / Supported documents card */}
        <div className="rounded-xl border border-border bg-card p-5 text-card-foreground shadow-xs">
          <div className="flex items-start gap-3">
            <Info className="size-5 text-primary shrink-0 mt-0.5" />
            <div className="space-y-2 text-sm">
              <h3 className="font-semibold text-foreground">คำแนะนำการใช้งาน</h3>
              <ul className="list-disc list-inside space-y-1 text-muted-foreground text-xs sm:text-sm">
                <li>
                  เลือกหรือลากไฟล์ PDF ทั้งหมดที่ต้องการกระทบยอดพร้อมกัน เช่น ไฟล์ สปส. 1-10 และ ภ.ง.ด. 1 ประจำเดือน ม.ค. - ธ.ค.
                </li>
                <li>
                  ระบบจะแยกประเภทเอกสาร (สปส. 1-10 และ ภ.ง.ด. 1) และช่วงเดือน/ปีโดยอัตโนมัติจากเนื้อหาในไฟล์
                </li>
                <li>
                  สามารถส่งออกผลลัพธ์เป็น <strong>Excel (.xlsx)</strong> หรือ <strong>CSV (UTF-8 with BOM)</strong> ได้ทันที
                </li>
              </ul>
            </div>
          </div>
        </div>

        {/* Upload Zone */}
        <div className="space-y-6">
          <Card className="border-border shadow-xs">
            <CardContent className="p-6">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h2 className="text-base font-semibold">อัปโหลดไฟล์เอกสาร PDF</h2>
                  <p className="text-xs text-muted-foreground">
                    รองรับไฟล์ สปส. 1-10 ส่วนที่ 2 และ ภ.ง.ด. 1 ใบแนบ
                  </p>
                </div>
                {selectedFiles.length > 0 && (
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={clearAllFiles}
                    className="text-xs text-muted-foreground hover:text-destructive"
                  >
                    ล้างทั้งหมด ({selectedFiles.length})
                  </Button>
                )}
              </div>

              {/* Drag and Drop Zone */}
              <div
                onDragEnter={handleDrag}
                onDragLeave={handleDrag}
                onDragOver={handleDrag}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all duration-200 ${
                  dragActive
                    ? "border-primary bg-primary/5"
                    : "border-border hover:border-primary/50 hover:bg-muted/30"
                }`}
              >
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple
                  accept="application/pdf,.pdf"
                  className="hidden"
                  onChange={(e) => handleFiles(e.target.files)}
                />

                <div className="flex flex-col items-center justify-center space-y-3">
                  <div className="p-3.5 rounded-full bg-primary/10 text-primary">
                    <UploadCloud className="size-8" />
                  </div>
                  <div className="space-y-1">
                    <p className="text-sm font-medium">
                      คลิกเพื่อเลือกไฟล์ หรือลากไฟล์มาวางที่นี่
                    </p>
                    <p className="text-xs text-muted-foreground">
                      สามารถเลือกได้หลายไฟล์พร้อมกัน (สปส. 1-10 และ ภ.ง.ด. 1)
                    </p>
                  </div>
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
                  {reconcileError && (
                    <div className="flex items-center gap-2 text-sm text-destructive bg-destructive/10 p-3 rounded-lg">
                      <AlertCircle className="size-4 shrink-0" />
                      <span>{reconcileError}</span>
                    </div>
                  )}

                  {reconcileSuccess && (
                    <div className="flex items-center gap-2 text-sm text-emerald-600 dark:text-emerald-400 bg-emerald-500/10 p-3 rounded-lg">
                      <CheckCircle2 className="size-4 shrink-0" />
                      <span>{reconcileSuccess}</span>
                    </div>
                  )}

                  {/* Actions */}
                  <div className="pt-2 flex flex-col sm:flex-row items-center justify-end gap-3">
                    <Button
                      variant="outline"
                      size="default"
                      disabled={isReconciling}
                      onClick={() => handleReconcile("csv")}
                      className="w-full sm:w-auto"
                    >
                      {isReconciling ? (
                        <Loader2 className="size-4 mr-2 animate-spin" />
                      ) : (
                        <FileSpreadsheet className="size-4 mr-2" />
                      )}
                      ส่งออกเป็น CSV
                    </Button>
                    <Button
                      size="default"
                      disabled={isReconciling}
                      onClick={() => handleReconcile("xlsx")}
                      className="w-full sm:w-auto"
                    >
                      {isReconciling ? (
                        <Loader2 className="size-4 mr-2 animate-spin" />
                      ) : (
                        <FileSpreadsheet className="size-4 mr-2" />
                      )}
                      กระทบยอดและส่งออกเป็น Excel (XLSX)
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
