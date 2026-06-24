export type ConvertMode = "scan" | "native";

export interface ScanFlags {
  no_highlight: boolean;
  word_order: boolean;
  iim: boolean;
  ink_bold: boolean;
  ocr_preprocess: boolean;
}

export interface AnalyzeResult {
  suggested: ConvertMode;
  confidence: "high" | "low";
  reason: string;
}

// Быстро определяет тип PDF (скан/нативный), чтобы подсветить рекомендованный режим.
export async function analyzePdf(file: File): Promise<AnalyzeResult> {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch("/convert/analyze", { method: "POST", body: form });
  if (!res.ok) {
    const detail = await res.json().catch(() => ({}));
    throw new Error(detail.detail || `Анализ не удался (${res.status})`);
  }
  return (await res.json()) as AnalyzeResult;
}

export interface JobStatus {
  job_id: string;
  status: "queued" | "running" | "done" | "error";
  stage: string;
  progress: number;
  filename: string;
  error: string | null;
}

// Запускает конвертацию. Для скана прикладывает флаги, для нативного — только файл.
export async function startConversion(
  mode: ConvertMode,
  file: File,
  flags?: ScanFlags,
): Promise<string> {
  const form = new FormData();
  form.append("file", file);
  if (mode === "scan" && flags) {
    form.append("no_highlight", String(flags.no_highlight));
    form.append("word_order", String(flags.word_order));
    form.append("iim", String(flags.iim));
    form.append("ink_bold", String(flags.ink_bold));
    form.append("ocr_preprocess", String(flags.ocr_preprocess));
  }
  const res = await fetch(`/convert/${mode}`, { method: "POST", body: form });
  if (!res.ok) {
    const detail = await res.json().catch(() => ({}));
    throw new Error(detail.detail || `Ошибка запуска (${res.status})`);
  }
  const data = await res.json();
  return data.job_id as string;
}

export async function getStatus(jobId: string): Promise<JobStatus> {
  const res = await fetch(`/convert/status/${jobId}`);
  if (!res.ok) throw new Error(`Статус недоступен (${res.status})`);
  return (await res.json()) as JobStatus;
}

// Опрашивает статус раз в секунду, пока не done/error. onUpdate — для прогресс-бара.
export async function pollUntilDone(
  jobId: string,
  onUpdate: (s: JobStatus) => void,
): Promise<JobStatus> {
  for (;;) {
    const s = await getStatus(jobId);
    onUpdate(s);
    if (s.status === "done" || s.status === "error") return s;
    await new Promise((r) => setTimeout(r, 1000));
  }
}

// Скачивает готовый DOCX и инициирует сохранение в браузере.
export async function downloadResult(jobId: string, filename: string): Promise<void> {
  const res = await fetch(`/convert/download/${jobId}`);
  if (!res.ok) throw new Error(`Скачивание не удалось (${res.status})`);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
