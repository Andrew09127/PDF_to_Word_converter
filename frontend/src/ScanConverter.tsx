import { useRef, useState } from "react";
import {
  Accordion,
  AccordionDetails,
  AccordionSummary,
  Alert,
  Box,
  Button,
  Checkbox,
  Chip,
  Divider,
  FormControlLabel,
  LinearProgress,
  Paper,
  Stack,
  Typography,
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import PictureAsPdfIcon from "@mui/icons-material/PictureAsPdf";
import PsychologyIcon from "@mui/icons-material/Psychology";
import DescriptionIcon from "@mui/icons-material/Description";
import CloseIcon from "@mui/icons-material/Close";
import {
  ConvertMode,
  ScanFlags,
  analyzePdf,
  downloadResult,
  pollUntilDone,
  startConversion,
} from "./api";

// Версия приложения (показывается в шапке и памятке).
const VERSION = "1.0.0";

// Описание настроек распознавания — человеческими словами для рядового пользователя.
//   invert: галочка показывает обратное значение флага (для «обратных» флагов
//           вроде no_highlight), чтобы «галочка стоит = функция включена».
//   cli:    техническое имя флага — мелкой подписью, на случай обращения в поддержку.
const FLAG_DEFS: {
  key: keyof ScanFlags;
  cli: string;
  title: string;
  desc: string;
  invert?: boolean;
}[] = [
  {
    key: "no_highlight",
    cli: "--no-highlight",
    invert: true,
    title: "Подсвечивать сомнительные слова",
    desc: "Выделяет жёлтым слова, в распознавании которых программа не уверена — удобно, чтобы быстро найти и проверить их вручную. Включено по умолчанию.",
  },
  {
    key: "iim",
    cli: "--iim",
    title: "Улучшать текст нейросетью (ИИ)",
    desc: "Локальная нейросеть исправляет оставшиеся ошибки распознавания и делает текст чище. Работает чуть дольше. Включено по умолчанию.",
  },
  {
    key: "ocr_preprocess",
    cli: "--ocr-preprocess",
    title: "Улучшать качество плохих сканов",
    desc: "Выравнивает наклон и повышает контраст перед распознаванием. Помогает тёмным, бледным и перекошенным сканам, но может ухудшить и без того чёткие. Включайте, если результат плохой.",
  },
  {
    key: "word_order",
    cli: "--word-order",
    title: "Восстанавливать порядок строк на сложных сканах",
    desc: "Помогает, когда колонки или строки в документе распознаются вперемешку. Обычно не требуется — порядок и так определяется верно.",
  },
  {
    key: "ink_bold",
    cli: "--ink-bold",
    title: "Определять жирный шрифт по толщине букв",
    desc: "Экспериментальная функция. На сканах срабатывает ненадёжно и может ошибаться. По умолчанию выключено.",
  },
];

const DEFAULT_FLAGS: ScanFlags = {
  no_highlight: false,
  word_order: false,
  iim: true,
  ink_bold: false,
  ocr_preprocess: false,
};

export default function ScanConverter() {
  const [file, setFile] = useState<File | null>(null);
  const [flags, setFlags] = useState<ScanFlags>(DEFAULT_FLAGS);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState("Ожидание выбора PDF-файла…");
  const [statusKind, setStatusKind] = useState<"info" | "success" | "error">("info");
  const [dragOver, setDragOver] = useState(false);
  // Рекомендованный анализатором режим — для подсветки кнопки. null — ещё не определён.
  const [suggestedMode, setSuggestedMode] = useState<ConvertMode | null>(null);
  const [analyzing, setAnalyzing] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  // Сбрасывает выбранный файл и связанное состояние (рекомендацию, прогресс).
  function clearFile() {
    setFile(null);
    setSuggestedMode(null);
    setProgress(0);
    setStatus("Ожидание выбора PDF-файла…");
    setStatusKind("info");
  }

  function pickFile(f: File | null) {
    if (!f) return;
    if (!f.name.toLowerCase().endsWith(".pdf")) {
      setStatus("Можно загружать только PDF-файлы.");
      setStatusKind("error");
      return;
    }
    setFile(f);
    setProgress(0);
    setSuggestedMode(null);
    setStatusKind("info");
    setStatus("Определяем тип документа…");
    setAnalyzing(true);
    // Лёгкий синхронный анализ на бэке (без OCR) — подсветит рекомендованный режим.
    analyzePdf(f)
      .then((res) => {
        setSuggestedMode(res.suggested);
        setStatus(res.reason);
      })
      .catch(() => {
        // Анализ не критичен: при сбое просто не подсвечиваем, режим выбирает пользователь.
        setStatus("PDF добавлен. Выберите режим конвертации.");
      })
      .finally(() => setAnalyzing(false));
  }

  async function convert(mode: ConvertMode) {
    if (!file || busy) return;
    setBusy(true);
    setProgress(0);
    setStatusKind("info");
    setStatus("Загрузка файла…");
    try {
      const jobId = await startConversion(mode, file, mode === "scan" ? flags : undefined);
      const final = await pollUntilDone(jobId, (s) => {
        setProgress(Math.round(s.progress * 100));
        setStatus(`Обработка: ${s.stage}`);
      });
      if (final.status === "done") {
        await downloadResult(jobId, final.filename);
        setProgress(100);
        setStatusKind("success");
        setStatus(`Готово: ${final.filename} скачан.`);
      } else {
        setStatusKind("error");
        setStatus(`Ошибка: ${final.error || "конвертация не удалась"}`);
      }
    } catch (e) {
      setStatusKind("error");
      setStatus(`Ошибка: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <Box sx={{ minHeight: "100vh", display: "flex", justifyContent: "center", py: 4, px: 2 }}>
      <Paper elevation={4} sx={{ width: 720, borderRadius: 4, overflow: "hidden" }}>
        {/* Шапка */}
        <Box sx={{ px: 3, py: 2, borderBottom: "1px solid #E5E7EB", display: "flex",
                   alignItems: "center", gap: 1.5 }}>
          <Box component="img" src="/logo.png" alt="logo"
               sx={{ width: 34, height: 34, objectFit: "contain" }} />
          <Typography variant="h6" fontWeight={700}>
            PDF → Word конвертер
          </Typography>
          <Chip label={`v${VERSION}`} size="small" color="primary" variant="outlined"
                sx={{ ml: 0.5, fontWeight: 600 }} />
          <Chip label="beta" size="small"
                sx={{ bgcolor: "#FEF3C7", color: "#92400E", fontWeight: 600 }} />
        </Box>

        <Stack spacing={2.5} sx={{ p: 3 }}>
          {/* Дропзона */}
          <Box
            onClick={() => inputRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              pickFile(e.dataTransfer.files?.[0] ?? null);
            }}
            sx={{
              border: "2px dashed",
              borderColor: dragOver ? "primary.main" : "#CBD5E1",
              bgcolor: dragOver ? "#E7F6EA" : "#fff",
              borderRadius: 3,
              py: 4,
              textAlign: "center",
              cursor: "pointer",
              transition: "all .15s",
            }}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".pdf"
              hidden
              onChange={(e) => pickFile(e.target.files?.[0] ?? null)}
            />
            <PictureAsPdfIcon color="primary" sx={{ fontSize: 44 }} />
            <Typography variant="h6" fontWeight={700} mt={1}>
              Нажмите или перетащите PDF
            </Typography>
            <Typography variant="body2" color="text.secondary">
              два режима: распознавание сканов (LLM) или нативный PDF
            </Typography>
            {file && (
              <Box mt={1.5}>
                <Chip
                  color="primary"
                  variant="outlined"
                  label={file.name}
                  onDelete={busy ? undefined : clearFile}
                  deleteIcon={<CloseIcon />}
                />
              </Box>
            )}
          </Box>

          {/* Флаги (применяются к скан-режиму) */}
          <Accordion disableGutters sx={{ borderRadius: 2, "&:before": { display: "none" } }}
                     variant="outlined">
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography fontWeight={700}>
                Дополнительные настройки распознавания
              </Typography>
            </AccordionSummary>
            <AccordionDetails>
              <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
                Эти настройки влияют только на режим «Отсканированный PDF». Для большинства
                документов менять их не нужно — значения по умолчанию подобраны оптимально.
              </Typography>
              <Stack spacing={1}>
                {FLAG_DEFS.map((f) => {
                  // Для «обратных» флагов галочка показывает противоположное значение.
                  const checked = f.invert ? !flags[f.key] : flags[f.key];
                  return (
                    <FormControlLabel
                      key={f.key}
                      sx={{ alignItems: "flex-start", m: 0 }}
                      control={
                        <Checkbox
                          checked={checked}
                          disabled={busy}
                          onChange={(e) => {
                            const v = f.invert ? !e.target.checked : e.target.checked;
                            setFlags((prev) => ({ ...prev, [f.key]: v }));
                          }}
                        />
                      }
                      label={
                        <Box sx={{ py: 0.5 }}>
                          <Typography variant="body2" fontWeight={700}>
                            {f.title}
                          </Typography>
                          <Typography variant="caption" color="text.secondary"
                                      component="p" sx={{ mt: 0.25 }}>
                            {f.desc}
                          </Typography>
                          <Typography variant="caption" sx={{ color: "#9CA3AF" }}>
                            {f.cli}
                          </Typography>
                        </Box>
                      }
                    />
                  );
                })}
              </Stack>
            </AccordionDetails>
          </Accordion>

          {/* Памятка */}
          <Accordion disableGutters sx={{ borderRadius: 2, "&:before": { display: "none" } }}
                     variant="outlined">
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography fontWeight={700}>Памятка</Typography>
            </AccordionSummary>
            <AccordionDetails>
              <Typography variant="body2" fontWeight={700} gutterBottom>
                Версия {VERSION} — первый публичный релиз (beta).
              </Typography>
              <Typography variant="body2" color="text.secondary" gutterBottom>
                Это первая версия конвертера, и она будет активно дорабатываться.
                Возможны неточности в распознавании текста и оформлении — особенно на
                сложных сканах. Готовый документ рекомендуем перед использованием
                просматривать и при необходимости править вручную.
              </Typography>
              <Typography variant="body2" color="text.secondary" gutterBottom>
                Все данные обрабатываются <b>полностью на вашем компьютере</b> — файлы
                никуда не загружаются и не передаются в интернет.
              </Typography>

              <Divider sx={{ my: 1.5 }} />

              <Typography variant="body2" fontWeight={700} gutterBottom>
                Как выбрать режим
              </Typography>
              <Typography variant="body2" color="text.secondary" gutterBottom>
                • <b>Отсканированный PDF</b> — если страницы являются изображениями
                (фото или скан документа) и текст в них нельзя выделить мышью.
              </Typography>
              <Typography variant="body2" color="text.secondary" gutterBottom>
                • <b>Неотсканированный PDF</b> — если в документе есть текстовый слой
                (текст выделяется и копируется). Работает быстро и точно.
              </Typography>
              <Typography variant="caption" color="text.secondary">
                Не уверены? Попробуйте сначала «Неотсканированный PDF»: если результат
                окажется пустым или сломанным — значит документ отсканированный, запустите
                второй режим.
              </Typography>

              <Divider sx={{ my: 1.5 }} />

              <Typography variant="body2" fontWeight={700} gutterBottom>
                Советы для лучшего результата
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Загружайте сканы хорошего качества: без сильного размытия, обрезанных
                страниц, перекосов и слишком тёмного фона. Если распознавание плохое —
                включите «Улучшать качество плохих сканов» в дополнительных настройках.
              </Typography>
            </AccordionDetails>
          </Accordion>

          {/* Прогресс */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", mb: 0.5 }}>
              <Typography variant="body2" color="text.secondary">Прогресс</Typography>
              <Typography variant="body2" color="text.secondary">{progress}%</Typography>
            </Box>
            <LinearProgress
              variant={busy && progress === 0 ? "indeterminate" : "determinate"}
              value={progress}
              color={statusKind === "success" ? "success" : "primary"}
              sx={{ height: 10, borderRadius: 5 }}
            />
          </Box>

          <Alert severity={statusKind} variant="outlined">{status}</Alert>

          {/* Две карточки режимов с пояснениями.
              Анализатор подсвечивает рекомендованную карточку зелёной рамкой и делает
              её кнопку «success»-зелёной; обе кнопки остаются активны — выбор за пользователем. */}
          <Stack direction={{ xs: "column", sm: "row" }} spacing={2} alignItems="stretch">
            {/* Скан-режим */}
            {(() => {
              const recommended = suggestedMode === "scan";
              return (
            <Box sx={{ flex: 1, p: 2, borderRadius: 2, display: "flex", flexDirection: "column",
                       border: "2px solid", borderColor: recommended ? "success.main" : "#E5E7EB",
                       bgcolor: recommended ? "#F0FBF2" : "#fff", transition: "all .2s" }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 0.5 }}>
                <PsychologyIcon color="primary" fontSize="small" />
                <Typography variant="subtitle2" fontWeight={700}>
                  Отсканированный PDF
                </Typography>
                {recommended && (
                  <Chip label="рекомендуем" size="small" color="success"
                        sx={{ ml: "auto", fontWeight: 600, height: 20 }} />
                )}
              </Box>
              <Typography variant="caption" color="text.secondary" sx={{ mb: 1.5, flexGrow: 1 }}>
                Для документов-изображений (фото и сканы), где текст нельзя выделить.
                Распознаёт текст и восстанавливает структуру с помощью ИИ. Точнее, но
                заметно медленнее — учитывает настройки выше.
              </Typography>
              <Button
                fullWidth
                size="large"
                variant={recommended ? "contained" : "outlined"}
                color={recommended ? "success" : "primary"}
                startIcon={<PsychologyIcon />}
                disabled={!file || busy || analyzing}
                onClick={() => convert("scan")}
              >
                Конвертировать скан
              </Button>
            </Box>
              );
            })()}

            {/* Нативный режим */}
            {(() => {
              const recommended = suggestedMode === "native";
              return (
            <Box sx={{ flex: 1, p: 2, borderRadius: 2, display: "flex", flexDirection: "column",
                       border: "2px solid", borderColor: recommended ? "success.main" : "#E5E7EB",
                       bgcolor: recommended ? "#F0FBF2" : "#fff", transition: "all .2s" }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 0.5 }}>
                <DescriptionIcon color="primary" fontSize="small" />
                <Typography variant="subtitle2" fontWeight={700}>
                  Неотсканированный PDF
                </Typography>
                {recommended && (
                  <Chip label="рекомендуем" size="small" color="success"
                        sx={{ ml: "auto", fontWeight: 600, height: 20 }} />
                )}
              </Box>
              <Typography variant="caption" color="text.secondary" sx={{ mb: 1.5, flexGrow: 1 }}>
                Для документов с текстовым слоем (текст выделяется и копируется).
                Переносит текст, таблицы и картинки напрямую — быстро и точно.
                Распознавание и настройки выше не используются.
              </Typography>
              <Button
                fullWidth
                size="large"
                variant={recommended ? "contained" : "outlined"}
                color={recommended ? "success" : "primary"}
                startIcon={<DescriptionIcon />}
                disabled={!file || busy || analyzing}
                onClick={() => convert("native")}
              >
                Конвертировать PDF
              </Button>
            </Box>
              );
            })()}
          </Stack>
        </Stack>
      </Paper>

      {/* Незаметная «подпись автора» — полупрозрачный логотип в правом нижнем углу.
          Не перехватывает клики, чуть проявляется при наведении. */}
      <Box
        component="img"
        src="/logo.png"
        alt=""
        title="PDF → Word конвертер"
        sx={{
          position: "fixed",
          right: 16,
          bottom: 16,
          width: 28,
          height: 28,
          objectFit: "contain",
          opacity: 0.25,
          pointerEvents: "none",
          userSelect: "none",
          transition: "opacity .2s",
          "&:hover": { opacity: 0.6 },
        }}
      />
    </Box>
  );
}
