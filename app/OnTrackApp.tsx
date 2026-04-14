"use client";

import { supabase } from "@/lib/supabase";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  BarChart3,
  CalendarCheck2,
  ChevronRight,
  FileSpreadsheet,
  LayoutGrid,
  MessageSquare,
  Search,
  ShieldCheck,
  Sparkles,
  Trash2,
  TrendingDown,
  TrendingUp,
  Upload,
  Users,
} from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Progress } from "@/components/ui/progress";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  PieChart,
  Pie,
  Cell,
} from "recharts";

type Trend = "improving" | "stable" | "worsening";
type TabKey = "overview" | "files" | "students" | "interventions";

type Student = {
  id: string;
  name: string;
  grade: string;
  school: string;
  attendance: number;
  gpa: number;
  risk: number;
  trend: Trend;
  drivers: string[];
  suggestions: string[];
  behavior: number;
  engagement: number;
  tardies: number;
  failedCourses: number;
  sourceFile: string;
};

type RawRow = Record<string, unknown>;

type UploadedDataset = {
  id: string;
  fileName: string;
  fileSize: number;
  lastModified: number;
  fingerprint: string;
  students: Student[];
};

type SortKey = "student" | "grade" | "attendance" | "gpa" | "risk" | "file";
type SortDirection = "asc" | "desc";

const riskRank: Record<string, number> = {
  low: 0,
  medium: 1,
  high: 2,
};

function parsePercent(value: unknown): number {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const cleaned = value.replace("%", "").trim();
    const num = Number(cleaned);
    return Number.isNaN(num) ? -Infinity : num;
  }
  return -Infinity;
}

function parseNumber(value: unknown): number {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const num = Number(value.trim());
    return Number.isNaN(num) ? -Infinity : num;
  }
  return -Infinity;
}

function compareText(a: unknown, b: unknown) {
  return String(a ?? "").localeCompare(String(b ?? ""), undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function normalizeKey(key: string) {
  return key.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findColumn(row: RawRow, aliases: string[]) {
  const entries = Object.keys(row).map((key) => ({
    original: key,
    normalized: normalizeKey(key),
  }));

  for (const alias of aliases) {
    const aliasNormalized = normalizeKey(alias);
    const exact = entries.find((entry) => entry.normalized === aliasNormalized);
    if (exact) return exact.original;
  }

  for (const alias of aliases) {
    const aliasNormalized = normalizeKey(alias);
    const partial = entries.find(
      (entry) =>
        entry.normalized.includes(aliasNormalized) ||
        aliasNormalized.includes(entry.normalized)
    );
    if (partial) return partial.original;
  }

  return undefined;
}

function toNumber(value: unknown, fallback = 0) {
  if (typeof value === "number" && !Number.isNaN(value)) return value;

  if (typeof value === "string") {
    const cleaned = value.replace(/[%,$]/g, "").trim();
    const parsed = Number(cleaned);
    if (!Number.isNaN(parsed)) return parsed;
  }

  return fallback;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function detectTrend(attendance: number, gpa: number, failedCourses: number, tardies: number): Trend {
  if (attendance < 80 || gpa < 2.0 || failedCourses >= 2 || tardies >= 10) return "worsening";
  if (attendance >= 92 && gpa >= 3.0 && failedCourses === 0 && tardies <= 2) return "improving";
  return "stable";
}

function buildDrivers(student: {
  attendance: number;
  gpa: number;
  behavior: number;
  engagement: number;
  tardies: number;
  failedCourses: number;
}) {
  const drivers: string[] = [];

  if (student.attendance < 80) drivers.push("Chronic absenteeism");
  else if (student.attendance < 90) drivers.push("Attendance below target");

  if (student.gpa < 2.0) drivers.push("Low GPA");
  else if (student.gpa < 2.5) drivers.push("GPA below healthy range");

  if (student.failedCourses >= 2) drivers.push("Multiple failed courses");
  else if (student.failedCourses === 1) drivers.push("One failed course");

  if (student.tardies >= 10) drivers.push("Repeated tardiness");
  else if (student.tardies >= 5) drivers.push("Frequent tardiness");

  if (student.engagement < 50) drivers.push("Low engagement");
  if (student.behavior < 50) drivers.push("Behavior concerns");

  if (drivers.length === 0) drivers.push("Low current concern");

  return drivers;
}

async function uploadFileToSupabase(file: File) {
    const filePath = `excel/${Date.now()}-${file.name}`;
  
    const { data, error } = await supabase.storage
      .from("uploads")
      .upload(filePath, file, { upsert: false });
  
    if (error) throw error;
    return data;
  }

function buildSuggestions(student: {
  attendance: number;
  gpa: number;
  behavior: number;
  engagement: number;
  tardies: number;
  failedCourses: number;
  risk: number;
}) {
  const suggestions: string[] = [];

  if (student.attendance < 85) {
    suggestions.push("Schedule a counselor check-in and family outreach about attendance.");
  }

  if (student.gpa < 2.5 || student.failedCourses >= 1) {
    suggestions.push("Assign tutoring, missing-work recovery, and weekly academic monitoring.");
  }

  if (student.engagement < 55) {
    suggestions.push("Match the student with a mentor and connect them to an activity or club.");
  }

  if (student.tardies >= 5) {
    suggestions.push("Review daily routines, transportation barriers, and first-period support.");
  }

  if (student.behavior < 55) {
    suggestions.push("Create a behavior support plan with a trusted adult and clear goals.");
  }

  if (student.risk >= 70) {
    suggestions.push("Escalate to a student-support team meeting within 1 week.");
  }

  if (suggestions.length === 0) {
    suggestions.push("Continue light-touch monitoring and positive reinforcement.");
  }

  return suggestions;
}

function calculateRisk(student: {
  attendance: number;
  gpa: number;
  behavior: number;
  engagement: number;
  tardies: number;
  failedCourses: number;
}) {
  const attendanceRisk = clamp((100 - student.attendance) * 0.9, 0, 35);
  const gpaRisk = clamp((4 - student.gpa) * 10, 0, 25);
  const failedCourseRisk = clamp(student.failedCourses * 12, 0, 24);
  const tardyRisk = clamp(student.tardies * 1.5, 0, 10);
  const engagementRisk = clamp((100 - student.engagement) * 0.08, 0, 8);
  const behaviorRisk = clamp((100 - student.behavior) * 0.06, 0, 6);

  return Math.round(
    clamp(
      attendanceRisk +
        gpaRisk +
        failedCourseRisk +
        tardyRisk +
        engagementRisk +
        behaviorRisk,
      0,
      100
    )
  );
}

function riskLabel(risk: number) {
  if (risk >= 70) return "High";
  if (risk >= 35) return "Moderate";
  return "Low";
}

function riskVariant(risk: number): "default" | "secondary" | "destructive" {
  if (risk >= 70) return "destructive";
  if (risk >= 35) return "secondary";
  return "default";
}

function parseStudents(rows: RawRow[], fileName: string): Student[] {
  if (!rows.length) return [];

  return rows.map((row, index) => {
    const idCol = findColumn(row, ["student id", "id", "studentid"]);
    const nameCol = findColumn(row, ["student name", "name", "fullname"]);
    const gradeCol = findColumn(row, ["grade", "class", "year"]);
    const schoolCol = findColumn(row, ["school", "school name"]);
    const attendanceCol = findColumn(row, ["attendance", "attendance %", "attendance percent"]);
    const gpaCol = findColumn(row, ["gpa", "grade point average"]);
    const behaviorCol = findColumn(row, ["behavior", "behavior score", "discipline"]);
    const engagementCol = findColumn(row, ["engagement", "engagement score", "participation"]);
    const tardiesCol = findColumn(row, ["tardies", "tardy", "late count"]);
    const failedCoursesCol = findColumn(row, ["failed courses", "failed course", "failures"]);

    const attendance = clamp(toNumber(attendanceCol ? row[attendanceCol] : undefined, 95), 0, 100);
    const gpa = clamp(toNumber(gpaCol ? row[gpaCol] : undefined, 3.0), 0, 4);
    const behavior = clamp(toNumber(behaviorCol ? row[behaviorCol] : undefined, 80), 0, 100);
    const engagement = clamp(toNumber(engagementCol ? row[engagementCol] : undefined, 75), 0, 100);
    const tardies = clamp(toNumber(tardiesCol ? row[tardiesCol] : undefined, 0), 0, 100);
    const failedCourses = clamp(
      toNumber(failedCoursesCol ? row[failedCoursesCol] : undefined, 0),
      0,
      20
    );

    const base = {
      attendance,
      gpa,
      behavior,
      engagement,
      tardies,
      failedCourses,
    };

    const risk = calculateRisk(base);
    const trend = detectTrend(attendance, gpa, failedCourses, tardies);
    const drivers = buildDrivers(base);
    const suggestions = buildSuggestions({ ...base, risk });

    return {
      id: String(
        idCol ? row[idCol] ?? `${fileName}-S-${1000 + index}` : `${fileName}-S-${1000 + index}`
      ),
      name: String(nameCol ? row[nameCol] ?? `Student ${index + 1}` : `Student ${index + 1}`),
      grade: String(gradeCol ? row[gradeCol] ?? "Unknown" : "Unknown"),
      school: String(schoolCol ? row[schoolCol] ?? "Unknown School" : "Unknown School"),
      attendance,
      gpa,
      risk,
      trend,
      drivers,
      suggestions,
      behavior,
      engagement,
      tardies,
      failedCourses,
      sourceFile: fileName,
    };
  });
}

function StatCard({
  title,
  value,
  subtitle,
  icon: Icon,
}: {
  title: string;
  value: string;
  subtitle: string;
  icon: React.ComponentType<{ className?: string }>;
}) {
  return (
    <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
      <CardContent className="p-5">
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-sm text-slate-500">{title}</p>
            <p className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">{value}</p>
            <p className="mt-1 text-sm text-slate-500">{subtitle}</p>
          </div>
          <div className="rounded-2xl border border-slate-200 bg-slate-50 p-3 text-slate-700">
            <Icon className="h-5 w-5" />
          </div>
        </div>
      </CardContent>
    </Card>
  );
}

function TabButton({
  active,
  onClick,
  children,
}: {
  active: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      className={`rounded-full px-4 py-2 text-sm font-medium transition ${
        active
          ? "bg-slate-950 text-white"
          : "text-slate-600 hover:bg-slate-100 hover:text-slate-950"
      }`}
    >
      {children}
    </button>
  );
}

const CHART_COLORS = ["#10b981", "#f59e0b", "#ef4444"];

function isDuplicateDatasetFile(
  file: File,
  existingDatasets: Array<
    Pick<UploadedDataset, "fingerprint" | "fileName" | "fileSize" | "lastModified">
  >,
  seenInThisBatch: Set<string>
) {
  const strictFingerprint = `${file.name}__${file.size}__${file.lastModified}`;
  const fallbackFingerprint = `${file.name}__${file.size}`;

  if (seenInThisBatch.has(strictFingerprint) || seenInThisBatch.has(fallbackFingerprint)) {
    return true;
  }

  return existingDatasets.some((dataset) => {
    const existingFallbackFingerprint = `${dataset.fileName}__${dataset.fileSize}`;

    return (
      dataset.fingerprint === strictFingerprint ||
      dataset.fingerprint === fallbackFingerprint ||
      existingFallbackFingerprint === fallbackFingerprint ||
      (dataset.fileName === file.name && dataset.fileSize === file.size) ||
      (dataset.fileName === file.name &&
        dataset.lastModified !== 0 &&
        dataset.lastModified === file.lastModified)
    );
  });
}

export default function OnTrackApp() {
  const [activeTab, setActiveTab] = useState<TabKey>("overview");
  const [datasets, setDatasets] = useState<UploadedDataset[]>([]);
  const [activeDatasetId, setActiveDatasetId] = useState<string>("all");
  const [viewMode, setViewMode] = useState<"combined" | "single">("combined");
  const [query, setQuery] = useState("");
  const [gradeFilter, setGradeFilter] = useState("all");
  const [selectedId, setSelectedId] = useState("");
  const [error, setError] = useState("");
  const [duplicateNotice, setDuplicateNotice] = useState<string[]>([]);
  const [studentNotes, setStudentNotes] = useState<Record<string, string>>({});
  const [user, setUser] = useState<any>(null);
  const [sortKey, setSortKey] = useState<SortKey>("student");
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");

  useEffect(() => {
    try {
      const savedNotes = localStorage.getItem("ontrack-student-notes");
      if (savedNotes) {
        setStudentNotes(JSON.parse(savedNotes));
      }
    } catch (err) {
      console.error("Failed to load notes from localStorage", err);
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem("ontrack-student-notes", JSON.stringify(studentNotes));
    } catch (err) {
      console.error("Failed to save notes to localStorage", err);
    }
  }, [studentNotes]);

  useEffect(() => {
    async function getUser() {
      const { data } = await supabase.auth.getUser();
      setUser(data.user);
    }
  
    getUser();
  }, []);

  useEffect(() => {
    async function loadSavedDatasets() {
      if (!user) return;

      const { data: savedRows, error } = await supabase
        .from("datasets")
        .select("id, file_name, storage_path")
        .eq("user_id", user.id);

      if (error) {
        console.error("Failed to load saved datasets:", error);
        return;
      }

      const restoredDatasets: UploadedDataset[] = [];

      for (const row of savedRows ?? []) {
        if (!row.storage_path) continue;

        const { data, error: downloadError } = await supabase.storage
          .from("uploads")
          .download(row.storage_path);

        if (downloadError || !data) {
          console.error("Failed to download file:", row.storage_path, downloadError);
          continue;
        }

        try {
          const buffer = await data.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          if (!firstSheetName) continue;

          const worksheet = workbook.Sheets[firstSheetName];
          const jsonRows = XLSX.utils.sheet_to_json<RawRow>(worksheet, { defval: "" });
          if (!jsonRows.length) continue;

          const students = parseStudents(jsonRows, row.file_name);
          if (!students.length) continue;

          restoredDatasets.push({
            id: row.id,
            fileName: row.file_name,
            fileSize: data.size ?? 0,
            lastModified: 0,
            fingerprint: `${row.file_name}__${data.size ?? 0}`,
            students,
          });
        } catch (err) {
          console.error("Failed to rebuild dataset from storage:", err);
        }
      }

      setDatasets(restoredDatasets);

      if (restoredDatasets.length && !selectedId) {
        setSelectedId(restoredDatasets[0].students[0]?.id ?? "");
      }
    }

    loadSavedDatasets();
  }, [user]);

  async function handleFilesUpload(event: React.ChangeEvent<HTMLInputElement>) {
  const files = Array.from(event.target.files ?? []);
  if (!files.length) return;

  setError("");
  setDuplicateNotice([]);

  const newDatasets: UploadedDataset[] = [];
  const skippedDuplicates: string[] = [];
  const uploadFailures: string[] = [];

  const seenInThisBatch = new Set<string>();
  const existingDatasetsSnapshot: UploadedDataset[] = [...datasets];

  for (const file of files) {
    const strictFingerprint = `${file.name}__${file.size}__${file.lastModified}`;
    const fallbackFingerprint = `${file.name}__${file.size}`;

    const alreadyExists = isDuplicateDatasetFile(
      file,
      existingDatasetsSnapshot,
      seenInThisBatch
    );

    if (alreadyExists) {
      skippedDuplicates.push(file.name);
      continue;
    }

    seenInThisBatch.add(strictFingerprint);
    seenInThisBatch.add(fallbackFingerprint);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];

      if (!firstSheetName) {
        uploadFailures.push(file.name);
        continue;
      }

      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json<RawRow>(worksheet, { defval: "" });

      if (!jsonRows.length) {
        uploadFailures.push(file.name);
        continue;
      }

      const students = parseStudents(jsonRows, file.name);

      if (!students.length) {
        uploadFailures.push(file.name);
        continue;
      }

      if (user) {
        const filePath = `excel/${Date.now()}-${Math.random()
          .toString(36)
          .slice(2, 8)}-${file.name}`;

        const { error: uploadError } = await supabase.storage
          .from("uploads")
          .upload(filePath, file, { upsert: false });

        if (uploadError) {
          console.error("UPLOAD FAILED:", uploadError);
          uploadFailures.push(file.name);
          continue;
        }

        const { error: dbError } = await supabase.from("datasets").insert({
          user_id: user.id,
          file_name: file.name,
          storage_path: filePath,
        });

        if (dbError) {
          console.error("DATABASE SAVE FAILED:", dbError);
        }
      }

      const createdDataset: UploadedDataset = {
        id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        fileName: file.name,
        fileSize: file.size,
        lastModified: file.lastModified,
        fingerprint: strictFingerprint,
        students,
      };

      newDatasets.push(createdDataset);
      existingDatasetsSnapshot.push(createdDataset);
    } catch (err) {
      console.error("File processing failed:", err);
      uploadFailures.push(file.name);
    }
  }

  if (newDatasets.length) {
    setDatasets((prev) => {
      const updated = [...prev, ...newDatasets];

      if (!selectedId && updated[0]?.students[0]?.id) {
        setSelectedId(updated[0].students[0].id);
      }

      return updated;
    });

    if (viewMode === "single" && activeDatasetId === "all") {
      setActiveDatasetId(newDatasets[0].id);
    }
  }

  if (skippedDuplicates.length) {
    setDuplicateNotice(skippedDuplicates);
  }

  if (uploadFailures.length) {
    setError(`Could not process: ${uploadFailures.join(", ")}`);
  }

  event.target.value = "";
}

  function deleteDataset(datasetId: string) {
    const dataset = datasets.find((item) => item.id === datasetId);

    setDatasets((prev) => {
      const updated = prev.filter((item) => item.id !== datasetId);

      if (activeDatasetId === datasetId) {
        if (viewMode === "combined") {
          setActiveDatasetId("all");
        } else {
          setActiveDatasetId(updated[0]?.id ?? "all");
        }
      }

      return updated;
    });

    if (dataset) {
      const removedStudentIds = new Set(dataset.students.map((student) => student.id));
      setStudentNotes((prev) => {
        const copy = { ...prev };
        for (const id of removedStudentIds) {
          delete copy[id];
        }
        return copy;
      });
    }
  }

  function handleSort(key: SortKey) {
  if (sortKey === key) {
    setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
  } else {
    setSortKey(key);
    setSortDirection("asc");
  }
  }
  const allStudents = useMemo(() => datasets.flatMap((dataset) => dataset.students), [datasets]);

  const visibleStudents = useMemo(() => {
    if (viewMode === "combined") return allStudents;
    const active = datasets.find((dataset) => dataset.id === activeDatasetId);
    return active?.students ?? [];
  }, [viewMode, activeDatasetId, datasets, allStudents]);

  const filteredStudents = useMemo(() => {
    return visibleStudents.filter((student) => {
      const matchesQuery = `${student.name} ${student.id} ${student.school} ${student.sourceFile}`
        .toLowerCase()
        .includes(query.toLowerCase());

      const matchesGrade = gradeFilter === "all" || student.grade === gradeFilter;
      return matchesQuery && matchesGrade;
    });
  }, [visibleStudents, query, gradeFilter]);

  const sortedStudents = useMemo(() => {
    const getRiskLabel = (risk: number) => {
      if (risk >= 70) return "High";
      if (risk >= 35) return "Moderate";
      return "Low";
    };

    const riskOrder: Record<string, number> = {
      High: 0,
      Moderate: 1,
      Low: 2,
    };

    return [...filteredStudents].sort((a, b) => {
      let result = 0;

      switch (sortKey) {
        case "student":
          result = a.name.localeCompare(b.name, undefined, { sensitivity: "base", numeric: true });
          break;

        case "grade":
          result = String(a.grade).localeCompare(String(b.grade), undefined, {
            numeric: true,
            sensitivity: "base",
          });
          break;

        case "attendance":
          result = a.attendance - b.attendance;
          break;

        case "gpa":
          result = a.gpa - b.gpa;
          break;

        case "risk":
          result = riskOrder[getRiskLabel(a.risk)] - riskOrder[getRiskLabel(b.risk)];
          break;

        case "file":
          result = a.sourceFile.localeCompare(b.sourceFile, undefined, {
            sensitivity: "base",
            numeric: true,
          });
          break;

        default:
          result = 0;
      }

      return sortDirection === "asc" ? result : -result;
    });
  }, [filteredStudents, sortKey, sortDirection]);

  useEffect(() => {
    if (!sortedStudents.length) {
      setSelectedId("");
      return;
    }
    const exists = sortedStudents.some((student) => student.id === selectedId);
    if (!exists) setSelectedId(sortedStudents[0].id);
  }, [sortedStudents, selectedId]);

  const selectedStudent =
    sortedStudents.find((student) => student.id === selectedId) ||
    sortedStudents[0] ||
    visibleStudents[0];

  const selectedStudentNote = selectedStudent ? studentNotes[selectedStudent.id] ?? "" : "";

  const high = visibleStudents.filter((student) => student.risk >= 70).length;
  const moderate = visibleStudents.filter((student) => student.risk >= 35 && student.risk < 70).length;
  const low = visibleStudents.filter((student) => student.risk < 35).length;

  const avgAttendance = visibleStudents.length
    ? Math.round(
        visibleStudents.reduce((sum, student) => sum + student.attendance, 0) / visibleStudents.length
      )
    : 0;

  const pieData = [
    { name: "Low", value: low },
    { name: "Moderate", value: moderate },
    { name: "High", value: high },
  ];

  const barData = (() => {
    const grouped = visibleStudents.reduce<Record<string, { total: number; count: number }>>(
      (acc, student) => {
        const grade = student.grade || "Unknown";
        if (!acc[grade]) acc[grade] = { total: 0, count: 0 };
        acc[grade].total += student.risk;
        acc[grade].count += 1;
        return acc;
      },
      {}
    );

    return Object.keys(grouped)
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .map((grade) => ({
        grade,
        avgRisk: Math.round(grouped[grade].total / grouped[grade].count),
      }));
  })();

  const gradeOptions = [...new Set(visibleStudents.map((student) => student.grade))]
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  function updateStudentNote(studentId: string, value: string) {
    setStudentNotes((prev) => ({
      ...prev,
      [studentId]: value,
    }));
  }

  function renderOverviewTab() {
    return (
      <>
        <div className="grid gap-6 lg:grid-cols-[1.35fr_0.65fr]">
          <div className="rounded-[32px] border border-slate-200 bg-gradient-to-br from-slate-950 via-slate-900 to-slate-800 p-8 text-white shadow-sm md:p-10">
            <div className="flex flex-wrap gap-2">
              <Badge className="rounded-full border-0 bg-white/10 text-white hover:bg-white/10">
                Multi-file upload
              </Badge>
              <Badge className="rounded-full border-0 bg-white/10 text-white hover:bg-white/10">
                Student notes
              </Badge>
              <Badge className="rounded-full border-0 bg-white/10 text-white hover:bg-white/10">
                Early warning dashboard
              </Badge>
            </div>

            <div className="mt-6 max-w-3xl">
              <h1 className="text-4xl font-semibold tracking-tight md:text-5xl">
                Preventative Student Dropout Model
              </h1>
              <p className="mt-4 text-base text-slate-300 md:text-lg">
                Upload one or more student files, compare cohorts, and keep reccomendations attached
                to individual students.
              </p>
            </div>

            <div className="mt-8 flex flex-wrap gap-3">
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-full bg-white px-5 py-3 text-sm font-medium text-slate-950 hover:bg-slate-100">
                <Upload className="h-4 w-4" />
                Add student files
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  className="hidden"
                  onChange={handleFilesUpload}
                />
              </label>

              <div className="inline-flex items-center gap-2 rounded-full border border-white/15 px-5 py-3 text-sm text-slate-200">
                <ShieldCheck className="h-4 w-4" />
                Human review first
              </div>
            </div>

            <div className="mt-8 grid gap-3 sm:grid-cols-3">
              <div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                <p className="text-sm text-slate-300">Students</p>
                <p className="mt-1 text-2xl font-semibold">{visibleStudents.length}</p>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                <p className="text-sm text-slate-300">High support</p>
                <p className="mt-1 text-2xl font-semibold">{high}</p>
              </div>
              <div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                <p className="text-sm text-slate-300">Files loaded</p>
                <p className="mt-1 text-2xl font-semibold">{datasets.length}</p>
              </div>
            </div>
          </div>

          <Card className="rounded-[28px] border border-slate-200 bg-white shadow-sm">
            <CardHeader>
              <CardTitle>Workspace Controls</CardTitle>
              <CardDescription>Choose how uploaded files are shown</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div>
                <p className="mb-2 text-sm font-medium text-slate-700">View mode</p>
                <div className="grid grid-cols-2 gap-2">
                  <Button
                    variant={viewMode === "combined" ? "default" : "outline"}
                    className="rounded-2xl"
                    onClick={() => {
                      setViewMode("combined");
                      setActiveDatasetId("all");
                    }}
                  >
                    Combined
                  </Button>
                  <Button
                    variant={viewMode === "single" ? "default" : "outline"}
                    className="rounded-2xl"
                    onClick={() => {
                      setViewMode("single");
                      if (datasets.length && activeDatasetId === "all") {
                        setActiveDatasetId(datasets[0].id);
                      }
                    }}
                  >
                    Single file
                  </Button>
                </div>
              </div>

              <div>
                <p className="mb-2 text-sm font-medium text-slate-700">Active dataset</p>
                <Select
                  value={viewMode === "combined" ? "all" : activeDatasetId}
                  onValueChange={setActiveDatasetId}
                  disabled={viewMode === "combined" || !datasets.length}
                >
                  <SelectTrigger className="rounded-2xl">
                    <SelectValue placeholder="Select file" />
                  </SelectTrigger>
                  <SelectContent>
                    {datasets.map((dataset) => (
                      <SelectItem key={dataset.id} value={dataset.id}>
                        {dataset.fileName}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-600">
                Combined mode merges all uploaded students. Single file mode lets you inspect one
                dataset at a time.
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="mt-6 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <StatCard
            title="Students monitored"
            value={String(visibleStudents.length)}
            subtitle={viewMode === "combined" ? "Across all files" : "In selected file"}
            icon={Users}
          />
          <StatCard
            title="High-support students"
            value={String(high)}
            subtitle="Need immediate review"
            icon={AlertTriangle}
          />
          <StatCard
            title="Moderate-support students"
            value={String(moderate)}
            subtitle="Track closely"
            icon={TrendingDown}
          />
          <StatCard
            title="Average attendance"
            value={`${avgAttendance}%`}
            subtitle="Across current view"
            icon={CalendarCheck2}
          />
        </div>

        <div className="mt-8 grid gap-6 lg:grid-cols-2">
          <Card className="min-w-0 rounded-[24px] border border-slate-200 bg-white shadow-sm">
            <CardHeader>
              <CardTitle>Support distribution</CardTitle>
              <CardDescription>Low, moderate, and high support need</CardDescription>
            </CardHeader>
            <CardContent className="h-[320px] min-w-0">
              {visibleStudents.length ? (
                <div className="h-full w-full min-w-0">
                  <ResponsiveContainer width="100%" height="100%" minWidth={280} minHeight={280}>
                    <PieChart>
                      <Pie
                        data={pieData}
                        dataKey="value"
                        nameKey="name"
                        outerRadius={100}
                        innerRadius={55}
                        paddingAngle={4}
                      >
                        {pieData.map((_, index) => (
                          <Cell key={index} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              ) : (
                <div className="flex h-full items-center justify-center rounded-2xl border border-dashed border-slate-200 text-sm text-slate-500">
                  Upload files to see charts
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      </>
    );
  }

  function renderFilesTab() {
    return (
      <div className="grid gap-6 xl:grid-cols-[0.7fr_1.3fr]">
        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <CardTitle>Upload files</CardTitle>
            <CardDescription>Add Excel or CSV data files</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <label className="flex cursor-pointer flex-col items-center justify-center rounded-3xl border border-dashed border-slate-300 bg-slate-50 px-6 py-10 text-center">
              <Upload className="mb-4 h-8 w-8 text-slate-700" />
              <p className="font-medium text-slate-900">Upload student files</p>
              <p className="mt-1 text-sm text-slate-500">
                Supports .xlsx, .xls, and .csv
              </p>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                multiple
                className="hidden"
                onChange={handleFilesUpload}
              />
            </label>

            <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-600">
              Duplicate file detection is enabled. If the same file is uploaded again, it will be skipped.
            </div>
          </CardContent>
        </Card>

        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <CardTitle>Uploaded files</CardTitle>
            <CardDescription>Switch or delete datasets</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3">
            {datasets.length ? (
              datasets.map((dataset) => (
                <div
                  key={dataset.id}
                  className="flex items-center justify-between rounded-2xl border border-slate-200 p-4"
                >
                  <div className="min-w-0">
                    <p className="truncate font-medium text-slate-900">{dataset.fileName}</p>
                    <p className="text-sm text-slate-500">
                      {dataset.students.length} students · {(dataset.fileSize / 1024).toFixed(1)} KB
                    </p>
                  </div>

                  <div className="flex items-center gap-2">
                    <Button
                      variant="outline"
                      className="rounded-xl"
                      onClick={() => {
                        setViewMode("single");
                        setActiveDatasetId(dataset.id);
                        setActiveTab("students");
                      }}
                    >
                      View
                    </Button>
                    <Button
                      variant="outline"
                      size="icon"
                      className="rounded-xl"
                      onClick={() => deleteDataset(dataset.id)}
                    >
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  </div>
                </div>
              ))
            ) : (
              <div className="rounded-2xl border border-dashed border-slate-200 p-8 text-center text-sm text-slate-500">
                No files uploaded yet
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    );
  }

  function renderStudentsTab() {
    return (
      <div className="grid gap-6 xl:grid-cols-[1.1fr_0.9fr]">
        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <CardTitle>Students</CardTitle>
                <CardDescription>Search, filter, and inspect records</CardDescription>
              </div>

              <div className="flex flex-col gap-3 sm:flex-row">
                <div className="relative min-w-[220px]">
                  <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                  <Input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search by name, ID, school, or file"
                    className="rounded-2xl pl-9"
                  />
                </div>

                <Select value={gradeFilter} onValueChange={setGradeFilter}>
                  <SelectTrigger className="w-[150px] rounded-2xl">
                    <SelectValue placeholder="Grade" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All grades</SelectItem>
                    {gradeOptions.map((grade) => (
                      <SelectItem key={grade} value={grade}>
                        Grade {grade}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
          </CardHeader>

          <CardContent>
            <div className="overflow-hidden rounded-2xl border border-slate-200">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead onClick={() => handleSort("student")} className="cursor-pointer">
                      Student {sortKey === "student" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>

                    <TableHead onClick={() => handleSort("grade")} className="cursor-pointer">
                      Grade {sortKey === "grade" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>

                    <TableHead onClick={() => handleSort("attendance")} className="cursor-pointer">
                      Attendance {sortKey === "attendance" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>

                    <TableHead onClick={() => handleSort("gpa")} className="cursor-pointer">
                      GPA {sortKey === "gpa" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>

                    <TableHead onClick={() => handleSort("risk")} className="cursor-pointer">
                      Risk {sortKey === "risk" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>

                    <TableHead onClick={() => handleSort("file")} className="cursor-pointer">
                      File {sortKey === "file" ? (sortDirection === "asc" ? "↑" : "↓") : ""}
                    </TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredStudents.length ? (
                    sortedStudents.map((student) => (
                      <TableRow
                        key={student.id}
                        className={`cursor-pointer hover:bg-slate-50 ${
                          selectedStudent?.id === student.id ? "bg-slate-50" : ""
                        }`}
                        onClick={() => setSelectedId(student.id)}
                      >
                        <TableCell>
                          <div>
                            <p className="font-medium text-slate-900">{student.name}</p>
                            <p className="text-xs text-slate-500">{student.id}</p>
                          </div>
                        </TableCell>
                        <TableCell>{student.grade}</TableCell>
                        <TableCell>{student.attendance}%</TableCell>
                        <TableCell>{student.gpa}</TableCell>
                        <TableCell>
                          <Badge variant={riskVariant(student.risk)}>
                            {riskLabel(student.risk)} · {student.risk}
                          </Badge>
                        </TableCell>
                        <TableCell className="max-w-[160px] truncate text-slate-500">
                          {student.sourceFile}
                        </TableCell>
                      </TableRow>
                    ))
                  ) : (
                    <TableRow>
                      <TableCell colSpan={6} className="py-10 text-center text-slate-500">
                        Upload files to populate the table
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>

        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <CardTitle>Student profile</CardTitle>
            <CardDescription>Selected student summary and notes</CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {selectedStudent ? (
              <>
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <h3 className="text-2xl font-semibold text-slate-950">{selectedStudent.name}</h3>
                    <p className="text-sm text-slate-500">
                      {selectedStudent.id} · Grade {selectedStudent.grade} · {selectedStudent.school}
                    </p>
                    <p className="mt-1 text-xs text-slate-400">
                      Source file: {selectedStudent.sourceFile}
                    </p>
                  </div>
                  <Badge variant={riskVariant(selectedStudent.risk)} className="rounded-full px-3 py-1">
                    {riskLabel(selectedStudent.risk)} support
                  </Badge>
                </div>

                <div>
                  <div className="mb-2 flex items-center justify-between text-sm">
                    <span>Support risk score</span>
                    <span className="font-medium">{selectedStudent.risk}/100</span>
                  </div>
                  <Progress value={selectedStudent.risk} />
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <p className="text-sm text-slate-500">Attendance</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.attendance}%
                    </p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <p className="text-sm text-slate-500">GPA</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.gpa}
                    </p>
                  </div>
                </div>

                <div>
                  <p className="text-sm font-medium text-slate-950">Notes / comments</p>
                  <div className="mt-3">
                  <textarea
                    value={selectedStudentNote}
                    onChange={(e) => updateStudentNote(selectedStudent.id, e.target.value)}
                    placeholder="Add counselor notes, outreach updates, family contact summaries, or follow-up plans..."
                    className="min-h-[180px] w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none ring-0 placeholder:text-slate-400 focus:border-slate-400"
                  />
                  </div>
                  <p className="mt-2 text-xs text-slate-500">
                    Notes are saved in this browser using localStorage.
                  </p>
                </div>
              </>
            ) : (
              <div className="rounded-2xl border border-dashed border-slate-200 p-8 text-center text-sm text-slate-500">
                Upload files and select a student
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    );
  }

  function renderInterventionsTab() {
    return (
      <div className="grid gap-6 xl:grid-cols-[1fr_1fr]">
        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <CardTitle>Student profile</CardTitle>
            <CardDescription>Main drivers behind the current risk level</CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {selectedStudent ? (
              <>
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <h3 className="text-2xl font-semibold text-slate-950">{selectedStudent.name}</h3>
                    <p className="text-sm text-slate-500">
                      {selectedStudent.id} · Grade {selectedStudent.grade} · {selectedStudent.school}
                    </p>
                  </div>
                  <Badge variant={riskVariant(selectedStudent.risk)} className="rounded-full px-3 py-1">
                    {riskLabel(selectedStudent.risk)} support
                  </Badge>
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <p className="text-sm text-slate-500">Engagement</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.engagement}
                    </p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <p className="text-sm text-slate-500">Behavior</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.behavior}
                    </p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-white p-4">
                    <p className="text-sm text-slate-500">Tardies</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.tardies}
                    </p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-white p-4">
                    <p className="text-sm text-slate-500">Failed courses</p>
                    <p className="mt-1 text-2xl font-semibold text-slate-950">
                      {selectedStudent.failedCourses}
                    </p>
                  </div>
                </div>

                <div>
                  <p className="text-sm font-medium text-slate-950">Main drivers</p>
                  <div className="mt-3 space-y-2">
                    {selectedStudent.drivers.map((driver) => (
                      <div
                        key={driver}
                        className="rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                      >
                        {driver}
                      </div>
                    ))}
                  </div>
                </div>
              </>
            ) : (
              <div className="rounded-2xl border border-dashed border-slate-200 p-8 text-center text-sm text-slate-500">
                Select a student to review drivers
              </div>
            )}
          </CardContent>
        </Card>

        <Card className="rounded-[24px] border border-slate-200 bg-white shadow-sm">
          <CardHeader>
            <CardTitle>Suggested interventions</CardTitle>
            <CardDescription>Action ideas for the selected student</CardDescription>
          </CardHeader>
          <CardContent>
            {selectedStudent ? (
              <div className="space-y-3">
                {selectedStudent.suggestions.map((suggestion, index) => (
                  <div
                    key={`${suggestion}-${index}`}
                    className="flex items-start gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-4"
                  >
                    <div className="mt-0.5 rounded-full bg-slate-950 p-1 text-white">
                      <ChevronRight className="h-3 w-3" />
                    </div>
                    <p className="text-sm text-slate-700">{suggestion}</p>
                  </div>
                ))}

                <div className="mt-4 rounded-2xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900">
                  Recommendation: use these suggestions to support planning, not as automatic decisions.
                </div>
              </div>
            ) : (
              <div className="rounded-2xl border border-dashed border-slate-200 p-8 text-center text-sm text-slate-500">
                Select a student to see intervention suggestions
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-white text-slate-950">
      <div className="border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex max-w-7xl flex-col gap-4 px-4 py-4 md:px-8">
          <div className="flex items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="rounded-2xl border border-slate-200 bg-slate-950 p-2 text-white shadow-sm">
                <Sparkles className="h-5 w-5" />
              </div>
              <div>
                <div className="text-lg font-semibold tracking-tight">OnTrack</div>
                <div className="text-sm text-slate-500">Student Support Intelligence</div>
              </div>
            </div>
  
            <div className="flex items-center gap-3">
              {user ? (
                <span className="hidden text-sm text-slate-500 md:inline">{user.email}</span>
              ) : (
                <a
                  href="/login"
                  className="rounded-full border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
                >
                  Log in to save
                </a>
              )}
  
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-full bg-slate-950 px-4 py-2 text-sm font-medium text-white shadow-sm hover:bg-slate-800">
                <Upload className="h-4 w-4" />
                Upload files
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  className="hidden"
                  onChange={handleFilesUpload}
                />
              </label>
  
              {user ? (
                <button
                  onClick={async () => {
                    await supabase.auth.signOut();
                    window.location.reload();
                  }}
                  className="rounded-full border border-slate-300 px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
                >
                  Sign out
                </button>
              ) : null}
            </div>
          </div>
  
          <div className="flex flex-wrap items-center gap-2">
            <TabButton active={activeTab === "overview"} onClick={() => setActiveTab("overview")}>
              <LayoutGrid className="mr-2 inline h-4 w-4" />
              Overview
            </TabButton>
  
            <TabButton active={activeTab === "files"} onClick={() => setActiveTab("files")}>
              <FileSpreadsheet className="mr-2 inline h-4 w-4" />
              Files
            </TabButton>
  
            <TabButton active={activeTab === "students"} onClick={() => setActiveTab("students")}>
              <Users className="mr-2 inline h-4 w-4" />
              Students
            </TabButton>
  
            <TabButton
              active={activeTab === "interventions"}
              onClick={() => setActiveTab("interventions")}
            >
              <MessageSquare className="mr-2 inline h-4 w-4" />
              Interventions
            </TabButton>
          </div>
        </div>
      </div>
  
      <div className="mx-auto max-w-7xl px-4 py-8 md:px-8">
        {error ? (
          <Card className="mb-6 rounded-[24px] border border-red-200 bg-red-50 shadow-sm">
            <CardContent className="p-4 text-sm text-red-700">{error}</CardContent>
          </Card>
        ) : null}
  
        {duplicateNotice.length ? (
          <Card className="mb-6 rounded-[24px] border border-amber-200 bg-amber-50 shadow-sm">
            <CardContent className="p-4 text-sm text-amber-800">
              Skipped duplicate file{duplicateNotice.length > 1 ? "s" : ""}:{" "}
              <span className="font-medium">{duplicateNotice.join(", ")}</span>
            </CardContent>
          </Card>
        ) : null}
  
        {activeTab === "overview" && renderOverviewTab()}
        {activeTab === "files" && renderFilesTab()}
        {activeTab === "students" && renderStudentsTab()}
        {activeTab === "interventions" && renderInterventionsTab()}
      </div>
    </div>
  );
}