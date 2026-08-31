import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Dumbbell, Upload, Users, Play, ArrowLeft, Search, Plus, X, LogOut,
  Download, ChevronRight, FileSpreadsheet, Check, Trash2, Pencil, Save,
  AlertTriangle, CalendarDays, Clock, UserPlus, CheckCircle2, MessageCircle, TrendingUp
} from "lucide-react";
import { initializeApp } from "firebase/app";
import {
  getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  onAuthStateChanged, signOut,
} from "firebase/auth";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, deleteDoc,
  collection, getDocs, query, where, serverTimestamp, arrayUnion,
} from "firebase/firestore";

/* ------------------------------------------------------------------ *
 *  TRAINING CLUB — dados reais no Firebase (Auth + Firestore)
 *  Login de professor e aluno, e os treinos ficam salvos na nuvem:
 *  qualquer pessoa que abrir o link com a própria conta vê os
 *  mesmos dados, mesmo saindo e entrando de novo depois.
 * ------------------------------------------------------------------ */

const firebaseConfig = {
  apiKey: "AIzaSyAFWPHG6uqt5aWs4rCpBGapTHyugWrhVNI",
  authDomain: "training-club-47aa2.firebaseapp.com",
  projectId: "training-club-47aa2",
  storageBucket: "training-club-47aa2.firebasestorage.app",
  messagingSenderId: "972073779965",
  appId: "1:972073779965:web:075c45c6f60475ffcaf494",
};

const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);

// App Firebase secundário: usado só para criar login de aluno sem
// derrubar a sessão do professor que está logado no app principal.
const secondaryApp = initializeApp(firebaseConfig, "Secondary");
const secondaryAuth = getAuth(secondaryApp);

const LOGO_SRC = "data:image/png;base64,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";

const uid = () => "id-" + Math.random().toString(36).slice(2, 10);
const stripAccents = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
const normHeader = (s) => stripAccents(String(s)).toLowerCase().trim().replace(/\s+/g, " ");

const HEADER_MAP = {
  "nome aluno": "nomeAluno", "aluno": "nomeAluno", "nome do aluno": "nomeAluno",
  "dia": "dia", "dia do treino": "dia",
  "exercicio": "exercicio", "nome do exercicio": "exercicio",
  "series": "series", "serie": "series",
  "repeticoes": "repeticoes", "repeticao": "repeticoes", "reps": "repeticoes",
  "link video": "linkVideo", "video": "linkVideo", "link do video": "linkVideo",
  "observacoes": "observacoes", "observacao": "observacoes", "obs": "observacoes",
};

function formatDate(ts) {
  if (!ts) return "—";
  return new Date(ts).toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" });
}
function initials(nome) {
  return nome.split(" ").filter(Boolean).slice(0, 2).map((w) => w[0]).join("").toUpperCase();
}
const PACOTES = { bronze: "Bronze", prata: "Prata", gold: "Gold" };
function acessoExpirado(acessoAte) {
  return !!acessoAte && Date.now() > acessoAte;
}
function acessoStatusLabel(acessoAte) {
  if (!acessoAte) return "sem prazo definido";
  const dias = Math.ceil((acessoAte - Date.now()) / 86400000);
  if (dias < 0) return `expirado há ${Math.abs(dias)} dia(s)`;
  if (dias === 0) return "expira hoje";
  return `expira em ${dias} dia(s)`;
}


export default function App() {
  const [authLoading, setAuthLoading] = useState(true);
  const [authUser, setAuthUser] = useState(null);
  const [authRole, setAuthRole] = useState(null); // "admin" | "aluno" | null
  const [screen, setScreen] = useState("role");
  const [students, setStudents] = useState([]);
  const [studentsLoading, setStudentsLoading] = useState(false);
  const [currentStudentId, setCurrentStudentId] = useState(null);
  const [viewStudentId, setViewStudentId] = useState(null);
  const [currentDiaIdx, setCurrentDiaIdx] = useState(0);
  const [search, setSearch] = useState("");
  const [uploadPreview, setUploadPreview] = useState(null);
  const [toast, setToast] = useState(null);
  const [showAddStudent, setShowAddStudent] = useState(false);
  const [newName, setNewName] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newSenha, setNewSenha] = useState("");
  const [newPacote, setNewPacote] = useState("bronze");
  const [newAcessoAte, setNewAcessoAte] = useState("");
  const [config, setConfig] = useState({ whatsappProfessor: "" });
  const [editingKey, setEditingKey] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [completed, setCompleted] = useState({});
  const [notes, setNotes] = useState({});
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap";
    document.head.appendChild(link);
    return () => document.head.removeChild(link);
  }, []);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (user) => {
      if (!user) {
        setAuthUser(null); setAuthRole(null); setAuthLoading(false);
        return;
      }
      try {
        const profDoc = await getDoc(doc(db, "professores", user.uid));
        if (profDoc.exists()) {
          setAuthUser(user); setAuthRole("admin");
          setScreen("adminDashboard");
        } else {
          const alunoDoc = await getDoc(doc(db, "alunos", user.uid));
          if (alunoDoc.exists()) {
            setAuthUser(user); setAuthRole("aluno");
            setCurrentStudentId(user.uid);
            const alunoData = alunoDoc.data();
            setStudents((prev) => mergeStudent(prev, { id: user.uid, ...alunoData }));
            const prog = alunoData.progresso || {};
            const completedObj = {}; const notesObj = {};
            Object.entries(prog.completed || {}).forEach(([k, v]) => { completedObj[`${user.uid}-${k.replace("_", "-")}`] = v; });
            Object.entries(prog.notes || {}).forEach(([k, v]) => { notesObj[`${user.uid}-${k.replace("_", "-")}`] = v; });
            setCompleted(completedObj); setNotes(notesObj);
            setScreen("studentDashboard");
          } else {
            setAuthUser(user); setAuthRole(null);
          }
        }
      } catch (e) {
        setAuthUser(user); setAuthRole(null);
      }
      setAuthLoading(false);
    });
    return () => unsub();
  }, []);

  function mergeStudent(list, student) {
    const idx = list.findIndex((s) => s.id === student.id);
    if (idx < 0) return [...list, student];
    const next = [...list];
    next[idx] = student;
    return next;
  }

  async function loadAllStudents() {
    setStudentsLoading(true);
    try {
      const snap = await getDocs(collection(db, "alunos"));
      setStudents(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    } catch (e) {
      showToast("error", "Não foi possível carregar os alunos.");
    }
    setStudentsLoading(false);
  }

  useEffect(() => {
    if (authRole === "admin") loadAllStudents();
  }, [authRole]);

  async function loadConfig() {
    try {
      const snap = await getDoc(doc(db, "config", "geral"));
      if (snap.exists()) setConfig(snap.data());
    } catch (e) {}
  }

  useEffect(() => {
    if (authRole === "admin" || authRole === "aluno") loadConfig();
  }, [authRole]);

  async function saveWhatsapp(numero) {
    try {
      await setDoc(doc(db, "config", "geral"), { whatsappProfessor: numero }, { merge: true });
      setConfig((prev) => ({ ...prev, whatsappProfessor: numero }));
      showToast("success", "WhatsApp de suporte atualizado.");
    } catch (e) {
      showToast("error", "Não foi possível salvar o WhatsApp de suporte.");
    }
  }

  const showToast = (type, msg) => setToast({ type, msg });

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 3200);
    return () => clearTimeout(t);
  }, [toast]);

  const currentStudent = students.find((s) => s.id === currentStudentId) || null;
  const viewStudent = students.find((s) => s.id === viewStudentId) || null;

  /* ---------------- Upload / parse da planilha ---------------- */

  function buildFieldMap(keys) {
    const map = {};
    keys.forEach((k) => {
      const canon = HEADER_MAP[normHeader(k)];
      if (canon) map[canon] = k;
    });
    return map;
  }

  async function handleFile(file) {
    if (!file) return;
    const okExt = /\.(xlsx|xls|csv)$/i.test(file.name);
    if (!okExt) {
      showToast("error", "Formato inválido. Envie um arquivo .xlsx, .xls ou .csv.");
      return;
    }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) {
        showToast("error", "A planilha está vazia.");
        return;
      }
      const map = buildFieldMap(Object.keys(rows[0]));
      if (!map.nomeAluno || !map.dia || !map.exercicio) {
        showToast("error", "Não encontrei as colunas obrigatórias (Nome Aluno, Dia, Exercício). Confira o modelo.");
        return;
      }
      const errors = [];
      const byStudent = {};
      rows.forEach((row, i) => {
        const nomeAluno = String(row[map.nomeAluno] ?? "").trim();
        const diaRaw = String(row[map.dia] ?? "").trim();
        const exercicio = String(row[map.exercicio] ?? "").trim();
        if (!nomeAluno || !diaRaw || !exercicio) {
          errors.push(`Linha ${i + 2}: faltam dados obrigatórios (aluno, dia ou exercício).`);
          return;
        }
        const series = Number(row[map.series]) || 0;
        const repeticoes = Number(row[map.repeticoes]) || 0;
        const linkVideo = String(row[map.linkVideo] ?? "").trim();
        const observacoes = String(row[map.observacoes] ?? "").trim();
        const diaNumMatch = diaRaw.match(/\d+/);
        if (!byStudent[nomeAluno]) byStudent[nomeAluno] = {};
        if (!byStudent[nomeAluno][diaRaw]) {
          byStudent[nomeAluno][diaRaw] = {
            dia: diaNumMatch ? parseInt(diaNumMatch[0], 10) : Object.keys(byStudent[nomeAluno]).length + 1,
            nome: diaRaw, exercicios: [],
          };
        }
        byStudent[nomeAluno][diaRaw].exercicios.push({ nome: exercicio, series, repeticoes, linkVideo, observacoes });
      });
      const preview = Object.entries(byStudent).map(([nome, diasObj]) => {
        const dias = Object.values(diasObj).sort((a, b) => a.dia - b.dia);
        const existing = students.find((s) => s.nome.toLowerCase() === nome.toLowerCase());
        return {
          nome, isNew: !existing,
          totalExercicios: dias.reduce((a, d) => a + d.exercicios.length, 0),
          dias,
        };
      });
      if (!preview.length) {
        showToast("error", "Nenhum registro válido encontrado na planilha.");
        return;
      }
      setUploadPreview({ entries: preview, errors, fileName: file.name });
      setScreen("adminUpload");
    } catch (e) {
      showToast("error", "Não foi possível ler o arquivo. Verifique o formato.");
    }
  }

  async function confirmUpload() {
    const naoVinculados = uploadPreview.entries.filter((e) => e.isNew);
    if (naoVinculados.length) {
      showToast("error", `${naoVinculados.length} aluno(s) da planilha ainda não têm conta cadastrada. Cadastre-os primeiro em "Novo aluno" (com e-mail e senha) e suba a planilha de novo.`);
    }
    const vinculados = uploadPreview.entries.filter((e) => !e.isNew);
    try {
      for (const entry of vinculados) {
        const alvo = students.find((s) => s.nome.toLowerCase() === entry.nome.toLowerCase());
        if (!alvo) continue;
        await updateDoc(doc(db, "alunos", alvo.id), {
          treino: { dias: entry.dias, ultimaAtualizacao: Date.now() },
        });
      }
      if (vinculados.length) showToast("success", `Treino atualizado para ${vinculados.length} aluno(s).`);
      await loadAllStudents();
    } catch (e) {
      showToast("error", "Não foi possível salvar o treino no banco de dados.");
    }
    setUploadPreview(null);
    setScreen("adminDashboard");
  }

  function baixarModelo() {
    const data = [
      { "Nome Aluno": "João Silva", "Dia": "Dia 1", "Exercício": "Supino Reto", "Séries": 4, "Repetições": 12, "Link Vídeo": "https://youtube.com/results?search_query=supino+reto", "Observações": "Executar lento" },
      { "Nome Aluno": "João Silva", "Dia": "Dia 1", "Exercício": "Crucifixo", "Séries": 4, "Repetições": 15, "Link Vídeo": "https://youtube.com/results?search_query=crucifixo", "Observações": "" },
      { "Nome Aluno": "João Silva", "Dia": "Dia 2", "Exercício": "Puxada Frontal", "Séries": 4, "Repetições": 12, "Link Vídeo": "https://youtube.com/results?search_query=puxada+frontal", "Observações": "" },
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Treinos");
    XLSX.writeFile(wb, "planilha-modelo-treino.xlsx");
  }

  /* ---------------- CRUD alunos / exercícios (admin) ---------------- */

  async function addStudent() {
    if (!newName.trim() || !newEmail.trim() || !newSenha.trim()) {
      showToast("error", "Preencha nome, e-mail e senha do aluno.");
      return;
    }
    if (newSenha.trim().length < 6) {
      showToast("error", "A senha precisa ter pelo menos 6 caracteres.");
      return;
    }
    try {
      const cred = await createUserWithEmailAndPassword(secondaryAuth, newEmail.trim(), newSenha.trim());
      const newUid = cred.user.uid;
      await setDoc(doc(db, "alunos", newUid), {
        nome: newName.trim(), email: newEmail.trim(), pacote: newPacote,
        acessoAte: newAcessoAte ? new Date(newAcessoAte + "T23:59:59").getTime() : null,
        treino: { dias: [], ultimaAtualizacao: null },
      });
      await signOut(secondaryAuth);
      setNewName(""); setNewEmail(""); setNewSenha(""); setNewPacote("bronze"); setNewAcessoAte(""); setShowAddStudent(false);
      showToast("success", "Aluno cadastrado. Já pode entrar com o e-mail e senha dele.");
      await loadAllStudents();
    } catch (e) {
      const msg = e.code === "auth/email-already-in-use"
        ? "Esse e-mail já está cadastrado."
        : e.code === "auth/invalid-email"
        ? "E-mail inválido."
        : "Não foi possível cadastrar o aluno.";
      showToast("error", msg);
    }
  }

  async function removeStudent(id) {
    try {
      await deleteDoc(doc(db, "alunos", id));
      setStudents((prev) => prev.filter((s) => s.id !== id));
      if (viewStudentId === id) { setViewStudentId(null); setScreen("adminDashboard"); }
      showToast("success", "Aluno removido do sistema. (O login dele continua existindo no Firebase Auth — apague manualmente lá se quiser revogar o acesso.)");
    } catch (e) {
      showToast("error", "Não foi possível remover o aluno.");
    }
  }

  async function saveTreino(studentId, updater) {
    const student = students.find((s) => s.id === studentId);
    if (!student) return;
    const novoTreino = updater(student.treino);
    try {
      await updateDoc(doc(db, "alunos", studentId), { treino: novoTreino });
      setStudents((prev) => prev.map((s) => s.id === studentId ? { ...s, treino: novoTreino } : s));
    } catch (e) {
      showToast("error", "Não foi possível salvar a alteração.");
    }
  }

  function removeDia(studentId, diaIdx) {
    saveTreino(studentId, (treino) => ({
      ...treino, dias: treino.dias.filter((_, i) => i !== diaIdx),
    }));
  }

  function removeExercicio(studentId, diaIdx, exIdx) {
    saveTreino(studentId, (treino) => ({
      ...treino, dias: treino.dias.map((d, di) => di !== diaIdx ? d : {
        ...d, exercicios: d.exercicios.filter((_, i) => i !== exIdx),
      }),
    }));
  }

  function startEdit(diaIdx, exIdx, ex) {
    setEditingKey(`${diaIdx}-${exIdx}`);
    setEditDraft({ ...ex });
  }
  function saveEdit(studentId, diaIdx, exIdx) {
    saveTreino(studentId, (treino) => ({
      ...treino, dias: treino.dias.map((d, di) => di !== diaIdx ? d : {
        ...d, exercicios: d.exercicios.map((e, ei) => ei !== exIdx ? e : {
          ...editDraft, series: Number(editDraft.series) || 0, repeticoes: Number(editDraft.repeticoes) || 0,
        }),
      }),
    })).then(() => showToast("success", "Exercício atualizado."));
    setEditingKey(null); setEditDraft(null);
  }

  /* ---------------- Aluno: concluído / anotações ---------------- */

  function toggleCompleted(diaIdx, exIdx) {
    const key = `${currentStudentId}-${diaIdx}-${exIdx}`;
    const value = !completed[key];
    setCompleted((prev) => ({ ...prev, [key]: value }));
    updateDoc(doc(db, "alunos", currentStudentId), { [`progresso.completed.${diaIdx}_${exIdx}`]: value }).catch(() => {});
  }
  function setNote(diaIdx, exIdx, value) {
    const key = `${currentStudentId}-${diaIdx}-${exIdx}`;
    setNotes((prev) => ({ ...prev, [key]: value }));
    updateDoc(doc(db, "alunos", currentStudentId), { [`progresso.notes.${diaIdx}_${exIdx}`]: value }).catch(() => {});
  }

  async function setPacote(studentId, pacote) {
    setStudents((prev) => prev.map((s) => s.id !== studentId ? s : { ...s, pacote }));
    try {
      await updateDoc(doc(db, "alunos", studentId), { pacote });
      showToast("success", "Pacote atualizado.");
    } catch (e) {
      showToast("error", "Não foi possível atualizar o pacote.");
    }
  }

  async function setAcessoAte(studentId, dateStr) {
    const ts = dateStr ? new Date(dateStr + "T23:59:59").getTime() : null;
    setStudents((prev) => prev.map((s) => s.id !== studentId ? s : { ...s, acessoAte: ts }));
    try {
      await updateDoc(doc(db, "alunos", studentId), { acessoAte: ts });
      showToast("success", "Validade de acesso atualizada.");
    } catch (e) {
      showToast("error", "Não foi possível atualizar a validade.");
    }
  }

  async function registrarCarga(diaIdx, exIdx, valorStr) {
    const valor = Number(valorStr);
    if (!valor || valor <= 0) { showToast("error", "Informe um valor de carga válido."); return; }
    const key = `${diaIdx}_${exIdx}`;
    const entry = { valor, data: Date.now() };
    setStudents((prev) => prev.map((s) => s.id !== currentStudentId ? s : {
      ...s, progresso: { ...(s.progresso || {}), cargas: {
        ...(s.progresso?.cargas || {}), [key]: [...(s.progresso?.cargas?.[key] || []), entry],
      }},
    }));
    try {
      await updateDoc(doc(db, "alunos", currentStudentId), { [`progresso.cargas.${key}`]: arrayUnion(entry) });
    } catch (e) {
      showToast("error", "Não foi possível salvar a carga.");
    }
  }

  function setFeedbackAluno(diaIdx, texto) {
    setStudents((prev) => prev.map((s) => s.id !== currentStudentId ? s : {
      ...s, progresso: { ...(s.progresso || {}), feedback: {
        ...(s.progresso?.feedback || {}), [diaIdx]: { ...(s.progresso?.feedback?.[diaIdx] || {}), aluno: texto },
      }},
    }));
    updateDoc(doc(db, "alunos", currentStudentId), { [`progresso.feedback.${diaIdx}.aluno`]: texto }).catch(() => {});
  }

  function setFeedbackProfessor(studentId, diaIdx, texto) {
    setStudents((prev) => prev.map((s) => s.id !== studentId ? s : {
      ...s, progresso: { ...(s.progresso || {}), feedback: {
        ...(s.progresso?.feedback || {}), [diaIdx]: { ...(s.progresso?.feedback?.[diaIdx] || {}), professor: texto },
      }},
    }));
    updateDoc(doc(db, "alunos", studentId), { [`progresso.feedback.${diaIdx}.professor`]: texto }).catch(() => {});
  }

  function logout() {
    signOut(auth);
    setCurrentStudentId(null); setViewStudentId(null); setUploadPreview(null);
    setStudents([]); setCompleted({}); setNotes({});
    setScreen("role");
  }

  const filteredStudents = students.filter((s) => s.nome.toLowerCase().includes(search.toLowerCase()));

  /* ==================================================================== */

  return (
    <div className="ta-root">
      <Style />
      {toast && (
        <div className={`ta-toast ${toast.type}`}>
          {toast.type === "error" ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />}
          <span>{toast.msg}</span>
        </div>
      )}

      {authLoading && (
        <div className="ta-center">
          <img src={LOGO_SRC} alt="Training Club" className="ta-hero-logo" style={{ opacity: 0.7 }} />
        </div>
      )}

      {!authLoading && authUser && authRole === null && (
        <div className="ta-center">
          <div className="ta-auth-card" style={{ textAlign: "center" }}>
            <p className="ta-sub small">Essa conta não está vinculada a um professor nem a um aluno cadastrado no Training Club.</p>
            <button className="ta-btn ta-btn-primary ta-btn-block" onClick={logout}>Sair</button>
          </div>
        </div>
      )}

      {!authLoading && !authUser && screen === "role" && <RoleSelect onPick={(r) => setScreen(r === "admin" ? "adminLogin" : "studentLogin")} />}

      {!authUser && screen === "adminLogin" && (
        <AdminLogin
          onBack={() => setScreen("role")}
          onLogin={async (email, senha) => {
            try {
              await signInWithEmailAndPassword(auth, email, senha);
            } catch (e) {
              showToast("error", "E-mail ou senha incorretos.");
            }
          }}
        />
      )}

      {!authUser && screen === "studentLogin" && (
        <StudentLogin
          onBack={() => setScreen("role")}
          onLogin={async (email, senha) => {
            try {
              await signInWithEmailAndPassword(auth, email, senha);
            } catch (e) {
              showToast("error", "E-mail ou senha incorretos.");
            }
          }}
        />
      )}

      {authRole === "admin" && screen === "adminDashboard" && (
        <AdminDashboard
          students={filteredStudents}
          studentsLoading={studentsLoading}
          search={search} setSearch={setSearch}
          onLogout={logout}
          onOpenStudent={(id) => { setViewStudentId(id); setScreen("adminStudentView"); }}
          onRemoveStudent={removeStudent}
          onUploadClick={() => fileInputRef.current?.click()}
          onDropFile={handleFile}
          dragOver={dragOver} setDragOver={setDragOver}
          onDownloadModel={baixarModelo}
          showAddStudent={showAddStudent} setShowAddStudent={setShowAddStudent}
          newName={newName} setNewName={setNewName}
          newEmail={newEmail} setNewEmail={setNewEmail}
          newSenha={newSenha} setNewSenha={setNewSenha}
          newPacote={newPacote} setNewPacote={setNewPacote}
          newAcessoAte={newAcessoAte} setNewAcessoAte={setNewAcessoAte}
          onAddStudent={addStudent}
          config={config} onSaveWhatsapp={saveWhatsapp}
        />
      )}
      <input
        ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
        onChange={(e) => { handleFile(e.target.files?.[0]); e.target.value = ""; }}
      />

      {authRole === "admin" && screen === "adminUpload" && uploadPreview && (
        <AdminUploadPreview
          preview={uploadPreview}
          onCancel={() => { setUploadPreview(null); setScreen("adminDashboard"); }}
          onConfirm={confirmUpload}
        />
      )}

      {authRole === "admin" && screen === "adminStudentView" && viewStudent && (
        <AdminStudentView
          student={viewStudent}
          onBack={() => setScreen("adminDashboard")}
          editingKey={editingKey} editDraft={editDraft} setEditDraft={setEditDraft}
          onStartEdit={startEdit}
          onSaveEdit={(di, ei) => saveEdit(viewStudent.id, di, ei)}
          onCancelEdit={() => { setEditingKey(null); setEditDraft(null); }}
          onRemoveExercicio={(di, ei) => removeExercicio(viewStudent.id, di, ei)}
          onRemoveDia={(di) => removeDia(viewStudent.id, di)}
          onSetPacote={(pacote) => setPacote(viewStudent.id, pacote)}
          onSetAcessoAte={(dateStr) => setAcessoAte(viewStudent.id, dateStr)}
          onSetFeedbackProfessor={(di, texto) => setFeedbackProfessor(viewStudent.id, di, texto)}
        />
      )}

      {authRole === "aluno" && currentStudent && acessoExpirado(currentStudent.acessoAte) && (
        <AcessoExpirado student={currentStudent} config={config} onLogout={logout} />
      )}

      {authRole === "aluno" && screen === "studentDashboard" && currentStudent && !acessoExpirado(currentStudent.acessoAte) && (
        <StudentDashboard
          student={currentStudent}
          completed={completed}
          onLogout={logout}
          onOpenDia={(idx) => { setCurrentDiaIdx(idx); setScreen("studentDay"); }}
          config={config}
        />
      )}

      {authRole === "aluno" && screen === "studentDay" && currentStudent && !acessoExpirado(currentStudent.acessoAte) && currentStudent.treino.dias[currentDiaIdx] && (
        <StudentDayDetail
          student={currentStudent}
          diaIdx={currentDiaIdx}
          dia={currentStudent.treino.dias[currentDiaIdx]}
          totalDias={currentStudent.treino.dias.length}
          completed={completed}
          notes={notes}
          onToggleCompleted={toggleCompleted}
          onSetNote={setNote}
          onRegistrarCarga={registrarCarga}
          onSetFeedbackAluno={setFeedbackAluno}
          onBack={() => setScreen("studentDashboard")}
        />
      )}
    </div>
  );
}

/* ============================== TELAS ============================== */

function RoleSelect({ onPick }) {
  return (
    <div className="ta-center">
      <div className="ta-hero">
        <img src={LOGO_SRC} alt="Training Club" className="ta-hero-logo" />
        <h1 className="ta-h1">TRAINING&nbsp;<span className="ta-accent">CLUB</span></h1>
        <p className="ta-tagline">Assessoria de treino online</p>
        <div className="ta-role-grid">
          <button className="ta-role-card" onClick={() => onPick("admin")}>
            <Users size={26} />
            <span className="ta-role-title">Sou professor(a)</span>
            <span className="ta-role-desc">Gerenciar alunos e enviar planilhas de treino</span>
            <ChevronRight className="ta-role-chevron" size={18} />
          </button>
          <button className="ta-role-card" onClick={() => onPick("student")}>
            <Dumbbell size={26} />
            <span className="ta-role-title">Sou aluno(a)</span>
            <span className="ta-role-desc">Ver meus dias de treino e exercícios</span>
            <ChevronRight className="ta-role-chevron" size={18} />
          </button>
        </div>
      </div>
    </div>
  );
}

function AdminLogin({ onBack, onLogin }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [loading, setLoading] = useState(false);
  const handleEnter = async () => {
    if (!email.trim() || !senha.trim()) return;
    setLoading(true);
    await onLogin(email.trim(), senha.trim());
    setLoading(false);
  };
  return (
    <div className="ta-center">
      <div className="ta-auth-card">
        <button type="button" className="ta-back" onClick={onBack}><ArrowLeft size={16} /> voltar</button>
        <div className="ta-auth-icon"><Users size={22} /></div>
        <h2 className="ta-h2">Acesso do professor</h2>
        <p className="ta-sub small">Entre para gerenciar alunos e treinos.</p>
        <label className="ta-label">E-mail</label>
        <input className="ta-input" type="email" value={email} onChange={(e) => setEmail(e.target.value)} />
        <label className="ta-label">Senha</label>
        <input
          className="ta-input" type="password" placeholder="••••••••" value={senha}
          onChange={(e) => setSenha(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") handleEnter(); }}
        />
        <button className="ta-btn ta-btn-primary ta-btn-block" type="button" onClick={handleEnter} disabled={loading}>
          {loading ? "Entrando..." : "Entrar"}
        </button>
      </div>
    </div>
  );
}

function StudentLogin({ onBack, onLogin }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [loading, setLoading] = useState(false);
  const handleEnter = async () => {
    if (!email.trim() || !senha.trim()) return;
    setLoading(true);
    await onLogin(email.trim(), senha.trim());
    setLoading(false);
  };
  return (
    <div className="ta-center">
      <div className="ta-auth-card">
        <button type="button" className="ta-back" onClick={onBack}><ArrowLeft size={16} /> voltar</button>
        <div className="ta-auth-icon"><Dumbbell size={22} /></div>
        <h2 className="ta-h2">Acesso do aluno</h2>
        <p className="ta-sub small">Entre com o e-mail e senha cadastrados pelo seu professor.</p>
        <label className="ta-label">E-mail</label>
        <input className="ta-input" type="email" value={email} onChange={(e) => setEmail(e.target.value)} />
        <label className="ta-label">Senha</label>
        <input
          className="ta-input" type="password" placeholder="••••••••" value={senha}
          onChange={(e) => setSenha(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") handleEnter(); }}
        />
        <button className="ta-btn ta-btn-primary ta-btn-block" type="button" onClick={handleEnter} disabled={loading}>
          {loading ? "Entrando..." : "Entrar"}
        </button>
      </div>
    </div>
  );
}

function AdminDashboard(props) {
  const {
    students, studentsLoading, search, setSearch, onLogout, onOpenStudent, onRemoveStudent,
    onUploadClick, onDropFile, dragOver, setDragOver, onDownloadModel,
    showAddStudent, setShowAddStudent, newName, setNewName, newEmail, setNewEmail,
    newSenha, setNewSenha, newPacote, setNewPacote, newAcessoAte, setNewAcessoAte, onAddStudent,
    config, onSaveWhatsapp,
  } = props;
  const [showSuporteConfig, setShowSuporteConfig] = useState(false);
  const [whatsDraft, setWhatsDraft] = useState(config?.whatsappProfessor || "");
  useEffect(() => { setWhatsDraft(config?.whatsappProfessor || ""); }, [config?.whatsappProfessor]);

  return (
    <div className="ta-page">
      <TopBar title="Painel do professor" icon={<Users size={18} />} onLogout={onLogout} />
      <div className="ta-wrap">
        <div
          className={`ta-dropzone ${dragOver ? "over" : ""}`}
          onClick={onUploadClick}
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); onDropFile(e.dataTransfer.files?.[0]); }}
        >
          <FileSpreadsheet size={26} />
          <div>
            <div className="ta-dz-title">Enviar planilha de treinos</div>
            <div className="ta-dz-sub">Arraste um arquivo aqui ou clique para escolher — .xlsx, .xls ou .csv</div>
          </div>
          <button type="button" className="ta-btn ta-btn-ghost" onClick={(e) => { e.stopPropagation(); onDownloadModel(); }}>
            <Download size={15} /> Planilha modelo
          </button>
        </div>

        <div className="ta-row-between">
          <div className="ta-search">
            <Search size={16} />
            <input placeholder="Buscar aluno..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <button className="ta-btn ta-btn-ghost" onClick={() => setShowSuporteConfig((v) => !v)}>WhatsApp de suporte</button>
          <button className="ta-btn ta-btn-primary" onClick={() => setShowAddStudent(true)}><UserPlus size={16} /> Novo aluno</button>
        </div>

        {showSuporteConfig && (
          <div className="ta-card ta-add-form">
            <p className="ta-sub small" style={{ margin: 0 }}>Número que aparece pros alunos dos planos Prata e Gold entrarem em contato direto (com código do país, só números — ex: 5541999999999).</p>
            <input className="ta-input" placeholder="55DDDNÚMERO" value={whatsDraft} onChange={(e) => setWhatsDraft(e.target.value)} />
            <div className="ta-row-end">
              <button className="ta-btn ta-btn-ghost" onClick={() => setShowSuporteConfig(false)}>Fechar</button>
              <button className="ta-btn ta-btn-primary" onClick={() => onSaveWhatsapp(whatsDraft.trim())}>Salvar</button>
            </div>
          </div>
        )}

        {showAddStudent && (
          <div className="ta-card ta-add-form">
            <input className="ta-input" placeholder="Nome do aluno" value={newName} onChange={(e) => setNewName(e.target.value)} autoFocus />
            <input className="ta-input" type="email" placeholder="E-mail do aluno" value={newEmail} onChange={(e) => setNewEmail(e.target.value)} />
            <input className="ta-input" type="password" placeholder="Senha (mín. 6 caracteres)" value={newSenha} onChange={(e) => setNewSenha(e.target.value)} />
            <label className="ta-label">Pacote</label>
            <div className="ta-pacote-picker">
              {Object.entries(PACOTES).map(([key, label]) => (
                <button
                  type="button" key={key}
                  className={`ta-pacote-option ${key} ${newPacote === key ? "selected" : ""}`}
                  onClick={() => setNewPacote(key)}
                >{label}</button>
              ))}
            </div>
            <p className="ta-hint">Esse e-mail e senha serão o login do aluno no app. Suporte direto por WhatsApp é liberado nos planos Prata e Gold.</p>
            <label className="ta-label">Acesso válido até (opcional — deixe em branco para sem prazo)</label>
            <input className="ta-input" type="date" value={newAcessoAte} onChange={(e) => setNewAcessoAte(e.target.value)} />
            <button
              type="button" className="ta-btn ta-btn-ghost"
              onClick={() => { const d = new Date(); d.setDate(d.getDate() + 7); setNewAcessoAte(d.toISOString().slice(0, 10)); }}
            >Usar teste grátis de 7 dias</button>
            <div className="ta-row-end">
              <button className="ta-btn ta-btn-ghost" onClick={() => setShowAddStudent(false)}>Cancelar</button>
              <button className="ta-btn ta-btn-primary" onClick={onAddStudent}>Salvar aluno</button>
            </div>
          </div>
        )}

        <div className="ta-student-list">
          {studentsLoading && <EmptyState msg="Carregando alunos..." />}
          {!studentsLoading && students.length === 0 && <EmptyState msg="Nenhum aluno encontrado." />}
          {students.map((s) => (
            <div key={s.id} className="ta-student-row">
              <div className="ta-avatar">{initials(s.nome)}</div>
              <div className="ta-student-info">
                <div className="ta-student-name">
                  {s.nome} <span className={`ta-badge pacote-${s.pacote || "bronze"}`}>{PACOTES[s.pacote] || "Bronze"}</span>
                  {acessoExpirado(s.acessoAte) && <span className="ta-badge expirado">expirado</span>}
                </div>
                <div className="ta-student-meta">
                  {s.treino.dias.length} dia(s) de treino · atualizado em {formatDate(s.treino.ultimaAtualizacao)} · acesso {acessoStatusLabel(s.acessoAte)}
                </div>
              </div>
              <button className="ta-btn ta-btn-ghost" onClick={() => onOpenStudent(s.id)}>Ver treino</button>
              <button className="ta-icon-btn danger" onClick={() => onRemoveStudent(s.id)} title="Remover aluno"><Trash2 size={16} /></button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function AdminUploadPreview({ preview, onCancel, onConfirm }) {
  return (
    <div className="ta-page">
      <TopBar title="Pré-visualização do upload" icon={<FileSpreadsheet size={18} />} />
      <div className="ta-wrap">
        <div className="ta-card">
          <div className="ta-preview-file"><FileSpreadsheet size={16} /> {preview.fileName}</div>
          <p className="ta-sub small">Confira os dados antes de confirmar. O treino atual de cada aluno listado será substituído.</p>
        </div>

        {preview.errors.length > 0 && (
          <div className="ta-warning-box">
            <AlertTriangle size={16} />
            <div>
              <div className="ta-warning-title">{preview.errors.length} linha(s) ignorada(s)</div>
              <ul>{preview.errors.slice(0, 6).map((e, i) => <li key={i}>{e}</li>)}</ul>
            </div>
          </div>
        )}

        <div className="ta-preview-list">
          {preview.entries.map((entry, i) => (
            <div className="ta-card ta-preview-card" key={i}>
              <div className="ta-row-between">
                <div className="ta-student-name">{entry.nome}</div>
                <span className={`ta-badge ${entry.isNew ? "new" : "update"}`}>{entry.isNew ? "sem conta — não será salvo" : "atualização"}</span>
              </div>
              <div className="ta-student-meta">{entry.dias.length} dia(s) · {entry.totalExercicios} exercício(s)</div>
              <div className="ta-preview-dias">
                {entry.dias.map((d, di) => (
                  <span key={di} className="ta-chip">{d.nome} <b>({d.exercicios.length})</b></span>
                ))}
              </div>
            </div>
          ))}
        </div>

        <div className="ta-row-end">
          <button className="ta-btn ta-btn-ghost" onClick={onCancel}>Cancelar</button>
          <button className="ta-btn ta-btn-primary" onClick={onConfirm}><Check size={16} /> Confirmar upload</button>
        </div>
      </div>
    </div>
  );
}

function AdminStudentView(props) {
  const { student, onBack, editingKey, editDraft, setEditDraft, onStartEdit, onSaveEdit, onCancelEdit, onRemoveExercicio, onRemoveDia, onSetPacote, onSetFeedbackProfessor, onSetAcessoAte } = props;
  const acessoDateValue = student.acessoAte ? new Date(student.acessoAte).toISOString().slice(0, 10) : "";
  return (
    <div className="ta-page">
      <TopBar title={student.nome} icon={<Users size={18} />} onBack={onBack} />
      <div className="ta-wrap">
        <div className="ta-row-between" style={{ marginBottom: 14 }}>
          <div className="ta-student-meta">
            Última atualização: {formatDate(student.treino.ultimaAtualizacao)} · {student.treino.dias.length} dia(s) · acesso {acessoStatusLabel(student.acessoAte)}
          </div>
          <div className="ta-pacote-picker">
            {Object.entries(PACOTES).map(([key, label]) => (
              <button
                type="button" key={key}
                className={`ta-pacote-option ${key} ${(student.pacote || "bronze") === key ? "selected" : ""}`}
                onClick={() => onSetPacote(key)}
              >{label}</button>
            ))}
          </div>
        </div>
        <div className="ta-row-between" style={{ marginTop: -6 }}>
          <label className="ta-label" style={{ margin: 0 }}>Acesso válido até</label>
          <input
            className="ta-input" type="date" style={{ maxWidth: 170 }}
            value={acessoDateValue}
            onChange={(e) => onSetAcessoAte(e.target.value)}
          />
        </div>
        {student.treino.dias.length === 0 && <EmptyState msg="Este aluno ainda não tem treino cadastrado. Envie uma planilha com o nome dele." />}
        {student.treino.dias.map((dia, di) => {
          const fb = student.progresso?.feedback?.[di] || {};
          return (
          <div className="ta-card ta-day-block" key={di}>
            <div className="ta-row-between">
              <div className="ta-day-title">{dia.nome}</div>
              <button className="ta-icon-btn danger" onClick={() => onRemoveDia(di)} title="Remover dia"><Trash2 size={15} /></button>
            </div>
            {dia.exercicios.map((ex, ei) => {
              const key = `${di}-${ei}`;
              const editing = editingKey === key;
              const cargas = student.progresso?.cargas?.[`${di}_${ei}`] || [];
              return (
                <div className="ta-ex-row" key={ei}>
                  {editing ? (
                    <div className="ta-ex-edit">
                      <input className="ta-input" value={editDraft.nome} onChange={(e) => setEditDraft({ ...editDraft, nome: e.target.value })} placeholder="Exercício" />
                      <div className="ta-ex-edit-grid">
                        <input className="ta-input" type="number" value={editDraft.series} onChange={(e) => setEditDraft({ ...editDraft, series: e.target.value })} placeholder="Séries" />
                        <input className="ta-input" type="number" value={editDraft.repeticoes} onChange={(e) => setEditDraft({ ...editDraft, repeticoes: e.target.value })} placeholder="Repetições" />
                      </div>
                      <input className="ta-input" value={editDraft.linkVideo} onChange={(e) => setEditDraft({ ...editDraft, linkVideo: e.target.value })} placeholder="Link do vídeo" />
                      <input className="ta-input" value={editDraft.observacoes} onChange={(e) => setEditDraft({ ...editDraft, observacoes: e.target.value })} placeholder="Observações" />
                      <div className="ta-row-end">
                        <button className="ta-btn ta-btn-ghost" onClick={onCancelEdit}>Cancelar</button>
                        <button className="ta-btn ta-btn-primary" onClick={() => onSaveEdit(di, ei)}><Save size={14} /> Salvar</button>
                      </div>
                    </div>
                  ) : (
                    <>
                      <div className="ta-ex-info">
                        <div className="ta-ex-name">{ex.nome}</div>
                        <div className="ta-ex-tags">
                          <span className="ta-badge">{ex.series} séries</span>
                          <span className="ta-badge">{ex.repeticoes} reps</span>
                          {ex.observacoes && <span className="ta-badge muted">{ex.observacoes}</span>}
                        </div>
                        {cargas.length > 0 && (
                          <div className="ta-carga-history">
                            Cargas: {cargas.slice(-5).map((c, i) => `${c.valor}kg (${formatDate(c.data)})`).join(" · ")}
                          </div>
                        )}
                      </div>
                      <div className="ta-ex-actions">
                        <button className="ta-icon-btn" onClick={() => onStartEdit(di, ei, ex)} title="Editar"><Pencil size={15} /></button>
                        <button className="ta-icon-btn danger" onClick={() => onRemoveExercicio(di, ei)} title="Remover"><Trash2 size={15} /></button>
                      </div>
                    </>
                  )}
                </div>
              );
            })}
            <div className="ta-feedback-block">
              <div className="ta-feedback-label">Como o aluno se sentiu nesse treino</div>
              <div className="ta-feedback-aluno">{fb.aluno ? fb.aluno : <span className="ta-sub small" style={{ margin: 0 }}>Ainda sem feedback do aluno.</span>}</div>
              <label className="ta-label">Sua resposta (o aluno vê isso)</label>
              <textarea
                className="ta-textarea" placeholder="Escreva uma orientação ou resposta pro aluno..."
                value={fb.professor || ""} onChange={(e) => onSetFeedbackProfessor(di, e.target.value)}
              />
            </div>
          </div>
          );
        })}
      </div>
    </div>
  );
}

function StudentDashboard({ student, completed, onLogout, onOpenDia, config }) {
  const dias = student.treino.dias;
  return (
    <div className="ta-page">
      <TopBar title="Meu treino" icon={<Dumbbell size={18} />} onLogout={onLogout} />
      <div className="ta-wrap">
        <div className="ta-welcome">
          <div className="ta-avatar big">{initials(student.nome)}</div>
          <div>
            <div className="ta-welcome-hi">Bem-vindo(a) de volta,</div>
            <div className="ta-welcome-name">{student.nome}</div>
          </div>
        </div>

        <div className="ta-stat-grid">
          <div className="ta-stat-card">
            <CalendarDays size={18} />
            <div className="ta-stat-num">{dias.length}</div>
            <div className="ta-stat-label">dias de treino</div>
          </div>
          <div className="ta-stat-card">
            <Dumbbell size={18} />
            <div className="ta-stat-num">{dias[0]?.nome.replace(/^Dia \d+\s*[—-]?\s*/i, "") || "—"}</div>
            <div className="ta-stat-label">próximo treino</div>
          </div>
          <div className="ta-stat-card">
            <Clock size={18} />
            <div className="ta-stat-num">{formatDate(student.treino.ultimaAtualizacao)}</div>
            <div className="ta-stat-label">última atualização</div>
          </div>
        </div>

        {dias.length === 0 && <EmptyState msg="Seu professor ainda não enviou seu treino. Volte em breve!" />}

        <SuporteCard pacote={student.pacote || "bronze"} whatsapp={config?.whatsappProfessor} />

        <div className="ta-plate-list">
          {dias.map((dia, idx) => {
            const total = dia.exercicios.length;
            const done = dia.exercicios.filter((_, ei) => completed[`${student.id}-${idx}-${ei}`]).length;
            const pct = total ? Math.round((done / total) * 100) : 0;
            return (
              <button className="ta-plate-bar" key={idx} onClick={() => onOpenDia(idx)}>
                <div className="ta-plate-disc">{dia.dia}</div>
                <div className="ta-plate-body">
                  <div className="ta-plate-name">{dia.nome}</div>
                  <div className="ta-plate-sub">{total} exercício(s){done > 0 ? ` · ${done}/${total} concluídos` : ""}</div>
                  <div className="ta-plate-track"><div className="ta-plate-fill" style={{ width: `${pct}%` }} /></div>
                </div>
                <ChevronRight size={20} className="ta-plate-chevron" />
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function SuporteCard({ pacote, whatsapp }) {
  const liberado = pacote === "prata" || pacote === "gold";
  const numero = (whatsapp || "").replace(/\D/g, "");
  if (!liberado) {
    return (
      <div className="ta-card ta-suporte-card locked">
        <MessageCircle size={20} />
        <div>
          <div className="ta-suporte-title">Suporte direto com o professor</div>
          <div className="ta-suporte-sub">Disponível nos planos Prata e Gold. Fale com seu professor para saber mais.</div>
        </div>
      </div>
    );
  }
  return (
    <a
      className="ta-card ta-suporte-card"
      href={numero ? `https://wa.me/${numero}?text=${encodeURIComponent("Oi! Tenho uma dúvida/dor/incômodo sobre meu treino.")}` : undefined}
      target="_blank" rel="noopener noreferrer"
      onClick={(e) => { if (!numero) e.preventDefault(); }}
    >
      <MessageCircle size={20} />
      <div>
        <div className="ta-suporte-title">Alguma dúvida, dor ou incômodo no treino?</div>
        <div className="ta-suporte-sub">{numero ? "Fale direto com seu professor pelo WhatsApp" : "Suporte ainda não configurado pelo professor"}</div>
      </div>
    </a>
  );
}

function StudentDayDetail({ student, diaIdx, dia, totalDias, completed, notes, onToggleCompleted, onSetNote, onRegistrarCarga, onSetFeedbackAluno, onBack }) {
  const [cargaDraft, setCargaDraft] = useState({});
  const fb = student.progresso?.feedback?.[diaIdx] || {};
  return (
    <div className="ta-page">
      <TopBar title={dia.nome} icon={<Dumbbell size={18} />} onBack={onBack} />
      <div className="ta-wrap">
        <div className="ta-progress-line">Dia {diaIdx + 1} de {totalDias}</div>
        <div className="ta-ex-list">
          {dia.exercicios.map((ex, ei) => {
            const key = `${student.id}-${diaIdx}-${ei}`;
            const done = !!completed[key];
            const cargaKey = `${diaIdx}_${ei}`;
            const cargas = student.progresso?.cargas?.[cargaKey] || [];
            return (
              <div className={`ta-ex-card ${done ? "done" : ""}`} key={ei}>
                <div className="ta-ex-card-top">
                  <button className={`ta-check ${done ? "checked" : ""}`} onClick={() => onToggleCompleted(diaIdx, ei)} aria-label="Marcar como concluído">
                    {done && <Check size={14} />}
                  </button>
                  <div className="ta-ex-card-name">{ex.nome}</div>
                </div>
                <div className="ta-ex-tags">
                  <span className="ta-badge">{ex.series} séries</span>
                  <span className="ta-badge">{ex.repeticoes} repetições</span>
                </div>
                {ex.observacoes && <div className="ta-ex-obs">{ex.observacoes}</div>}
                {ex.linkVideo && (
                  <a className="ta-btn ta-btn-video" href={ex.linkVideo} target="_blank" rel="noopener noreferrer">
                    <Play size={14} /> Ver vídeo explicativo
                  </a>
                )}

                <div className="ta-carga-block">
                  <div className="ta-carga-label"><TrendingUp size={13} /> Evolução de carga</div>
                  {cargas.length > 0 && (
                    <div className="ta-carga-history">
                      {cargas.slice(-5).map((c, i) => `${c.valor}kg (${formatDate(c.data)})`).join(" · ")}
                    </div>
                  )}
                  <div className="ta-carga-add">
                    <input
                      className="ta-input" type="number" placeholder="Carga usada (kg)"
                      value={cargaDraft[cargaKey] || ""}
                      onChange={(e) => setCargaDraft({ ...cargaDraft, [cargaKey]: e.target.value })}
                    />
                    <button
                      className="ta-btn ta-btn-ghost"
                      onClick={() => { onRegistrarCarga(diaIdx, ei, cargaDraft[cargaKey]); setCargaDraft({ ...cargaDraft, [cargaKey]: "" }); }}
                    >Registrar</button>
                  </div>
                </div>

                <textarea
                  className="ta-textarea"
                  placeholder="Suas anotações sobre este exercício..."
                  value={notes[key] || ""}
                  onChange={(e) => onSetNote(diaIdx, ei, e.target.value)}
                />
              </div>
            );
          })}
        </div>

        <div className="ta-card ta-feedback-block">
          <div className="ta-feedback-label">Como você se sentiu nesse treino?</div>
          <textarea
            className="ta-textarea" placeholder="Doeu algo? Achou fácil, difícil, algo que não gostou..."
            value={fb.aluno || ""} onChange={(e) => onSetFeedbackAluno(diaIdx, e.target.value)}
          />
          {fb.professor && (
            <div className="ta-feedback-professor">
              <strong>Seu professor respondeu:</strong> {fb.professor}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== PARTES REUTILIZÁVEIS ============================== */

function AcessoExpirado({ student, config, onLogout }) {
  const numero = (config?.whatsappProfessor || "").replace(/\D/g, "");
  const podeSuporte = student.pacote === "prata" || student.pacote === "gold";
  return (
    <div className="ta-center">
      <div className="ta-auth-card" style={{ textAlign: "center" }}>
        <div className="ta-auth-icon" style={{ margin: "0 auto 12px" }}><Clock size={22} /></div>
        <h2 className="ta-h2">Seu acesso expirou</h2>
        <p className="ta-sub small">Fale com seu professor para renovar e continuar acompanhando seu treino.</p>
        {podeSuporte && numero && (
          <a
            className="ta-btn ta-btn-primary ta-btn-block"
            href={`https://wa.me/${numero}?text=${encodeURIComponent("Oi! Meu acesso ao Training Club expirou, gostaria de renovar.")}`}
            target="_blank" rel="noopener noreferrer"
          >Falar com o professor</a>
        )}
        <button className="ta-btn ta-btn-ghost ta-btn-block" onClick={onLogout} style={{ marginTop: 10 }}>Sair</button>
      </div>
    </div>
  );
}

function TopBar({ title, icon, onBack, onLogout }) {
  return (
    <div className="ta-topbar">
      <div className="ta-topbar-left">
        {onBack && <button className="ta-icon-btn" onClick={onBack}><ArrowLeft size={18} /></button>}
        <div className="ta-topbar-mark">{icon}</div>
        <div className="ta-topbar-title">{title}</div>
      </div>
      {onLogout && <button className="ta-btn ta-btn-ghost" onClick={onLogout}><LogOut size={15} /> Sair</button>}
    </div>
  );
}

function EmptyState({ msg }) {
  return <div className="ta-empty">{msg}</div>;
}

/* ============================== ESTILO ============================== */

function Style() {
  return (
    <style>{`
      .ta-root {
        --bg: #201f1c;
        --surface: #29271f;
        --surface-alt: #322f26;
        --border: #443f32;
        --text: #f5f1e6;
        --text-muted: #a89e8a;
        --accent: #c9a24b;
        --accent-dark: #e0bd6c;
        --accent-soft: #c9a24b26;
        --teal: #6f9575;
        --teal-soft: #6f957526;
        --danger: #ff8080;
        --danger-soft: #ff80801f;
        --radius: 14px;
        font-family: 'Inter', system-ui, sans-serif;
        background: var(--bg);
        color: var(--text);
        min-height: 100vh;
        width: 100%;
        box-sizing: border-box;
      }
      .ta-root * { box-sizing: border-box; }
      .ta-root h1, .ta-root h2, .ta-root .ta-display {
        font-family: 'Montserrat', 'Arial Narrow', sans-serif;
        letter-spacing: 0.02em;
        text-transform: uppercase;
      }

      .ta-center { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px; }
      .ta-page { min-height: 100vh; padding-bottom: 40px; }
      .ta-wrap { max-width: 640px; margin: 0 auto; padding: 20px 18px 40px; }

      .ta-hero { text-align: center; max-width: 420px; }
      .ta-hero-logo { width: 96px; height: 96px; object-fit: contain; margin: 0 auto 10px; display: block; }
      .ta-h1 { font-size: 26px; font-weight: 700; margin: 0 0 4px; letter-spacing: 0.08em; }
      .ta-tagline { color: var(--accent-dark); font-size: 11px; letter-spacing: 0.22em; text-transform: uppercase; margin: 0 0 24px; }
      .ta-accent { color: var(--accent-dark); }
      .ta-sub { color: var(--text-muted); font-size: 14.5px; margin: 0 0 28px; line-height: 1.5; }
      .ta-sub.small { margin-bottom: 18px; }

      .ta-role-grid { display: flex; flex-direction: column; gap: 12px; }
      .ta-role-card {
        position: relative; display: flex; flex-direction: column; align-items: flex-start; gap: 6px;
        background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
        padding: 18px 44px 18px 18px; color: var(--text); cursor: pointer; text-align: left;
        transition: border-color .15s ease, transform .15s ease;
      }
      .ta-role-card:hover { border-color: var(--accent); transform: translateY(-1px); }
      .ta-role-card svg:first-child { color: var(--accent); }
      .ta-role-title { font-weight: 600; font-size: 16px; }
      .ta-role-desc { color: var(--text-muted); font-size: 13px; }
      .ta-role-chevron { position: absolute; right: 16px; top: 50%; transform: translateY(-50%); color: var(--text-muted); }

      .ta-auth-card { width: 100%; max-width: 380px; background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 24px; }
      .ta-back { display: inline-flex; align-items: center; gap: 6px; background: none; border: none; color: var(--text-muted); font-size: 13px; cursor: pointer; padding: 0; margin-bottom: 16px; }
      .ta-auth-icon { width: 42px; height: 42px; border-radius: 12px; background: var(--accent-soft); color: var(--accent); display: flex; align-items: center; justify-content: center; margin-bottom: 12px; }
      .ta-h2 { font-size: 20px; margin: 0 0 4px; }
      .ta-label { display: block; font-size: 12.5px; color: var(--text-muted); margin: 14px 0 6px; }
      .ta-input {
        width: 100%; background: var(--surface-alt); border: 1px solid var(--border); border-radius: 9px;
        padding: 10px 12px; color: var(--text); font-size: 14px; font-family: inherit;
      }
      .ta-input:focus { outline: none; border-color: var(--accent); }
      .ta-btn-block { width: 100%; margin-top: 20px; justify-content: center; }
      .ta-hint { color: var(--text-muted); font-size: 11.5px; margin: 12px 0 0; line-height: 1.5; }

      .ta-btn {
        display: inline-flex; align-items: center; gap: 6px; border-radius: 9px; border: 1px solid transparent;
        padding: 9px 14px; font-size: 13.5px; font-weight: 600; cursor: pointer; font-family: inherit;
        text-decoration: none; white-space: nowrap;
      }
      .ta-btn-primary { background: var(--accent); color: #17130f; }
      .ta-btn-primary:hover { background: #ff7d4c; }
      .ta-btn-primary:disabled { opacity: .5; cursor: not-allowed; }
      .ta-btn-ghost { background: transparent; border-color: var(--border); color: var(--text); }
      .ta-btn-ghost:hover { border-color: var(--accent); color: var(--accent); }
      .ta-btn-video { background: var(--teal-soft); color: var(--teal); border-color: transparent; margin-top: 10px; }
      .ta-btn-video:hover { background: var(--teal); color: #0b1613; }
      .ta-icon-btn {
        display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px;
        border-radius: 8px; background: transparent; border: 1px solid var(--border); color: var(--text-muted); cursor: pointer;
      }
      .ta-icon-btn:hover { color: var(--text); border-color: var(--text-muted); }
      .ta-icon-btn.danger:hover { color: var(--danger); border-color: var(--danger); }

      .ta-topbar { display: flex; align-items: center; justify-content: space-between; padding: 16px 18px; border-bottom: 1px solid var(--border); position: sticky; top: 0; background: var(--bg); z-index: 5; }
      .ta-topbar-left { display: flex; align-items: center; gap: 10px; }
      .ta-topbar-mark { color: var(--accent); display: flex; }
      .ta-topbar-title { font-family: 'Montserrat', sans-serif; text-transform: uppercase; letter-spacing: .02em; font-size: 15px; font-weight: 600; }

      .ta-dropzone {
        display: flex; align-items: center; gap: 14px; background: var(--surface); border: 1.5px dashed var(--border);
        border-radius: var(--radius); padding: 18px; cursor: pointer; margin-bottom: 20px; transition: border-color .15s, background .15s;
      }
      .ta-dropzone svg:first-child { color: var(--accent); flex-shrink: 0; }
      .ta-dropzone.over { border-color: var(--accent); background: var(--accent-soft); }
      .ta-dz-title { font-weight: 600; font-size: 14.5px; }
      .ta-dz-sub { color: var(--text-muted); font-size: 12.5px; margin-top: 2px; }
      .ta-dropzone .ta-btn { margin-left: auto; flex-shrink: 0; }

      .ta-row-between { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
      .ta-row-end { display: flex; justify-content: flex-end; gap: 8px; margin-top: 14px; }
      .ta-search { display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border); border-radius: 9px; padding: 8px 12px; flex: 1; color: var(--text-muted); }
      .ta-search input { background: none; border: none; color: var(--text); font-size: 13.5px; outline: none; width: 100%; font-family: inherit; }

      .ta-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px; margin-bottom: 12px; }
      .ta-add-form { display: flex; flex-direction: column; gap: 10px; }

      .ta-student-list { display: flex; flex-direction: column; gap: 10px; }
      .ta-student-row { display: flex; align-items: center; gap: 12px; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px; }
      .ta-avatar { width: 38px; height: 38px; border-radius: 50%; background: var(--surface-alt); border: 1px solid var(--border); display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px; color: var(--accent); flex-shrink: 0; }
      .ta-avatar.big { width: 52px; height: 52px; font-size: 16px; }
      .ta-student-info { flex: 1; min-width: 0; }
      .ta-student-name { font-weight: 600; font-size: 14.5px; }
      .ta-student-meta { color: var(--text-muted); font-size: 12px; margin-top: 2px; }

      .ta-preview-file { display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 14px; color: var(--accent); }
      .ta-warning-box { display: flex; gap: 10px; background: var(--danger-soft); border: 1px solid var(--danger); border-radius: 12px; padding: 12px 14px; margin-bottom: 14px; color: var(--danger); }
      .ta-warning-box svg { flex-shrink: 0; margin-top: 2px; color: var(--danger); }
      .ta-warning-title { font-weight: 600; font-size: 13px; margin-bottom: 4px; color: var(--text); }
      .ta-warning-box ul { margin: 0; padding-left: 16px; font-size: 12px; }
      .ta-preview-list { display: flex; flex-direction: column; gap: 10px; margin-bottom: 8px; }
      .ta-preview-dias { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
      .ta-chip { background: var(--surface-alt); border: 1px solid var(--border); border-radius: 20px; padding: 4px 10px; font-size: 11.5px; color: var(--text-muted); }
      .ta-chip b { color: var(--text); }

      .ta-badge { display: inline-flex; align-items: center; background: var(--surface-alt); border: 1px solid var(--border); border-radius: 20px; padding: 3px 10px; font-size: 11.5px; color: var(--text); }
      .ta-badge.new { background: var(--teal-soft); color: var(--teal); border-color: transparent; }
      .ta-badge.update { background: var(--accent-soft); color: var(--accent); border-color: transparent; }
      .ta-badge.muted { color: var(--text-muted); }

      .ta-day-block { }
      .ta-day-title { font-family: 'Montserrat', sans-serif; text-transform: uppercase; font-size: 14px; letter-spacing: .02em; color: var(--accent); margin-bottom: 10px; }
      .ta-ex-row { border-top: 1px solid var(--border); padding: 12px 0; display: flex; align-items: flex-start; justify-content: space-between; gap: 10px; }
      .ta-ex-row:first-of-type { border-top: none; padding-top: 0; }
      .ta-ex-info { flex: 1; }
      .ta-ex-name { font-weight: 600; font-size: 14px; margin-bottom: 6px; }
      .ta-ex-tags { display: flex; flex-wrap: wrap; gap: 6px; }
      .ta-ex-actions { display: flex; gap: 6px; flex-shrink: 0; }
      .ta-ex-edit { width: 100%; display: flex; flex-direction: column; gap: 8px; }
      .ta-ex-edit-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }

      .ta-welcome { display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }
      .ta-welcome-hi { color: var(--text-muted); font-size: 12.5px; }
      .ta-welcome-name { font-family: 'Montserrat', sans-serif; text-transform: uppercase; font-size: 19px; font-weight: 600; }

      .ta-stat-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 22px; }
      .ta-stat-card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 12px 10px; text-align: center; }
      .ta-stat-card svg { color: var(--accent); margin-bottom: 6px; }
      .ta-stat-num { font-family: 'Montserrat', sans-serif; font-size: 13.5px; font-weight: 600; text-transform: uppercase; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .ta-stat-label { color: var(--text-muted); font-size: 10.5px; margin-top: 2px; text-transform: uppercase; letter-spacing: .03em; }

      .ta-plate-list { display: flex; flex-direction: column; gap: 12px; }
      .ta-plate-bar {
        display: flex; align-items: center; gap: 14px; background: var(--surface); border: 1px solid var(--border);
        border-radius: var(--radius); padding: 14px 16px; cursor: pointer; text-align: left; width: 100%; font-family: inherit;
        transition: border-color .15s, transform .15s;
      }
      .ta-plate-bar:hover { border-color: var(--accent); transform: translateX(2px); }
      .ta-plate-disc {
        width: 44px; height: 44px; border-radius: 50%; flex-shrink: 0; background: radial-gradient(circle at 35% 30%, #fffaf0, #ecdfbf 70%);
        border: 3px solid var(--accent); display: flex; align-items: center; justify-content: center;
        font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 17px; color: #221d14;
      }
      .ta-plate-body { flex: 1; min-width: 0; }
      .ta-plate-name { font-weight: 600; font-size: 14.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .ta-plate-sub { color: var(--text-muted); font-size: 12px; margin: 2px 0 8px; }
      .ta-plate-track { height: 4px; background: var(--border); border-radius: 4px; overflow: hidden; }
      .ta-plate-fill { height: 100%; background: linear-gradient(90deg, var(--accent), var(--teal)); border-radius: 4px; transition: width .3s ease; }
      .ta-plate-chevron { color: var(--text-muted); flex-shrink: 0; }

      .ta-progress-line { color: var(--text-muted); font-size: 12.5px; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 16px; }
      .ta-ex-list { display: flex; flex-direction: column; gap: 12px; }
      .ta-ex-card { background: var(--surface); border: 1px solid var(--border); border-left: 3px solid var(--accent); border-radius: var(--radius); padding: 14px 16px; }
      .ta-ex-card.done { border-left-color: var(--teal); opacity: .8; }
      .ta-ex-card-top { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
      .ta-ex-card-name { font-weight: 600; font-size: 15px; }
      .ta-check {
        width: 22px; height: 22px; border-radius: 6px; border: 1.5px solid var(--border); background: var(--surface-alt);
        display: flex; align-items: center; justify-content: center; cursor: pointer; flex-shrink: 0; color: #0b1613;
      }
      .ta-check.checked { background: var(--teal); border-color: var(--teal); }
      .ta-ex-obs { color: var(--text-muted); font-size: 12.5px; margin-top: 8px; font-style: italic; }
      .ta-textarea {
        width: 100%; margin-top: 12px; background: var(--surface-alt); border: 1px solid var(--border); border-radius: 9px;
        padding: 9px 11px; color: var(--text); font-size: 12.5px; font-family: inherit; min-height: 54px; resize: vertical;
      }
      .ta-textarea:focus { outline: none; border-color: var(--teal); }

      .ta-empty { text-align: center; color: var(--text-muted); font-size: 13.5px; padding: 32px 16px; border: 1px dashed var(--border); border-radius: var(--radius); }

      .ta-toast {
        position: fixed; top: 18px; left: 50%; transform: translateX(-50%); z-index: 50;
        display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border);
        border-radius: 10px; padding: 10px 16px; font-size: 13px; box-shadow: 0 8px 24px rgba(0,0,0,.45);
      }
      .ta-toast.success { border-color: var(--teal); color: var(--teal); }
      .ta-toast.error { border-color: var(--danger); color: var(--danger); }

      .ta-pacote-picker { display: flex; gap: 8px; flex-wrap: wrap; }
      .ta-pacote-option {
        border: 1.5px solid var(--border); background: var(--surface-alt); color: var(--text-muted);
        border-radius: 20px; padding: 6px 14px; font-size: 12.5px; font-weight: 600; cursor: pointer; font-family: inherit;
      }
      .ta-pacote-option.selected { color: #221d14; border-color: transparent; }
      .ta-pacote-option.bronze.selected { background: linear-gradient(135deg, #cd7f32, #a86428); }
      .ta-pacote-option.prata.selected { background: linear-gradient(135deg, #d7d7de, #a7a7b3); }
      .ta-pacote-option.gold.selected { background: linear-gradient(135deg, var(--accent), var(--accent-dark)); }
      .ta-badge.pacote-bronze { background: #cd7f321f; color: #cd9a5f; border-color: transparent; }
      .ta-badge.pacote-prata { background: #d7d7de1f; color: #c7c7d1; border-color: transparent; }
      .ta-badge.pacote-gold { background: var(--accent-soft); color: var(--accent-dark); border-color: transparent; }
      .ta-badge.expirado { background: var(--danger-soft); color: var(--danger); border-color: transparent; margin-left: 6px; }

      .ta-carga-block { margin-top: 10px; background: var(--surface-alt); border: 1px solid var(--border); border-radius: 9px; padding: 10px 12px; }
      .ta-carga-label { display: flex; align-items: center; gap: 6px; font-size: 11.5px; text-transform: uppercase; letter-spacing: .03em; color: var(--text-muted); margin-bottom: 6px; }
      .ta-carga-history { font-size: 12px; color: var(--text); margin-bottom: 8px; }
      .ta-carga-add { display: flex; gap: 8px; }
      .ta-carga-add .ta-input { flex: 1; }

      .ta-feedback-block { margin-top: 14px; }
      .ta-feedback-label { font-weight: 600; font-size: 13.5px; margin-bottom: 8px; }
      .ta-feedback-aluno { font-size: 13px; color: var(--text); background: var(--surface-alt); border-radius: 9px; padding: 10px 12px; margin-bottom: 10px; }
      .ta-feedback-professor { margin-top: 10px; background: var(--accent-soft); border-radius: 9px; padding: 10px 12px; font-size: 12.5px; color: var(--text); }

      .ta-suporte-card { display: flex; align-items: center; gap: 12px; text-decoration: none; color: var(--text); margin-bottom: 16px; transition: border-color .15s; }
      .ta-suporte-card:not(.locked):hover { border-color: var(--teal); }
      .ta-suporte-card svg { color: var(--teal); flex-shrink: 0; }
      .ta-suporte-card.locked svg { color: var(--text-muted); }
      .ta-suporte-title { font-weight: 600; font-size: 13.5px; }
      .ta-suporte-sub { color: var(--text-muted); font-size: 12px; margin-top: 2px; }
    `}</style>
  );
}
