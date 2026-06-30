let allQuestions = [], quizQuestions = [], currentIndex = 0, score = 0, results = [], questionStates = [];
let quizTitle = "Deutsch Heimkehr", historical = [], selectedQuestionCount = 0, currentResultSaved = false;
let retryMode = false, retryQuestions = [], retryResults = [], originalResults = [], retryScore = 0, originalQuizQuestions = [];
let workbookZip = null, workbookFileName = "deutsch_heimkehr.xlsx", resultsSheetPath = "";
let workbookLoadedAt = null, workbookLastModified = null;
let selectedAnswer = "", answered = false;
let lessonLibrary = [];
let selectedLibraryIndex = -1;
let currentLessonSummary = null;
let courseManifest = { courses: [] };
let libraryView = "levels";
let selectedLevel = "";
let selectedUnit = "";
let lessonData = { Lesson: [], Mission: [], Vocabulary: [], Grammar: [], Dialogue: [], Practice: [] };
let dynamicLessonSections = null;

const chooseFolderVisualBtn = document.getElementById("chooseFolderVisualBtn");
const lessonFolderInput = document.getElementById("lessonFolderInput");
const backToUnitsBtn = document.getElementById("backToUnitsBtn");
const lessonLibraryGrid = document.getElementById("lessonLibraryGrid");
const libraryStatus = document.getElementById("libraryStatus");
const libraryBreadcrumb = document.getElementById("libraryBreadcrumb");
const learnerNameInput = document.getElementById("learnerNameInput");
const saveLearnerBtn = document.getElementById("saveLearnerBtn");
const learnerStatus = document.getElementById("learnerStatus");
let lessonSections = ["Lesson", "Vocabulary", "Grammar", "Dialogue", "Practice"];
let currentLessonSectionIndex = 0;

const pageTitle = document.getElementById("pageTitle");
const quizHeaderTitle = document.getElementById("quizHeaderTitle");
const courseHomeBtn = document.getElementById("courseHomeBtn");
const currentLessonLabel = document.getElementById("currentLessonLabel");
const setupHeader = document.getElementById("setupHeader");
const quizHeader = document.getElementById("quizHeader");
const xlsxFile = document.getElementById("xlsxFile");
const loadBtn = document.getElementById("loadBtn");
const loadStatus = document.getElementById("loadStatus");
const statsArea = document.getElementById("statsArea");
const quizStartArea = document.getElementById("quizStartArea");
const questionCountText = document.getElementById("questionCountText");
const setupPanel = document.getElementById("setupPanel");
const lessonPanel = document.getElementById("lessonPanel");
const lessonSectionTitle = document.getElementById("lessonSectionTitle");
const lessonNav = document.getElementById("lessonNav");
const quizLessonNav = document.getElementById("quizLessonNav");
const lessonContent = document.getElementById("lessonContent");
const prevLessonBtn = document.getElementById("prevLessonBtn");
const nextLessonBtn = document.getElementById("nextLessonBtn");
const startQuizFromLessonBtn = document.getElementById("startQuizFromLessonBtn");
const quizPanel = document.getElementById("quizPanel");
const resultsPanel = document.getElementById("resultsPanel");
const progressText = document.getElementById("progressText");
const progressBar = document.getElementById("progressBar");
const scoreText = document.getElementById("scoreText");
const promptLabel = document.getElementById("promptLabel");
const contextText = document.getElementById("contextText");
const taskText = document.getElementById("taskText");
const targetText = document.getElementById("targetText");
const choicesArea = document.getElementById("choicesArea");
const fibArea = document.getElementById("fibArea");
const fibInput = document.getElementById("fibInput");
const feedback = document.getElementById("feedback");
const checkBtn = document.getElementById("checkBtn");
const nextBtn = document.getElementById("nextBtn");
const quitBtn = document.getElementById("quitBtn");
const finalScore = document.getElementById("finalScore");
const finalSummary = document.getElementById("finalSummary");
const saveStatus = document.getElementById("saveStatus");
const safetyDetails = document.getElementById("safetyDetails");
const confirmDownloadBtn = document.getElementById("confirmDownloadBtn");
const cancelDownloadBtn = document.getElementById("cancelDownloadBtn");
const restartBtn = document.getElementById("restartBtn");
const missedBtn = document.getElementById("missedBtn");
const retryMissedBtn = document.getElementById("retryMissedBtn");
const returnCourseBtn = document.getElementById("returnCourseBtn");
const reviewArea = document.getElementById("reviewArea");
const resultsLessonNav = document.getElementById("resultsLessonNav");
const dashTotal = document.getElementById("dashTotal");
const dashCorrect = document.getElementById("dashCorrect");
const dashIncorrect = document.getElementById("dashIncorrect");
const dashScore = document.getElementById("dashScore");
const questionMap = document.getElementById("questionMap");

document.getElementById("quiz10Btn").addEventListener("click", () => startQuiz(10));
document.getElementById("quiz25Btn").addEventListener("click", () => startQuiz(25));
document.getElementById("quizAllBtn").addEventListener("click", () => startQuiz(allQuestions.length));
document.getElementById("beginLessonBtn").addEventListener("click", beginLessonFlow);
if(chooseFolderVisualBtn && lessonFolderInput){
  chooseFolderVisualBtn.addEventListener("click", () => lessonFolderInput.click());
  lessonFolderInput.addEventListener("change", buildLessonLibraryFromFolder);
}
backToUnitsBtn.addEventListener("click", handleLibraryBack);

function normalize(value){ return String(value || "").trim().toLowerCase(); }
function normalizeGerman(value){
  return normalize(value)
    .replaceAll("ä","ae")
    .replaceAll("ö","oe")
    .replaceAll("ü","ue")
    .replaceAll("ß","ss");
}
function escapeHtml(value){ return String(value || "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;"); }
function xmlEscape(value){ return String(value).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&apos;"); }
function shuffle(array){ const copy=[...array]; for(let i=copy.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [copy[i],copy[j]]=[copy[j],copy[i]]; } return copy; }
function colIndex(cellRef){ const letters = String(cellRef || "").replace(/[0-9]/g, ""); let n=0; for(const ch of letters){ n = n*26 + ch.charCodeAt(0)-64; } return n-1; }
function textContent(node){ return node ? node.textContent : ""; }
function nodesByLocalName(parent, localName){ return Array.from(parent.getElementsByTagName("*")).filter(n => n.localName === localName); }
function firstNodeByLocalName(parent, localName){ return nodesByLocalName(parent, localName)[0] || null; }
function attr(node, name){ return node ? node.getAttribute(name) : null; }

function insertAtCursor(input, text){
  const start = input.selectionStart ?? input.value.length;
  const end = input.selectionEnd ?? input.value.length;
  input.value = input.value.slice(0, start) + text + input.value.slice(end);
  const newPos = start + text.length;
  input.focus();
  input.setSelectionRange(newPos, newPos);
}
document.querySelectorAll(".key-btn").forEach(btn => btn.addEventListener("click", () => insertAtCursor(fibInput, btn.dataset.symbol || "")));

async function parseWorkbook(arrayBuffer){
  workbookZip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const workbookFile = workbookZip.file("xl/workbook.xml");
  const relsFile = workbookZip.file("xl/_rels/workbook.xml.rels");
  if(!workbookFile || !relsFile) throw new Error("Missing workbook.xml or relationships.");

  const workbookXml = parser.parseFromString(await workbookFile.async("text"), "application/xml");
  const relsXml = parser.parseFromString(await relsFile.async("text"), "application/xml");

  const relMap = {};
  nodesByLocalName(relsXml, "Relationship").forEach(r => {
    const id = attr(r, "Id");
    let target = attr(r, "Target") || "";
    target = target.replace(/^\//, "");
    relMap[id] = target.startsWith("xl/") ? target : "xl/" + target;
  });

  const sheets = nodesByLocalName(workbookXml, "sheet").map(s => ({ name: attr(s,"name") || "", id: attr(s,"r:id") || attr(s,"id") }));
  const sharedStrings = await loadSharedStrings(parser);

  const isFinalReviewWorkbook =
    !!(currentLessonSummary && currentLessonSummary.isFinalReview) ||
    /final[_\s-]*review/i.test(workbookFileName || "");

  dynamicLessonSections = null;
  resultsSheetPath = "";

  if(isFinalReviewWorkbook){
    lessonData = {};
    dynamicLessonSections = [];

    const qSheet = sheets.find(s => s.name.toLowerCase() === "questions");

    for(const sheetInfo of sheets){
      const sheetName = sheetInfo.name || "";
      const lower = sheetName.toLowerCase();
      if(!sheetName || !relMap[sheetInfo.id]) continue;

      // Do not display Questions as a normal workbook tab.
      // It powers the interactive Knowledge Check instead.
      if(lower === "questions") continue;

      lessonData[sheetName] = await readSheetRows(parser, relMap[sheetInfo.id], sharedStrings);
      dynamicLessonSections.push(sheetName);
    }

    if(qSheet && relMap[qSheet.id]){
      const questionRows = await readSheetRows(parser, relMap[qSheet.id], sharedStrings);
      loadQuestionsFromRows(questionRows);
    } else {
      allQuestions = [];
      quizTitle = currentLessonSummary?.title || "A1 Comprehensive Review & Assessment";
      pageTitle.textContent = quizTitle;
      quizHeaderTitle.textContent = quizTitle;
      document.title = quizTitle;
      questionCountText.textContent = "No interactive assessment is available for this workbook.";
      quizStartArea.classList.add("hidden");
    }

    loadStatus.textContent = `Opened ${workbookFileName}.`;
    historical = [];
    updateStats();
    return;
  }

  lessonData = { Lesson: [], Mission: [], Vocabulary: [], Grammar: [], Dialogue: [], Practice: [] };

  const qSheet = sheets.find(s => s.name.toLowerCase() === "questions");
  if(!qSheet) throw new Error("Workbook must contain a Questions sheet.");

  const questionsSheetPath = relMap[qSheet.id];

  const lessonSheet = sheets.find(s => s.name.toLowerCase() === "lesson") || sheets.find(s => s.name.toLowerCase() === "mission");
  if(lessonSheet && relMap[lessonSheet.id]){
    lessonData.Lesson = await readSheetRows(parser, relMap[lessonSheet.id], sharedStrings);
  }
  for(const section of ["Vocabulary", "Grammar", "Dialogue", "Practice"]){
    const sheet = sheets.find(s => s.name.toLowerCase() === section.toLowerCase());
    if(sheet && relMap[sheet.id]){
      lessonData[section] = await readSheetRows(parser, relMap[sheet.id], sharedStrings);
    }
  }

  const questionRows = await readSheetRows(parser, questionsSheetPath, sharedStrings);
  loadQuestionsFromRows(questionRows);
  historical = [];
  updateStats();

  if(!currentLessonSummary || currentLessonSummary.title === workbookFileName){
    const meta = lessonRowsToMetadata(lessonData.Lesson || []);
    const heading = (lessonData.Lesson && lessonData.Lesson[0] && lessonData.Lesson[0][1]) ? lessonData.Lesson[0][1] : "";
    currentLessonSummary = {
      title: meta.Title || heading || quizTitle || workbookFileName,
      fileName: workbookFileName,
      level: inferFromFilename(workbookFileName, "A"),
      unit: inferFromFilename(workbookFileName, "U"),
      lesson: inferFromFilename(workbookFileName, "L")
    };
  }
}

async function loadSharedStrings(parser){
  const f = workbookZip.file("xl/sharedStrings.xml");
  if(!f) return [];
  const xml = parser.parseFromString(await f.async("text"), "application/xml");
  return nodesByLocalName(xml, "si").map(si => nodesByLocalName(si, "t").map(t => t.textContent).join(""));
}

async function readSheetRows(parser, path, sharedStrings){
  const sheetFile = workbookZip.file(path);
  if(!sheetFile) throw new Error("Missing sheet XML: " + path);
  const xml = parser.parseFromString(await sheetFile.async("text"), "application/xml");
  const rows = [];
  nodesByLocalName(xml, "row").forEach(rowNode => {
    const row = [];
    nodesByLocalName(rowNode, "c").forEach(c => {
      const idx = colIndex(attr(c, "r"));
      const t = attr(c, "t");
      let val = "";
      if(t === "s") val = sharedStrings[parseInt(textContent(firstNodeByLocalName(c, "v")) || "0", 10)] || "";
      else if(t === "inlineStr") val = textContent(firstNodeByLocalName(c, "t"));
      else val = textContent(firstNodeByLocalName(c, "v"));
      row[idx] = val;
    });
    rows.push(row.map(v => v || ""));
  });
  return rows;
}

function loadQuestionsFromRows(rows){
  if(!rows || rows.length < 2) throw new Error("Questions sheet does not contain expected rows.");
  if(normalize(rows[0][0]) === "title") quizTitle = rows[0][1] || "Deutsch Heimkehr";
  pageTitle.textContent = quizTitle; quizHeaderTitle.textContent = quizTitle; document.title = quizTitle;

  const headerRowIndex = rows.findIndex(r => normalize(r[0]) === "id");
  if(headerRowIndex < 0) throw new Error("Questions sheet must have an ID header row.");
  const headers = rows[headerRowIndex].map(h => String(h || "").trim());

  allQuestions = rows.slice(headerRowIndex + 1).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]).trim() : "");
    return normalizeLoadedQuestion(obj);
  }).filter(q => q.ID && q.Type && (q.Question || q.Exercise || q.Phrase || q.Instruction || q.Target) && q.Answer !== "");

  if(allQuestions.length === 0) throw new Error("No valid questions found.");
  loadStatus.textContent = `Opened ${allQuestions.length} questions from ${workbookFileName}.`;
  questionCountText.textContent = `${allQuestions.length} practice questions are ready.`;
  quizStartArea.classList.remove("hidden");
}

function normalizeLoadedQuestion(q){
  q.Type = String(q.Type || "").trim().toUpperCase();

  // Deutsch Heimkehr v1.3 format:
  // Exercise = heading, Phrase = large text, Instruction = smaller italic guidance.
  // Backward compatibility is kept for Prompt/Context/Task and old Question-only files.
  q.Exercise = q.Exercise || q.Prompt || "";
  q.Phrase = q.Phrase || q.Context || "";
  q.Instruction = q.Instruction || q.Task || "";
  q.Target = q.Target || "";

  if(!q.Exercise && q.Question){
    q.Exercise = q.Question;
  }

  // Support old A/B/C/D headers and old ChoiceA/ChoiceB/etc.
  q.ChoiceA = q.ChoiceA || q.A || "";
  q.ChoiceB = q.ChoiceB || q.B || "";
  q.ChoiceC = q.ChoiceC || q.C || "";
  q.ChoiceD = q.ChoiceD || q.D || "";

  if(q.Type === "TF"){
    const key = String(q.Answer || "").trim().toUpperCase();
    if(key === "1" || key === "TRUE") q.Answer = "True";
    else if(key === "0" || key === "FALSE") q.Answer = "False";
  }

  return q;
}

function loadResultsFromRows(rows){
  historical = [];
  if(!rows || rows.length < 2){ updateStats(); return; }
  const headers = rows[0].map(h => String(h || "").trim());
  historical = rows.slice(1).map(row => {
    const item = {};
    headers.forEach((h,i)=> item[h] = row[i] ? String(row[i]).trim() : "");
    return { Date:item.Date||"", QuestionCount:item.QuestionCount||item.Total||"", Score:parseFloat(item.Score||0), Total:parseFloat(item.Total||0), Percent:parseFloat(item.Percent||0) };
  }).filter(x => x.Date && !isNaN(x.Percent));
  updateStats();
}

function updateStats(){
  if(historical.length === 0){ statsArea.innerHTML = ""; return; }
  const last = historical[historical.length - 1];
  const best = historical.reduce((a,b)=> b.Percent > a.Percent ? b : a, historical[0]);
  const avg = (historical.reduce((s,i)=>s+i.Percent,0)/historical.length).toFixed(1);
  statsArea.innerHTML = `<div class="file-note"><strong>Previous Results</strong><br>Attempts: ${historical.length}<br>Last Attempt: ${last.QuestionCount} questions — ${last.Score}/${last.Total} (${last.Percent}%)<br>Best Attempt: ${best.QuestionCount} questions — ${best.Score}/${best.Total} (${best.Percent}%)<br>Average: ${avg}%</div>`;
}

if(loadBtn){ loadBtn.addEventListener("click", async () => {
  const file = xlsxFile.files[0];
  if(!file){ loadStatus.textContent = "Please choose a workbook first."; return; }
  try{
    workbookFileName = file.name;
    currentLessonSummary = { title: file.name, fileName: file.name };
    workbookLoadedAt = new Date();
    workbookLastModified = new Date(file.lastModified);
    loadStatus.textContent = "Opening lesson workbook...";
    const buffer = await file.arrayBuffer();
    await parseWorkbook(buffer);
  }catch(err){
    console.error(err);
    loadStatus.textContent = "There was a problem reading the lesson workbook. Make sure it has a Questions sheet.";
  }
});
}





const DH_PROGRESS_PREFIX = "deutschHeimkehr.progress.";
let learnerName = localStorage.getItem("deutschHeimkehr.activeLearner") || "Guest";

function sanitizeLearnerName(name){
  return String(name || "").trim().replace(/\s+/g, " ").slice(0, 40) || "Guest";
}

function progressKey(){
  return DH_PROGRESS_PREFIX + learnerName.toLowerCase();
}

function getProgress(){
  try{
    return JSON.parse(localStorage.getItem(progressKey()) || "{}");
  }catch(e){
    return {};
  }
}

function saveProgress(progress){
  localStorage.setItem(progressKey(), JSON.stringify(progress || {}));
}

function lessonProgressId(summary){
  const s = summary || currentLessonSummary || {};
  const level = s.level || inferFromFilename(workbookFileName || "", "A") || "A1";
  const unit = s.unit || inferFromFilename(workbookFileName || "", "U") || "1";
  const lesson = s.lesson || inferFromFilename(workbookFileName || "", "L") || "1";
  return `${level}-U${unit}-L${lesson}`;
}

function updateLearnerUI(){
  if(learnerNameInput) learnerNameInput.value = learnerName === "Guest" ? "" : learnerName;
  if(learnerStatus){
    learnerStatus.textContent = learnerName === "Guest"
      ? "Using Guest profile. Enter a name to keep progress separate on this device."
      : `👋 Welcome back, ${learnerName}! Progress is saved in this browser.`;
  }
}

function setLearnerName(){
  learnerName = sanitizeLearnerName(learnerNameInput ? learnerNameInput.value : learnerName);
  localStorage.setItem("deutschHeimkehr.activeLearner", learnerName);
  updateLearnerUI();
  renderLessonLibrary();
}

function saveCurrentLessonResult(){
  if(!currentLessonSummary || !quizQuestions || !quizQuestions.length) return;

  const progress = getProgress();
  const id = lessonProgressId(currentLessonSummary);
  const total = quizQuestions.length;
  const percent = total ? Math.round((score / total) * 100) : 0;
  const previous = progress[id] || {};

  progress[id] = {
    level: currentLessonSummary.level || "",
    unit: currentLessonSummary.unit || "",
    lesson: currentLessonSummary.lesson || "",
    title: currentLessonSummary.title || workbookFileName || "",
    attempts: (previous.attempts || 0) + 1,
    lastScore: score,
    lastTotal: total,
    lastPercent: percent,
    bestPercent: Math.max(previous.bestPercent || 0, percent),
    completed: percent >= 70 || previous.completed || false,
    updatedAt: new Date().toISOString()
  };

  saveProgress(progress);
}

function progressBadgeFor(summary){
  const progress = getProgress();
  const record = progress[lessonProgressId(summary)];
  if(!record) return "";
  if(record.completed){
    return `<div class="progress-badge">✓ Completed • Best ${record.bestPercent}%</div>`;
  }
  return `<div class="progress-badge">In progress • Best ${record.bestPercent || record.lastPercent || 0}%</div>`;
}

if(saveLearnerBtn){
  saveLearnerBtn.addEventListener("click", setLearnerName);
}
if(learnerNameInput){
  learnerNameInput.addEventListener("keydown", (e) => {
    if(e.key === "Enter") setLearnerName();
  });
}
updateLearnerUI();


async function loadCourseManifest(){
  lessonLibrary = [];
  selectedLibraryIndex = -1;
  lessonLibraryGrid.innerHTML = "";
  libraryStatus.textContent = "Loading course…";

  try{
    const response = await fetch("source/course.json", { cache: "no-store" });
    if(!response.ok) throw new Error("source/course.json was not found.");

    courseManifest = await response.json();
    const items = [];

    (courseManifest.courses || []).forEach(course => {
      (course.units || []).forEach(unit => {
        const unitPath = unit.path || "";
        (unit.lessons || []).forEach(lesson => {
          const lessonObj = typeof lesson === "string" ? { file: lesson } : (lesson || {});
          const fileName = lessonObj.file || lessonObj.fileName || "";
          if(!fileName || lessonObj.status === "coming-soon") return;
          items.push({
            level: lessonObj.level || course.level || "",
            unit: String(lessonObj.unit || unit.unit || ""),
            lessonNumber: lessonObj.lesson || lessonObj.lessonNumber || "",
            unitTitle: unit.title || "",
            title: lessonObj.title || "",
            subtitle: lessonObj.subtitle || lessonObj.description || "",
            fileName,
            url: lessonObj.path || (unitPath ? `${unitPath}/${fileName}` : fileName),
            isFinalReview: false
          });
        });
      });

      const review = course.finalReview;
      if(review && review.status !== "coming-soon"){
        const fileName = review.file || review.fileName || "";
        if(fileName){
          items.push({
            level: course.level || "",
            unit: "final",
            lessonNumber: "final",
            unitTitle: "Comprehensive Review",
            title: review.title || `${course.level || ""} Comprehensive Review`,
            subtitle: review.subtitle || "",
            fileName,
            url: review.path ? `${review.path}/${fileName}` : fileName,
            isFinalReview: true
          });
        }
      }
    });

    if(items.length){
      libraryStatus.textContent = `Loading ${items.length} workbook(s)...`;
    }

    for(const item of items){
      try{
        const lessonResponse = await fetch(item.url, { cache: "no-store" });
        if(!lessonResponse.ok) throw new Error(`Could not download ${item.url}`);
        const buffer = await lessonResponse.arrayBuffer();
        const pseudoFile = { name: item.fileName || item.url.split("/").pop(), lastModified: Date.now() };
        const summary = await readWorkbookSummary(buffer.slice(0), pseudoFile);

        if(item.level) summary.level = item.level;
        if(item.unit) summary.unit = item.unit;
        if(item.lessonNumber) summary.lesson = String(item.lessonNumber);
        if(item.title) summary.title = item.title;
        if(item.subtitle) summary.subtitle = item.subtitle;
        if(item.unitTitle) summary.unitTitle = item.unitTitle;
        if(item.isFinalReview) summary.isFinalReview = true;

        lessonLibrary.push({
          url: item.url,
          fileName: pseudoFile.name,
          file: pseudoFile,
          buffer,
          summary
        });
      }catch(lessonError){
        console.warn("Skipping unavailable workbook:", item.url, lessonError);
      }
    }

    lessonLibrary.sort((a,b) => naturalLessonSort(a.summary, b.summary));
    renderLessonLibrary();
  }catch(err){
    console.warn("Could not load course manifest.", err);
    libraryStatus.textContent = "Course could not load. Please refresh the page.";
    libraryBreadcrumb.textContent = "";
  }
}

async function buildLessonLibraryFromFolder(){
  const files = Array.from(lessonFolderInput.files || []).filter(f => {
  const n = f.name;
  return n.toLowerCase().endsWith(".xlsx") &&
         !n.startsWith("~$") &&
         !n.startsWith(".");
});
  lessonLibrary = [];
  selectedLibraryIndex = -1;
  lessonLibraryGrid.innerHTML = "";
  if(!files.length){
    libraryStatus.textContent = "No valid lesson workbooks found in that folder.";
    return;
  }
  libraryStatus.textContent = `Reading ${files.length} workbook(s)...`;

  for(const file of files){
    try{
      const buffer = await file.arrayBuffer();
      const summary = await readWorkbookSummary(buffer, file);
      lessonLibrary.push({ file, summary });
    }catch(e){
      console.warn("Could not read workbook", file.name, e);
      console.warn("Skipping invalid workbook:", file.name);
    }
  }

  lessonLibrary.sort((a,b) => naturalLessonSort(a.summary, b.summary));
  renderLessonLibrary();
}

function naturalLessonSort(a,b){
  const av = [a.level || "", a.isFinalReview ? 999 : Number(a.unit || 0), a.isFinalReview ? 999 : Number(a.lesson || 0), a.title || ""].join("-");
  const bv = [b.level || "", b.isFinalReview ? 999 : Number(b.unit || 0), b.isFinalReview ? 999 : Number(b.lesson || 0), b.title || ""].join("-");
  return av.localeCompare(bv, undefined, { numeric:true });
}

async function readWorkbookSummary(arrayBuffer, file){
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const workbookFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if(!workbookFile || !relsFile) throw new Error("Missing workbook structure.");

  const workbookXml = parser.parseFromString(await workbookFile.async("text"), "application/xml");
  const relsXml = parser.parseFromString(await relsFile.async("text"), "application/xml");

  const relMap = {};
  nodesByLocalName(relsXml, "Relationship").forEach(r => {
    const id = attr(r, "Id");
    let target = attr(r, "Target") || "";
    target = target.replace(/^\//, "");
    relMap[id] = target.startsWith("xl/") ? target : "xl/" + target;
  });

  const sheets = nodesByLocalName(workbookXml, "sheet").map(s => ({ name: attr(s,"name") || "", id: attr(s,"r:id") || attr(s,"id") }));
  const sharedStrings = await loadSharedStringsFromZip(zip, parser);

  let lessonRows = [];
  let vocabRows = [];
  let questionRows = [];

  const lessonSheet = sheets.find(s => s.name.toLowerCase() === "lesson") || sheets.find(s => s.name.toLowerCase() === "mission");
  const vocabSheet = sheets.find(s => s.name.toLowerCase() === "vocabulary");
  const qSheet = sheets.find(s => s.name.toLowerCase() === "questions");

  if(lessonSheet && relMap[lessonSheet.id]) lessonRows = await readSheetRowsFromZip(zip, parser, relMap[lessonSheet.id], sharedStrings);
  if(vocabSheet && relMap[vocabSheet.id]) vocabRows = await readSheetRowsFromZip(zip, parser, relMap[vocabSheet.id], sharedStrings);
  if(qSheet && relMap[qSheet.id]) questionRows = await readSheetRowsFromZip(zip, parser, relMap[qSheet.id], sharedStrings);

  const metadata = lessonRowsToMetadata(lessonRows);
  const titleFromSheet = questionRows && questionRows[0] && normalize(questionRows[0][0]) === "title" ? questionRows[0][1] : "";
  const qHeaderIndex = questionRows.findIndex(r => normalize(r[0]) === "id");
  const questionCount = qHeaderIndex >= 0 ? questionRows.slice(qHeaderIndex + 1).filter(r => r[0]).length : 0;
  const vocabCount = vocabRows.length > 1 ? vocabRows.slice(1).filter(r => r.some(Boolean)).length : 0;

  const firstHeading = (lessonRows[0] && lessonRows[0][1]) ? lessonRows[0][1] : "";
  const parsed = parseLessonHeading(firstHeading || file.name);
  return {
    level: metadata.Course || metadata.Level || parsed.level || inferFromFilename(file.name, "A"),
    unit: metadata.Unit || parsed.unit || inferFromFilename(file.name, "U"),
    lesson: metadata.Lesson || parsed.lesson || inferFromFilename(file.name, "L"),
    title: metadata.Title || metadata.LessonTitle || parsed.title || (lessonRows[1] && lessonRows[1][1]) || titleFromSheet || file.name,
    subtitle: metadata.Introduction || metadata.Overview || metadata.Subtitle || metadata.Goal || metadata["Mission Goal"] || metadata["Lesson Goal"] || "",
    estimated: metadata["Estimated Time"] || metadata.EstimatedTime || "",
    grammar: metadata["Grammar Focus"] || "",
    questions: questionCount,
    vocab: vocabCount,
    fileName: file.name
  };
}

async function loadSharedStringsFromZip(zip, parser){
  const f = zip.file("xl/sharedStrings.xml");
  if(!f) return [];
  const xml = parser.parseFromString(await f.async("text"), "application/xml");
  return nodesByLocalName(xml, "si").map(si => nodesByLocalName(si, "t").map(t => t.textContent).join(""));
}

async function readSheetRowsFromZip(zip, parser, path, sharedStrings){
  const sheetFile = zip.file(path);
  if(!sheetFile) return [];
  const xml = parser.parseFromString(await sheetFile.async("text"), "application/xml");
  const rows = [];
  nodesByLocalName(xml, "row").forEach(rowNode => {
    const row = [];
    nodesByLocalName(rowNode, "c").forEach(c => {
      const idx = colIndex(attr(c, "r"));
      const t = attr(c, "t");
      let val = "";
      if(t === "s") val = sharedStrings[parseInt(textContent(firstNodeByLocalName(c, "v")) || "0", 10)] || "";
      else if(t === "inlineStr") val = textContent(firstNodeByLocalName(c, "t"));
      else val = textContent(firstNodeByLocalName(c, "v"));
      row[idx] = val;
    });
    rows.push(row.map(v => v || ""));
  });
  return rows;
}

function lessonRowsToMetadata(rows){
  const meta = {};
  rows.forEach(row => {
    if(row[0] && row[1]){
      meta[String(row[0]).trim()] = String(row[1]).trim();
    }
  });
  return meta;
}


function parseLessonHeading(text){
  const raw = String(text || "");
  const result = { level:"", unit:"", lesson:"", title:"" };
  const levelMatch = raw.match(/\b(A1|A2|B1|B2|C1|C2)\b/i);
  const unitMatch = raw.match(/Unit\s+(\d+)/i);
  const lessonMatch = raw.match(/Lesson\s+(\d+)/i);
  if(levelMatch) result.level = levelMatch[1].toUpperCase();
  if(unitMatch) result.unit = unitMatch[1];
  if(lessonMatch) result.lesson = lessonMatch[1];

  const titleParts = raw.split(/\s[-–—]\s/);
  if(titleParts.length > 1){
    result.title = titleParts.slice(1).join(" - ").trim();
  }
  return result;
}

function inferFromFilename(name, letter){
  const pattern = new RegExp(letter + "(\\d+)", "i");
  const match = String(name).match(pattern);
  return match ? match[1] : "";
}


function getCourseByLevel(level){
  return (courseManifest.courses || []).find(c => String(c.level || "") === String(level || ""));
}

function getAvailableLessonEntry(level, unit, lesson){
  return lessonLibrary
    .map((entry, idx) => ({ entry, idx }))
    .find(x =>
      String(x.entry.summary.level || "Course") === String(level || "") &&
      String(x.entry.summary.unit || "1") === String(unit || "") &&
      String(x.entry.summary.lesson || "") === String(lesson || "")
    ) || null;
}

function getFinalReviewEntry(level){
  return lessonLibrary
    .map((entry, idx) => ({ entry, idx }))
    .find(x =>
      x.entry.summary &&
      x.entry.summary.isFinalReview &&
      String(x.entry.summary.level || "") === String(level || "")
    ) || null;
}

function getUnitGroups(levelFilter){
  const groups = [];
  const courses = levelFilter ? [getCourseByLevel(levelFilter)].filter(Boolean) : (courseManifest.courses || []);

  courses.forEach(course => {
    (course.units || []).forEach(unit => {
      const lessons = (unit.lessons || []).map(lesson => {
        const lessonObj = typeof lesson === "string" ? { file: lesson } : (lesson || {});
        const lessonNumber = String(lessonObj.lesson || lessonObj.lessonNumber || inferFromFilename(lessonObj.file || lessonObj.fileName || "", "L") || "");
        const available = getAvailableLessonEntry(course.level, unit.unit, lessonNumber);
        return { manifest: lessonObj, available, lessonNumber };
      });
      groups.push({
        level: course.level || "",
        unit: String(unit.unit || ""),
        unitTitle: unit.title || unitTitle(course.level, unit.unit),
        subtitle: unit.subtitle || "",
        lessons
      });
    });
  });

  return groups.sort((a,b) => Number(a.unit || 0) - Number(b.unit || 0));
}

function unitTitle(level, unit){
  const titles = {
    "A1|1": "Getting Started",
    "A1|2": "Around Me",
    "A1|3": "Everyday Life",
    "A1|4": "Shopping & Food",
    "A1|5": "Travel & Directions",
    "A1|6": "Final Review"
  };
  return titles[`${level}|${unit}`] || `Unit ${unit}`;
}

function handleLibraryBack(){
  if(libraryView === "lessons") renderUnitLibrary(selectedLevel);
  else renderLevelLibrary();
}

function renderLevelLibrary(){
  libraryView = "levels";
  selectedLevel = "";
  selectedUnit = "";
  selectedLibraryIndex = -1;
  lessonLibraryGrid.innerHTML = "";
  backToUnitsBtn.classList.add("hidden");
  libraryBreadcrumb.textContent = "Course Library";

  const courses = courseManifest.courses || [];
  if(!courses.length){
    libraryStatus.textContent = "No course levels are available yet.";
    lessonLibraryGrid.innerHTML = '<div class="lesson-card">No course levels were found.</div>';
    return;
  }

  libraryStatus.textContent = "Choose a level to begin.";
  courses.forEach(course => {
    const units = course.units || [];
    const actualLessons = lessonLibrary.filter(x => String(x.summary.level || "") === String(course.level || "")).length;
    const comingSoon = course.status === "coming-soon" || !units.length;
    const card = document.createElement("div");
    card.className = "lesson-library-card level-card" + (comingSoon ? " disabled" : "");
    card.title = comingSoon ? "Coming soon." : "Click to view units.";
    card.innerHTML = `
      <div class="unit-card-title">${escapeHtml(course.title || course.level || "Level")}</div>
      <div class="lesson-card-desc">${escapeHtml(course.subtitle || "")}</div>
      <div class="unit-card-desc">${comingSoon ? "Coming soon" : `${units.length} unit${units.length === 1 ? "" : "s"} • ${actualLessons} available lesson${actualLessons === 1 ? "" : "s"}`}</div>
      <div class="lesson-card-file">${comingSoon ? "Not available yet." : "Click to view this level."}</div>
    `;
    if(!comingSoon){
      card.addEventListener("click", () => renderUnitLibrary(course.level));
    }
    lessonLibraryGrid.appendChild(card);
  });
}

function renderUnitLibrary(level){
  libraryView = "units";
  selectedLevel = level || selectedLevel || "";
  selectedUnit = "";
  selectedLibraryIndex = -1;
  lessonLibraryGrid.innerHTML = "";
  backToUnitsBtn.classList.remove("hidden");

  const course = getCourseByLevel(selectedLevel);
  const groups = getUnitGroups(selectedLevel);
  if(!course || !groups.length){
    libraryStatus.textContent = "No units are available for this level yet.";
    libraryBreadcrumb.textContent = selectedLevel || "Course";
    lessonLibraryGrid.innerHTML = '<div class="lesson-card">No units are available yet.</div>';
    return;
  }

  libraryStatus.textContent = `${groups.length} unit(s) available. Choose a unit.`;
  libraryBreadcrumb.textContent = course.title || selectedLevel;

  groups.forEach(group => {
    const availableLessons = group.lessons.filter(x => x.available).length;
    const plannedLessons = group.lessons.length;
    const card = document.createElement("div");
    card.className = "lesson-library-card";
    card.title = "Click to view lessons.";
    card.innerHTML = `
      <div class="unit-card-title">${escapeHtml(group.level ? group.level + " • Unit " + group.unit : "Unit " + group.unit)}</div>
      <div class="lesson-card-meta">${escapeHtml(group.unitTitle || unitTitle(group.level, group.unit))}</div>
      ${group.subtitle ? `<div class="unit-card-desc">${escapeHtml(group.subtitle)}</div>` : ""}
      <div class="unit-card-desc">${availableLessons} available • ${plannedLessons} planned</div>
      <div class="lesson-card-file">Click to view this unit.</div>
    `;
    card.addEventListener("click", () => renderLessonsForUnit(group.level, group.unit));
    lessonLibraryGrid.appendChild(card);
  });

  if(course.finalReview){
    const review = course.finalReview || {};
    const finalEntry = getFinalReviewEntry(selectedLevel);
    const isAvailable = review.status !== "coming-soon" && !!finalEntry;
    const card = document.createElement("div");
    card.className = "lesson-library-card final-review-card" + (!isAvailable ? " disabled" : "");
    card.title = isAvailable ? "Click to begin the comprehensive review." : "Coming soon.";
    card.innerHTML = `
      <div class="unit-card-title">🎓 ${escapeHtml(review.title || "Comprehensive Review")}</div>
      <div class="lesson-card-meta">${escapeHtml(selectedLevel)} • Final Review</div>
      ${review.subtitle ? `<div class="unit-card-desc">${escapeHtml(review.subtitle)}</div>` : ""}
      ${isAvailable ? progressBadgeFor(finalEntry.entry.summary) : '<div class="coming-soon-badge">Coming soon</div>'}
      <div class="lesson-card-file">${isAvailable ? "Click to begin the final review." : "This review workbook has not been added yet."}</div>
    `;
    if(isAvailable){
      card.addEventListener("click", async () => {
        await loadLessonFromLibrary(finalEntry.idx);
        beginLessonFlow();
      });
    }
    lessonLibraryGrid.appendChild(card);
  }
}

function renderLessonsForUnit(level, unit){
  libraryView = "lessons";
  selectedLevel = level;
  selectedUnit = unit;
  lessonLibraryGrid.innerHTML = "";
  backToUnitsBtn.classList.remove("hidden");

  const course = getCourseByLevel(level);
  const unitObj = (course?.units || []).find(u => String(u.unit || "") === String(unit || ""));
  const planned = (unitObj?.lessons || []).map(lesson => {
    const lessonObj = typeof lesson === "string" ? { file: lesson } : (lesson || {});
    const lessonNumber = String(lessonObj.lesson || lessonObj.lessonNumber || inferFromFilename(lessonObj.file || lessonObj.fileName || "", "L") || "");
    const available = getAvailableLessonEntry(level, unit, lessonNumber);
    return { lessonObj, lessonNumber, available };
  });

  const title = unitObj?.title || unitTitle(level, unit);
  const availableCount = planned.filter(x => x.available).length;
  libraryStatus.textContent = `${availableCount} lesson(s) available. Choose any available lesson.`;
  libraryBreadcrumb.textContent = `${level} • Unit ${unit} - ${title}`;

  if(!planned.length){
    lessonLibraryGrid.innerHTML = '<div class="lesson-card">No lessons are available yet.</div>';
    return;
  }

  planned.forEach(({lessonObj, lessonNumber, available}) => {
    const entry = available?.entry;
    const idx = available?.idx;
    const s = entry?.summary || {};
    const isAvailable = !!available;
    const displayTitle = lessonObj.title || s.title || (lessonNumber === "review" ? "Unit Review" : `Lesson ${lessonNumber}`);
    const displaySubtitle = lessonObj.subtitle || s.subtitle || "";
    const card = document.createElement("div");
    card.className = "lesson-library-card" + (idx === selectedLibraryIndex ? " selected" : "") + (!isAvailable ? " disabled" : "");
    card.title = isAvailable ? "Click to begin lesson." : "Coming soon.";
    const metaParts = [level, `Unit ${unit}`];
    if(lessonNumber) metaParts.push(lessonNumber === "review" ? "Review" : `Lesson ${lessonNumber}`);
    card.innerHTML = `
      <div class="lesson-card-title">${escapeHtml(displayTitle)}</div>
      <div class="lesson-card-meta">${escapeHtml(metaParts.join(" • "))}</div>
      ${displaySubtitle ? `<div class="lesson-card-desc">${escapeHtml(displaySubtitle)}</div>` : ""}
      ${isAvailable ? `<div class="lesson-card-meta">${escapeHtml([s.vocab ? s.vocab + " vocabulary" : "", s.questions ? s.questions + " questions" : ""].filter(Boolean).join(" • "))}</div>` : '<div class="coming-soon-badge">Coming soon</div>'}
      ${isAvailable ? progressBadgeFor(s) : ""}
      <div class="lesson-card-file">${isAvailable ? "Click to begin this lesson." : "This lesson workbook has not been added yet."}</div>
    `;
    if(isAvailable){
      card.addEventListener("click", async () => {
        await loadLessonFromLibrary(idx);
        beginLessonFlow();
      });
    }
    lessonLibraryGrid.appendChild(card);
  });
}

function renderLessonLibrary(){
  renderLevelLibrary();
}

async function loadLessonFromLibrary(index){
  selectedLibraryIndex = index;
  const entry = lessonLibrary[index];
  try{
    workbookFileName = entry.file?.name || entry.fileName || entry.url || "Lesson workbook";
    currentLessonSummary = entry.summary;
    workbookLoadedAt = new Date();
    workbookLastModified = entry.file?.lastModified ? new Date(entry.file.lastModified) : new Date();

    loadStatus.textContent = "Opening selected lesson...";
    let buffer;
    if(entry.buffer){
      buffer = entry.buffer.slice(0);
    } else if(entry.url){
      const response = await fetch(entry.url, { cache: "no-store" });
      if(!response.ok) throw new Error(`Could not download lesson: ${entry.url}`);
      buffer = await response.arrayBuffer();
      entry.buffer = buffer.slice(0);
    } else if(entry.file && entry.file.arrayBuffer){
      buffer = await entry.file.arrayBuffer();
    } else {
      throw new Error("Lesson entry does not contain a file or URL.");
    }

    await parseWorkbook(buffer);
    loadStatus.textContent = `Selected lesson: ${entry.summary?.title || workbookFileName}`;
    quizStartArea.classList.remove("hidden");
  }catch(e){
    console.error(e);
    loadStatus.textContent = "There was a problem opening that lesson workbook.";
  }
}



function getCurrentLessonDisplayName(){
  if(currentLessonSummary){
    const parts = [];
    if(currentLessonSummary.level) parts.push(currentLessonSummary.level);
    if(currentLessonSummary.unit) parts.push(`Unit ${currentLessonSummary.unit}`);
    if(currentLessonSummary.lesson) parts.push(`Lesson ${currentLessonSummary.lesson}`);
    const prefix = parts.join(" • ");
    const title = currentLessonSummary.title || workbookFileName || "";
    return prefix ? `${prefix} - ${title}` : title;
  }
  return workbookFileName || "";
}

function updateCourseHeader(){
  if(currentLessonLabel){
    currentLessonLabel.textContent = getCurrentLessonDisplayName();
  }
}

function showCourseHome(){
  dynamicLessonSections = null;
  quizPanel.classList.add("hidden");
  lessonPanel.classList.add("hidden");
  resultsPanel.classList.add("hidden");
  setupPanel.classList.remove("hidden");
  quizHeader.classList.add("hidden");
  setupHeader.classList.remove("hidden");
  feedback.className = "feedback";
  feedback.innerHTML = "";
  updateDashboard();
  if(courseManifest && (courseManifest.courses || []).length){
    renderLevelLibrary();
  }
}

courseHomeBtn.addEventListener("click", showCourseHome);


function beginLessonFlow(){
  updateCourseHeader();
  currentLessonSectionIndex = 0;
  setupPanel.classList.add("hidden");
  resultsPanel.classList.add("hidden");
  quizPanel.classList.add("hidden");
  lessonPanel.classList.remove("hidden");
  setupHeader.classList.add("hidden");
  quizHeader.classList.remove("hidden");
  renderLessonSection();
}

function renderLessonNav(){
  lessonNav.innerHTML = "";
  const activeSections = dynamicLessonSections || lessonSections;

  const courseItem = document.createElement("span");
  courseItem.className = "lesson-step course-step";
  courseItem.textContent = "🏠 Course";
  courseItem.addEventListener("click", showCourseHome);
  lessonNav.appendChild(courseItem);

  activeSections.forEach((section, idx) => {
    const item = document.createElement("span");
    item.className = "lesson-step" + (idx === currentLessonSectionIndex ? " active" : "");
    item.textContent = section;
    item.addEventListener("click", () => {
      currentLessonSectionIndex = idx;
      renderLessonSection();
    });
    lessonNav.appendChild(item);
  });

  if(allQuestions && allQuestions.length){
    const quizItem = document.createElement("span");
    quizItem.className = "lesson-step";
    quizItem.textContent = "Knowledge Check";
    quizItem.addEventListener("click", () => startQuiz(allQuestions.length));
    lessonNav.appendChild(quizItem);
  }
}


function renderResultsNav(){
  if(!resultsLessonNav) return;
  resultsLessonNav.innerHTML = "";
  const activeSections = dynamicLessonSections || lessonSections;

  const courseItem = document.createElement("span");
  courseItem.className = "lesson-step course-step";
  courseItem.textContent = "🏠 Course";
  courseItem.addEventListener("click", showCourseHome);
  resultsLessonNav.appendChild(courseItem);

  activeSections.forEach((section, idx) => {
    const item = document.createElement("span");
    item.className = "lesson-step";
    item.textContent = section;
    item.addEventListener("click", () => {
      currentLessonSectionIndex = idx;
      resultsPanel.classList.add("hidden");
      quizPanel.classList.add("hidden");
      lessonPanel.classList.remove("hidden");
      renderLessonSection();
    });
    resultsLessonNav.appendChild(item);
  });

  const resultsItem = document.createElement("span");
  resultsItem.className = "lesson-step active";
  resultsItem.textContent = "Results";
  resultsLessonNav.appendChild(resultsItem);
}


function renderQuizNav(){
  if(!quizLessonNav) return;
  quizLessonNav.innerHTML = "";
  const activeSections = dynamicLessonSections || lessonSections;

  const courseItem = document.createElement("span");
  courseItem.className = "lesson-step course-step";
  courseItem.textContent = "🏠 Course";
  courseItem.addEventListener("click", showCourseHome);
  quizLessonNav.appendChild(courseItem);

  activeSections.forEach((section, idx) => {
    const item = document.createElement("span");
    item.className = "lesson-step";
    item.textContent = section;
    item.addEventListener("click", () => {
      currentLessonSectionIndex = idx;
      quizPanel.classList.add("hidden");
      resultsPanel.classList.add("hidden");
      lessonPanel.classList.remove("hidden");
      renderLessonSection();
    });
    quizLessonNav.appendChild(item);
  });

  const quizItem = document.createElement("span");
  quizItem.className = "lesson-step active";
  quizItem.textContent = "Knowledge Check";
  quizLessonNav.appendChild(quizItem);
}


function renderLessonSection(){
  const activeSections = dynamicLessonSections || lessonSections;
  const section = activeSections[currentLessonSectionIndex];
  lessonSectionTitle.textContent = section === "Mission" ? "Lesson" : section;
  renderLessonNav();
  lessonContent.innerHTML = "";

  const rows = lessonData[section] || [];
  if(section === "Lesson") renderLesson(rows);
  else if(section === "Vocabulary") renderVocabulary(rows);
  else if(section === "Grammar") renderGenericTable(rows);
  else if(section === "Dialogue") renderDialogue(rows);
  else if(section === "Practice") renderGenericTable(rows);
  else renderGenericTable(rows);

  prevLessonBtn.disabled = currentLessonSectionIndex === 0;
  nextLessonBtn.classList.toggle("hidden", currentLessonSectionIndex === activeSections.length - 1);
  startQuizFromLessonBtn.classList.toggle("hidden", currentLessonSectionIndex !== activeSections.length - 1 || !allQuestions.length);
}


function renderLesson(rows){
  if(!rows.length){
    lessonContent.innerHTML = '<div class="lesson-card"><h3>Lesson</h3><p>No lesson content found.</p></div>';
    return;
  }
  rows.forEach(row => {
    if(!row.some(Boolean)) return;
    const title = row[0] || "";
    const body = row.slice(1).filter(Boolean).join("\\n");
    const card = document.createElement("div");
    card.className = "lesson-card";
    card.innerHTML = body
      ? `<h3>${escapeHtml(title)}</h3><p>${escapeHtml(body)}</p>`
      : `<p>${escapeHtml(title)}</p>`;
    lessonContent.appendChild(card);
  });
}

function renderVocabulary(rows){
  if(rows.length <= 1){
    lessonContent.innerHTML = '<div class="lesson-card"><h3>Vocabulary</h3><p>No Vocabulary sheet found.</p></div>';
    return;
  }
  const grid = document.createElement("div");
  grid.className = "vocab-grid";
  rows.slice(1).forEach(row => {
    if(!row.some(Boolean)) return;
    const german = row[0] || "";
    const english = row[1] || "";
    const example = row[2] || "";
    const notes = row[3] || "";
    const card = document.createElement("div");
    card.className = "lesson-card";
    card.innerHTML = `
      <div class="vocab-term">${escapeHtml(german)}</div>
      <div class="vocab-meaning">${escapeHtml(english)}</div>
      ${example ? `<p><strong>Example:</strong> ${escapeHtml(example)}</p>` : ""}
      ${notes ? `<p class="small">${escapeHtml(notes)}</p>` : ""}
    `;
    grid.appendChild(card);
  });
  lessonContent.appendChild(grid);
}

function renderDialogue(rows){
  if(rows.length <= 1){
    lessonContent.innerHTML = '<div class="lesson-card"><h3>Dialogue</h3><p>No Dialogue sheet found.</p></div>';
    return;
  }
  const card = document.createElement("div");
  card.className = "lesson-card";
  rows.slice(1).forEach(row => {
    if(!row.some(Boolean)) return;
    const speaker = row[0] || "";
    const german = row[1] || "";
    const english = row[2] || "";
    const line = document.createElement("div");
    line.className = "dialogue-line";
    line.innerHTML = `<div class="speaker">${escapeHtml(speaker)}</div><div><strong>${escapeHtml(german)}</strong>${english ? `<br><span class="small">${escapeHtml(english)}</span>` : ""}</div>`;
    card.appendChild(line);
  });
  lessonContent.appendChild(card);
}

function renderGenericTable(rows){
  if(!rows.length){
    lessonContent.innerHTML = '<div class="lesson-card"><p>No content found for this section.</p></div>';
    return;
  }
  const table = document.createElement("table");
  table.className = "lesson-table";
  rows.forEach((row, idx) => {
    if(!row.some(Boolean)) return;
    const tr = document.createElement("tr");
    row.forEach(cell => {
      const el = document.createElement(idx === 0 ? "th" : "td");
      el.textContent = cell || "";
      tr.appendChild(el);
    });
    table.appendChild(tr);
  });
  const card = document.createElement("div");
  card.className = "lesson-card";
  card.appendChild(table);
  lessonContent.appendChild(card);
}

prevLessonBtn.addEventListener("click", () => {
  if(currentLessonSectionIndex > 0){
    currentLessonSectionIndex--;
    renderLessonSection();
  }
});

nextLessonBtn.addEventListener("click", () => {
  if(currentLessonSectionIndex < (dynamicLessonSections || lessonSections).length - 1){
    currentLessonSectionIndex++;
    renderLessonSection();
  }
});

startQuizFromLessonBtn.addEventListener("click", () => startQuiz(allQuestions.length));


function startQuiz(count){
  updateCourseHeader();
  const requestedCount = Math.min(count, allQuestions.length);
  selectedQuestionCount = requestedCount;
  quizQuestions = shuffle(allQuestions).slice(0, requestedCount);
  currentIndex = 0; score = 0; results = []; retryMode = false; retryScore = 0; retryResults = [];
  originalQuizQuestions = [...quizQuestions];
  questionStates = quizQuestions.map(() => ({ answered:false, userAnswer:"", correct:false, feedbackHtml:"", feedbackClass:"feedback" }));
  currentResultSaved = false;

  setupPanel.classList.add("hidden");
  resultsPanel.classList.add("hidden");
  lessonPanel.classList.add("hidden");
  quizPanel.classList.remove("hidden");
  renderQuizNav();
  setupHeader.classList.add("hidden");
  quizHeader.classList.remove("hidden");
  showQuestion();
}

function currentState(){ return questionStates[currentIndex] || {answered:false}; }
function answeredCount(){ return questionStates.filter(s => s && s.answered).length; }

function showQuestion(){
  answered = false; selectedAnswer = "";
  const q = quizQuestions[currentIndex];
  const state = currentState();

  progressText.textContent = retryMode ? `Retry Question ${currentIndex+1} of ${quizQuestions.length}` : `Question ${currentIndex+1} of ${quizQuestions.length}`;
  progressBar.style.width = `${(currentIndex/quizQuestions.length)*100}%`;
  scoreText.textContent = retryMode ? `Retry Score: ${retryScore}` : `Score: ${score}`;

  promptLabel.textContent = q.Exercise || "Practice";
  contextText.textContent = q.Phrase || "";
  taskText.textContent = q.Instruction || "";
  if(q.Target){
    targetText.textContent = q.Target;
    targetText.classList.remove("hidden");
  } else {
    targetText.classList.add("hidden");
    targetText.textContent = "";
  }

  choicesArea.innerHTML = "";
  fibInput.value = "";
  fibInput.disabled = false;
  feedback.className = "feedback";
  feedback.innerHTML = "";
  checkBtn.disabled = false;
  nextBtn.classList.add("hidden");

  if(q.Type === "FIB"){
    fibArea.classList.remove("hidden");
    fibInput.focus();
  } else {
    fibArea.classList.add("hidden");
    const choices = q.Type === "TF" ? ["True","False"] : [q.ChoiceA,q.ChoiceB,q.ChoiceC,q.ChoiceD].filter(Boolean);
    choices.forEach(choice => {
      if(q.Type === "MS"){
        const label = document.createElement("label");
        label.className = "choice";
        label.dataset.value = choice;
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        const span = document.createElement("span");
        span.textContent = choice;
        label.appendChild(checkbox);
        label.appendChild(span);
        checkbox.addEventListener("change", () => {
          if(answered){ checkbox.checked = !checkbox.checked; return; }
          label.classList.toggle("selected", checkbox.checked);
        });
        choicesArea.appendChild(label);
      } else {
        const button = document.createElement("button");
        button.className = "choice";
        button.dataset.value = choice;
        button.textContent = choice;
        button.addEventListener("click", () => {
          if(answered) return;
          document.querySelectorAll(".choice").forEach(btn => btn.classList.remove("selected"));
          button.classList.add("selected");
          selectedAnswer = choice;
        });
        choicesArea.appendChild(button);
      }
    });
  }

  if(state.answered){
    restoreAnsweredState(q, state);
  }

  updateDashboard();
}

function restoreAnsweredState(q, state){
  answered = true;
  selectedAnswer = state.userAnswer || "";
  if(q.Type === "FIB"){
    fibInput.value = state.userAnswer || "";
    fibInput.disabled = true;
  } else {
    document.querySelectorAll(".choice").forEach(btn => {
      const value = btn.dataset.value || btn.textContent;
      const selectedValues = splitAnswers(state.userAnswer);
      if(q.Type === "MS"){
        const cb = btn.querySelector('input[type="checkbox"]');
        if(selectedValues.includes(normalizeGerman(value))){
          btn.classList.add("selected");
          if(cb) cb.checked = true;
        }
        if(cb) cb.disabled = true;
      } else if(normalizeGerman(value) === normalizeGerman(state.userAnswer)){
        btn.classList.add("selected");
      }
      const isCorrectChoice = normalizedCorrectAnswers(q).includes(normalizeGerman(value));
      const wasSelected = q.Type === "MS" ? selectedValues.includes(normalizeGerman(value)) : normalizeGerman(value) === normalizeGerman(state.userAnswer);
      if(isCorrectChoice) btn.classList.add("correct");
      if(wasSelected && !isCorrectChoice) btn.classList.add("incorrect");
    });
  }
  feedback.className = state.feedbackClass;
  feedback.innerHTML = state.feedbackHtml;
  checkBtn.disabled = true;
  nextBtn.classList.remove("hidden");
}

function answerTokenToText(q, token){
  const raw = String(token ?? "").trim();
  const key = raw.toUpperCase();
  if(q.Type === "TF"){
    if(key === "1" || key === "TRUE") return "True";
    if(key === "0" || key === "FALSE") return "False";
  }
  const map = { "A": q.ChoiceA, "B": q.ChoiceB, "C": q.ChoiceC, "D": q.ChoiceD };
  return map[key] ? String(map[key]).trim() : raw;
}

function correctAnswerTexts(q){
  if(q.Type === "MS"){
    return String(q.Answer || "").split(/[|;]/).map(token => answerTokenToText(q, token)).filter(Boolean);
  }
  return [answerTokenToText(q, q.Answer)];
}

function canonicalAnswer(q){
  const first = String(q.Answer || "").split(/[|;]/)[0].trim();
  return answerTokenToText(q, first);
}

function normalizedCorrectAnswers(q){
  return correctAnswerTexts(q).map(v => normalizeGerman(v)).filter(Boolean).sort((a,b)=>a.localeCompare(b));
}

function splitAnswers(value){
  return String(value || "").split(/[|;]/).map(v => normalizeGerman(v)).filter(Boolean).sort((a,b)=>a.localeCompare(b));
}

function splitDisplayAnswers(value){
  return String(value || "").split(/[|;]/).map(v => v.trim()).filter(Boolean);
}

function answersMatch(userAnswers, correctAnswers){
  if(userAnswers.length !== correctAnswers.length) return false;
  return userAnswers.every((v,i)=>v===correctAnswers[i]);
}

function fibAnswerMatches(userAnswer, correctAnswer){
  const user = normalizeGerman(userAnswer);
  const correctOptions = String(correctAnswer || "").split(/[|;]/).map(v => normalizeGerman(v)).filter(Boolean);
  return correctOptions.includes(user);
}

function needsStandardGermanNote(userAnswer, answer){
  return userAnswer && answer &&
         normalize(userAnswer) !== normalize(answer) &&
         normalizeGerman(userAnswer) === normalizeGerman(answer);
}

function checkAnswer(){
  if(answered) return;
  const q = quizQuestions[currentIndex];
  let userAnswer = "";

  if(q.Type === "FIB") userAnswer = fibInput.value.trim();
  else if(q.Type === "MS") userAnswer = Array.from(document.querySelectorAll(".choice.selected")).map(btn => btn.dataset.value).join("|");
  else userAnswer = selectedAnswer;

  if(!userAnswer){
    feedback.className = "feedback incorrect";
    feedback.innerHTML = "<strong>Please enter or select an answer first.</strong>";
    return;
  }

  answered = true;
  let isCorrect = false;
  if(q.Type === "MS") isCorrect = answersMatch(splitAnswers(userAnswer), normalizedCorrectAnswers(q));
  else if(q.Type === "FIB") isCorrect = fibAnswerMatches(userAnswer, q.Answer);
  else isCorrect = normalizeGerman(userAnswer) === normalizedCorrectAnswers(q)[0];

  if(retryMode){ if(isCorrect) retryScore++; }
  else { if(isCorrect) score++; }

  document.querySelectorAll(".choice").forEach(btn => {
    const value = btn.dataset.value || btn.textContent;
    const isCorrectChoice = normalizedCorrectAnswers(q).includes(normalizeGerman(value));
    const wasSelected = q.Type === "MS" ? btn.classList.contains("selected") : normalizeGerman(value) === normalizeGerman(userAnswer);
    if(isCorrectChoice) btn.classList.add("correct");
    if(wasSelected && !isCorrectChoice) btn.classList.add("incorrect");
    const cb = btn.querySelector('input[type="checkbox"]');
    if(cb) cb.disabled = true;
  });

  const answerDisplay = canonicalAnswer(q);
  const userAnswerDisplay = q.Type === "MS" ? splitDisplayAnswers(userAnswer).join(", ") : userAnswer;
  const standardNote = (isCorrect && q.Type === "FIB" && needsStandardGermanNote(userAnswerDisplay, answerDisplay))
    ? `<br><strong>Standard German spelling:</strong> ${escapeHtml(answerDisplay)}`
    : "";

  feedback.className = isCorrect ? "feedback correct" : "feedback incorrect";
  feedback.innerHTML = `<strong>${isCorrect ? "Correct!" : "Not quite."}</strong><br>${!isCorrect ? `<strong>Your Answer:</strong> ${escapeHtml(userAnswerDisplay)}<br>` : ""}<strong>Answer:</strong> ${escapeHtml(answerDisplay)}${standardNote}<br><strong>Explanation:</strong> ${escapeHtml(q.Explanation || "No explanation provided.")}`;

  const attemptResult = { question: buildReviewQuestion(q), yourAnswer: userAnswerDisplay, correctAnswer: answerDisplay, explanation: q.Explanation, correct: isCorrect };
  if(retryMode) retryResults.push(attemptResult);
  else results.push(attemptResult);

  questionStates[currentIndex] = {
    answered:true,
    userAnswer:userAnswer,
    correct:isCorrect,
    feedbackHtml:feedback.innerHTML,
    feedbackClass:feedback.className
  };

  checkBtn.disabled = true;
  nextBtn.classList.remove("hidden");
  scoreText.textContent = retryMode ? `Retry Score: ${retryScore}` : `Score: ${score}`;
  updateDashboard();
}

function buildReviewQuestion(q){
  return [q.Exercise, q.Phrase, q.Instruction, q.Target].filter(Boolean).join(" | ");
}

checkBtn.addEventListener("click", checkAnswer);
fibInput.addEventListener("keydown", e => { if(e.key === "Enter") checkAnswer(); });
nextBtn.addEventListener("click", () => {
  currentIndex++;
  if(currentIndex >= quizQuestions.length){ if(retryMode) showRetryResults(); else showResults(); }
  else showQuestion();
});
quitBtn.addEventListener("click", showResults);

function updateDashboard(){
  const total = quizQuestions.length || allQuestions.length || 0;
  const currentScore = retryMode ? retryScore : score;
  const answered = answeredCount();
  const incorrect = Math.max(answered - currentScore, 0);
  const accuracy = answered ? Math.round((currentScore / answered) * 100) : 0;
  dashTotal.textContent = total;
  dashCorrect.textContent = currentScore;
  dashIncorrect.textContent = incorrect;
  dashScore.textContent = accuracy + "%";

  questionMap.innerHTML = "";
  for(let i=0;i<total;i++){
    const state = questionStates[i] || {};
    const box = document.createElement("div");
    box.className = "legend-box";
    if(i === currentIndex && !quizPanel.classList.contains("hidden")) box.classList.add("current");
    if(state.answered) box.classList.add(state.correct ? "correct" : "incorrect");
    box.textContent = i + 1;
    if(state.answered || i === currentIndex){
      box.style.cursor = "pointer";
      box.addEventListener("click", () => { currentIndex = i; showQuestion(); });
    }
    questionMap.appendChild(box);
  }
}

function getResultMessage(percent){
  if(percent >= 95) return "Ausgezeichnet! Excellent work.";
  if(percent >= 90) return "Sehr gut. You are ready to move forward.";
  if(percent >= 80) return "Good progress. Review the missed questions.";
  if(percent >= 70) return "Solid start. More practice is recommended.";
  return "Review this mission before moving forward.";
}

function showResults(){
  progressBar.style.width = "100%";
  saveCurrentLessonResult();
  quizPanel.classList.add("hidden");
  resultsPanel.classList.remove("hidden");
  renderResultsNav();
  quizHeader.classList.add("hidden");
  setupHeader.classList.remove("hidden");
  const total = results.length || quizQuestions.length;
  const percent = total ? Math.round((score/total)*100) : 0;
  finalScore.textContent = `${score} / ${total} (${percent}%)`;
  finalSummary.textContent = getResultMessage(percent);
  saveStatus.classList.remove("hidden");
  saveStatus.innerHTML = learnerName === "Guest"
    ? "Practice finished. Your progress was saved to the Guest profile in this browser."
    : `Practice finished. Progress saved for ${escapeHtml(learnerName)} in this browser.`;
  reviewArea.innerHTML = "";
  updateDashboard();
}


restartBtn.addEventListener("click", () => {
  questionStates = [];
  updateDashboard();
  resultsPanel.classList.add("hidden");
  lessonPanel.classList.add("hidden");
  setupPanel.classList.remove("hidden");
  quizHeader.classList.add("hidden");
  setupHeader.classList.remove("hidden");
  reviewArea.innerHTML = "";
  saveStatus.classList.add("hidden");
  retryMode = false;
  retryQuestions = [];
  retryResults = [];
  originalResults = [];
  retryScore = 0;
  originalQuizQuestions = [];
});

missedBtn.addEventListener("click", () => {
  const missed = results.filter(r => !r.correct);
  if(missed.length === 0){ reviewArea.innerHTML = "<p><strong>No missed questions. Excellent!</strong></p>"; return; }
  const rows = missed.map(r => `<tr><td>${escapeHtml(r.question)}</td><td>${escapeHtml(r.yourAnswer)}</td><td>${escapeHtml(r.correctAnswer)}</td><td>${escapeHtml(r.explanation || "")}</td></tr>`).join("");
  reviewArea.innerHTML = `<h3>Missed Questions</h3><table><thead><tr><th>Question</th><th>Your Answer</th><th>Correct Answer</th><th>Explanation</th></tr></thead><tbody>${rows}</tbody></table>`;
});

retryMissedBtn.addEventListener("click", startRetryMissedQuestions);
returnCourseBtn.addEventListener("click", showCourseHome);

function startRetryMissedQuestions(){
  const missed = results.filter(r => !r.correct);
  if(missed.length === 0){
    reviewArea.innerHTML = "<p><strong>No missed questions to retry. Excellent!</strong></p>";
    return;
  }

  originalResults = [...results];
  originalQuizQuestions = [...quizQuestions];

  retryQuestions = missed.map(missedItem => originalQuizQuestions.find(q => buildReviewQuestion(q) === missedItem.question)).filter(Boolean);
  if(retryQuestions.length === 0){
    reviewArea.innerHTML = "<p><strong>Unable to build retry round from missed questions.</strong></p>";
    return;
  }

  retryMode = true;
  retryResults = [];
  retryScore = 0;
  currentIndex = 0;
  quizQuestions = retryQuestions;
  questionStates = quizQuestions.map(() => ({ answered:false, userAnswer:"", correct:false, feedbackHtml:"", feedbackClass:"feedback" }));

  resultsPanel.classList.add("hidden");
  setupPanel.classList.add("hidden");
  quizPanel.classList.remove("hidden");
  setupHeader.classList.add("hidden");
  quizHeader.classList.remove("hidden");
  showQuestion();
}

function showRetryResults(){
  progressBar.style.width = "100%";
  quizPanel.classList.add("hidden");
  resultsPanel.classList.remove("hidden");
  renderResultsNav();
  quizHeader.classList.add("hidden");
  setupHeader.classList.remove("hidden");

  const total = retryResults.length || retryQuestions.length;
  const percent = total ? Math.round((retryScore/total)*100) : 0;

  finalScore.textContent = `Retry Round: ${retryScore} / ${total} (${percent}%)`;
  finalSummary.textContent = retryScore === total ? "Excellent. You corrected all missed questions." : "Review any questions still missed below.";
  saveStatus.classList.add("hidden");

  const stillMissed = retryResults.filter(r => !r.correct);
  if(stillMissed.length === 0){
    reviewArea.innerHTML = "<p><strong>No missed questions remain from the retry round.</strong></p>";
  } else {
    const rows = stillMissed.map(r => `<tr><td>${escapeHtml(r.question)}</td><td>${escapeHtml(r.yourAnswer)}</td><td>${escapeHtml(r.correctAnswer)}</td><td>${escapeHtml(r.explanation || "")}</td></tr>`).join("");
    reviewArea.innerHTML = `<h3>Still Missed After Retry</h3><table><thead><tr><th>Question</th><th>Your Retry Answer</th><th>Correct Answer</th><th>Explanation</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  results = originalResults;
  quizQuestions = originalQuizQuestions;
  retryMode = false;
}

// GitHub Pages / web mode: load course manifest automatically.
if(document.readyState === "loading"){
  document.addEventListener("DOMContentLoaded", loadCourseManifest);
} else {
  loadCourseManifest();
}