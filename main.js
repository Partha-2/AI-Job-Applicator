import './style.css';
import { createIcons, icons } from 'lucide';
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';
import mammoth from 'mammoth';
import Papa from 'papaparse';
import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import * as xlsx from 'xlsx';

GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/legacy/build/pdf.worker.mjs', import.meta.url).toString();

createIcons({ icons });

const APPLIED_CACHE_KEY = 'jobApplicator.appliedRecords';
const THEME_KEY = 'jobApplicator.theme';
const CLIENT_ID_KEY = 'jobApplicator.clientId';
const OUTREACH_DRAFT_KEY = 'jobApplicator.outreachDraft';
const ATS_TAILOR_KEY = 'jobApplicator.atsTailor';
const GMAIL_APP_USER_KEY = 'jobApplicator.gmailAppUser';
const GMAIL_APP_PASSWORD_KEY = 'jobApplicator.gmailAppPassword';
const DRAFT_DB_NAME = 'jobApplicatorDrafts';
const DRAFT_FILE_STORE = 'files';
const MANUAL_ATTACHMENT_KEY = 'manualAttachments';
const BULK_RESUME_KEY = 'bulkResume';
const APPLIED_STATUSES = ['Applied', 'Interviewing', 'Follow-up', 'Rejected', 'Offer'];

let jobs = [];
let appliedRecords = [];
let currentPage = 1;
const pageSize = 14;
let currentUser = null;
let mailCapability = { canSendMail: false, mode: 'unavailable', senderEmail: '', requiresGoogleAuth: false };
let gmailAppUser = '';
let gmailAppPassword = '';
let selectedJob = null;
let selectedAppliedJob = null;
let outreachContacts = [];
let outreachVariants = [];
let savedManualAttachments = [];
let savedResumeFiles = [];
let atsAnalysis = null;
let atsResumeFileName = '';

const clientId = ensureClientId();

const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const dashboardSection = document.getElementById('dashboardSection');
const jobList = document.getElementById('jobList');
const jobCount = document.getElementById('jobCount');
const resetBtn = document.getElementById('resetBtn');
const paginationControls = document.getElementById('paginationControls');
const jobDetailsEmptyState = document.getElementById('detailsEmptyState');
const jobDetailsContent = document.getElementById('detailsContent');
const appStatus = document.getElementById('appStatus');

const heroPendingCount = document.getElementById('heroPendingCount');
const heroAppliedCount = document.getElementById('heroAppliedCount');
const appliedCount = document.getElementById('appliedCount');
const appliedTotalMetric = document.getElementById('appliedTotalMetric');
const interviewMetric = document.getElementById('interviewMetric');
const offerMetric = document.getElementById('offerMetric');
const appliedUpdatedAt = document.getElementById('appliedUpdatedAt');
const appliedList = document.getElementById('appliedList');
const appliedDetailsEmpty = document.getElementById('appliedDetailsEmpty');
const appliedDetailsContent = document.getElementById('appliedDetailsContent');

const tabBtns = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.tab-panel');

const outreachFileInput = document.getElementById('outreachFileInput');
const outreachList = document.getElementById('outreachList');
const outreachEmptyState = document.getElementById('outreachEmptyState');
const outreachStatus = document.getElementById('outreachStatus');
const manualStatus = document.getElementById('manualStatus');
const scrapeStatus = document.getElementById('scrapeStatus');
const bulkSenderStatus = document.getElementById('bulkSenderStatus');
const bulkLoginHint = document.getElementById('bulkLoginHint');
const bulkAppPasswordBtn = document.getElementById('bulkAppPasswordBtn');
const manualSenderStatus = document.getElementById('manualSenderStatus');
const manualLoginHint = document.getElementById('manualLoginHint');
const manualAppPasswordBtn = document.getElementById('manualAppPasswordBtn');
const resumeUpload = document.getElementById('resumeUpload');
const bulkSubjectInput = document.getElementById('bulkSubject');
const bulkBodyInput = document.getElementById('bulkBody');
const bulkPreview = document.getElementById('bulkPreview');
const manualToEmailInput = document.getElementById('manualToEmail');
const manualSubjectInput = document.getElementById('manualSubject');
const manualBodyInput = document.getElementById('manualBody');
const manualAttachmentsInput = document.getElementById('manualAttachments');
const scrapeUrlInput = document.getElementById('scrapeUrl');
const savedAttachmentInfo = document.getElementById('savedAttachmentInfo');
const savedResumeInfo = document.getElementById('savedResumeInfo');
const fireAllMailsBtn = document.getElementById('fireAllMailsBtn');
const sendManualBtn = document.getElementById('sendManualBtn');
const headerLoginBtn = document.getElementById('loginBtn');
const atsResumeInput = document.getElementById('atsResumeInput');
const atsResumeFileInput = document.getElementById('atsResumeFileInput');
const atsResumeFileInfo = document.getElementById('atsResumeFileInfo');
const atsJobInput = document.getElementById('atsJobInput');
const generateAtsBtn = document.getElementById('generateAtsBtn');
const clearAtsBtn = document.getElementById('clearAtsBtn');
const downloadAtsDocxBtn = document.getElementById('downloadAtsDocxBtn');
const atsStatus = document.getElementById('atsStatus');
const atsMatchScore = document.getElementById('atsMatchScore');
const atsMatchedCount = document.getElementById('atsMatchedCount');
const atsMissingCount = document.getElementById('atsMissingCount');
const atsMatchedKeywords = document.getElementById('atsMatchedKeywords');
const atsMissingKeywords = document.getElementById('atsMissingKeywords');
const atsSummaryOutput = document.getElementById('atsSummaryOutput');
const atsBulletsOutput = document.getElementById('atsBulletsOutput');

const walkInList = document.getElementById('walkInList');
const walkInMeta = document.getElementById('walkInMeta');

setupTabs();
setupTheme();
bindEvents();
restoreOutreachDraft();
restoreGmailAppPassword();
restoreAtsDraft();
initializeStoredFiles();
updateSenderUi();
checkUserSession();
renderAtsAnalysis();

function ensureClientId() {
  const existing = localStorage.getItem(CLIENT_ID_KEY);
  if (existing) return existing;

  const generated = globalThis.crypto?.randomUUID?.() || `client-${Math.random().toString(36).slice(2)}${Date.now()}`;
  localStorage.setItem(CLIENT_ID_KEY, generated);
  return generated;
}

function setupTabs() {
  tabBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target;
      tabBtns.forEach((item) => item.classList.toggle('active', item === btn));
      tabPanels.forEach((panel) => panel.classList.toggle('hidden', panel.id !== target));
    });
  });
}

function setupTheme() {
  const savedTheme = localStorage.getItem(THEME_KEY) || 'theme-dark';
  document.body.classList.remove('theme-dark', 'theme-light');
  document.body.classList.add(savedTheme);

  document.getElementById('themeToggle').addEventListener('click', () => {
    const nextTheme = document.body.classList.contains('theme-dark') ? 'theme-light' : 'theme-dark';
    document.body.classList.remove('theme-dark', 'theme-light');
    document.body.classList.add(nextTheme);
    localStorage.setItem(THEME_KEY, nextTheme);
  });
}

function bindEvents() {
  fileInput.addEventListener('change', handleFileUpload);
  uploadSection.addEventListener('dragover', (event) => {
    event.preventDefault();
    uploadSection.classList.add('dragover');
  });
  uploadSection.addEventListener('dragleave', () => uploadSection.classList.remove('dragover'));
  uploadSection.addEventListener('drop', (event) => {
    event.preventDefault();
    uploadSection.classList.remove('dragover');
    if (event.dataTransfer.files.length) {
      fileInput.files = event.dataTransfer.files;
      handleFileUpload();
    }
  });

  resetBtn.addEventListener('click', () => {
    jobs = [];
    selectedJob = null;
    fileInput.value = '';
    uploadSection.classList.remove('hidden');
    dashboardSection.classList.add('hidden');
    renderJobs();
    renderJobDetails();
    showBanner('Upload reset. You can import a fresh jobs file now.', 'info');
  });

  document.getElementById('prevPageBtn').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage -= 1;
      renderJobs();
    }
  });

  document.getElementById('nextPageBtn').addEventListener('click', () => {
    const pendingJobs = jobs.filter((job) => !job.applied);
    if (currentPage * pageSize < pendingJobs.length) {
      currentPage += 1;
      renderJobs();
    }
  });

  outreachFileInput.addEventListener('change', handleOutreachUpload);
  document.getElementById('fireAllMailsBtn').addEventListener('click', sendBulkOutreach);
  document.getElementById('sendManualBtn').addEventListener('click', sendManualEmail);
  document.getElementById('scrapeBtn').addEventListener('click', scrapeEmailsFromUrl);
  document.getElementById('scanWalkinsBtn').addEventListener('click', scanWalkins);
  bulkAppPasswordBtn?.addEventListener('click', openGmailAppPasswordModal);
  manualAppPasswordBtn?.addEventListener('click', openGmailAppPasswordModal);
  generateAtsBtn?.addEventListener('click', generateAtsTailor);
  clearAtsBtn?.addEventListener('click', clearAtsTailor);
  downloadAtsDocxBtn?.addEventListener('click', downloadAtsDocx);
  atsResumeFileInput?.addEventListener('change', handleAtsResumeUpload);
  headerLoginBtn.addEventListener('click', () => {
    window.location.href = '/auth/google';
  });

  [bulkSubjectInput, bulkBodyInput, manualToEmailInput, manualSubjectInput, manualBodyInput, scrapeUrlInput].forEach((element) => {
    element.addEventListener('input', persistOutreachDraft);
  });

  [bulkSubjectInput, bulkBodyInput].forEach((element) => {
    element.addEventListener('input', renderBulkPreview);
  });

  [atsResumeInput, atsJobInput].forEach((element) => {
    element?.addEventListener('input', persistAtsDraft);
  });

  resumeUpload.addEventListener('change', async () => {
    savedResumeFiles = Array.from(resumeUpload.files || []);
    await saveStoredFiles(BULK_RESUME_KEY, savedResumeFiles);
    renderSavedFileInfo();
  });

  manualAttachmentsInput.addEventListener('change', async () => {
    savedManualAttachments = Array.from(manualAttachmentsInput.files || []);
    await saveStoredFiles(MANUAL_ATTACHMENT_KEY, savedManualAttachments);
    renderSavedFileInfo();
  });
}

async function apiFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set('x-client-id', clientId);

  return fetch(url, {
    credentials: 'include',
    ...options,
    headers
  });
}

async function safeJson(response) {
  const text = await response.text();
  if (!text) {
    if (!response.ok) {
      throw new Error(`Backend request failed (${response.status}). Make sure the API server is running on port 3000.`);
    }
    throw new Error('Backend returned an empty response. Make sure the API server is running on port 3000.');
  }

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Backend returned invalid JSON for ${new URL(response.url).pathname || 'this request'}.`);
  }
}

async function checkUserSession() {
  try {
    const healthResponse = await apiFetch('/api/health');
    const health = await safeJson(healthResponse);
    mailCapability = health.mail || mailCapability;

    const userResponse = await apiFetch('/api/user');
    const user = await safeJson(userResponse);
    currentUser = user;

    if (user) {
      document.getElementById('userProfile').classList.remove('hidden');
      document.getElementById('userName').innerText = user.displayName;
    }
  } catch (error) {
    console.warn('Backend session check failed.', error);
    showBanner('Backend unavailable. Start `node server.js` to enable applied history, outreach, and walk-in scan.', 'warning');
  }

  updateSenderUi();

  await loadAppliedRecords();
}

async function loadAppliedRecords() {
  try {
    const response = await apiFetch('/api/applied-jobs');
    if (!response.ok) throw new Error('Unable to load applied history.');
    appliedRecords = await safeJson(response);
    persistAppliedCache();
  } catch (error) {
    appliedRecords = readAppliedCache();
  }

  if (!selectedAppliedJob && appliedRecords.length > 0) {
    selectedAppliedJob = appliedRecords[0];
  }

  if (jobs.length > 0) {
    const appliedIds = new Set(appliedRecords.map((record) => record.id));
    jobs = jobs.map((job) => ({ ...job, applied: appliedIds.has(job.id) }));
    if (selectedJob?.id && appliedIds.has(selectedJob.id)) {
      selectedJob = jobs.find((job) => !job.applied) || null;
    }
    renderJobs();
    renderJobDetails();
  }

  renderAppliedTracker();
  refreshCounts();
}

function readAppliedCache() {
  try {
    return JSON.parse(localStorage.getItem(APPLIED_CACHE_KEY) || '[]');
  } catch {
    return [];
  }
}

function persistAppliedCache() {
  localStorage.setItem(APPLIED_CACHE_KEY, JSON.stringify(appliedRecords));
}

function getDefaultManualBody() {
  return 'Hi [Recruiter Name],\n\nI wanted to share my profile for opportunities at [Company]. Please find my resume attached.\n\nBest regards,';
}

function getDefaultBulkSubject() {
  return 'Application for relevant opportunity';
}

function getDefaultBulkBody() {
  return getDefaultManualBody();
}

function restoreOutreachDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem(OUTREACH_DRAFT_KEY) || '{}');
    bulkSubjectInput.value = draft.bulkSubject || getDefaultBulkSubject();
    bulkBodyInput.value = draft.bulkBody || getDefaultBulkBody();
    manualToEmailInput.value = draft.manualToEmail || '';
    manualSubjectInput.value = draft.manualSubject || 'Application for relevant opportunity';
    manualBodyInput.value = draft.manualBody || getDefaultManualBody();
    scrapeUrlInput.value = draft.scrapeUrl || '';
  } catch {
    bulkSubjectInput.value = getDefaultBulkSubject();
    bulkBodyInput.value = getDefaultBulkBody();
    manualSubjectInput.value = 'Application for relevant opportunity';
    manualBodyInput.value = getDefaultManualBody();
  }

  renderBulkPreview();
}

function persistOutreachDraft() {
  localStorage.setItem(OUTREACH_DRAFT_KEY, JSON.stringify({
    bulkSubject: bulkSubjectInput.value.trim(),
    bulkBody: bulkBodyInput.value,
    manualToEmail: manualToEmailInput.value.trim(),
    manualSubject: manualSubjectInput.value.trim(),
    manualBody: manualBodyInput.value,
    scrapeUrl: scrapeUrlInput.value.trim()
  }));
}

function restoreAtsDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem(ATS_TAILOR_KEY) || '{}');
    atsResumeInput.value = draft.resume || '';
    atsJobInput.value = draft.jobDescription || '';
    atsAnalysis = draft.analysis || null;
    atsResumeFileName = draft.resumeFileName || '';
    atsResumeFileInfo.textContent = atsResumeFileName
      ? `Last uploaded resume: ${atsResumeFileName}`
      : 'Upload `.txt`, `.docx`, or `.pdf`, or paste resume text manually below.';
  } catch {
    atsAnalysis = null;
    atsResumeFileName = '';
  }
}

function persistAtsDraft() {
  localStorage.setItem(ATS_TAILOR_KEY, JSON.stringify({
    resume: atsResumeInput.value,
    jobDescription: atsJobInput.value,
    analysis: atsAnalysis,
    resumeFileName: atsResumeFileName
  }));
}

function persistAtsResumeFileName(fileName = '') {
  atsResumeFileName = fileName;
  localStorage.setItem(ATS_TAILOR_KEY, JSON.stringify({
    resume: atsResumeInput.value,
    jobDescription: atsJobInput.value,
    analysis: atsAnalysis,
    resumeFileName: fileName
  }));
}

async function extractPdfText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await getDocument({ data: arrayBuffer }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => ('str' in item ? item.str : '')).join(' '));
  }

  return pages.join('\n').trim();
}

async function extractDocxText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value.trim();
}

async function extractTextFromResumeFile(file) {
  const lowerName = file.name.toLowerCase();

  if (lowerName.endsWith('.txt')) {
    return file.text();
  }

  if (lowerName.endsWith('.docx')) {
    return extractDocxText(file);
  }

  if (lowerName.endsWith('.pdf')) {
    return extractPdfText(file);
  }

  throw new Error('Unsupported file type. Upload `.txt`, `.docx`, or `.pdf`.');
}

async function handleAtsResumeUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  atsResumeFileInfo.textContent = `Reading ${file.name}...`;
  setInlineStatus(atsStatus, `Reading resume from ${file.name}...`, 'info');

  try {
    const extractedText = (await extractTextFromResumeFile(file)).trim();
    if (!extractedText) {
      throw new Error('Could not extract readable text from that file.');
    }

    atsResumeInput.value = extractedText;
    persistAtsResumeFileName(file.name);
    atsResumeFileInfo.textContent = `Loaded resume: ${file.name}`;
    setInlineStatus(atsStatus, `Resume uploaded from ${file.name}. Review the text, then generate the ATS tailor.`, 'success');
  } catch (error) {
    atsResumeFileInfo.textContent = `Failed to read ${file.name}`;
    setInlineStatus(atsStatus, error.message || 'Resume upload failed.', 'error');
  }
}

const ATS_KEYWORDS = [
  'java', 'python', 'javascript', 'typescript', 'react', 'node.js', 'spring boot', 'sql', 'aws',
  'azure', 'gcp', 'excel', 'power bi', 'tableau', 'analytics', 'reporting', 'stakeholder management',
  'product strategy', 'roadmap', 'user research', 'ui/ux', 'figma', 'seo', 'content marketing',
  'performance marketing', 'sales pipeline', 'crm', 'lead generation', 'customer success',
  'operations', 'process improvement', 'vendor management', 'communication', 'leadership',
  'project management', 'agile', 'scrum', 'problem solving', 'cross-functional collaboration'
];
const ATS_STOP_WORDS = new Set([
  'and', 'the', 'with', 'from', 'into', 'that', 'this', 'will', 'your', 'their', 'have', 'has',
  'had', 'for', 'our', 'you', 'are', 'but', 'not', 'all', 'any', 'can', 'able', 'using', 'use',
  'used', 'about', 'than', 'then', 'them', 'they', 'who', 'what', 'when', 'where', 'while',
  'work', 'works', 'working', 'role', 'team', 'years', 'year', 'experience', 'skills', 'skill',
  'requirements', 'preferred', 'strong', 'good', 'excellent', 'knowledge', 'understanding'
]);

function normalizeKeywordText(text) {
  return (text || '').toLowerCase().replace(/[^a-z0-9+.#/\s-]/g, ' ');
}

function getRelevantAtsKeywords(jobText) {
  const normalized = normalizeKeywordText(jobText);
  const matchedLibraryTerms = ATS_KEYWORDS.filter((keyword) => normalized.includes(keyword));
  const phraseMatches = Array.from(new Set(
    (jobText.match(/\b[A-Za-z0-9+/#.-]{3,}(?:\s+[A-Za-z0-9+/#.-]{3,}){0,2}\b/g) || [])
      .map((phrase) => phrase.toLowerCase().trim())
      .filter((phrase) => {
        const words = phrase.split(/\s+/);
        return words.some((word) => word.length > 3 && !ATS_STOP_WORDS.has(word));
      })
      .filter((phrase) => !ATS_STOP_WORDS.has(phrase))
  )).slice(0, 20);

  return Array.from(new Set([...matchedLibraryTerms, ...phraseMatches]))
    .filter((keyword) => keyword.length > 2)
    .slice(0, 18);
}

function formatKeywordList(keywords) {
  if (!keywords.length) return '';
  if (keywords.length === 1) return keywords[0];
  if (keywords.length === 2) return `${keywords[0]} and ${keywords[1]}`;
  return `${keywords.slice(0, -1).join(', ')}, and ${keywords.at(-1)}`;
}

function buildTailoredSummary(jobText, matchedKeywords, missingKeywords) {
  const primaryFocus = matchedKeywords.slice(0, 4);
  const growthAreas = missingKeywords.slice(0, 3);
  const jdLine = jobText.split(/\n+/).find((line) => line.trim().length > 30) || 'high-impact roles aligned with your experience';
  const firstSentence = primaryFocus.length
    ? `Professional profile aligned with ${formatKeywordList(primaryFocus)} and positioned around the highest-signal skills from this target role.`
    : 'Professional profile positioned around measurable outcomes, relevant tools, and clear business impact for the target role.';
  const secondSentence = `Tailored for a role centered on ${jdLine.trim().slice(0, 120)}${jdLine.trim().length > 120 ? '...' : ''}`;
  const thirdSentence = growthAreas.length
    ? `Resume emphasis should clearly show hands-on exposure to ${formatKeywordList(growthAreas)} where relevant and truthful.`
    : 'Resume emphasis should highlight measurable outcomes, ownership, and collaboration with cross-functional teams.';

  return [firstSentence, secondSentence, thirdSentence].join(' ');
}

function buildTailoredBullets(matchedKeywords, missingKeywords) {
  const bulletSeeds = [
    `Tailor one bullet around ${formatKeywordList(matchedKeywords.slice(0, 3)) || 'the strongest overlapping skills from the job description'} with a measurable result.`,
    `Add one bullet that shows impact through metrics, ownership, delivery speed, revenue influence, efficiency gains, or customer outcomes.`,
    `Show cross-functional collaboration with the teams most relevant to the role, such as product, design, data, operations, sales, or engineering.`,
    `Only add keywords like ${formatKeywordList(missingKeywords.slice(0, 3)) || 'target role terminology'} where your real experience can support them honestly.`
  ];

  return bulletSeeds.map((line) => `• ${line}`).join('\n');
}

function renderKeywordChips(container, keywords, emptyLabel) {
  container.innerHTML = '';
  if (!keywords.length) {
    container.innerHTML = `<span class="empty-chip">${emptyLabel}</span>`;
    return;
  }

  keywords.forEach((keyword) => {
    const chip = document.createElement('span');
    chip.className = 'keyword-chip';
    chip.textContent = keyword;
    container.appendChild(chip);
  });
}

function renderAtsAnalysis() {
  const analysis = atsAnalysis || {
    score: 0,
    matchedKeywords: [],
    missingKeywords: [],
    tailoredSummary: '',
    tailoredBullets: ''
  };

  atsMatchScore.textContent = `${analysis.score || 0}%`;
  atsMatchedCount.textContent = `${analysis.matchedKeywords?.length || 0}`;
  atsMissingCount.textContent = `${analysis.missingKeywords?.length || 0}`;
  atsSummaryOutput.value = analysis.tailoredSummary || '';
  atsBulletsOutput.value = analysis.tailoredBullets || '';
  renderKeywordChips(atsMatchedKeywords, analysis.matchedKeywords || [], 'Generate analysis to see matched skills.');
  renderKeywordChips(atsMissingKeywords, analysis.missingKeywords || [], 'Missing skills will appear here.');
  downloadAtsDocxBtn.disabled = !atsAnalysis;
}

function generateAtsTailor() {
  const resumeText = atsResumeInput.value.trim();
  const jobText = atsJobInput.value.trim();

  if (!resumeText || !jobText) {
    setInlineStatus(atsStatus, 'Paste both your resume text and the target job description first.', 'warning');
    return;
  }

  const jobKeywords = getRelevantAtsKeywords(jobText);
  const normalizedResume = normalizeKeywordText(resumeText);
  const matchedKeywords = jobKeywords.filter((keyword) => normalizedResume.includes(keyword));
  const missingKeywords = jobKeywords.filter((keyword) => !normalizedResume.includes(keyword));
  const scoreBase = jobKeywords.length ? Math.round((matchedKeywords.length / jobKeywords.length) * 100) : 0;

  atsAnalysis = {
    score: Math.min(98, Math.max(12, scoreBase)),
    matchedKeywords,
    missingKeywords,
    tailoredSummary: buildTailoredSummary(jobText, matchedKeywords, missingKeywords),
    tailoredBullets: buildTailoredBullets(matchedKeywords, missingKeywords),
    sourceResume: resumeText,
    sourceJobDescription: jobText
  };

  persistAtsDraft();
  renderAtsAnalysis();
  setInlineStatus(atsStatus, `ATS tailoring complete. Matched ${matchedKeywords.length} keywords and highlighted ${missingKeywords.length} gaps.`, 'success');
}

function clearAtsTailor() {
  atsResumeInput.value = '';
  atsJobInput.value = '';
  atsResumeFileInput.value = '';
  atsAnalysis = null;
  atsResumeFileName = '';
  atsResumeFileInfo.textContent = 'Upload `.txt`, `.docx`, or `.pdf`, or paste resume text manually below.';
  localStorage.removeItem(ATS_TAILOR_KEY);
  renderAtsAnalysis();
  setInlineStatus(atsStatus, 'ATS tailor inputs and results cleared.', 'info');
}

async function downloadAtsDocx() {
  if (!atsAnalysis) {
    setInlineStatus(atsStatus, 'Generate ATS analysis before downloading the DOCX file.', 'warning');
    return;
  }

  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          text: 'ATS Tailored Resume Draft',
          heading: HeadingLevel.TITLE
        }),
        new Paragraph({
          children: [new TextRun({ text: `Match Score: ${atsAnalysis.score}%`, bold: true })]
        }),
        new Paragraph({ text: 'Tailored Summary', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: atsAnalysis.tailoredSummary }),
        new Paragraph({ text: 'Tailored Experience Bullets', heading: HeadingLevel.HEADING_1 }),
        ...atsAnalysis.tailoredBullets.split('\n').filter(Boolean).map((line) => new Paragraph({ text: line.replace(/^•\s*/, ''), bullet: { level: 0 } })),
        new Paragraph({ text: 'Matched Keywords', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: atsAnalysis.matchedKeywords.join(', ') || 'None detected' }),
        new Paragraph({ text: 'Missing Keywords To Add Truthfully', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: atsAnalysis.missingKeywords.join(', ') || 'No obvious gaps detected' }),
        new Paragraph({ text: 'Original Resume Text', heading: HeadingLevel.HEADING_1 }),
        ...atsResumeInput.value.split('\n').filter((line) => line.trim()).map((line) => new Paragraph({ text: line.trim() }))
      ]
    }]
  });

  const blob = await Packer.toBlob(doc);
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = 'ats-tailored-resume.docx';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(downloadUrl);
  setInlineStatus(atsStatus, 'DOCX resume draft downloaded successfully.', 'success');
}

function restoreGmailAppPassword() {
  gmailAppUser = localStorage.getItem(GMAIL_APP_USER_KEY) || '';
  gmailAppPassword = localStorage.getItem(GMAIL_APP_PASSWORD_KEY) || '';
}

function getHasGmailAppPassword() {
  return Boolean(gmailAppUser && gmailAppPassword);
}

function openGmailAppPasswordModal() {
  document.querySelector('.modal-overlay')?.remove();

  const modal = document.createElement('div');
  modal.className = 'modal-overlay';
  modal.innerHTML = `
    <div class="modal-card">
      <div class="modal-header">
        <div>
          <h3>Use Your Gmail</h3>
          <p class="muted-text">Save a Gmail address and 16-character Google App Password on this browser only.</p>
        </div>
        <button type="button" class="icon-btn modal-close-btn" aria-label="Close"><i data-lucide="x"></i></button>
      </div>
      <div class="modal-body">
        <div class="form-group">
          <label>Gmail Address</label>
          <input type="email" id="gmailAppUserInput" class="input-field" placeholder="your-email@gmail.com" value="${gmailAppUser}" />
        </div>
        <div class="form-group">
          <label>Google App Password</label>
          <input type="password" id="gmailAppPasswordInput" class="input-field" placeholder="xxxx xxxx xxxx xxxx" />
          <div class="saved-draft">Use a 16-character Google App Password, not your normal Gmail password. This is stored only in this browser.</div>
        </div>
      </div>
      <div class="modal-footer">
        ${getHasGmailAppPassword() ? '<button type="button" id="clearGmailAppPasswordBtn" class="secondary-btn">Clear Saved Gmail</button>' : ''}
        <button type="button" id="saveGmailAppPasswordBtn" class="primary-btn">Save Gmail Sender</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);
  createIcons({ icons });

  modal.addEventListener('click', (event) => {
    if (event.target === modal) {
      modal.remove();
    }
  });

  modal.querySelector('.modal-close-btn')?.addEventListener('click', () => modal.remove());
  modal.querySelector('#saveGmailAppPasswordBtn')?.addEventListener('click', saveGmailAppPassword);
  modal.querySelector('#clearGmailAppPasswordBtn')?.addEventListener('click', clearGmailAppPassword);
}

function saveGmailAppPassword() {
  const userInput = document.getElementById('gmailAppUserInput')?.value.trim();
  const passwordInput = document.getElementById('gmailAppPasswordInput')?.value.replace(/\s/g, '');

  if (!userInput || !passwordInput) {
    setInlineStatus(manualStatus, 'Enter both Gmail address and App Password.', 'warning');
    return;
  }

  gmailAppUser = userInput;
  gmailAppPassword = passwordInput;
  localStorage.setItem(GMAIL_APP_USER_KEY, gmailAppUser);
  localStorage.setItem(GMAIL_APP_PASSWORD_KEY, gmailAppPassword);
  document.querySelector('.modal-overlay')?.remove();
  updateSenderUi();
  setInlineStatus(outreachStatus, `Gmail App Password saved for ${gmailAppUser}.`, 'success');
  setInlineStatus(manualStatus, `Gmail App Password saved for ${gmailAppUser}.`, 'success');
}

function clearGmailAppPassword() {
  gmailAppUser = '';
  gmailAppPassword = '';
  localStorage.removeItem(GMAIL_APP_USER_KEY);
  localStorage.removeItem(GMAIL_APP_PASSWORD_KEY);
  document.querySelector('.modal-overlay')?.remove();
  updateSenderUi();
  setInlineStatus(outreachStatus, 'Saved Gmail App Password cleared from this browser.', 'info');
  setInlineStatus(manualStatus, 'Saved Gmail App Password cleared from this browser.', 'info');
}

function getEffectiveMailCapability() {
  if (getHasGmailAppPassword()) {
    return {
      canSendMail: true,
      mode: 'gmail-app-password',
      senderEmail: gmailAppUser,
      requiresGoogleAuth: false
    };
  }

  return {
    canSendMail: Boolean(currentUser?.canSendMail || mailCapability.canSendMail),
    mode: currentUser?.mailMode || mailCapability.mode,
    senderEmail: currentUser?.senderEmail || mailCapability.senderEmail || '',
    requiresGoogleAuth: currentUser?.requiresGoogleAuth ?? mailCapability.requiresGoogleAuth
  };
}

function getMailBlockedMessage() {
  const capability = getEffectiveMailCapability();

  if (capability.requiresGoogleAuth) {
    return currentUser?.email
      ? 'Connect Gmail sender or use Gmail App Password before sending mail.'
      : 'Mail sending needs Gmail App Password, Google connect, or a configured server sender.';
  }

  return 'Mail sending is not configured on the server yet.';
}

function updateSenderUi() {
  const capability = getEffectiveMailCapability();
  const hasDirectSender = capability.mode === 'server-smtp' || capability.mode === 'resend-api' || capability.mode === 'gmail-app-password';
  const senderLabel = hasDirectSender
    ? `${capability.mode === 'resend-api' ? 'Resend sender' : capability.mode === 'gmail-app-password' ? 'Your Gmail' : 'Server sender'} ${capability.senderEmail || 'configured'}`
    : currentUser?.email
      ? `${currentUser.displayName} (${currentUser.email})`
      : 'Login with Google to save your tracker. Gmail connect is optional.';

  const permissionHint = capability.canSendMail
    ? hasDirectSender
      ? capability.mode === 'gmail-app-password'
        ? 'Your Gmail App Password is ready for sending from this browser.'
        : 'Direct mail is ready. Gmail connect is not needed.'
      : 'Mail access is ready.'
    : currentUser?.email
      ? 'Connect Gmail sender only if you want to send mail.'
      : capability.requiresGoogleAuth
        ? 'Mail sending is disabled until you log in or connect Gmail.'
        : 'Mail sending is disabled until a server sender is configured.';

  const senderActionHref = currentUser?.email ? '/auth/google-gmail' : '/auth/google';
  const senderActionLabel = capability.canSendMail && !hasDirectSender
    ? 'Reconnect Gmail Sender'
    : currentUser?.email
      ? 'Connect Gmail Sender'
      : 'Login with Google';

  bulkSenderStatus.textContent = `${senderLabel} ${permissionHint}`;
  manualSenderStatus.textContent = `${senderLabel} ${permissionHint}`;
  bulkLoginHint.classList.toggle('hidden', hasDirectSender || !capability.requiresGoogleAuth);
  manualLoginHint.classList.toggle('hidden', hasDirectSender || !capability.requiresGoogleAuth);
  bulkLoginHint.textContent = senderActionLabel;
  bulkLoginHint.href = senderActionHref;
  manualLoginHint.textContent = senderActionLabel;
  manualLoginHint.href = senderActionHref;
  fireAllMailsBtn.disabled = false;
  sendManualBtn.disabled = false;
  fireAllMailsBtn.title = capability.canSendMail ? 'Send bulk outreach emails' : getMailBlockedMessage();
  sendManualBtn.title = capability.canSendMail ? 'Send manual email' : getMailBlockedMessage();
  fireAllMailsBtn.setAttribute('aria-disabled', capability.canSendMail ? 'false' : 'true');
  sendManualBtn.setAttribute('aria-disabled', capability.canSendMail ? 'false' : 'true');

  document.getElementById('userProfile').classList.toggle('hidden', !currentUser?.email);
  headerLoginBtn.classList.toggle('hidden', Boolean(currentUser?.email));
  createIcons({ icons });
}

function openDraftDatabase() {
  if (openDraftDatabase.promise) return openDraftDatabase.promise;

  openDraftDatabase.promise = new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DRAFT_DB_NAME, 1);

    request.onupgradeneeded = () => {
      request.result.createObjectStore(DRAFT_FILE_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  return openDraftDatabase.promise;
}

async function saveStoredFiles(key, files) {
  const db = await openDraftDatabase();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(DRAFT_FILE_STORE, 'readwrite');
    tx.objectStore(DRAFT_FILE_STORE).put(files, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadStoredFiles(key) {
  const db = await openDraftDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DRAFT_FILE_STORE, 'readonly');
    const request = tx.objectStore(DRAFT_FILE_STORE).get(key);
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

async function initializeStoredFiles() {
  try {
    savedManualAttachments = normalizeStoredFiles(await loadStoredFiles(MANUAL_ATTACHMENT_KEY));
    savedResumeFiles = normalizeStoredFiles(await loadStoredFiles(BULK_RESUME_KEY));
  } catch (error) {
    console.warn('Could not restore saved files.', error);
    savedManualAttachments = [];
    savedResumeFiles = [];
  }

  renderSavedFileInfo();
}

function normalizeStoredFiles(files) {
  return (Array.isArray(files) ? files : []).map((file, index) => {
    if (file instanceof File) return file;
    return new File([file], file.name || `attachment-${index + 1}`, {
      type: file.type || 'application/octet-stream',
      lastModified: file.lastModified || Date.now()
    });
  });
}

function renderSavedFileInfo() {
  savedAttachmentInfo.textContent = savedManualAttachments.length
    ? `Saved attachments: ${savedManualAttachments.map((file) => file.name).join(', ')}`
    : 'No attachments saved yet.';

  savedResumeInfo.textContent = savedResumeFiles[0]
    ? `Saved resume: ${savedResumeFiles[0].name}`
    : 'No resume saved yet.';
}

function refreshCounts() {
  const pendingCount = jobs.filter((job) => !job.applied).length;
  heroPendingCount.textContent = pendingCount;
  heroAppliedCount.textContent = appliedRecords.length;
}

function showBanner(message, tone = 'info') {
  appStatus.textContent = message;
  appStatus.className = `status-banner ${tone}`;
  appStatus.classList.remove('hidden');

  window.clearTimeout(showBanner.timer);
  showBanner.timer = window.setTimeout(() => {
    appStatus.classList.add('hidden');
  }, 3200);
}

function setInlineStatus(element, message, tone = 'info') {
  element.textContent = message;
  element.className = `inline-status ${tone}`;
  element.classList.remove('hidden');
}

function handleFileUpload() {
  const file = fileInput.files[0];
  if (!file) return;

  const extension = file.name.includes('.') ? file.name.split('.').pop().toLowerCase() : '';
  if (extension === 'xlsx' || extension === 'xls') {
    parseExcel(file);
  } else {
    parseCSV(file);
  }
}

function parseCSV(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => processExtractedData(results.data),
    error: () => showBanner('Could not parse the CSV file.', 'error')
  });
}

function parseExcel(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = xlsx.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const json = xlsx.utils.sheet_to_json(workbook.Sheets[firstSheet]);
      processExtractedData(json);
    } catch {
      showBanner('Could not parse the Excel file.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function processExtractedData(data) {
  if (!data || data.length === 0) {
    showBanner('The uploaded file is empty.', 'warning');
    return;
  }

  jobs = data.map(normalizeJobRow);
  currentPage = 1;
  selectedJob = jobs.find((job) => !job.applied) || jobs[0] || null;

  uploadSection.classList.add('hidden');
  dashboardSection.classList.remove('hidden');
  renderJobs();
  renderJobDetails();
  refreshCounts();
  showBanner(`Imported ${jobs.length} jobs. Pending roles are ready to review.`, 'success');
}

function normalizeJobRow(row) {
  const getValue = (keys) => {
    for (const key of keys) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') return row[key];
    }

    const fuzzyKey = Object.keys(row).find((key) => {
      const normalized = key.toLowerCase().replace(/[^a-z0-9]/g, '');
      return keys.some((candidate) => normalized.includes(candidate.toLowerCase().replace(/[^a-z0-9]/g, '')));
    });

    return fuzzyKey ? row[fuzzyKey] : null;
  };

  const id = getValue(['id', 'jobId']) || `${getValue(['company', 'employer']) || 'job'}-${getValue(['title', 'jobtitle', 'role']) || 'role'}`.replace(/\s+/g, '-').toLowerCase();
  const title = getValue(['title', 'jobtitle', 'role', 'position', 'designation']) || 'Unknown Title';
  const company = getValue(['company/name', 'company', 'employer', 'org', 'organization']) || 'Unknown Company';
  const description = getValue(['descriptionText', 'descriptionHtml', 'description', 'jd', 'detail', 'summary']) || 'No description provided.';
  const logo = getValue(['company/logos/0/url', 'company/logo', 'logoUrl']);
  const linkedinUrl = getValue(['linkedinUrl', 'linkedin_url']);
  const easyApplyUrl = getValue(['applyMethod/easyApplyUrl', 'easyApplyUrl', 'query/easyApply']);
  const companyApplyUrl = getValue(['applyMethod/companyApplyUrl', 'company_url']);

  const generalLinkRaw = getValue(['applylink', 'link', 'url', 'applyurl', 'joblink', 'joburl', 'naukriurl', 'application_url']);
  let generalLink = (generalLinkRaw && (generalLinkRaw.toLowerCase().includes('ambitionbox.com') || generalLinkRaw.toLowerCase().includes('/reviews/'))) ? null : generalLinkRaw;

  if (!generalLink && description) {
    const urlRegex = /https?:\/\/[^\s"'<>()[\]]+/g;
    const links = description.match(urlRegex) || [];
    generalLink = links.find((link) => !link.toLowerCase().includes('ambitionbox.com') && !link.toLowerCase().includes('/reviews/')) || '';
  }

  const location = getValue(['location/linkedinText', 'location/parsed/text', 'location/parsed/city', 'company/locations/0/city', 'location', 'city']) || '';
  const workplaceType = getValue(['workplaceType', 'workRemoteAllowed', 'remote']) || '';
  const employmentType = getValue(['employmentType', 'query/employmentType/0', 'type']) || '';
  const experienceLevel = getValue(['experienceLevel', 'query/experienceLevel/0', 'exp', 'experience']) || '';
  const hrProfile = getValue(['hiringTeam/0/name', 'hr', 'contact', 'hrName', 'recruiter']) || '';
  const hrEmail = getValue(['email', 'hrEmail', 'recruiterEmail', 'mail', 'contactEmail']) || '';
  const hrPhone = getValue(['phone', 'contactNumber', 'mobile', 'hrPhone', 'recruiterPhone', 'contactNo']) || '';

  const salaryMin = getValue(['salary/min']);
  const salaryMax = getValue(['salary/max']);
  const salaryCurrency = getValue(['salary/currency']) || '';
  let salaryStr = getValue(['salary/text', 'salary', 'ctc']) || '';
  if (!salaryStr && salaryMin && salaryMax) {
    salaryStr = `${salaryCurrency}${salaryMin} - ${salaryCurrency}${salaryMax}`;
  }

  const appliedMatch = appliedRecords.find((record) => record.id === id);

  return {
    id,
    title,
    company,
    description,
    logo,
    linkedinUrl,
    easyApplyUrl,
    companyApplyUrl,
    generalLink,
    location,
    workplaceType,
    employmentType,
    experienceLevel,
    hrProfile,
    hrEmail,
    hrPhone,
    salaryStr,
    applied: Boolean(appliedMatch),
    appliedStatus: appliedMatch?.status || ''
  };
}

function renderJobs() {
  jobList.innerHTML = '';

  const pendingJobs = jobs.filter((job) => !job.applied);
  jobCount.textContent = pendingJobs.length;
  heroPendingCount.textContent = pendingJobs.length;

  if (pendingJobs.length === 0) {
    jobList.innerHTML = '<div class="empty-state">No pending jobs right now. Marked items move into Applied Tracker automatically.</div>';
    paginationControls.classList.add('hidden');
    renderJobDetails();
    return;
  }

  const totalPages = Math.ceil(pendingJobs.length / pageSize);
  currentPage = Math.min(currentPage, totalPages);
  const visibleJobs = pendingJobs.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  paginationControls.classList.toggle('hidden', totalPages <= 1);
  document.getElementById('pageInfo').textContent = `${currentPage}/${totalPages}`;
  document.getElementById('prevPageBtn').disabled = currentPage === 1;
  document.getElementById('nextPageBtn').disabled = currentPage === totalPages;

  visibleJobs.forEach((job) => {
    const card = document.createElement('button');
    card.className = `compact-job-card ${selectedJob?.id === job.id ? 'active' : ''}`;
    card.type = 'button';
    card.innerHTML = `
      <div class="company-brand">
        ${job.logo ? `<img src="${job.logo}" alt="" class="company-logo" onerror="this.style.display='none'" />` : '<i data-lucide="building-2"></i>'}
        <div class="card-copy">
          <h3 class="job-title">${truncate(job.title, 48)}</h3>
          <p class="job-company">${truncate(job.company, 34)}</p>
        </div>
      </div>
      <div class="mini-meta">
        ${job.location ? `<span>${job.location}</span>` : '<span>Location pending</span>'}
        ${job.salaryStr ? `<strong>${truncate(job.salaryStr, 22)}</strong>` : ''}
      </div>
    `;
    card.addEventListener('click', () => {
      selectedJob = job;
      renderJobs();
      renderJobDetails();
    });
    jobList.appendChild(card);
  });

  createIcons({ icons });
}

function renderJobDetails() {
  if (!selectedJob || selectedJob.applied) {
    jobDetailsEmptyState.classList.remove('hidden');
    jobDetailsContent.classList.add('hidden');
    return;
  }

  jobDetailsEmptyState.classList.add('hidden');
  jobDetailsContent.classList.remove('hidden');

  const badges = [
    renderBadge('map-pin', selectedJob.location),
    renderBadge('monitor', selectedJob.workplaceType),
    renderBadge('briefcase', selectedJob.employmentType),
    renderBadge('trending-up', selectedJob.experienceLevel),
    renderBadge('banknote', selectedJob.salaryStr, true)
  ].filter(Boolean).join('');

  const actions = [
    selectedJob.easyApplyUrl ? renderLinkButton(selectedJob.easyApplyUrl, 'Easy Apply', 'zap', 'primary-btn') : '',
    selectedJob.companyApplyUrl ? renderLinkButton(selectedJob.companyApplyUrl, 'Company Site', 'external-link', 'secondary-btn') : '',
    selectedJob.linkedinUrl ? renderLinkButton(selectedJob.linkedinUrl, 'LinkedIn', 'linkedin', 'secondary-btn') : '',
    !selectedJob.easyApplyUrl && !selectedJob.companyApplyUrl && !selectedJob.linkedinUrl && selectedJob.generalLink
      ? renderLinkButton(selectedJob.generalLink, 'Apply Now', 'external-link', 'primary-btn')
      : '',
    !selectedJob.easyApplyUrl && !selectedJob.companyApplyUrl && !selectedJob.linkedinUrl && !selectedJob.generalLink
      ? renderLinkButton(`https://www.google.com/search?q=${encodeURIComponent(`${selectedJob.company} ${selectedJob.title} career apply`)}`, 'Search Role', 'search', 'secondary-btn')
      : ''
  ].filter(Boolean).join('');

  const recruiter = [
    renderBadge('user', selectedJob.hrProfile ? `HR: ${selectedJob.hrProfile}` : ''),
    selectedJob.hrEmail ? `<a href="mailto:${selectedJob.hrEmail}" class="badge highlight-badge clickable-badge"><i data-lucide="mail"></i><span>${selectedJob.hrEmail}</span></a>` : '',
    selectedJob.hrPhone ? `<a href="tel:${selectedJob.hrPhone}" class="badge clickable-badge"><i data-lucide="phone"></i><span>${selectedJob.hrPhone}</span></a>` : ''
  ].filter(Boolean).join('');

  jobDetailsContent.innerHTML = `
    <div class="details-header">
      <div class="details-company-info">
        ${selectedJob.logo ? `<img src="${selectedJob.logo}" alt="Logo" class="details-logo" onerror="this.style.display='none'" />` : '<div class="details-logo fallback-logo"><i data-lucide="building-2"></i></div>'}
        <div class="details-title">
          <p class="eyebrow">Pending role</p>
          <h1>${selectedJob.title}</h1>
          <p class="details-company">${selectedJob.company}</p>
        </div>
      </div>
      <button class="primary-btn" id="detailApplyBtn"><i data-lucide="check-check"></i> Mark Applied</button>
    </div>

    <div class="details-badges">${badges || '<span class="muted-text">No extra metadata in this export.</span>'}</div>

    <div class="details-section">
      <h3>Application Actions</h3>
      <div class="details-actions">${actions}</div>
    </div>

    <div class="details-section">
      <h3>Description</h3>
      <div class="details-body">${selectedJob.description}</div>
    </div>

    ${recruiter ? `<div class="details-section"><h3>Recruiter Details</h3><div class="details-badges">${recruiter}</div></div>` : ''}
  `;

  document.getElementById('detailApplyBtn').addEventListener('click', () => markAsApplied(selectedJob.id));
  createIcons({ icons });
}

function renderAppliedTracker() {
  const interviews = appliedRecords.filter((record) => record.status === 'Interviewing').length;
  const offers = appliedRecords.filter((record) => record.status === 'Offer').length;
  const latestUpdate = appliedRecords[0]?.lastUpdatedAt || appliedRecords[0]?.appliedAt;

  appliedCount.textContent = appliedRecords.length;
  appliedTotalMetric.textContent = appliedRecords.length;
  interviewMetric.textContent = interviews;
  offerMetric.textContent = offers;
  heroAppliedCount.textContent = appliedRecords.length;
  appliedUpdatedAt.textContent = latestUpdate ? `Updated ${formatDateTime(latestUpdate)}` : 'No updates yet';

  appliedList.innerHTML = '';

  if (appliedRecords.length === 0) {
    appliedList.innerHTML = '<div class="empty-state">No applied roles tracked yet. Mark a pending role as applied to move it here.</div>';
    appliedDetailsEmpty.classList.remove('hidden');
    appliedDetailsContent.classList.add('hidden');
    refreshCounts();
    return;
  }

  if (!selectedAppliedJob || !appliedRecords.some((record) => record.id === selectedAppliedJob.id)) {
    selectedAppliedJob = appliedRecords[0];
  }

  appliedRecords.forEach((record) => {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = `compact-job-card applied-card ${selectedAppliedJob?.id === record.id ? 'active' : ''}`;
    card.innerHTML = `
      <div class="company-brand">
        ${record.logo ? `<img src="${record.logo}" alt="" class="company-logo" onerror="this.style.display='none'" />` : '<i data-lucide="folder-check"></i>'}
        <div class="card-copy">
          <h3 class="job-title">${truncate(record.title, 42)}</h3>
          <p class="job-company">${truncate(record.company, 28)}</p>
        </div>
      </div>
      <div class="mini-meta">
        <span>${record.status}</span>
        <strong>${formatShortDate(record.appliedAt)}</strong>
      </div>
    `;
    card.addEventListener('click', () => {
      selectedAppliedJob = record;
      renderAppliedTracker();
    });
    appliedList.appendChild(card);
  });

  renderAppliedDetails();
  refreshCounts();
  createIcons({ icons });
}

function renderAppliedDetails() {
  if (!selectedAppliedJob) {
    appliedDetailsEmpty.classList.remove('hidden');
    appliedDetailsContent.classList.add('hidden');
    return;
  }

  appliedDetailsEmpty.classList.add('hidden');
  appliedDetailsContent.classList.remove('hidden');

  const links = [
    selectedAppliedJob.easyApplyUrl ? renderLinkButton(selectedAppliedJob.easyApplyUrl, 'Easy Apply', 'zap', 'primary-btn') : '',
    selectedAppliedJob.companyApplyUrl ? renderLinkButton(selectedAppliedJob.companyApplyUrl, 'Company Site', 'external-link', 'secondary-btn') : '',
    selectedAppliedJob.linkedinUrl ? renderLinkButton(selectedAppliedJob.linkedinUrl, 'LinkedIn', 'linkedin', 'secondary-btn') : '',
    selectedAppliedJob.generalLink ? renderLinkButton(selectedAppliedJob.generalLink, 'Original Link', 'link', 'secondary-btn') : ''
  ].filter(Boolean).join('');

  appliedDetailsContent.innerHTML = `
    <div class="details-header">
      <div class="details-company-info">
        ${selectedAppliedJob.logo ? `<img src="${selectedAppliedJob.logo}" alt="Logo" class="details-logo" onerror="this.style.display='none'" />` : '<div class="details-logo fallback-logo"><i data-lucide="folder-check"></i></div>'}
        <div class="details-title">
          <p class="eyebrow">Applied role</p>
          <h1>${selectedAppliedJob.title}</h1>
          <p class="details-company">${selectedAppliedJob.company}</p>
        </div>
      </div>
      <span class="pill">${selectedAppliedJob.status}</span>
    </div>

    <div class="details-badges">
      ${renderBadge('calendar-days', `Applied ${formatDateTime(selectedAppliedJob.appliedAt)}`)}
      ${selectedAppliedJob.location ? renderBadge('map-pin', selectedAppliedJob.location) : ''}
      ${selectedAppliedJob.salaryStr ? renderBadge('banknote', selectedAppliedJob.salaryStr, true) : ''}
    </div>

    <div class="details-section">
      <h3>Pipeline Update</h3>
      <div class="update-grid">
        <div class="form-group">
          <label>Status</label>
          <select id="appliedStatusSelect" class="input-field">
            ${APPLIED_STATUSES.map((status) => `<option value="${status}" ${selectedAppliedJob.status === status ? 'selected' : ''}>${status}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label>Notes</label>
          <textarea id="appliedNotes" class="input-field textarea" rows="5" placeholder="Add recruiter updates, next steps, interview dates...">${selectedAppliedJob.notes || ''}</textarea>
        </div>
      </div>
      <div class="cta-row">
        <button class="primary-btn" id="saveAppliedUpdateBtn"><i data-lucide="save"></i> Save Update</button>
        <span class="muted-text">Last updated ${formatDateTime(selectedAppliedJob.lastUpdatedAt || selectedAppliedJob.appliedAt)}</span>
      </div>
    </div>

    ${links ? `<div class="details-section"><h3>Quick Links</h3><div class="details-actions">${links}</div></div>` : ''}
    <div class="details-section">
      <h3>Stored Description</h3>
      <div class="details-body">${selectedAppliedJob.description || 'No description stored.'}</div>
    </div>
  `;

  document.getElementById('saveAppliedUpdateBtn').addEventListener('click', async () => {
    const status = document.getElementById('appliedStatusSelect').value;
    const notes = document.getElementById('appliedNotes').value.trim();
    await updateAppliedRecord(selectedAppliedJob.id, { status, notes });
  });

  createIcons({ icons });
}

async function markAsApplied(id) {
  const job = jobs.find((item) => item.id === id);
  if (!job) return;

  const record = {
    ...job,
    status: 'Applied',
    notes: '',
    appliedAt: new Date().toISOString()
  };

  job.applied = true;
  appliedRecords = upsertAppliedRecord(record);
  persistAppliedCache();

  try {
    const response = await apiFetch('/api/applied-jobs', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ job: record })
    });
    if (response.ok) {
      const data = await safeJson(response);
      appliedRecords = data.records;
      persistAppliedCache();
    }
  } catch (error) {
    console.warn('Applied job sync failed.', error);
  }

  const pendingJobs = jobs.filter((item) => !item.applied);
  selectedJob = pendingJobs[0] || null;
  selectedAppliedJob = appliedRecords[0] || selectedAppliedJob;

  renderJobs();
  renderJobDetails();
  renderAppliedTracker();
  showBanner(`${job.company} moved to Applied Tracker.`, 'success');
}

async function updateAppliedRecord(id, updates) {
  appliedRecords = appliedRecords.map((record) => (
    record.id === id
      ? { ...record, ...updates, lastUpdatedAt: new Date().toISOString() }
      : record
  ));

  selectedAppliedJob = appliedRecords.find((record) => record.id === id) || selectedAppliedJob;
  persistAppliedCache();
  renderAppliedTracker();

  try {
    const response = await apiFetch(`/api/applied-jobs/${id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(updates)
    });

    if (response.ok) {
      const data = await safeJson(response);
      appliedRecords = appliedRecords.map((record) => record.id === id ? data.record : record);
      selectedAppliedJob = data.record;
      persistAppliedCache();
      renderAppliedTracker();
    }
  } catch (error) {
    console.warn('Applied update sync failed.', error);
  }

  showBanner('Applied tracker updated.', 'success');
}

function upsertAppliedRecord(record) {
  const existingIndex = appliedRecords.findIndex((item) => item.id === record.id);
  const normalized = {
    ...record,
    lastUpdatedAt: new Date().toISOString()
  };

  if (existingIndex >= 0) {
    const clone = [...appliedRecords];
    clone[existingIndex] = { ...clone[existingIndex], ...normalized };
    return clone;
  }

  return [normalized, ...appliedRecords];
}

function handleOutreachUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (loadEvent) => parseOutreachHTML(loadEvent.target.result);
  reader.readAsText(file);
}

function parseOutreachHTML(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, 'text/html');

  outreachContacts = extractContactsFromDocument(doc);
  outreachVariants = extractTemplatesFromDocument(doc);

  if (outreachVariants.length === 0) {
    outreachVariants = [{
      id: 'default-template',
      subject: 'Application for relevant opportunity',
      bodyTemplate: 'Hi [Recruiter Name],\n\nI am interested in opportunities at [Company]. Please find my resume attached.\n\nBest regards,'
    }];
  }

  applyBulkTemplate(outreachVariants[0]);

  if (outreachContacts.length === 0) {
    setInlineStatus(outreachStatus, 'No recruiter emails were detected in that HTML file.', 'warning');
    return;
  }

  outreachEmptyState.classList.add('hidden');
  outreachList.classList.remove('hidden');
  document.getElementById('automationConfig').classList.remove('hidden');
  renderOutreachContacts();
  setInlineStatus(outreachStatus, `Loaded ${outreachContacts.length} contacts and ${outreachVariants.length} email template(s).`, 'success');
}

function extractContactsFromDocument(doc) {
  const contactsMap = new Map();

  doc.querySelectorAll('.cc').forEach((card) => {
    const email = card.querySelector('.cc-email a[href^="mailto:"]')?.getAttribute('href')?.replace(/^mailto:/i, '').trim();
    if (!email) return;

    contactsMap.set(email.toLowerCase(), {
      email,
      name: card.querySelector('.cc-name')?.textContent?.trim() || 'Recruiter',
      company: card.querySelector('.cc-company')?.textContent?.trim() || ''
    });
  });

  doc.querySelectorAll('a[href^="mailto:"]').forEach((anchor) => {
    const email = anchor.getAttribute('href').replace(/^mailto:/i, '').trim();
    if (!email) return;

    if (contactsMap.has(email.toLowerCase())) return;

    const card = anchor.closest('.cc, .contact-card, .today-card, .outreach-item, tr, li, div');
    const surroundingText = card?.textContent?.replace(/\s+/g, ' ').trim() || anchor.textContent.trim();
    const companyMatch = surroundingText.match(/at\s+([A-Za-z0-9 &.-]+)/i);

    contactsMap.set(email.toLowerCase(), {
      email,
      name: anchor.textContent.trim() && !anchor.textContent.includes('@') ? anchor.textContent.trim() : surroundingText.split('|')[0].trim() || 'Recruiter',
      company: companyMatch?.[1]?.trim() || ''
    });
  });

  doc.querySelectorAll('.contact-card').forEach((card) => {
    const email = card.querySelector('.contact-email a')?.textContent?.trim();
    if (!email) return;

    contactsMap.set(email.toLowerCase(), {
      email,
      name: card.querySelector('.contact-name')?.textContent?.trim() || 'Recruiter',
      company: card.querySelector('.contact-company')?.textContent?.trim() || ''
    });
  });

  return Array.from(contactsMap.values());
}

function extractTemplatesFromDocument(doc) {
  const templates = [];

  doc.querySelectorAll('.oc').forEach((card, index) => {
    const subject = card.querySelector('.oc-subj')?.textContent?.replace(/^Subject:\s*/i, '').trim()
      || card.querySelector('.oc-varname')?.textContent?.trim()
      || `Outreach Template ${index + 1}`;
    const body = card.querySelector('.oc-text')?.textContent?.trim() || '';

    if (body) {
      templates.push({ id: `template-oc-${index + 1}`, subject, bodyTemplate: body });
    }
  });

  doc.querySelectorAll('.outreach-card, [data-outreach-template]').forEach((card, index) => {
    const subject = card.querySelector('.outreach-subject')?.textContent?.replace('Subject:', '').trim() || `Outreach Template ${index + 1}`;
    const body = card.querySelector('.outreach-text')?.textContent?.trim() || card.getAttribute('data-outreach-template') || '';
    if (body) {
      templates.push({ id: `template-${index + 1}`, subject, bodyTemplate: body });
    }
  });

  return templates;
}

function renderOutreachContacts() {
  outreachList.innerHTML = '';

  outreachContacts.forEach((contact) => {
    const row = document.createElement('div');
    row.className = 'outreach-item';
    row.innerHTML = `
      <div class="outreach-person">
        <h4>${contact.name || 'Recruiter'}</h4>
        <p>${contact.company || 'Company not detected'}</p>
        <p class="accent-copy">${contact.email}</p>
      </div>
      <button class="secondary-btn btn-sm copy-email-btn" type="button" data-email="${contact.email}">
        <i data-lucide="copy"></i> Copy
      </button>
    `;
    outreachList.appendChild(row);
  });

  outreachList.querySelectorAll('.copy-email-btn').forEach((button) => {
    button.addEventListener('click', () => {
      navigator.clipboard.writeText(button.dataset.email);
      setInlineStatus(outreachStatus, `${button.dataset.email} copied to clipboard.`, 'info');
    });
  });

  createIcons({ icons });
  renderBulkPreview();
}

function applyBulkTemplate(template) {
  bulkSubjectInput.value = template?.subject || getDefaultBulkSubject();
  bulkBodyInput.value = template?.bodyTemplate || getDefaultBulkBody();
  persistOutreachDraft();
  renderBulkPreview();
}

function renderBulkPreview() {
  if (!bulkPreview) return;

  const subjectTemplate = bulkSubjectInput.value.trim() || getDefaultBulkSubject();
  const bodyTemplate = bulkBodyInput.value || getDefaultBulkBody();

  if (!outreachContacts.length) {
    bulkPreview.textContent = 'Load recruiter HTML to preview the first personalized email.';
    return;
  }

  const contact = outreachContacts[0];
  const previewSubject = personalizeTemplate(subjectTemplate, contact.email);
  const previewBody = personalizeTemplate(bodyTemplate, contact.email);
  bulkPreview.textContent = `Preview for ${contact.name || contact.email}\nSubject: ${previewSubject}\n\n${previewBody}`;
}

async function sendBulkOutreach() {
  const resume = savedResumeFiles[0] || resumeUpload.files[0];
  const capability = getEffectiveMailCapability();
  const useGmailAppPassword = capability.mode === 'gmail-app-password' && getHasGmailAppPassword();

  if (!outreachContacts.length) {
    setInlineStatus(outreachStatus, 'Load recruiter HTML first so there are contacts to send to.', 'warning');
    return;
  }

  if (!capability.canSendMail) {
    setInlineStatus(outreachStatus, getMailBlockedMessage(), 'warning');
    return;
  }

  if (!resume) {
    setInlineStatus(outreachStatus, 'Add a resume PDF first. It will stay saved until you change it.', 'warning');
    return;
  }

  const subjectTemplate = bulkSubjectInput.value.trim() || outreachVariants[0]?.subject || getDefaultBulkSubject();
  const bodyTemplate = bulkBodyInput.value || outreachVariants[0]?.bodyTemplate || getDefaultBulkBody();

  if (!subjectTemplate || !bodyTemplate.trim()) {
    setInlineStatus(outreachStatus, 'Add both bulk subject and bulk body before sending.', 'warning');
    return;
  }

  const contacts = outreachContacts.map((contact) => ({
    to: contact.email,
    subject: personalizeTemplate(subjectTemplate, contact.email),
    body: personalizeTemplate(bodyTemplate, contact.email)
  }));

  const payload = new FormData();
  payload.append('contacts', JSON.stringify(contacts));
  payload.append('resume', resume, resume.name);
  if (useGmailAppPassword) {
    payload.append('gmailUser', gmailAppUser);
    payload.append('gmailAppPassword', gmailAppPassword);
  }

  const button = document.getElementById('fireAllMailsBtn');
  button.disabled = true;
  button.textContent = 'Sending...';
  setInlineStatus(outreachStatus, `Sending ${contacts.length} outreach emails...`, 'info');

  try {
    const response = await apiFetch('/api/send-cold-emails', {
      method: 'POST',
      body: payload
    });
    const data = await safeJson(response);

    if (!data.success) throw new Error(data.error || 'Bulk send failed.');

    const successCount = data.results.filter((item) => item.status === 'Sent').length;
    const failureCount = data.results.length - successCount;
    const senderLabel = useGmailAppPassword ? gmailAppUser : data.senderEmail;
    setInlineStatus(outreachStatus, `Bulk outreach finished from ${senderLabel}. Sent: ${successCount}, failed: ${failureCount}.`, failureCount ? 'warning' : 'success');
  } catch (error) {
    setInlineStatus(outreachStatus, error.message || 'Bulk outreach failed.', 'error');
  } finally {
    button.disabled = false;
    button.innerHTML = '<i data-lucide="zap"></i> Send Bulk Outreach';
    createIcons({ icons });
  }
}

async function sendManualEmail() {
  const to = manualToEmailInput.value.trim();
  const subject = personalizeTemplate(manualSubjectInput.value.trim(), to);
  const body = personalizeTemplate(manualBodyInput.value.trim(), to);
  const files = savedManualAttachments.length ? savedManualAttachments : Array.from(manualAttachmentsInput.files || []);
  const capability = getEffectiveMailCapability();
  const useGmailAppPassword = capability.mode === 'gmail-app-password' && getHasGmailAppPassword();

  if (!capability.canSendMail) {
    setInlineStatus(manualStatus, getMailBlockedMessage(), 'warning');
    return;
  }

  if (!to || !subject || !body) {
    setInlineStatus(manualStatus, 'Fill receiver email, subject, and body before sending.', 'warning');
    return;
  }

  const payload = new FormData();
  payload.append('to', to);
  payload.append('subject', subject);
  payload.append('body', body);
  files.forEach((file) => payload.append('attachments', file, file.name));
  if (useGmailAppPassword) {
    payload.append('gmailUser', gmailAppUser);
    payload.append('gmailAppPassword', gmailAppPassword);
  }

  const button = document.getElementById('sendManualBtn');
  button.disabled = true;
  button.textContent = 'Sending...';

  try {
    const response = await apiFetch('/api/send-single-email', {
      method: 'POST',
      body: payload
    });
    const data = await safeJson(response);
    if (!data.success) throw new Error(data.error || 'Manual send failed.');
    const senderLabel = useGmailAppPassword ? gmailAppUser : data.senderEmail;
    setInlineStatus(manualStatus, `Email sent successfully from ${senderLabel} to ${to}.`, 'success');
  } catch (error) {
    setInlineStatus(manualStatus, error.message || 'Manual email failed.', 'error');
  } finally {
    button.disabled = false;
    button.innerHTML = '<i data-lucide="send"></i> Send Email';
    createIcons({ icons });
  }
}

async function scrapeEmailsFromUrl() {
  const url = document.getElementById('scrapeUrl').value.trim();
  if (!url) {
    setInlineStatus(scrapeStatus, 'Enter a URL first.', 'warning');
    return;
  }

  const button = document.getElementById('scrapeBtn');
  button.disabled = true;
  button.textContent = 'Scraping...';

  try {
    const response = await apiFetch('/api/scrape-emails', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url })
    });
    const data = await safeJson(response);
    if (!data.success) throw new Error(data.error || 'Scrape failed.');

    const list = document.getElementById('scrapedEmailsList');
    list.innerHTML = '';

    if (!data.emails.length) {
      setInlineStatus(scrapeStatus, 'No email IDs were found on that page.', 'warning');
      document.getElementById('scrapeResults').classList.add('hidden');
      return;
    }

    data.emails.forEach((email) => {
      const item = document.createElement('li');
      item.innerHTML = `<span>${email}</span><button class="secondary-btn btn-sm" type="button">Use</button>`;
      item.querySelector('button').addEventListener('click', () => {
        manualToEmailInput.value = email;
        persistOutreachDraft();
        setInlineStatus(manualStatus, `${email} copied into the manual receiver field.`, 'info');
      });
      list.appendChild(item);
    });

    document.getElementById('scrapeResults').classList.remove('hidden');
    setInlineStatus(scrapeStatus, `Found ${data.emails.length} email ID(s).`, 'success');
  } catch (error) {
    setInlineStatus(scrapeStatus, error.message || 'Scraper failed.', 'error');
  } finally {
    button.disabled = false;
    button.innerHTML = '<i data-lucide="search"></i> Scrape';
    createIcons({ icons });
  }
}

async function scanWalkins() {
  const role = document.getElementById('walkInRole').value.trim() || 'Java Backend';
  const location = document.getElementById('walkInLocation').value.trim() || 'Bangalore';

  walkInList.innerHTML = '<div class="empty-state"><i data-lucide="loader-circle"></i> Scanning live job sources...</div>';
  setInlineStatus(walkInMeta, `Searching live links for ${role} in ${location}...`, 'info');
  createIcons({ icons });

  try {
    const response = await apiFetch(`/api/walkins?role=${encodeURIComponent(role)}&location=${encodeURIComponent(location)}&ts=${Date.now()}`, {
      cache: 'no-store'
    });
    const data = await safeJson(response);

    if (!response.ok) throw new Error(data.error || 'Walk-in scan failed.');

    walkInList.innerHTML = '';
    const providerLabel = data.provider === 'serpapi-google-jobs' ? 'SerpApi / Google Jobs' : 'search scan';

    if (!data.walkins.length) {
      walkInList.innerHTML = '<div class="empty-state">No live jobs were found for this role and location right now.</div>';
      setInlineStatus(walkInMeta, `Checked ${formatDateTime(data.generatedAt)} via ${providerLabel}. No live jobs were found yet.`, 'warning');
      return;
    }

    data.walkins.forEach((item) => {
      const card = document.createElement('article');
      card.className = 'walkin-card glass-card';
      card.innerHTML = `
        <div class="walkin-header">
          <div>
            <p class="eyebrow">Live source</p>
            <h3>${item.title}</h3>
          </div>
          <span class="pill">${item.source}</span>
        </div>
        <div class="details-badges">
          ${renderBadge('map-pin', item.location)}
          ${renderBadge('calendar-days', item.dateLabel)}
        </div>
        <p class="walkin-copy">${item.description || item.snippet || 'Live listing link found from current search results.'}</p>
        <div class="saved-draft">Scanned ${formatDateTime(data.generatedAt)} via ${providerLabel}${item.publishedAt ? ` • listed ${formatDateTime(item.publishedAt)}` : ''}</div>
        <a href="${item.verifyUrl}" target="_blank" rel="noopener noreferrer" class="primary-btn w-full"><i data-lucide="external-link"></i> Open Real Link</a>
      `;
      walkInList.appendChild(card);
    });

    setInlineStatus(walkInMeta, `Updated ${formatDateTime(data.generatedAt)} via ${providerLabel} with ${data.walkins.length} live result(s).`, 'success');
    createIcons({ icons });
  } catch (error) {
    walkInList.innerHTML = '<div class="empty-state">Live scan failed. Check the backend server and internet access for the API.</div>';
    setInlineStatus(walkInMeta, error.message || 'Walk-in scan failed.', 'error');
  }
}

function renderBadge(icon, text, highlighted = false) {
  if (!text) return '';
  return `<span class="badge ${highlighted ? 'highlight-badge' : ''}"><i data-lucide="${icon}"></i><span>${text}</span></span>`;
}

function renderLinkButton(url, label, icon, className) {
  return `<a href="${url}" target="_blank" rel="noopener noreferrer" class="${className}"><i data-lucide="${icon}"></i>${label}</a>`;
}

function truncate(value, length) {
  if (!value) return '';
  return value.length > length ? `${value.slice(0, length - 1)}...` : value;
}

function firstName(name) {
  return (name || 'Recruiter').trim().split(/\s+/)[0];
}

function greetingName(name) {
  const resolved = firstName(name);
  const genericGreetings = new Set(['recruiter', 'hiring', 'careers', 'jobs', 'talent', 'team', 'admin', 'info', 'hr']);
  return resolved && !genericGreetings.has(resolved.toLowerCase()) ? resolved : 'Sir/Madam';
}

function findContactByEmail(email) {
  const normalized = email.toLowerCase();
  return outreachContacts.find((contact) => contact.email.toLowerCase() === normalized) || null;
}

function deriveNameFromEmail(email) {
  const localPart = email.split('@')[0] || 'Recruiter';
  return localPart
    .replace(/[._-]+/g, ' ')
    .replace(/\d+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(' ') || 'Recruiter';
}

function deriveCompanyFromEmail(email) {
  const domain = email.split('@')[1] || '';
  const root = domain.split('.')[0] || '';
  return root
    .replace(/[-_]+/g, ' ')
    .replace(/\b\w/g, (char) => char.toUpperCase()) || 'your company';
}

function personalizeTemplate(template, email) {
  const contact = findContactByEmail(email);
  const resolvedName = contact?.name || deriveNameFromEmail(email);
  const resolvedCompany = contact?.company || deriveCompanyFromEmail(email);
  const resolvedGreetingName = greetingName(resolvedName);
  const resolvedFullName = resolvedGreetingName === 'Sir/Madam' ? 'Sir/Madam' : resolvedName;

  let value = template
    .replace(/\[\s*(Recruiter Name|First Name|Name|Recruiter)\s*\]/gi, resolvedGreetingName)
    .replace(/\[\s*Full Name\s*\]/gi, resolvedFullName)
    .replace(/\[\s*Company\s*\]/gi, resolvedCompany);

  value = value.replace(/^\s*(Hi|Hello)\s*,/i, `Hi ${resolvedGreetingName},`);
  value = value.replace(/^\s*(Hi|Hello)\s+(Sir\/Madam|Recruiter|Hiring Manager)\s*,/i, `Hi ${resolvedGreetingName},`);
  value = value.replace(/^\s*(Hi|Hello)\s+\[\s*(Recruiter Name|First Name|Name|Recruiter)\s*\]\s*,/i, `Hi ${resolvedGreetingName},`);
  value = value.replace(/^\s*Dear\s*,/i, `Dear ${resolvedGreetingName},`);
  value = value.replace(/^\s*Dear\s+(Sir\/Madam|Recruiter|Hiring Manager)\s*,/i, `Dear ${resolvedGreetingName},`);
  return value;
}

function formatShortDate(value) {
  if (!value) return 'Today';
  return new Date(value).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' });
}

function formatDateTime(value) {
  if (!value) return 'just now';
  return new Date(value).toLocaleString('en-IN', {
    day: 'numeric',
    month: 'short',
    hour: 'numeric',
    minute: '2-digit'
  });
}
