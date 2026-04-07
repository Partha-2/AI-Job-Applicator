import './style.css';
import { createIcons, icons } from 'lucide';
import Papa from 'papaparse';
import * as xlsx from 'xlsx';

createIcons({ icons });

// State
let jobs = [];
let currentPage = 1;
const pageSize = 12;
let currentUser = null;
let serverAppliedJobs = [];

let outreachContacts = [];
let outreachVariants = [];

// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const dashboardSection = document.getElementById('dashboardSection');
const jobList = document.getElementById('jobList');
const jobCount = document.getElementById('jobCount');
const resetBtn = document.getElementById('resetBtn');
const tabBtns = document.querySelectorAll('.tab-btn');
const jobsTabContent = document.getElementById('jobsTabContent');
const outreachSection = document.getElementById('outreachSection');
const walkInSection = document.getElementById('walkInSection');
const outreachFileInput = document.getElementById('outreachFileInput');
const outreachList = document.getElementById('outreachList');
const outreachEmptyState = document.getElementById('outreachEmptyState');
const paginationControls = document.getElementById('paginationControls');

// Tab Logic
tabBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    tabBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    
    jobsTabContent.classList.add('hidden');
    outreachSection.classList.add('hidden');
    walkInSection.classList.add('hidden');
    
    const target = btn.dataset.target;
    document.getElementById(target).classList.remove('hidden');
  });
});

// OAuth & User State
async function checkUserSession() {
  try {
    const res = await fetch('/api/user');
    const user = await res.json();
    if (user) {
      currentUser = user;
      document.getElementById('loginBtn').classList.add('hidden');
      document.getElementById('userProfile').classList.remove('hidden');
      document.getElementById('userName').innerText = user.displayName;
      
      // Fetch applied jobs from server
      const appliedRes = await fetch('/api/applied-jobs');
      serverAppliedJobs = await appliedRes.json();
    }
  } catch (e) {
    console.warn("Backend session check failed.");
  }
}

checkUserSession();

// File Upload Event Listeners for Jobs
fileInput.addEventListener('change', handleFileUpload);
uploadSection.addEventListener('dragover', (e) => { e.preventDefault(); uploadSection.classList.add('dragover'); });
uploadSection.addEventListener('dragleave', () => { uploadSection.classList.remove('dragover'); });
uploadSection.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadSection.classList.remove('dragover');
  if (e.dataTransfer.files.length) { fileInput.files = e.dataTransfer.files; handleFileUpload(); }
});

resetBtn.addEventListener('click', () => {
  jobs = [];
  dashboardSection.classList.add('hidden');
  uploadSection.classList.remove('hidden');
  fileInput.value = '';
});

function handleFileUpload() {
  const file = fileInput.files[0];
  if (!file) return;
  const ext = file.name.includes('.') ? file.name.split('.').pop().toLowerCase() : '';
  if (ext === 'xlsx' || ext === 'xls') parseExcel(file); else parseCSV(file);
}

function parseCSV(file) {
  Papa.parse(file, {
    header: true, skipEmptyLines: true,
    complete: (results) => processExtractedData(results.data),
    error: (err) => { console.error(err); alert('Error parsing CSV file'); }
  });
}

function parseExcel(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = xlsx.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const json = xlsx.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
      processExtractedData(json);
    } catch(err) { alert('Error parsing Excel file'); }
  };
  reader.readAsArrayBuffer(file);
}

function processExtractedData(data) {
  if (!data || data.length === 0) return alert('The uploaded file is empty.');
  jobs = data.map(normalizeJobRow);
  if (jobs.length === 0) return alert('Could not find any valid jobs.');
  uploadSection.classList.add('hidden');
  dashboardSection.classList.remove('hidden');
  tabBtns[0].click();
  currentPage = 1;
  renderJobs();
}

function normalizeJobRow(row) {
  const getVal = (keys) => {
    for (let key of keys) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') return row[key];
    }
    const foundKey = Object.keys(row).find(k => {
      const lowerK = k.toLowerCase().replace(/[^a-z0-9]/g, '');
      return keys.some(kw => lowerK.includes(kw.replace(/[^a-z0-9]/g, '')));
    });
    return foundKey ? row[foundKey] : null;
  };

  const id = getVal(['id', 'jobId']) || Math.random().toString(36).substr(2, 9);
  const title = getVal(['title', 'jobtitle', 'role', 'position']) || 'Unknown Title';
  const company = getVal(['company/name', 'company', 'employer', 'org']) || 'Unknown Company';
  const description = getVal(['descriptionText', 'descriptionHtml', 'description', 'jd', 'detail']) || 'No description provided.';
  const logo = getVal(['company/logos/0/url', 'company/logo', 'logoUrl']);
  const linkedinUrl = getVal(['linkedinUrl']);
  const easyApplyUrl = getVal(['applyMethod/easyApplyUrl', 'easyApplyUrl', 'query/easyApply']);
  const companyApplyUrl = getVal(['applyMethod/companyApplyUrl']);
  const generalLink = getVal(['applylink', 'link', 'url', 'applyurl']);
  const location = getVal(['location/linkedinText', 'location/parsed/text', 'location/parsed/city', 'company/locations/0/city']);
  const workplaceType = getVal(['workplaceType', 'workRemoteAllowed']);
  const employmentType = getVal(['employmentType', 'query/employmentType/0']);
  const experienceLevel = getVal(['experienceLevel', 'query/experienceLevel/0']);
  const hrProfile = getVal(['hiringTeam/0/name', 'hiringTeam/0/linkedinUrl', 'hr', 'contact']);
  const salaryMin = getVal(['salary/min']);
  const salaryMax = getVal(['salary/max']);
  const salaryCurrency = getVal(['salary/currency']) || '';
  let salaryStr = getVal(['salary/text']);
  if (!salaryStr && salaryMin && salaryMax) salaryStr = `${salaryCurrency}${salaryMin} - ${salaryCurrency}${salaryMax}`;

  // Priority to server tracking, fallback to localStorage
  const applied = serverAppliedJobs.includes(id) || JSON.parse(localStorage.getItem('appliedJobs') || '[]').includes(id);

  return { id, title, company, description, logo, linkedinUrl, easyApplyUrl, companyApplyUrl, generalLink, location, workplaceType, employmentType, experienceLevel, hrProfile, salaryStr, applied };
}

document.getElementById('prevPageBtn').addEventListener('click', () => { if (currentPage > 1) { currentPage--; renderJobs(); }});
document.getElementById('nextPageBtn').addEventListener('click', () => { 
  const pendingJobs = jobs.filter(j => !j.applied);
  if (currentPage * pageSize < pendingJobs.length) { currentPage++; renderJobs(); }
});

function renderJobs() {
  jobList.innerHTML = '';
  const pendingJobs = jobs.filter(j => !j.applied);
  jobCount.innerText = pendingJobs.length;

  if (pendingJobs.length === 0) {
    jobList.innerHTML = '<div class="empty-state" style="grid-column: 1 / -1; padding: 4rem; text-align: center;">No pending jobs found! 🎉</div>';
    paginationControls.classList.add('hidden');
    return;
  }
  
  const totalPages = Math.ceil(pendingJobs.length / pageSize);
  if (currentPage > totalPages) currentPage = totalPages;
  const startIndex = (currentPage - 1) * pageSize;
  const currentJobs = pendingJobs.slice(startIndex, startIndex + pageSize);

  if (totalPages > 1) {
    paginationControls.classList.remove('hidden');
    document.getElementById('pageInfo').innerText = `Page ${currentPage} of ${totalPages}`;
    document.getElementById('prevPageBtn').disabled = currentPage === 1;
    document.getElementById('nextPageBtn').disabled = currentPage === totalPages;
  } else {
    paginationControls.classList.add('hidden');
  }

  currentJobs.forEach(job => {
    const card = document.createElement('div');
    card.className = 'job-card glass-card';
    
    let badgesHtml = '';
    if (job.location) badgesHtml += `<span class="badge"><i data-lucide="map-pin"></i> ${job.location}</span>`;
    if (job.workplaceType) badgesHtml += `<span class="badge"><i data-lucide="monitor"></i> ${job.workplaceType}</span>`;
    if (job.employmentType) badgesHtml += `<span class="badge"><i data-lucide="briefcase"></i> ${job.employmentType}</span>`;
    if (job.experienceLevel) badgesHtml += `<span class="badge"><i data-lucide="trending-up"></i> ${job.experienceLevel}</span>`;
    if (job.salaryStr) badgesHtml += `<span class="badge highlight-badge"><i data-lucide="banknote"></i> ${job.salaryStr}</span>`;
    if (job.hrProfile) badgesHtml += `<span class="badge"><i data-lucide="user"></i> HR: ${job.hrProfile}</span>`;

    let linksHtml = '';
    if (job.easyApplyUrl) linksHtml += `<a href="${job.easyApplyUrl}" target="_blank" class="primary-btn apply-btn" data-id="${job.id}"><i data-lucide="zap"></i> Easy Apply</a>`;
    if (job.companyApplyUrl) linksHtml += `<a href="${job.companyApplyUrl}" target="_blank" class="secondary-btn apply-btn" data-id="${job.id}"><i data-lucide="external-link"></i> Company Site</a>`;
    if (job.linkedinUrl) linksHtml += `<a href="${job.linkedinUrl}" target="_blank" class="secondary-btn apply-btn" data-id="${job.id}"><i data-lucide="linkedin"></i> LinkedIn</a>`;
    if (!linksHtml && job.generalLink) linksHtml += `<a href="${job.generalLink}" target="_blank" class="primary-btn apply-btn" data-id="${job.id}"><i data-lucide="external-link"></i> Apply Now</a>`;
    if (!linksHtml) linksHtml = `<span class="no-link">No Application Links Found</span>`;

    card.innerHTML = `
      <div class="job-card-header">
        <div class="company-brand">
          ${job.logo ? `<img src="${job.logo}" alt="Company Logo" class="company-logo" onerror="this.style.display='none'" />` : ''}
          <div>
            <h3 class="job-title">${job.title}</h3>
            <p class="job-company"><i data-lucide="building-2"></i> ${job.company}</p>
          </div>
        </div>
        <button class="icon-btn toggle-applied-btn" data-id="${job.id}" title="Mark Applied"><i data-lucide="check-square"></i></button>
      </div>
      ${badgesHtml ? `<div class="job-badges">${badgesHtml}</div>` : ''}
      <div class="job-card-body">
        <p class="job-description">${truncate(job.description, 200)}</p>
      </div>
      <div class="job-card-actions">
        <div class="job-links-group">
          ${linksHtml}
        </div>
      </div>
    `;
    jobList.appendChild(card);
  });
  
  createIcons({ icons });

  document.querySelectorAll('.toggle-applied-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const id = e.target.closest('.toggle-applied-btn').dataset.id;
      markAsApplied(id);
    });
  });
}

async function markAsApplied(id) {
  const job = jobs.find(j => j.id === id);
  if (job) {
    job.applied = true;
    
    // Save to localStorage
    const local = JSON.parse(localStorage.getItem('appliedJobs') || '[]');
    if (!local.includes(id)) {
      local.push(id);
      localStorage.setItem('appliedJobs', JSON.stringify(local));
    }
    
    // Attempt save to server if logged in
    if (currentUser) {
      try {
        await fetch('/api/applied-jobs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ jobId: id })
        });
      } catch (e) {
        console.error("Failed to sync applied job to server.");
      }
    }
    
    renderJobs();
  }
}

function truncate(str, n) { if (!str) return ''; return (str.length > n) ? str.slice(0, n - 1) + '...' : str; }

// --- OUTREACH MANAGER LOGIC ---

outreachFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => { parseOutreachHTML(ev.target.result); };
  reader.readAsText(file);
});

function parseOutreachHTML(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, 'text/html');
  outreachContacts = [];
  outreachVariants = [];

  doc.querySelectorAll('.contact-card').forEach(card => {
    const nameEl = card.querySelector('.contact-name');
    const companyEl = card.querySelector('.contact-company');
    const emailEl = card.querySelector('.contact-email a'); 
    if (nameEl && emailEl) {
       outreachContacts.push({ name: nameEl.innerText.trim(), company: companyEl ? companyEl.innerText.trim() : '', email: emailEl.innerText.trim() });
    }
  });

  doc.querySelectorAll('.outreach-card').forEach((card, idx) => {
    const subjectEl = card.querySelector('.outreach-subject');
    const textBox = card.querySelector('.outreach-text');
    let subject = subjectEl ? subjectEl.innerText.replace('Subject:', '').trim() : 'Application for Open Role';
    let text = textBox ? textBox.innerText.trim() : '';
    if (text) outreachVariants.push({ id: 'v' + (idx + 1), subject, bodyTemplate: text });
  });

  if (outreachContacts.length === 0) return alert("No valid HR contacts found in the HTML.");

  outreachEmptyState.classList.add('hidden');
  document.getElementById('automationConfig').classList.remove('hidden');
  outreachList.classList.remove('hidden');
  renderOutreachContacts();
}

function renderOutreachContacts() {
  outreachList.innerHTML = '';
  outreachContacts.forEach((contact, idx) => {
     const row = document.createElement('div');
     row.className = 'outreach-item glass-card';
     row.innerHTML = `
        <div class="outreach-person">
          <h4>${contact.name}</h4>
          <p><i data-lucide="building"></i> ${contact.company}</p>
          <p style="color: #818cf8;"><i data-lucide="mail"></i> ${contact.email}</p>
        </div>
        <button class="secondary-btn btn-sm copy-email-btn" data-email="${contact.email}"><i data-lucide="copy"></i></button>
     `;
     outreachList.appendChild(row);
  });
  createIcons({ icons });
  document.querySelectorAll('.copy-email-btn').forEach(btn => btn.addEventListener('click', (e) => {
    navigator.clipboard.writeText(e.target.closest('.copy-email-btn').dataset.email);
    alert('Email copied!');
  }));
}

// BULK AUTOMATION
document.getElementById('fireAllMailsBtn').addEventListener('click', async () => {
   const email = document.getElementById('gmailAddress').value;
   const password = document.getElementById('gmailPassword').value;
   const resume = document.getElementById('resumeUpload').files[0];
   if (!email || !password || !resume) return alert("Fill Sender Gmail, App Password & Resume PDF first!");
   
   const payloadContacts = outreachContacts.map(c => {
     const v = outreachVariants[0];
     return {
       to: c.email,
       subject: v.subject,
       body: v.bodyTemplate.replace(/\[Recruiter Name\]/gi, c.name.split(' ')[0]).replace(/\[Company\]/gi, c.company || 'your company')
     };
   });

   const fd = new FormData();
   fd.append('email', email); fd.append('password', password); fd.append('contacts', JSON.stringify(payloadContacts)); fd.append('resume', resume);

   const btn = document.getElementById('fireAllMailsBtn');
   btn.disabled = true; btn.innerText = "Sending...";

   try {
     const res = await fetch('/api/send-cold-emails', { method: 'POST', body: fd });
     const data = await res.json();
     if (data.success) alert("Bulk outreach complete!"); else alert("Failed: " + data.error);
   } catch(e) { alert("Backend unreachable."); }
   finally { btn.disabled = false; btn.innerText = "Fire All Emails Automatically"; }
});

// MANUAL SINGLE EMAIL
document.getElementById('sendManualBtn').addEventListener('click', async () => {
    const fromEmail = document.getElementById('manualFromEmail').value;
    const fromPass = document.getElementById('manualFromPass').value;
    const to = document.getElementById('manualToEmail').value;
    const subject = document.getElementById('manualSubject').value;
    const body = document.getElementById('manualBody').value;
    const files = document.getElementById('manualAttachments').files;

    if(!fromEmail || !fromPass || !to || !subject || !body) return alert("Please fill all manual email fields.");

    const fd = new FormData();
    fd.append('fromEmail', fromEmail);
    fd.append('fromPass', fromPass);
    fd.append('to', to);
    fd.append('subject', subject);
    fd.append('body', body);
    for(let f of files) fd.append('attachments', f);

    const btn = document.getElementById('sendManualBtn');
    btn.disabled = true; btn.innerText = "Sending...";

    try {
        const res = await fetch('/api/send-single-email', { method: 'POST', body: fd });
        const data = await res.json();
        if(data.success) alert("Manual email sent!"); else alert("Error: " + data.error);
    } catch(e) { alert("Backend unreachable."); }
    finally { btn.disabled = false; btn.innerText = "Send Email"; }
});

// URL SCRAPER
document.getElementById('scrapeBtn').addEventListener('click', async () => {
    const url = document.getElementById('scrapeUrl').value;
    if(!url) return alert("Enter a URL first.");
    const btn = document.getElementById('scrapeBtn');
    btn.disabled = true; btn.innerText = "Scraping...";

    try {
        const res = await fetch('/api/scrape-emails', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url })
        });
        const data = await res.json();
        if(data.success && data.emails.length > 0) {
            document.getElementById('scrapeResults').classList.remove('hidden');
            const list = document.getElementById('scrapedEmailsList');
            list.innerHTML = data.emails.map(e => `<li>${e} <button class="primary-btn btn-sm" onclick="document.getElementById('manualToEmail').value='${e}'; document.getElementById('outreachSection').scrollIntoView();">Use</button></li>`).join('');
        } else {
            alert("No emails found on that page.");
        }
    } catch(e) { alert("Scraper failed. Site might be blocking crawlers."); }
    finally { btn.disabled = false; btn.innerText = "Scrape"; }
});


// WALK IN SCANNER
document.getElementById('scanWalkinsBtn').addEventListener('click', async () => {
    const role = document.getElementById('walkInRole').value;
    const loc = document.getElementById('walkInLocation').value;
    const list = document.getElementById('walkInList');
    list.innerHTML = `<div class="empty-state" style="grid-column: 1 / -1;"><i data-lucide="loader"></i> Scanning job networks...</div>`;
    createIcons({ icons });

    try {
      const res = await fetch(`/api/walkins?role=${encodeURIComponent(role)}&location=${encodeURIComponent(loc)}`);
      const data = await res.json();
      list.innerHTML = '';
      if(data.walkins && data.walkins.length > 0) {
        data.walkins.forEach(w => {
           const card = document.createElement('div');
           card.className = 'job-card glass-card';
           card.innerHTML = `
             <div class="job-card-header">
               <div>
                 <h3 class="job-title"><i data-lucide="calendar"></i> ${w.date}</h3>
                 <p class="job-company">${w.company} - ${w.role}</p>
               </div>
             </div>
             <div class="job-badges">
                <span class="badge highlight-badge"><i data-lucide="map-pin"></i> ${w.location}</span>
                <span class="badge"><i data-lucide="clock"></i> ${w.time}</span>
             </div>
             <div class="job-card-body" style="margin-top: 1rem;">
               <p style="font-size: 0.875rem; color: #e2e8f0; margin-bottom: 0.5rem;"><strong>Address:</strong> ${w.address}</p>
               <p class="job-description">${w.description}</p>
             </div>
             <div class="job-card-actions">
               <a href="${w.verifyUrl}" target="_blank" class="secondary-btn w-full"><i data-lucide="external-link"></i> Verify on Naukri/LinkedIn</a>
             </div>
           `;
           list.appendChild(card);
        });
        createIcons({ icons });
      } else {
        list.innerHTML = `<div class="empty-state" style="grid-column: 1 / -1;">No walk-ins found today.</div>`;
      }
    } catch(e) {
        list.innerHTML = `<div class="empty-state" style="grid-column: 1 / -1;">Backend unreachable. Run 'node server.js'.</div>`;
    }
});
