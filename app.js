import express from 'express';
import nodemailer from 'nodemailer';
import multer from 'multer';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import passport from 'passport';
import { Strategy as GoogleStrategy } from 'passport-google-oauth20';
import session from 'express-session';
import axios from 'axios';
import * as cheerio from 'cheerio';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DATA_DIR = path.join(__dirname, '.data');
const APPLIED_JOBS_FILE = path.join(DATA_DIR, 'applied-jobs.json');
const WALKIN_ALLOWED_HOSTS = ['naukri.com', 'linkedin.com', 'foundit.in', 'indeed.com', 'timesjobs.com'];

function ensureDataFile() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  if (!fs.existsSync(APPLIED_JOBS_FILE)) {
    fs.writeFileSync(APPLIED_JOBS_FILE, JSON.stringify({}, null, 2));
  }
}

function readAppliedJobsStore() {
  try {
    ensureDataFile();
    return JSON.parse(fs.readFileSync(APPLIED_JOBS_FILE, 'utf8'));
  } catch (error) {
    console.warn('Failed to read applied jobs store, using empty store.', error.message);
    return {};
  }
}

function writeAppliedJobsStore(store) {
  ensureDataFile();
  fs.writeFileSync(APPLIED_JOBS_FILE, JSON.stringify(store, null, 2));
}

function getOwnerKey(req) {
  const sessionUser = getSessionUser(req);
  const userEmail = sessionUser?.email || sessionUser?.emails?.[0]?.value;
  if (userEmail) {
    return `user:${userEmail.toLowerCase()}`;
  }

  const clientId = req.get('x-client-id') || req.query.clientId || req.body?.clientId;
  if (clientId) {
    return `guest:${clientId}`;
  }

  return null;
}

function buildAppliedRecord(payload) {
  return {
    id: payload.id,
    title: payload.title || 'Unknown Title',
    company: payload.company || 'Unknown Company',
    description: payload.description || 'No description provided.',
    logo: payload.logo || '',
    linkedinUrl: payload.linkedinUrl || '',
    easyApplyUrl: payload.easyApplyUrl || '',
    companyApplyUrl: payload.companyApplyUrl || '',
    generalLink: payload.generalLink || '',
    location: payload.location || '',
    workplaceType: payload.workplaceType || '',
    employmentType: payload.employmentType || '',
    experienceLevel: payload.experienceLevel || '',
    hrProfile: payload.hrProfile || '',
    hrEmail: payload.hrEmail || '',
    hrPhone: payload.hrPhone || '',
    salaryStr: payload.salaryStr || '',
    status: payload.status || 'Applied',
    notes: payload.notes || '',
    appliedAt: payload.appliedAt || new Date().toISOString(),
    lastUpdatedAt: new Date().toISOString()
  };
}

function getSessionUser(req) {
  return req.user || req.session?.localUser || null;
}

function buildPublicUser(user) {
  if (!user) return null;

  return {
    id: user.id,
    displayName: user.displayName || 'User',
    email: user.email || user.emails?.[0]?.value || '',
    photo: user.photo || user.photos?.[0]?.value || '',
    canSendMail: Boolean(user.refreshToken && process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET)
  };
}

function createAuthenticatedGmailTransport(req) {
  if (!req.user) {
    const error = new Error('Sign in with Google first to send email.');
    error.statusCode = 401;
    throw error;
  }

  if (!process.env.GOOGLE_CLIENT_ID || !process.env.GOOGLE_CLIENT_SECRET) {
    const error = new Error('Google mail sending is not configured on the server.');
    error.statusCode = 500;
    throw error;
  }

  if (!req.user.refreshToken) {
    const error = new Error('Your Google session does not include Gmail send access. Log out and sign in again, then approve Gmail access.');
    error.statusCode = 401;
    throw error;
  }

  const senderEmail = req.user.email || req.user.emails?.[0]?.value;
  if (!senderEmail) {
    const error = new Error('Could not determine the logged-in sender email.');
    error.statusCode = 400;
    throw error;
  }

  return {
    senderEmail,
    transporter: nodemailer.createTransport({
      service: 'gmail',
      auth: {
        type: 'OAuth2',
        user: senderEmail,
        clientId: process.env.GOOGLE_CLIENT_ID,
        clientSecret: process.env.GOOGLE_CLIENT_SECRET,
        refreshToken: req.user.refreshToken
      }
    })
  };
}

function normalizeMailerError(error) {
  const rawMessage = error?.message || 'Email sending failed.';
  const normalized = rawMessage.toLowerCase();

  if (normalized.includes('gmail send access')) {
    return rawMessage;
  }

  if (normalized.includes('sign in with google first')) {
    return rawMessage;
  }

  if (error?.code === 'EAUTH' || normalized.includes('535-5.7.8') || normalized.includes('badcredentials')) {
    return 'Google rejected mail access for this session. Log out, sign in again with Google, and approve Gmail send access.';
  }

  if (normalized.includes('invalid login')) {
    return 'Google mail login failed. Log out, sign in again with Google, and approve Gmail send access for this account.';
  }

  if (normalized.includes('daily user sending quota exceeded')) {
    return 'Gmail sending quota has been exceeded for this account today. Try again later or use a different account.';
  }

  if (normalized.includes('message rejected') || normalized.includes('unauthenticated')) {
    return 'Google rejected the message. Re-check the logged-in sender account and attachment size.';
  }

  return rawMessage;
}

function decodeDuckDuckGoUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl, 'https://duckduckgo.com');
    const uddg = parsed.searchParams.get('uddg');
    return uddg ? decodeURIComponent(uddg) : rawUrl;
  } catch {
    return rawUrl;
  }
}

function isAllowedWalkinHost(url) {
  try {
    const hostname = new URL(url).hostname.replace(/^www\./, '');
    return WALKIN_ALLOWED_HOSTS.some((host) => hostname === host || hostname.endsWith(`.${host}`));
  } catch {
    return false;
  }
}

function createSearchFallback(role, location) {
  return [];
}

async function searchWalkinQueryViaBing(query, location) {
  const response = await axios.get('https://www.bing.com/search', {
    params: { format: 'rss', q: query },
    timeout: 12000,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36'
    }
  });

  const $ = cheerio.load(response.data, { xmlMode: true });
  const results = [];

  $('item').each((index, element) => {
    if (index >= 8) return false;

    const verifyUrl = $(element).find('link').first().text().trim();
    if (!verifyUrl || !isAllowedWalkinHost(verifyUrl)) {
      return;
    }

    const title = $(element).find('title').first().text().trim();
    const snippet = $(element).find('description').first().text().trim();
    const source = (() => {
      try {
        return new URL(verifyUrl).hostname.replace(/^www\./, '');
      } catch {
        return 'External';
      }
    })();

    const companyMatch = title.match(/^(.*?)\s[-|–]/);

    results.push({
      id: `${source}-bing-${index}-${Buffer.from(verifyUrl).toString('base64').slice(0, 12)}`,
      title: title || `${query} result`,
      company: companyMatch?.[1]?.trim() || source,
      location,
      dateLabel: new Date().toLocaleDateString('en-IN', {
        day: 'numeric',
        month: 'short',
        year: 'numeric'
      }),
      description: snippet || 'Live listing discovered from indexed search results.',
      verifyUrl,
      source,
      snippet
    });
  });

  return results;
}

async function searchWalkinQuery(query, location) {
  const liveResults = [];

  try {
    const bingResults = await searchWalkinQueryViaBing(query, location);
    liveResults.push(...bingResults);
  } catch {
    // Fall through to DuckDuckGo HTML search.
  }

  if (liveResults.length > 0) {
    return liveResults;
  }

  const response = await axios.get('https://html.duckduckgo.com/html/', {
    params: { q: query },
    timeout: 12000,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36'
    }
  });

  const $ = cheerio.load(response.data);
  const results = [];

  $('div.result').each((index, element) => {
    if (index >= 6) return false;

    const anchor = $(element).find('a.result__a').first();
    const rawHref = anchor.attr('href');
    const verifyUrl = decodeDuckDuckGoUrl(rawHref || '');

    if (!verifyUrl || !isAllowedWalkinHost(verifyUrl)) {
      return;
    }

    const title = anchor.text().trim();
    const snippet = $(element).find('.result__snippet').text().trim();
    const source = (() => {
      try {
        return new URL(verifyUrl).hostname.replace(/^www\./, '');
      } catch {
        return 'External';
      }
    })();

    const companyMatch = title.match(/^(.*?)\s[-|–]/);

    results.push({
      id: `${source}-${index}-${Buffer.from(verifyUrl).toString('base64').slice(0, 12)}`,
      title: title || `${query} result`,
      company: companyMatch?.[1]?.trim() || source,
      location,
      dateLabel: new Date().toLocaleDateString('en-IN', {
        day: 'numeric',
        month: 'short',
        year: 'numeric'
      }),
      description: snippet || 'Live listing discovered from a search result.',
      verifyUrl,
      source,
      snippet
    });
  });

  return results;
}

async function fetchLiveWalkins(role, location) {
  const queries = [
    `"walk in interview" ${role} ${location} site:linkedin.com/jobs`,
    `"walk in" ${role} ${location} site:linkedin.com`,
    `"walk in" ${role} ${location} site:foundit.in`,
    `"walk in interview" ${role} ${location} site:naukri.com`,
    `${role} hiring drive ${location} site:linkedin.com`
  ];

  const settled = await Promise.allSettled(queries.map((query) => searchWalkinQuery(query, location)));
  const deduped = new Map();

  settled.forEach((result) => {
    if (result.status !== 'fulfilled') return;

    result.value.forEach((item) => {
      if (!deduped.has(item.verifyUrl)) {
        deduped.set(item.verifyUrl, item);
      }
    });
  });

  const liveResults = Array.from(deduped.values()).slice(0, 12);
  return liveResults.length > 0 ? liveResults : createSearchFallback(role, location);
}

export function createApp() {
  const app = express();
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';

  app.use(cors({ origin: process.env.FRONTEND_URL || 'http://localhost:5173', credentials: true }));
  app.use(express.json());
  app.use(session({
    secret: process.env.SESSION_SECRET || 'job-applicator-secret',
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false }
  }));
  app.use(passport.initialize());
  app.use(passport.session());

  const upload = multer({ dest: process.env.VERCEL ? '/tmp' : 'uploads/' });

  passport.serializeUser((user, done) => done(null, user));
  passport.deserializeUser((user, done) => done(null, user));

  if (process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_ID !== 'your_google_client_id_here') {
    passport.use(new GoogleStrategy({
      clientID: process.env.GOOGLE_CLIENT_ID,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      callbackURL: process.env.GOOGLE_CALLBACK_URL || 'https://ai-job-applicator.vercel.app/auth/google/callback'
    }, (accessToken, refreshToken, profile, done) => done(null, {
      id: profile.id,
      displayName: profile.displayName,
      email: profile.emails?.[0]?.value || '',
      photo: profile.photos?.[0]?.value || '',
      accessToken,
      refreshToken
    })));

    app.get('/auth/google', passport.authenticate('google', {
      scope: ['profile', 'email']
    }));

    app.get('/auth/google-gmail', passport.authenticate('google', {
      scope: ['profile', 'email', 'https://www.googleapis.com/auth/gmail.send'],
      accessType: 'offline',
      prompt: 'consent',
      includeGrantedScopes: true
    }));

    app.get('/auth/google/callback',
      passport.authenticate('google', { failureRedirect: `${frontendUrl}/login?error=true` }),
      (req, res) => {
        res.redirect(frontendUrl);
      }
    );
  } else {
    console.warn('Google OAuth keys missing or default. Authentication features disabled.');
    app.get('/auth/google', (req, res) => res.status(501).json({ error: 'OAuth not configured. Please add keys to .env' }));
    app.get('/auth/google-gmail', (req, res) => res.status(501).json({ error: 'OAuth not configured. Please add keys to .env' }));
  }

  app.get('/api/user', (req, res) => {
    res.json(buildPublicUser(getSessionUser(req)));
  });

  app.get('/api/health', (req, res) => {
    res.json({
      ok: true,
      timestamp: new Date().toISOString()
    });
  });

  app.post('/auth/local-login', (req, res) => {
    const name = (req.body?.name || '').toString().trim();
    const email = (req.body?.email || '').toString().trim().toLowerCase();

    if (!name || !email) {
      return res.status(400).json({ error: 'Name and email are required.' });
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ error: 'Enter a valid email address.' });
    }

    req.session.localUser = {
      id: `local:${email}`,
      displayName: name,
      email,
      provider: 'local'
    };

    res.json({ success: true, user: buildPublicUser(req.session.localUser) });
  });

  app.get('/auth/logout', (req, res, next) => {
    if (req.session) {
      req.session.localUser = null;
    }

    if (req.user && typeof req.logout === 'function') {
      req.logout((err) => {
        if (err) return next(err);
        res.redirect(frontendUrl);
      });
      return;
    }

    req.session?.destroy?.(() => {
      res.redirect(frontendUrl);
    });
  });

  app.get('/api/applied-jobs', (req, res) => {
    const ownerKey = getOwnerKey(req);
    if (!ownerKey) return res.json([]);

    const store = readAppliedJobsStore();
    res.json(store[ownerKey] || []);
  });

  app.post('/api/applied-jobs', (req, res) => {
    const ownerKey = getOwnerKey(req);
    const { job } = req.body;

    if (!ownerKey) {
      return res.status(400).json({ error: 'Missing client identity.' });
    }
    if (!job?.id) {
      return res.status(400).json({ error: 'Job payload with id is required.' });
    }

    const store = readAppliedJobsStore();
    const records = store[ownerKey] || [];
    const normalized = buildAppliedRecord(job);
    const existingIndex = records.findIndex((record) => record.id === normalized.id);

    if (existingIndex >= 0) {
      records[existingIndex] = {
        ...records[existingIndex],
        ...normalized,
        appliedAt: records[existingIndex].appliedAt || normalized.appliedAt
      };
    } else {
      records.unshift(normalized);
    }

    store[ownerKey] = records;
    writeAppliedJobsStore(store);

    res.json({ success: true, records });
  });

  app.patch('/api/applied-jobs/:id', (req, res) => {
    const ownerKey = getOwnerKey(req);
    if (!ownerKey) {
      return res.status(400).json({ error: 'Missing client identity.' });
    }

    const store = readAppliedJobsStore();
    const records = store[ownerKey] || [];
    const index = records.findIndex((record) => record.id === req.params.id);

    if (index < 0) {
      return res.status(404).json({ error: 'Applied job not found.' });
    }

    records[index] = {
      ...records[index],
      status: req.body.status || records[index].status,
      notes: req.body.notes ?? records[index].notes,
      lastUpdatedAt: new Date().toISOString()
    };

    store[ownerKey] = records;
    writeAppliedJobsStore(store);

    res.json({ success: true, record: records[index], records });
  });

  app.get('/api/walkins', async (req, res) => {
    const location = (req.query.location || 'Bangalore').toString();
    const role = (req.query.role || 'Java Backend').toString();

    try {
      const walkins = await fetchLiveWalkins(role, location);
      res.json({
        walkins,
        generatedAt: new Date().toISOString(),
        location,
        role,
        liveResultsFound: walkins.length > 0
      });
    } catch (error) {
      res.status(500).json({
        success: false,
        error: 'Unable to fetch live walk-in links right now.',
        details: error.message
      });
    }
  });

  app.post('/api/send-cold-emails', upload.single('resume'), async (req, res) => {
    try {
      const { contacts } = req.body;
      const parsedContacts = JSON.parse(contacts || '[]');
      const { senderEmail, transporter } = createAuthenticatedGmailTransport(req);

      if (parsedContacts.length === 0) {
        return res.status(400).json({ success: false, error: 'Missing required fields.' });
      }

      const results = [];
      for (const contact of parsedContacts) {
        try {
          await transporter.sendMail({
            from: senderEmail,
            to: contact.to,
            subject: contact.subject,
            text: contact.body,
            attachments: req.file ? [{ filename: req.file.originalname, path: req.file.path }] : []
          });
          results.push({ email: contact.to, status: 'Sent' });
        } catch (error) {
          results.push({ email: contact.to, status: 'Failed', error: normalizeMailerError(error) });
        }
      }

      if (req.file) fs.unlinkSync(req.file.path);
      res.json({ success: true, results, senderEmail });
    } catch (error) {
      res.status(error.statusCode || 500).json({ success: false, error: normalizeMailerError(error) });
    }
  });

  app.post('/api/send-single-email', upload.array('attachments'), async (req, res) => {
    try {
      const { to, subject, body } = req.body;
      const { senderEmail, transporter } = createAuthenticatedGmailTransport(req);

      if (!to || !subject || !body) {
        return res.status(400).json({ success: false, error: 'Missing required email fields.' });
      }

      const attachments = req.files ? req.files.map((file) => ({ filename: file.originalname, path: file.path })) : [];
      await transporter.sendMail({ from: senderEmail, to, subject, text: body, attachments });

      if (req.files) req.files.forEach((file) => fs.unlinkSync(file.path));
      res.json({ success: true, senderEmail });
    } catch (error) {
      res.status(error.statusCode || 500).json({ success: false, error: normalizeMailerError(error) });
    }
  });

  app.post('/api/scrape-emails', async (req, res) => {
    try {
      const { url } = req.body;
      if (!url) {
        return res.status(400).json({ error: 'URL is required' });
      }

      const parsedUrl = new URL(url);
      if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
        return res.status(400).json({ error: 'Only http/https URLs are supported.' });
      }

      const response = await axios.get(parsedUrl.toString(), {
        timeout: 12000,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });

      const $ = cheerio.load(response.data);
      const text = $('body').text();
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
      const emails = Array.from(new Set(text.match(emailRegex) || []));

      res.json({ success: true, emails, count: emails.length });
    } catch (error) {
      res.status(500).json({ success: false, error: error.message });
    }
  });

  return app;
}
