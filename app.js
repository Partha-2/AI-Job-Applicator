import express from 'express';
import nodemailer from 'nodemailer';
import multer from 'multer';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import passport from 'passport';
import { Strategy as GoogleStrategy } from 'passport-google-oauth20';
import jwt from 'jsonwebtoken';
import cookieParser from 'cookie-parser';
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
const SERPAPI_SEARCH_URL = 'https://serpapi.com/search';

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

const JWT_SECRET = process.env.SESSION_SECRET || 'job-applicator-jwt-secret';

function signUserToken(user) {
  return jwt.sign({ user }, JWT_SECRET, { expiresIn: '7d' });
}

function getUserFromToken(req) {
  try {
    const token = req.cookies?.auth_token;
    if (!token) return null;
    const decoded = jwt.verify(token, JWT_SECRET);
    return decoded.user || null;
  } catch {
    return null;
  }
}

function setAuthCookie(res, user) {
  const token = signUserToken(user);
  res.cookie('auth_token', token, {
    httpOnly: true,
    secure: process.env.VERCEL ? true : false,
    sameSite: process.env.VERCEL ? 'none' : 'lax',
    maxAge: 7 * 24 * 60 * 60 * 1000 // 7 days
  });
}

function clearAuthCookie(res) {
  res.clearCookie('auth_token', {
    httpOnly: true,
    secure: process.env.VERCEL ? true : false,
    sameSite: process.env.VERCEL ? 'none' : 'lax'
  });
}

function getSessionUser(req) {
  return getUserFromToken(req);
}

function extractEmailAddress(value) {
  if (!value) return '';
  const match = value.match(/<([^>]+)>/);
  return (match?.[1] || value).trim();
}

function resolveReplyTo(req) {
  const userEmail = req.user?.email || req.user?.emails?.[0]?.value || '';
  const requestedEmail = (req.body?.senderEmail || req.body?.replyTo || req.body?.gmailUser || '').toString().trim();
  return extractEmailAddress(userEmail || requestedEmail);
}

function getResendConfig() {
  const apiKey = process.env.RESEND_API_KEY || process.env.RESEND_KEY || '';
  if (!apiKey) return null;

  const from = process.env.RESEND_FROM_EMAIL
    || process.env.RESEND_FROM
    || process.env.RESEND_SENDER
    || 'AI Job Applicator <onboarding@resend.dev>';

  return {
    apiKey,
    from,
    senderEmail: extractEmailAddress(from)
  };
}

function getServerMailConfig() {
  const host = process.env.SMTP_HOST || process.env.MAIL_HOST || '';
  const configuredService = process.env.SMTP_SERVICE || process.env.MAIL_SERVICE || '';
  const service = configuredService || (!host ? 'gmail' : '');
  const user = process.env.SMTP_USER || process.env.MAIL_USER || process.env.EMAIL_USER || '';
  const pass = process.env.SMTP_PASS || process.env.MAIL_PASS || process.env.EMAIL_PASS || '';

  if (!user || !pass) return null;

  const parsedPort = Number(process.env.SMTP_PORT || process.env.MAIL_PORT || '');
  const secureFlag = (process.env.SMTP_SECURE || process.env.MAIL_SECURE || '').toLowerCase();
  const secure = secureFlag ? secureFlag === 'true' : parsedPort === 465;

  return {
    host,
    service,
    port: Number.isFinite(parsedPort) && parsedPort > 0 ? parsedPort : (secure ? 465 : 587),
    secure,
    user,
    pass,
    senderEmail: process.env.SMTP_FROM || process.env.MAIL_FROM || process.env.EMAIL_FROM || user
  };
}

function buildMailCapability(user) {
  const serverMail = getServerMailConfig();
  if (serverMail) {
    return {
      canSendMail: true,
      mode: 'server-smtp',
      senderEmail: serverMail.senderEmail,
      requiresGoogleAuth: false
    };
  }

  const resendConfig = getResendConfig();
  if (resendConfig) {
    return {
      canSendMail: true,
      mode: 'resend-api',
      senderEmail: resendConfig.senderEmail,
      requiresGoogleAuth: false
    };
  }

  const gmailClientId = process.env.GOOGLE_GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
  const gmailClientSecret = process.env.GOOGLE_GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;
  const senderEmail = user?.email || user?.emails?.[0]?.value || '';

  return {
    canSendMail: Boolean(user?.refreshToken && gmailClientId && gmailClientSecret),
    mode: 'google-oauth',
    senderEmail,
    requiresGoogleAuth: Boolean(gmailClientId && gmailClientSecret)
  };
}

function buildPublicUser(user) {
  if (!user) return null;
  const mailCapability = buildMailCapability(user);

  return {
    id: user.id,
    displayName: user.displayName || 'User',
    email: user.email || user.emails?.[0]?.value || '',
    photo: user.photo || user.photos?.[0]?.value || '',
    canSendMail: mailCapability.canSendMail,
    mailMode: mailCapability.mode,
    senderEmail: mailCapability.senderEmail,
    requiresGoogleAuth: mailCapability.requiresGoogleAuth
  };
}

function textToHtml(text) {
  const safeText = (text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  return `<div style="font-family:Arial,sans-serif;line-height:1.6;color:#111827;white-space:normal;">${safeText.replace(/\n/g, '<br />')}</div>`;
}

function buildResendAttachments(attachments = []) {
  return attachments.map((attachment) => ({
    filename: attachment.filename,
    content: fs.readFileSync(attachment.path).toString('base64')
  }));
}

function createResendMailTransport() {
  const resendConfig = getResendConfig();
  if (!resendConfig) return null;

  return {
    senderEmail: resendConfig.senderEmail,
    transporter: {
      async sendMail(message) {
        const payload = {
          from: resendConfig.from,
          to: Array.isArray(message.to) ? message.to : [message.to],
          subject: message.subject,
          html: message.html || textToHtml(message.text || ''),
          text: message.text || '',
          reply_to: message.replyTo || undefined,
          attachments: message.attachments?.length ? buildResendAttachments(message.attachments) : undefined
        };

        return axios.post('https://api.resend.com/emails', payload, {
          headers: {
            Authorization: `Bearer ${resendConfig.apiKey}`,
            'Content-Type': 'application/json'
          },
          timeout: 15000
        });
      }
    }
  };
}

function createServerMailTransport() {
  const serverMail = getServerMailConfig();
  if (!serverMail) return null;

  const transportConfig = serverMail.service
    ? {
        service: serverMail.service,
        auth: {
          user: serverMail.user,
          pass: serverMail.pass
        }
      }
    : {
        host: serverMail.host,
        port: serverMail.port,
        secure: serverMail.secure,
        auth: {
          user: serverMail.user,
          pass: serverMail.pass
        }
      };

  return {
    senderEmail: serverMail.senderEmail,
    transporter: nodemailer.createTransport(transportConfig)
  };
}

function createAuthenticatedGmailTransport(req) {
  if (!req.user) {
    const error = new Error('Sign in with Google first to send email.');
    error.statusCode = 401;
    throw error;
  }

  const gmailClientId = process.env.GOOGLE_GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
  const gmailClientSecret = process.env.GOOGLE_GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;

  if (!gmailClientId || !gmailClientSecret) {
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
        clientId: gmailClientId,
        clientSecret: gmailClientSecret,
        refreshToken: req.user.refreshToken
      }
    })
  };
}

function createRequestGmailAppPasswordTransport(req) {
  const gmailUser = (req.body?.gmailUser || req.query?.gmailUser || '').toString().trim();
  const gmailAppPassword = (req.body?.gmailAppPassword || req.query?.gmailAppPassword || '').toString().replace(/\s/g, '');

  if (!gmailUser || !gmailAppPassword) return null;

  return {
    senderEmail: gmailUser,
    transporter: nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: gmailUser,
        pass: gmailAppPassword
      }
    })
  };
}

function createMailTransport(req) {
  return createRequestGmailAppPasswordTransport(req)
    || createServerMailTransport()
    || createResendMailTransport()
    || createAuthenticatedGmailTransport(req);
}

function buildOAuthUser(existingUser, profile, accessToken, refreshToken) {
  return {
    id: existingUser?.id || profile.id,
    displayName: profile.displayName || existingUser?.displayName || 'User',
    email: profile.emails?.[0]?.value || existingUser?.email || '',
    photo: profile.photos?.[0]?.value || existingUser?.photo || '',
    accessToken: accessToken || existingUser?.accessToken || '',
    refreshToken: refreshToken || existingUser?.refreshToken || '',
    provider: 'google'
  };
}

function normalizeMailerError(error) {
  const rawMessage = error?.response?.data?.message || error?.response?.data?.error || error?.message || 'Email sending failed.';
  const normalized = rawMessage.toLowerCase();
  const hasResendSender = Boolean(getResendConfig());
  const hasServerSender = Boolean(getServerMailConfig());

  if (hasResendSender && normalized.includes('api key')) {
    return 'Resend rejected the API key. Update RESEND_API_KEY on the server.';
  }

  if (hasResendSender && normalized.includes('domain') && normalized.includes('not verified')) {
    return 'Resend sender domain is not verified. Verify the domain in Resend and update RESEND_FROM_EMAIL.';
  }

  if (hasResendSender && normalized.includes('you can only send testing emails')) {
    return 'Resend can only send test emails to your own address while using resend.dev. Verify a domain in Resend and set RESEND_FROM_EMAIL to that domain.';
  }

  if (normalized.includes('gmail send access')) {
    return rawMessage;
  }

  if (normalized.includes('sign in with google first')) {
    return 'Sign in with Google or provide your Gmail App Password before sending mail.';
  }

  if (error?.code === 'EAUTH' || normalized.includes('535-5.7.8') || normalized.includes('badcredentials')) {
    if (hasServerSender) {
      return 'Server mail sender login failed. Check SMTP credentials on the server and use an app password if this is Gmail.';
    }
    return 'Gmail login failed. Re-check the Gmail address and 16-character App Password, or reconnect Google sender.';
  }

  if (normalized.includes('invalid login')) {
    if (hasServerSender) {
      return 'Server mail sender rejected the configured login. Update SMTP credentials on the server.';
    }
    return 'Gmail login failed. Make sure you are using a 16-character Google App Password, not your normal Gmail password.';
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

function getSerpApiKey() {
  return (process.env.SERPAPI_API_KEY || process.env.SERP_API_KEY || '').trim();
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

function parseRelativePostedAt(rawValue) {
  const text = (rawValue || '').toString().trim().toLowerCase();
  if (!text) return null;
  if (text.includes('just posted') || text.includes('today')) return new Date();
  if (text.includes('yesterday')) {
    const date = new Date();
    date.setDate(date.getDate() - 1);
    return date;
  }

  const match = text.match(/(\d+)\+?\s*(hour|day|week|month)s?\s*ago/);
  if (!match) return extractWalkinDate(rawValue);

  const amount = Number(match[1]);
  const unit = match[2];
  const date = new Date();

  if (unit === 'hour') {
    date.setHours(date.getHours() - amount);
  } else if (unit === 'day') {
    date.setDate(date.getDate() - amount);
  } else if (unit === 'week') {
    date.setDate(date.getDate() - (amount * 7));
  } else if (unit === 'month') {
    date.setMonth(date.getMonth() - amount);
  }

  return date;
}

function extractWalkinDate(rawValue) {
  if (!rawValue) return null;
  const directDate = new Date(rawValue);
  if (!Number.isNaN(directDate.getTime())) return directDate;

  const text = rawValue.toString();
  const dateMatch = text.match(/\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\b/);
  if (!dateMatch) return null;

  const parsed = new Date(dateMatch[1]);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatWalkinDateLabel(dateValue) {
  if (!dateValue) return 'Indexed recently';
  return dateValue.toLocaleDateString('en-IN', {
    day: 'numeric',
    month: 'short',
    year: 'numeric'
  });
}

function scoreWalkinResult({ title = '', snippet = '', verifyUrl = '', role = '', location = '' }) {
  const haystack = `${title} ${snippet} ${verifyUrl}`.toLowerCase();
  const roleTokens = role.toLowerCase().split(/\s+/).filter(Boolean);
  const locationTokens = location.toLowerCase().split(/\s+/).filter(Boolean);

  let score = 0;
  if (haystack.includes('walk in')) score += 5;
  if (haystack.includes('walk-in')) score += 5;
  if (haystack.includes('interview')) score += 3;
  if (haystack.includes('hiring drive')) score += 4;
  if (haystack.includes('career')) score += 1;

  roleTokens.forEach((token) => {
    if (haystack.includes(token)) score += 2;
  });

  locationTokens.forEach((token) => {
    if (haystack.includes(token)) score += 2;
  });

  return score;
}

function scoreSerpApiJob({ title = '', company = '', description = '', location = '', role = '', publishedAt = '' }) {
  const haystack = `${title} ${company} ${description} ${location}`.toLowerCase();
  const roleTokens = role.toLowerCase().split(/\s+/).filter(Boolean);
  const locationTokens = location.toLowerCase().split(/\s+/).filter(Boolean);
  let score = 8;

  roleTokens.forEach((token) => {
    if (title.toLowerCase().includes(token)) score += 3;
    else if (haystack.includes(token)) score += 1;
  });

  locationTokens.forEach((token) => {
    if (haystack.includes(token)) score += 2;
  });

  if (publishedAt) {
    const publishedTime = new Date(publishedAt).getTime();
    if (!Number.isNaN(publishedTime)) {
      const ageHours = Math.max(0, (Date.now() - publishedTime) / (1000 * 60 * 60));
      if (ageHours <= 24) score += 5;
      else if (ageHours <= 24 * 7) score += 3;
      else if (ageHours <= 24 * 30) score += 1;
    }
  }

  return score;
}

function hasClosedJobSignals(value) {
  const haystack = (value || '').toString().toLowerCase();
  return ['expired', 'no longer accepting applications', 'applications closed', 'job closed', 'position filled', 'posting closed']
    .some((token) => haystack.includes(token));
}

function extractSerpApiUrl(job) {
  const applyOptions = Array.isArray(job.apply_options) ? job.apply_options : [];
  const relatedLinks = Array.isArray(job.related_links) ? job.related_links : [];
  const candidates = [
    ...applyOptions.map((option) => option?.link),
    ...relatedLinks.map((option) => option?.link),
    job.share_link,
    job.link
  ];

  return candidates.find((candidate) => {
    try {
      const parsed = new URL(candidate);
      return parsed.protocol === 'http:' || parsed.protocol === 'https:';
    } catch {
      return false;
    }
  }) || '';
}

function extractSerpApiPostedLabel(job) {
  const detected = job?.detected_extensions || {};
  const fromDetected = detected.posted_at || detected.schedule_type || '';
  if (fromDetected) return fromDetected;
  const extensions = Array.isArray(job?.extensions) ? job.extensions : [];
  return extensions.find((value) => /posted|ago|today|yesterday/i.test(value || '')) || '';
}

function buildSerpApiJobResult(job, index, role, fallbackLocation) {
  const verifyUrl = extractSerpApiUrl(job);
  const title = (job?.title || '').trim();
  const company = (job?.company_name || job?.company || '').trim() || 'Unknown Company';
  const location = (job?.location || fallbackLocation || '').trim() || fallbackLocation;
  const rawPostedLabel = extractSerpApiPostedLabel(job);
  const publishedAt = parseRelativePostedAt(rawPostedLabel);
  const description = (job?.description || '').trim()
    || (Array.isArray(job?.extensions) ? job.extensions.filter(Boolean).join(' • ') : '')
    || 'Live Google Jobs result.';
  const source = (job?.via || 'Google Jobs').toString().trim();
  const closedSignals = `${title} ${description} ${rawPostedLabel} ${source}`;

  if (!title || !verifyUrl || hasClosedJobSignals(closedSignals)) {
    return null;
  }

  return {
    id: `serpapi-${index}-${job?.job_id || Buffer.from(verifyUrl).toString('base64').slice(0, 10)}`,
    title,
    company,
    location,
    dateLabel: rawPostedLabel || formatWalkinDateLabel(publishedAt),
    description,
    verifyUrl,
    source,
    snippet: description,
    publishedAt: publishedAt?.toISOString() || '',
    score: scoreSerpApiJob({ title, company, description, location, role, publishedAt: publishedAt?.toISOString() || '' })
  };
}

async function fetchSerpApiJobs(role, location) {
  const apiKey = getSerpApiKey();
  if (!apiKey) return [];

  const response = await axios.get(SERPAPI_SEARCH_URL, {
    params: {
      engine: 'google_jobs',
      q: role,
      location,
      gl: 'in',
      hl: 'en',
      no_cache: 'true',
      api_key: apiKey
    },
    timeout: 15000
  });

  const jobs = Array.isArray(response.data?.jobs_results) ? response.data.jobs_results : [];

  return jobs
    .map((job, index) => buildSerpApiJobResult(job, index, role, location))
    .filter(Boolean)
    .sort((left, right) => {
      const scoreDelta = (right.score || 0) - (left.score || 0);
      if (scoreDelta !== 0) return scoreDelta;
      const rightDate = right.publishedAt ? new Date(right.publishedAt).getTime() : 0;
      const leftDate = left.publishedAt ? new Date(left.publishedAt).getTime() : 0;
      return rightDate - leftDate;
    })
    .slice(0, 12);
}

function normalizeWalkinProviderError(error) {
  const status = error?.response?.status;
  const message = error?.response?.data?.error || error?.message || 'Unable to fetch live jobs right now.';

  if (status === 401 || status === 403) {
    return 'SerpApi rejected the configured API key. Add a valid SERPAPI_API_KEY on the server.';
  }

  if (message.toLowerCase().includes('api key')) {
    return 'SerpApi rejected the configured API key. Add a valid SERPAPI_API_KEY on the server.';
  }

  return message;
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
    if (index >= 12) return false;

    const verifyUrl = $(element).find('link').first().text().trim();
    if (!verifyUrl || !isAllowedWalkinHost(verifyUrl)) {
      return;
    }

    const title = $(element).find('title').first().text().trim();
    const snippet = $(element).find('description').first().text().trim();
    const publishedAt = extractWalkinDate($(element).find('pubDate').first().text().trim() || snippet || title);
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
      dateLabel: formatWalkinDateLabel(publishedAt),
      description: snippet || 'Live listing discovered from indexed search results.',
      verifyUrl,
      source,
      snippet,
      publishedAt: publishedAt?.toISOString() || '',
      score: scoreWalkinResult({ title, snippet, verifyUrl, role: query, location })
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
    if (index >= 10) return false;

    const anchor = $(element).find('a.result__a').first();
    const rawHref = anchor.attr('href');
    const verifyUrl = decodeDuckDuckGoUrl(rawHref || '');

    if (!verifyUrl || !isAllowedWalkinHost(verifyUrl)) {
      return;
    }

    const title = anchor.text().trim();
    const snippet = $(element).find('.result__snippet').text().trim();
    const publishedAt = extractWalkinDate(snippet || title);
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
      dateLabel: formatWalkinDateLabel(publishedAt),
      description: snippet || 'Live listing discovered from a search result.',
      verifyUrl,
      source,
      snippet,
      publishedAt: publishedAt?.toISOString() || '',
      score: scoreWalkinResult({ title, snippet, verifyUrl, role: query, location })
    });
  });

  return results;
}

async function fetchLiveWalkins(role, location) {
  if (getSerpApiKey()) {
    const serpApiResults = await fetchSerpApiJobs(role, location);
    return {
      provider: 'serpapi-google-jobs',
      realtime: true,
      results: serpApiResults
    };
  }

  const now = new Date();
  const monthLabel = now.toLocaleDateString('en-IN', { month: 'long' });
  const yearLabel = `${now.getFullYear()}`;
  const queries = [
    `"walk in interview" ${role} ${location} site:linkedin.com/jobs`,
    `"walk in" ${role} ${location} site:linkedin.com`,
    `"walk in" ${role} ${location} site:foundit.in`,
    `"walk in interview" ${role} ${location} site:naukri.com`,
    `${role} hiring drive ${location} site:linkedin.com`,
    `"walk in interview" ${role} ${location} ${monthLabel} ${yearLabel} site:naukri.com`,
    `"walk in" ${role} ${location} ${monthLabel} ${yearLabel} site:foundit.in`
  ];

  const settled = await Promise.allSettled(queries.map((query) => searchWalkinQuery(query, location)));
  const deduped = new Map();

  settled.forEach((result) => {
    if (result.status !== 'fulfilled') return;

    result.value.forEach((item) => {
      if (item.score < 4) return;
      if (!deduped.has(item.verifyUrl)) {
        deduped.set(item.verifyUrl, item);
      }
    });
  });

  const liveResults = Array.from(deduped.values())
    .sort((left, right) => {
      const scoreDelta = (right.score || 0) - (left.score || 0);
      if (scoreDelta !== 0) return scoreDelta;
      const rightDate = right.publishedAt ? new Date(right.publishedAt).getTime() : 0;
      const leftDate = left.publishedAt ? new Date(left.publishedAt).getTime() : 0;
      return rightDate - leftDate;
    })
    .slice(0, 12);
  return {
    provider: 'search-fallback',
    realtime: false,
    results: liveResults.length > 0 ? liveResults : createSearchFallback(role, location)
  };
}

export function createApp() {
  const app = express();
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';

  const isProduction = Boolean(process.env.VERCEL || process.env.NODE_ENV === 'production');
  const allowedOrigins = [
    process.env.FRONTEND_URL || 'http://localhost:5173',
    'https://ai-job-applicator.vercel.app'
  ];

  app.use(cors({
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
      callback(new Error('Not allowed by CORS'));
    },
    credentials: true
  }));
  app.use(express.json());
  app.use(cookieParser());
  app.use((req, _res, next) => {
    req.user = getUserFromToken(req);
    next();
  });
  app.use(passport.initialize());

  // Passport only used to facilitate the OAuth redirect/callback — user is then stored in a JWT cookie
  passport.serializeUser((user, done) => done(null, user));
  passport.deserializeUser((user, done) => done(null, user));

  const upload = multer({ dest: process.env.VERCEL ? '/tmp' : 'uploads/' });

  const loginClientId = process.env.GOOGLE_LOGIN_CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
  const loginClientSecret = process.env.GOOGLE_LOGIN_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;
  const loginCallbackUrl = process.env.GOOGLE_LOGIN_CALLBACK_URL || process.env.GOOGLE_CALLBACK_URL || 'https://ai-job-applicator.vercel.app/auth/google/callback';
  const gmailClientId = process.env.GOOGLE_GMAIL_CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
  const gmailClientSecret = process.env.GOOGLE_GMAIL_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;
  const gmailUsesSameClient = Boolean(gmailClientId && gmailClientId === loginClientId);
  const gmailCallbackUrl = process.env.GOOGLE_GMAIL_CALLBACK_URL
    || (gmailUsesSameClient ? loginCallbackUrl : 'https://ai-job-applicator.vercel.app/auth/google-gmail/callback');
  const loginConfigured = Boolean(loginClientId && loginClientId !== 'your_google_client_id_here' && loginClientSecret);
  const gmailConfigured = Boolean(gmailClientId && gmailClientId !== 'your_google_client_id_here' && gmailClientSecret);

  if (loginConfigured) {
    passport.use('google-login', new GoogleStrategy({
      clientID: loginClientId,
      clientSecret: loginClientSecret,
      callbackURL: loginCallbackUrl,
      passReqToCallback: true
    }, (req, accessToken, refreshToken, profile, done) => done(null, buildOAuthUser(req.user, profile, accessToken, refreshToken))));

    app.get('/auth/google', passport.authenticate('google-login', {
      scope: ['profile', 'email'],
      state: 'login',
      session: false
    }));
  } else {
    console.warn('Google login OAuth keys missing. Google sign-in disabled.');
    app.get('/auth/google', (req, res) => res.status(501).json({ error: 'OAuth not configured.' }));
  }

  if (gmailConfigured) {
    passport.use('google-gmail', new GoogleStrategy({
      clientID: gmailClientId,
      clientSecret: gmailClientSecret,
      callbackURL: gmailCallbackUrl,
      passReqToCallback: true
    }, (req, accessToken, refreshToken, profile, done) => done(null, buildOAuthUser(req.user, profile, accessToken, refreshToken))));

    app.get('/auth/google-gmail', passport.authenticate('google-gmail', {
      scope: ['profile', 'email', 'https://www.googleapis.com/auth/gmail.send'],
      accessType: 'offline',
      prompt: 'consent',
      includeGrantedScopes: true,
      state: 'gmail',
      session: false
    }));
  } else {
    console.warn('Gmail OAuth keys missing. Gmail sender connect disabled.');
    app.get('/auth/google-gmail', (req, res) => res.status(501).json({ error: 'Gmail OAuth not configured.' }));
  }

  const finishOauthRedirect = (req, res) => {
    setAuthCookie(res, req.user);
    res.redirect(frontendUrl);
  };

  app.get('/auth/google/callback', (req, res, next) => {
    const isGmailFlow = req.query.state === 'gmail';

    if (isGmailFlow && !gmailConfigured) {
      return res.status(501).json({ error: 'Gmail OAuth not configured.' });
    }

    if (!isGmailFlow && !loginConfigured) {
      return res.status(501).json({ error: 'OAuth not configured.' });
    }

    const strategy = isGmailFlow ? 'google-gmail' : 'google-login';
    const failureRedirect = frontendUrl + '?' + (isGmailFlow ? 'gmailError' : 'loginError') + '=oauth';
    passport.authenticate(strategy, { failureRedirect, session: false })(req, res, next);
  }, finishOauthRedirect);

  if (gmailConfigured) {
    app.get('/auth/google-gmail/callback',
      passport.authenticate('google-gmail', { failureRedirect: frontendUrl + '?gmailError=oauth', session: false }),
      finishOauthRedirect
    );
  } else {
    app.get('/auth/google-gmail/callback', (req, res) => res.status(501).json({ error: 'Gmail OAuth not configured.' }));
  }
  app.get('/api/user', (req, res) => {
    res.json(buildPublicUser(req.user));
  });

  app.get('/api/health', (req, res) => {
    res.json({
      ok: true,
      timestamp: new Date().toISOString(),
      mail: buildMailCapability(req.user),
      jobs: {
        provider: getSerpApiKey() ? 'serpapi-google-jobs' : 'search-fallback',
        realtime: Boolean(getSerpApiKey())
      }
    });
  });

  app.post('/auth/local-login', (req, res) => {
    const name = (req.body?.name || '').toString().trim();
    const email = (req.body?.email || '').toString().trim().toLowerCase();

    if (!name || !email) return res.status(400).json({ error: 'Name and email are required.' });
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) return res.status(400).json({ error: 'Enter a valid email address.' });

    const user = { id: `local:${email}`, displayName: name, email, provider: 'local' };
    setAuthCookie(res, user);
    res.json({ success: true, user: buildPublicUser(user) });
  });

  app.get('/auth/logout', (req, res) => {
    clearAuthCookie(res);
    res.redirect(frontendUrl);
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
    res.set('Cache-Control', 'no-store');

    try {
      const { results, provider, realtime } = await fetchLiveWalkins(role, location);
      res.json({
        walkins: results,
        generatedAt: new Date().toISOString(),
        location,
        role,
        provider,
        realtime,
        liveResultsFound: results.length > 0
      });
    } catch (error) {
      res.status(500).json({
        success: false,
        error: normalizeWalkinProviderError(error),
        details: error.message
      });
    }
  });

  app.post('/api/send-cold-emails', upload.single('resume'), async (req, res) => {
    try {
      const { contacts } = req.body;
      const parsedContacts = JSON.parse(contacts || '[]');
      const { senderEmail, transporter } = createMailTransport(req);
      const replyTo = resolveReplyTo(req);

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
            replyTo,
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
      const { senderEmail, transporter } = createMailTransport(req);
      const replyTo = resolveReplyTo(req);

      if (!to || !subject || !body) {
        return res.status(400).json({ success: false, error: 'Missing required email fields.' });
      }

      const attachments = req.files ? req.files.map((file) => ({ filename: file.originalname, path: file.path })) : [];
      await transporter.sendMail({ from: senderEmail, to, subject, text: body, replyTo, attachments });

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
