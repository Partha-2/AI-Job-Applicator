import express from 'express';
import nodemailer from 'nodemailer';
import multer from 'multer';
import cors from 'cors';
import fs from 'fs';
import passport from 'passport';
import { Strategy as GoogleStrategy } from 'passport-google-oauth20';
import session from 'express-session';
import axios from 'axios';
import * as cheerio from 'cheerio';
import dotenv from 'dotenv';
import path from 'path';

dotenv.config();

const app = express();
app.use(cors({ origin: 'http://localhost:5173', credentials: true }));
app.use(express.json());
app.use(session({
  secret: process.env.SESSION_SECRET || 'job-applicator-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false } // Set to true if using HTTPS
}));
app.use(passport.initialize());
app.use(passport.session());

const upload = multer({ dest: 'uploads/' });

// Passport Serialiazation
passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

// Google Strategy Configuration
if (process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_ID !== 'your_google_client_id_here') {
  passport.use(new GoogleStrategy({
    clientID: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    callbackURL: process.env.GOOGLE_CALLBACK_URL || 'http://localhost:3000/auth/google/callback'
  },
  (accessToken, refreshToken, profile, done) => {
    return done(null, profile);
  }));

  // OAuth Routes
  app.get('/auth/google', passport.authenticate('google', { scope: ['profile', 'email'] }));

  app.get('/auth/google/callback', 
    passport.authenticate('google', { failureRedirect: 'http://localhost:5173/login?error=true' }),
    (req, res) => {
      res.redirect('http://localhost:5173');
    }
  );

  app.get('/auth/logout', (req, res, next) => {
    req.logout((err) => {
      if (err) return next(err);
      res.redirect('http://localhost:5173');
    });
  });
} else {
  console.warn('⚠️ Google OAuth keys missing or default. Authentication features disabled.');
  app.get('/auth/google', (req, res) => res.status(501).json({ error: 'OAuth not configured. Please add keys to .env' }));
}

// Applied Jobs Tracker (In-Memory for now, can be File/DB later)
const userAppliedJobs = {}; 

app.get('/api/applied-jobs', (req, res) => {
  if (!req.user) return res.json([]);
  const email = req.user.emails[0].value;
  res.json(userAppliedJobs[email] || []);
});

app.post('/api/applied-jobs', (req, res) => {
  if (!req.user) return res.status(401).json({ error: 'Unauthorized' });
  const email = req.user.emails[0].value;
  const { jobId } = req.body;
  if (!userAppliedJobs[email]) userAppliedJobs[email] = [];
  if (!userAppliedJobs[email].includes(jobId)) userAppliedJobs[email].push(jobId);
  res.json({ success: true });
});

// WALK-IN API with Verification Links
app.get('/api/walkins', (req, res) => {
  const { location, role } = req.query;
  const today = new Date().toLocaleDateString();
  const loc = location || 'Bangalore';
  const r = role || 'Java Developer';
  
  const mockWalkins = [
    {
      id: 'w1',
      company: 'Tech Mahindra',
      role: r,
      location: loc,
      date: today + ' - Tomorrow',
      time: '10:00 AM - 4:00 PM',
      address: 'Phase 2, Electronic City, Bangalore',
      description: 'Mega Walk-in Drive for Immediate Joiners.',
      verifyUrl: `https://www.naukri.com/walkin-jobs-in-${loc.toLowerCase()}?k=${encodeURIComponent(r)}`
    },
    {
      id: 'w2',
      company: 'TCS',
      role: r,
      location: loc,
      date: today,
      time: '9:30 AM - 2:00 PM',
      address: 'Sahyadri Park, Hinjewadi Phase 3',
      description: 'Looking for candidates with 1-3 years experience.',
      verifyUrl: `https://www.linkedin.com/jobs/search/?keywords=${encodeURIComponent(r + ' walk-in ' + loc)}`
    }
  ];
  res.json({ walkins: mockWalkins });
});


// COLD EMAIL AUTOMATION API
app.post('/api/send-cold-emails', upload.single('resume'), async (req, res) => {
  try {
    const { email, password, contacts } = req.body;
    const parsedContacts = JSON.parse(contacts);

    if (!email || !password || !parsedContacts) {
      return res.status(400).json({ success: false, error: "Missing required fields." });
    }

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: email, pass: password }
    });

    const results = [];
    for (let contact of parsedContacts) {
      try {
        await transporter.sendMail({
          from: email,
          to: contact.to,
          subject: contact.subject,
          text: contact.body,
          attachments: req.file ? [{ filename: req.file.originalname, path: req.file.path }] : []
        });
        results.push({ email: contact.to, status: 'Sent' });
      } catch (err) {
        results.push({ email: contact.to, status: 'Failed', error: err.message });
      }
    }
    if (req.file) fs.unlinkSync(req.file.path);
    res.json({ success: true, results: results });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// SINGLE EMAIL API
app.post('/api/send-single-email', upload.array('attachments'), async (req, res) => {
  try {
    const { fromEmail, fromPass, to, subject, body } = req.body;
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: fromEmail, pass: fromPass }
    });

    const attachments = req.files ? req.files.map(f => ({ filename: f.originalname, path: f.path })) : [];

    await transporter.sendMail({ from: fromEmail, to, subject, text: body, attachments });
    
    // Cleanup
    if (req.files) req.files.forEach(f => fs.unlinkSync(f.path));

    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// EMAIL SCRAPER API
app.post('/api/scrape-emails', async (req, res) => {
  try {
    const { url } = req.body;
    if (!url) return res.status(400).json({ error: 'URL is required' });

    const response = await axios.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
    const $ = cheerio.load(response.data);
    const text = $('body').text();
    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    const emails = Array.from(new Set(text.match(emailRegex) || []));

    res.json({ success: true, emails });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Job Applicator V4 API running on http://localhost:${PORT}`);
});
