'use strict';

// ─────────────────────────────────────────────
//  OUTSEEK MAIL — script.js
// ─────────────────────────────────────────────

// ── Helpers ──────────────────────────────────
function uid() { return Math.random().toString(36).slice(2,10); }

function formatTime(ts) {
  const now  = new Date();
  const date = new Date(ts);
  const diff = now - date;
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const yestStart  = todayStart - 86400000;
  if (date.getTime() >= todayStart) {
    return date.getHours().toString().padStart(2,'0') + ':' + date.getMinutes().toString().padStart(2,'0');
  } else if (date.getTime() >= yestStart) {
    return 'Yesterday';
  } else {
    return date.toLocaleDateString('en-GB', { day:'2-digit', month:'short' });
  }
}

function avatarColor(initials) {
  const colors = ['#0078d4','#107c41','#8764b8','#c43e1c','#005b70','#1b6ec2','#7a7574','#0e7a0d'];
  let h = 0;
  for (let i = 0; i < initials.length; i++) h = (h * 31 + initials.charCodeAt(i)) % colors.length;
  return colors[h];
}

function initials(name) {
  const parts = name.trim().split(' ');
  return (parts[0][0] + (parts[parts.length-1][0] || '')).toUpperCase();
}

function htmlEscape(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── User profile ──────────────────────────────
function getUserFullName() {
  return localStorage.getItem('outseek_user_name') || 'David Jenner';
}
function getUserFirstName() {
  return getUserFullName().split(' ')[0];
}
function getUserInitials() {
  return initials(getUserFullName());
}
// Replace the default "David" in email body text with the user's first name,
// but leave sender names like "David Patel" or "David Jenner" untouched.
function personaliseText(text) {
  const first = getUserFirstName();
  if (first === 'David') return text;
  return text.replace(/\bDavid\b(?!\s+(?:Patel|Jenner))/g, first);
}
function updateUserNameUI() {
  const full = getUserFullName();
  const ini  = getUserInitials();
  const tb = document.querySelector('.title-bar-text');
  if (tb) tb.textContent = 'Outseek Mail \u2014 Inbox \u2014 ' + full;
  const sbSpans = document.querySelectorAll('.status-bar span');
  if (sbSpans.length >= 2) {
    sbSpans[1].textContent = getUserFirstName().toLowerCase() + '.jenner@meridiangroup.co.uk';
  }
  const av = document.getElementById('userAvatar');
  if (av) {
    const img = av.querySelector('img');
    const hasPic = img && img.style.display !== 'none' && img.complete && img.naturalWidth > 0;
    if (!hasPic) av.textContent = ini;
    av.title = full;
  }
}
window.setUserName = function(name) {
  name = name.trim();
  if (!name) return;
  localStorage.setItem('outseek_user_name', name);
  updateUserNameUI();
};

// ── Email store ───────────────────────────────
const Store = {
  _emails: [],
  load() {
    try { this._emails = JSON.parse(localStorage.getItem('outseek_emails') || '[]'); }
    catch(e) { this._emails = []; }
    return this._emails;
  },
  save() { localStorage.setItem('outseek_emails', JSON.stringify(this._emails)); },
  add(e) { this._emails.unshift(e); this.save(); },
  remove(id) { this._emails = this._emails.filter(e => e.id !== id); this.save(); },
  get(id) { return this._emails.find(e => e.id === id); },
  markRead(id) { const e = this.get(id); if(e){ e.read=true; this.save(); } },
  markUnread(id) { const e = this.get(id); if(e){ e.read=false; this.save(); } },
  toggleFlag(id) { const e = this.get(id); if(e){ e.flagged=!e.flagged; this.save(); } },
  unreadCount() { return this._emails.filter(e => !e.read && e.folder==='inbox').length; }
};

// ── Email database (realistic workplace) ─────
const SENDERS = [
  { name:'Lucy Harrison',    email:'l.harrison@meridiangroup.co.uk',    role:'Head of Strategy' },
  { name:'Sophie Bennett',   email:'s.bennett@meridiangroup.co.uk',     role:'PA to Managing Director' },
  { name:'Alex Kumar',       email:'a.kumar@meridiangroup.co.uk',       role:'Senior Project Manager' },
  { name:'Tom Wilson',       email:'t.wilson@meridiangroup.co.uk',      role:'Analyst' },
  { name:'Sarah Mitchell',   email:'sarah.mitchell@meridiangroup.co.uk',role:'HR Business Partner' },
  { name:'Richard Thompson', email:'r.thompson@meridiangroup.co.uk',    role:'Managing Director' },
  { name:'Emma Clarke',      email:'emma.clarke@meridiangroup.co.uk',   role:'Head of Digital' },
  { name:'David Patel',      email:'d.patel@meridiangroup.co.uk',       role:'Finance Director' },
  { name:'Chris Newman',     email:'c.newman@meridiangroup.co.uk',      role:'Business Development Director' },
  { name:'Kate Davies',      email:'k.davies@meridiangroup.co.uk',      role:'Marketing Manager' },
  { name:'IT Support',       email:'it.support@meridiangroup.co.uk',    role:'IT Helpdesk' },
  { name:'Meridian HR',      email:'noreply.hr@meridiangroup.co.uk',    role:'HR System' },
];

function sender(n) { return SENDERS.find(s => s.name === n) || SENDERS[0]; }

// Build the initial email set (realistic inbox from Day 1)
function buildInitialEmails() {
  const now = Date.now();
  const H = 3600000, M = 60000;

  return [
    {
      id: uid(), folder:'inbox', read:false, flagged:false,
      from: sender('Lucy Harrison'),
      subject: 'Re: Project Phoenix — Weekly Update',
      preview: 'Thanks for the update. The timeline looks good to me...',
      ts: now - 25*M,
      attachments: [],
      body: buildBody(
        sender('Lucy Harrison'),
        `Hi David,\n\nThanks for the update — the timeline looks good to me. One thing I'd flag is the dependency on the API integration. Sam's team mentioned there might be a two-day slip on their end. Can you chase that this morning?\n\nAlso, can you make sure the client deck is polished before Thursday? Richard will want to review it by noon Wednesday at the latest.\n\nOther than that, really solid progress this week. Keep it up!\n\nCheers,\nLucy`,
        `From: David Jenner\nSent: Monday, 28 October 2024 17:45\nTo: Lucy Harrison; Alex Kumar; Tom Wilson\nSubject: Project Phoenix — Weekly Update\n\nAll,\n\nQuick update from my end this week:\n- Completed data migration for Phase 2\n- Stakeholder sign-off received on requirements doc\n- UAT scheduled for w/c 4 November\n\nOverall RAG: Amber (API dependency outstanding)\n\nDavid`
      )
    },
    {
      id: uid(), folder:'inbox', read:false, flagged:false,
      from: sender('Sophie Bennett'),
      subject: 'All-Hands Meeting — Thursday 3:00 PM',
      preview: 'Richard will be sharing Q2 results and some exciting news...',
      ts: now - 62*M,
      attachments: [{ name:'All_Hands_Agenda.docx', size:'38 KB', type:'word' }],
      body: buildBody(
        sender('Sophie Bennett'),
        `Hi All,\n\nYou're invited to the Q2 All-Hands Meeting:\n\n📅  Thursday, 31 October 2024\n🕒  15:00 – 16:00\n📍  Board Room A / Teams link below\n\nhttps://teams.microsoft.com/l/meetup/meridiangroup/all-hands\n\nRichard will be walking us through Q2 results, celebrating some great wins, and sharing the company's direction for H2. There will also be time for open Q&A.\n\nPlease block your diaries — attendance is expected for all staff. If you're working remotely, the Teams link above will be live from 14:55.\n\nAgenda attached.\n\nKind regards,\nSophie Bennett\nPA to Managing Director`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:false, flagged:true,
      from: sender('Alex Kumar'),
      subject: 'ACTION REQUIRED: Sprint 4 Sign-off by EOD',
      preview: 'I need your approval on the attached sign-off document before close of business...',
      ts: now - 90*M,
      attachments: [{ name:'Sprint4_Signoff.xlsx', size:'112 KB', type:'excel' }],
      body: buildBody(
        sender('Alex Kumar'),
        `Hi David,\n\nHope you're well. I need your formal sign-off on the Sprint 4 deliverables before we can cut the release branch tonight.\n\nThe attached spreadsheet details:\n- All 14 completed story points\n- Outstanding defects (2 minor, non-blocking)\n- UAT results from last Friday\n\nPlease review and reply with your approval, or flag any blockers, by 17:00 today.\n\nIf you have questions give me a call on ext. 2214.\n\nThanks,\nAlex`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Tom Wilson'),
      subject: 'Lunch today? Pret or Wahaca?',
      preview: 'Heading out around 12:30 — fancy joining?',
      ts: now - 3*H - 15*M,
      attachments: [],
      body: buildBody(
        sender('Tom Wilson'),
        `Hey,\n\nHeading out around 12:30 — Pret or Wahaca? Kate's coming too.\n\nLet me know!\n\nTom`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('IT Support'),
      subject: 'Your password expires in 7 days',
      preview: 'Please update your Meridian Group network password before 4 November...',
      ts: now - 20*H,
      attachments: [],
      body: buildBody(
        sender('IT Support'),
        `Hi David,\n\nThis is an automated reminder from Meridian IT.\n\nYour network password is due to expire on 04 November 2024. Please update it before that date to avoid being locked out of your account.\n\nTo change your password:\n1. Press Ctrl + Alt + Delete\n2. Select "Change a password"\n3. Follow the on-screen instructions\n\nPasswords must be at least 10 characters and include a mix of uppercase, lowercase, numbers and symbols.\n\nIf you need assistance, raise a ticket at helpdesk.meridiangroup.co.uk or call ext. 5000.\n\nMeridian IT Support`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Richard Thompson'),
      subject: 'Meridian Group — Q2 Strategy Update',
      preview: 'I wanted to share a few thoughts on where we are as a business heading into H2...',
      ts: now - 22*H,
      attachments: [{ name:'Q2_Strategy_Deck.pdf', size:'2.4 MB', type:'pdf' }],
      body: buildBody(
        sender('Richard Thompson'),
        `Hi everyone,\n\nI wanted to share a few thoughts on where we are as a business heading into the second half of the year.\n\nQ2 was a strong quarter. Revenue came in at £4.2M — 8% ahead of forecast — and we secured three major contract renewals, including the Ashworth Group deal that many of you worked incredibly hard on. Well done to everyone involved.\n\nLooking ahead, our priorities for Q3 and Q4 are:\n\n1.  Grow our Digital practice — Emma's team has a strong pipeline and we want to resource it properly.\n2.  Operational efficiency — David Patel's finance review identified savings opportunities we'll be acting on.\n3.  Talent retention — Sarah is leading a compensation benchmarking exercise and we'll share findings next month.\n\nI've attached the full Q2 strategy deck. Please read it before Thursday's All-Hands where we'll discuss it as a group.\n\nAs ever, my door is open if you want to talk anything through.\n\nBest,\nRichard\n\nRichard Thompson\nManaging Director, Meridian Group`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Kate Davies'),
      subject: 'Brand refresh materials — please review by EOD',
      preview: "I've shared the updated brand guidelines in the Teams channel. Can you take a look and send any comments...",
      ts: now - 25*H,
      attachments: [{ name:'Meridian_Brand_Guidelines_v3.pdf', size:'8.7 MB', type:'pdf' }],
      body: buildBody(
        sender('Kate Davies'),
        `Hi David,\n\nI've shared the updated brand guidelines (v3) in the #brand-refresh Teams channel and attached the PDF here too.\n\nThe main changes from v2 are:\n- Updated primary colour palette (Midnight Navy replacing the old dark blue)\n- New typeface: Neue Haas Grotesk replacing Helvetica Neue\n- Refreshed photography guidelines and approved image library\n- New PowerPoint and Word templates (on SharePoint)\n\nCould you have a look and send any comments by EOD today? We're hoping to sign off with the agency tomorrow morning.\n\nThanks so much!\n\nKate\n\nKate Davies | Marketing Manager\nMeridian Group`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('David Patel'),
      subject: 'Expense Claim Approved — £342.50',
      preview: 'Your expense claim for October travel has been approved and will be paid in the next payroll run...',
      ts: now - 27*H,
      attachments: [],
      body: buildBody(
        sender('David Patel'),
        `Hi David,\n\nYour expense claim (Ref: EXP-2024-0892) for £342.50 has been approved.\n\nBreakdown:\n  Rail travel (London–Manchester return)    £189.00\n  Taxis (client site)                       £48.50\n  Client lunch (2 covers)                   £105.00\n  ──────────────────────────────────────────────\n  Total                                     £342.50\n\nThis will be included in the next payroll run on 31 October and paid directly to your bank account.\n\nIf you have any questions, contact finance@meridiangroup.co.uk.\n\nKind regards,\nDavid Patel\nFinance Director`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Sarah Mitchell'),
      subject: 'Mandatory Training: Data Protection Refresher — Due by 15 Nov',
      preview: 'As part of our annual compliance programme, all staff must complete the Data Protection refresher...',
      ts: now - 2*24*H,
      attachments: [],
      body: buildBody(
        sender('Sarah Mitchell'),
        `Hi David,\n\nAs part of our annual compliance programme, all staff are required to complete the Data Protection & GDPR Refresher module by 15 November 2024.\n\nThe module takes approximately 25 minutes and is available via the Learning Hub:\nhttps://learn.meridiangroup.co.uk/data-protection-2024\n\nUpon completion you'll receive a certificate which will be logged automatically against your HR record.\n\nThis is mandatory — failure to complete by the deadline will be escalated to line managers.\n\nIf you have any issues accessing the module, please contact it.support@meridiangroup.co.uk.\n\nThanks,\nSarah Mitchell\nHR Business Partner`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Emma Clarke'),
      subject: 'Fwd: Client Feedback — TechNova Partnership',
      preview: 'Forwarding this from Priya at TechNova — genuinely brilliant feedback, well done team...',
      ts: now - 2*24*H - 4*H,
      attachments: [],
      body: buildBody(
        sender('Emma Clarke'),
        `Hi David,\n\nJust forwarding this — I don't think feedback gets much better than this. Really well done on the TechNova project, you and the team should be very proud.\n\nEmma`,
        `---------- Forwarded message ---------\nFrom: Priya Sharma <p.sharma@technova.io>\nDate: Friday, 25 October 2024 at 16:14\nTo: Emma Clarke <emma.clarke@meridiangroup.co.uk>\nSubject: Re: Digital Transformation Project — Final Delivery\n\nEmma,\n\nI just wanted to take a moment to say how impressed we've been with the Meridian team throughout this engagement. The quality of work, the responsiveness, and the genuine care for our outcomes has been exceptional.\n\nIn particular, David's leadership of the data workstream made what could have been a painful migration completely seamless. Please pass on our thanks.\n\nWe'll absolutely be recommending Meridian to our network and hope to work together again in the new year.\n\nWarm regards,\nPriya Sharma\nCTO, TechNova`
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Chris Newman'),
      subject: 'New Lead — Hargreaves & Stone interested in a proposal',
      preview: "Just got off the phone with their procurement lead. They're actively looking and want a proposal by mid-November...",
      ts: now - 3*24*H,
      attachments: [],
      body: buildBody(
        sender('Chris Newman'),
        `Hi David,\n\nJust got off the phone with James Hargreaves at Hargreaves & Stone (mid-size property management firm, ~£80M turnover). They're actively looking for a digital transformation partner and specifically mentioned our work with Ashworth Group as a reference.\n\nThey want a proposal by 14 November covering:\n- CRM migration (Salesforce)\n- Reporting & analytics overhaul\n- Change management support\n\nInitial contract value could be £600–800K over 18 months.\n\nCan we get in a room this week to scope it out? I was thinking Thursday morning before the All-Hands.\n\nCheers,\nChris`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Alex Kumar'),
      subject: 'Sprint Planning — Agenda & Prep',
      preview: "I've shared the backlog ahead of tomorrow's sprint planning. Please review your assigned stories before the session...",
      ts: now - 3*24*H - 2*H,
      attachments: [{ name:'Sprint5_Backlog_Draft.xlsx', size:'88 KB', type:'excel' }],
      body: buildBody(
        sender('Alex Kumar'),
        `Hi team,\n\nAhead of tomorrow's Sprint 5 planning session (10:00, Meeting Room 2), please could you review the attached backlog draft and come prepared with your capacity estimates.\n\nA few things to flag:\n\n- Story PH-142 (Reporting Dashboard) has been broken into 3 sub-tasks — see tab 2\n- PH-138 is blocked pending sign-off from David (see my earlier email)\n- New story PH-151 has been added at Richard's request — details in the sheet\n\nSession agenda:\n  10:00  Velocity review (15 min)\n  10:15  Backlog refinement (30 min)\n  10:45  Sprint commitment (30 min)\n  11:15  Close\n\nSee you tomorrow.\n\nAlex`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Sarah Mitchell'),
      subject: 'Welcome to Meridian Group — Your First Week',
      preview: 'Welcome aboard! Here is everything you need to get started at Meridian...',
      ts: now - 5*24*H,
      attachments: [
        { name:'New_Starter_Guide.pdf', size:'1.2 MB', type:'pdf' },
        { name:'IT_Access_Form.docx', size:'44 KB', type:'word' },
      ],
      body: buildBody(
        sender('Sarah Mitchell'),
        `Hi David,\n\nWelcome to Meridian Group — we're really pleased to have you on board!\n\nHere's a quick summary of what to expect in your first week:\n\nMonday\n  09:00  Meet your line manager Lucy Harrison (her desk is on the 3rd floor, Strategy team)\n  10:30  IT induction with James Anderson (ext. 5001)\n  14:00  Building tour & security pass collection (Reception)\n\nTuesday\n  09:30  HR induction (Meeting Room 4) — please bring your Right to Work documents\n  11:00  Finance & expenses walkthrough with David Patel\n\nWednesday–Friday\n  You'll be joining the Project Phoenix team. Alex Kumar will brief you on Monday afternoon.\n\nUseful links:\n  Intranet: intranet.meridiangroup.co.uk\n  IT Helpdesk: ext. 5000 or helpdesk.meridiangroup.co.uk\n  HR Portal: hr.meridiangroup.co.uk\n\nDon't hesitate to reach out if you need anything at all.\n\nWarm regards,\nSarah Mitchell\nHR Business Partner\nMeridian Group`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Lucy Harrison'),
      subject: '1:1 Tomorrow — Agenda',
      preview: 'Just a quick note ahead of our 1:1 tomorrow. A few things I want to cover...',
      ts: now - 4*24*H,
      attachments: [],
      body: buildBody(
        sender('Lucy Harrison'),
        `Hi David,\n\nJust a quick note ahead of our 1:1 tomorrow (Friday, 15:00, my office).\n\nTopics I'm planning to cover:\n\n1.  Project Phoenix update — particularly the API dependency and Sprint 4 sign-off\n2.  Your development plan — I'd like to revisit the goals we set in September\n3.  Q3 workload & capacity going into November\n4.  Any concerns/blockers you want to raise\n\nFeel free to add anything else to the list before we meet. It's your time too.\n\nSee you tomorrow.\n\nLucy`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('David Patel'),
      subject: 'Month-End Reports — Due by Friday 17:00',
      preview: 'Friendly reminder that all month-end variance reports must be submitted by Friday afternoon...',
      ts: now - 4*24*H - 5*H,
      attachments: [{ name:'Monthly_Report_Template_Oct24.xlsx', size:'156 KB', type:'excel' }],
      body: buildBody(
        sender('David Patel'),
        `Hi all,\n\nFriendly reminder that October month-end reports are due by 17:00 this Friday (1 November).\n\nPlease use the attached updated template — there are two new columns in the Variance tab that reflect the board's reporting requirements.\n\nKey dates:\n  31 Oct (Friday)  17:00 — All reports submitted\n  4 Nov  (Monday)  09:00 — Finance consolidation\n  5 Nov  (Tuesday) 14:00 — Board pack sign-off\n\nIf you're going to miss the deadline please let me know ASAP so we can make arrangements.\n\nThanks,\nDavid Patel\nFinance Director`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Tom Wilson'),
      subject: "Fwd: Friday Drinks Quiz — you coming?",
      preview: "Forwarding from Kate — sounds like a good one, reckon you're in?",
      ts: now - 5*24*H - 3*H,
      attachments: [],
      body: buildBody(
        sender('Tom Wilson'),
        `Hey,\n\nForwarding from Kate — sounds like a good one, are you in?\n\nTom`,
        `---------- Forwarded message ---------\nFrom: Kate Davies\nDate: Wednesday, 23 October at 17:34\nSubject: Friday Drinks Quiz\n\nHi all,\n\nWho's up for the quiz at The Crown on Friday evening? Starts at 19:00. They do great burgers too.\n\nTeams of 4 max — currently it's me, Priya, James and Tom. One more spot if anyone wants in!\n\nKate`
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('Richard Thompson'),
      subject: 'Congratulations — Meridian wins Best Workplace Award 2024',
      preview: "I'm delighted to share that Meridian Group has been named one of the UK's Best Workplaces...",
      ts: now - 6*24*H,
      attachments: [],
      body: buildBody(
        sender('Richard Thompson'),
        `Hi everyone,\n\nI am absolutely delighted to share some wonderful news — Meridian Group has been named one of the UK's Best Workplaces for 2024 by Great Place to Work.\n\nThis award is based on anonymous employee surveys and reflects the culture, trust and pride that every single one of you brings to work each day. This is your achievement.\n\nWe'll be celebrating at the Christmas party in December, but in the meantime I wanted to say a heartfelt thank you. It genuinely means a great deal.\n\nMore details to follow from HR.\n\nWith gratitude,\nRichard\n\nRichard Thompson\nManaging Director, Meridian Group`,
        null
      )
    },
    {
      id: uid(), folder:'inbox', read:true, flagged:false,
      from: sender('IT Support'),
      subject: 'IT Service Desk — Ticket #4821 Resolved',
      preview: 'Your request to add a shared mailbox has been completed...',
      ts: now - 6*24*H - 2*H,
      attachments: [],
      body: buildBody(
        sender('IT Support'),
        `Hi David,\n\nYour service desk ticket #4821 has been resolved.\n\nRequest: Add shared mailbox 'projectphoenix@meridiangroup.co.uk' to your Outseek Mail profile\nResolution: Shared mailbox added and permissions granted\nResolved by: James Anderson\nResolved on: 24 October 2024 14:30\n\nThe shared mailbox should now appear in your left-hand folder panel. If you don't see it, please close and reopen Outseek Mail.\n\nIf your issue has not been resolved or you have further questions, please reply to this email to re-open the ticket.\n\nMeridian IT Support`,
        null
      )
    },
  ];
}

// ── Email body builder ────────────────────────
function buildBody(fromObj, bodyText, quotedText) {
  return {
    from: fromObj,
    text: bodyText,
    quoted: quotedText || null
  };
}

// ── Pool of new incoming emails (arrive over time) ──
const NEW_EMAIL_POOL = [
  { from:'Lucy Harrison',    subject:'Can you send me the latest dashboard?',
    preview:'When you get a chance — just need the Oct version for a client...',
    text:`Hi David,\n\nWhen you get a chance, could you send over the latest version of the reporting dashboard? I need the October snapshot for a client call this afternoon.\n\nNo rush, but before 2pm if possible!\n\nThanks,\nLucy`, atts:[] },
  { from:'Tom Wilson',       subject:'Re: Sprint 4 retrospective notes',
    preview:'Just added my comments to the shared doc — take a look when you get a moment...',
    text:`Hey,\n\nAdded my retrospective notes to the shared doc. A couple of things I wanted to flag that I didn't get a chance to raise in the session:\n\n- The daily standup format isn't working brilliantly for me, worth a chat?\n- The shared test environment keeps going down — can we raise with IT?\n\nNothing urgent, just food for thought for next sprint.\n\nTom`, atts:[] },
  { from:'Sarah Mitchell',   subject:'Annual Leave Reminder — book before 31 Dec',
    preview:"Just a reminder that any untaken leave must be booked and taken before 31 December...",
    text:`Hi David,\n\nJust a reminder that any remaining annual leave for 2024 must be booked and taken by 31 December — we aren't able to carry days forward into 2025 under the current policy.\n\nYou currently have 6 days remaining. Please book these in MyHR at your earliest convenience and let Lucy know your plans.\n\nIf you have exceptional circumstances preventing you from taking leave before year-end, please speak to me directly.\n\nThanks,\nSarah`, atts:[] },
  { from:'Alex Kumar',       subject:'Quick question on the data model',
    preview:"Hope you're not too busy — just want to double-check something before I update the spec...",
    text:`Hi David,\n\nHope you're not too buried. Quick one — in the current data model, does the CustomerID field in the legacy system map 1:1 to the AccountRef in Salesforce, or is there a transformation step I'm missing?\n\nI want to make sure the spec is accurate before we share it with the dev team.\n\nCheers,\nAlex`, atts:[] },
  { from:'Kate Davies',      subject:'New proposal: Meridian podcast series',
    preview:"I've been thinking — what if we launched an internal podcast? I've put together a one-pager...",
    text:`Hi David,\n\nI've been mulling over an idea and wanted a sanity check before I take it to Richard.\n\nWhat if Meridian launched a short internal podcast series — something like 10-minute interviews with people across the business on what they're working on, industry trends, etc.? Could be great for culture and onboarding.\n\nI've put together a rough one-pager (attached) — would love your honest thoughts.\n\nKate`, atts:[{ name:'Podcast_Concept_Note.docx', size:'56 KB', type:'word' }] },
  { from:'Richard Thompson', subject:'One to watch: AI in professional services',
    preview:"Sharing an article I found this morning — highly relevant to where we're taking the Digital practice...",
    text:`Hi David,\n\nSharing this article from McKinsey — I think it's highly relevant to the Digital practice strategy and aligns with what Emma is proposing for Q4.\n\nhttps://www.mckinsey.com/capabilities/quantumblack/our-insights/the-state-of-ai\n\nWould be good to discuss at our next team meeting.\n\nRichard`, atts:[] },
  { from:'David Patel',      subject:'Budget update: Q4 planning cycle opens Monday',
    preview:'Heads up that the Q4 budget planning cycle opens on Monday. Please have your initial submissions...',
    text:`Hi David,\n\nJust a heads-up that the Q4 budget planning cycle opens on Monday 4 November.\n\nPlease submit your initial headcount and opex requirements by 11 November using the template on SharePoint (Finance > Budget Planning > Q4 2024).\n\nI'll be holding budget clinics on 7 and 8 November — book a slot via my calendar.\n\nDavid Patel\nFinance Director`, atts:[] },
  { from:'Emma Clarke',      subject:'FWD: Conference invite — Digital Transformation Summit',
    preview:"I think this could be worth attending. Let me know if you want to go and I'll request the budget...",
    text:`Hi David,\n\nThis looks relevant to your work on Project Phoenix and the TechNova account. Let me know if you'd like to attend and I'll put in a budget request.\n\nEmma`,
    quoted:`---------- Forwarded message ---------\nFrom: events@digitaltransformationsummit.co.uk\nSubject: You're invited — Digital Transformation Summit, London, 6 Feb 2025\n\nJoin 2,000+ digital leaders for a day of keynotes, workshops and networking at The Brewery, London.\n\nTopics include:\n- AI & automation in professional services\n- Data-driven decision making\n- Agile at scale\n- CX transformation\n\nEarly bird pricing ends 30 November.`, atts:[] },
  { from:'Lucy Harrison',    subject:'Re: 1:1 Notes — Action items',
    preview:'Following up from our chat — three actions for you before next Friday...',
    text:`Hi David,\n\nGreat 1:1 earlier. As discussed, three actions for you before our next catch-up:\n\n1. Share updated development plan with me by Wednesday (template in HR portal)\n2. Set up a call with Alex to resolve the Sprint 4 blocker this week\n3. Review the Q4 capacity sheet and flag any concerns\n\nLet me know if you need anything from my side.\n\nLucy`, atts:[] },
  { from:'Chris Newman',     subject:'Hargreaves & Stone proposal — can you join Thursday?',
    preview:'Quick one — can you join me for the initial scoping call with Hargreaves & Stone on Thursday morning?',
    text:`Hi David,\n\nFollowing my earlier email about Hargreaves & Stone — I've set up a scoping call with James Hargreaves for Thursday at 09:30 (before the All-Hands).\n\nI'd really value having you on the call given your experience on the Ashworth project. Should be no more than 45 minutes.\n\nCan you confirm you're free?\n\nCheers,\nChris`, atts:[] },
  { from:'Tom Wilson',       subject:'URGENT: Test environment is down',
    preview:'The shared test environment has gone down again. Blocking UAT for Sprint 4...',
    text:`Hi David,\n\nThe shared test environment is down again. This is blocking UAT for Sprint 4 and we're supposed to be doing the demo in 2 hours.\n\nI've raised a ticket with IT (Ref: #4899) but they're saying 2-hour response SLA. Can you escalate? I don't think we can wait.\n\nTom`, atts:[] },
  { from:'Sarah Mitchell',   subject:'Christmas Party — Save the Date!',
    preview:"I'm thrilled to announce that this year's Christmas party will be held at The Shard on 12 December...",
    text:`Hi everyone,\n\nI'm absolutely thrilled to announce that this year's Meridian Christmas Party will be held at:\n\n🎉  The Shard, Level 31 — Thursday, 12 December 2024\n🕖  19:00 – midnight\n🥂  Three-course dinner, open bar, live music\n\nThis is a huge thank you from Richard and the board for an exceptional year.\n\nPartners and plus-ones welcome. Please RSVP to this email by 15 November — just reply with your name and whether you're bringing a guest.\n\nLooking forward to celebrating with you all!\n\nSarah\nHR Business Partner`, atts:[] },

  { from:'Sarah Mitchell',   subject:'Career Development: Interview Skills Workshop — Book Your Place',
    preview:"Following our PDP conversations, we're running a free internal workshop on interview technique and career progression...",
    text:`Hi David,\n\nFollowing on from our recent personal development plan conversations, I wanted to flag a brilliant opportunity.\n\nWe are running a free internal workshop:\n\n📌  Interview Skills & Career Progression\n📅  Wednesday, 6 November, 14:00 – 17:00\n📍  Training Room 2, 4th Floor\n\nThe session covers:\n\n✅  How to structure compelling answers using the STAR method (Situation, Task, Action, Result)\n✅  What interviewers are really looking for — and how to stand out\n✅  Salary negotiation and how to confidently discuss your worth\n✅  Common interview mistakes and how to avoid them\n✅  Questions YOU should be asking in interviews\n\nRegardless of whether you are actively job hunting or simply want to be prepared, these are skills that pay dividends throughout your career.\n\nSpaces are limited to 12. Please reply to this email to reserve your place.\n\nBest,\nSarah Mitchell\nHR Business Partner`, atts:[{ name:'Interview_Prep_Workbook.pdf', size:'1.4 MB', type:'pdf' }] },

  { from:'Lucy Harrison',    subject:'Fwd: Really good read — CV writing in 2024',
    preview:"Sharing this — some of it I knew, but the section on quantifying achievements is genuinely eye-opening...",
    text:`Hi David,\n\nCame across this during our L&D research and thought it was genuinely worth sharing. The section on quantifying your achievements (rather than just listing responsibilities) is something I wish someone had told me earlier in my career.\n\nKey takeaways from the article:\n\n1.  Lead with impact, not duties\n    Instead of "Responsible for managing project timelines", write "Delivered 3 projects on time and under budget, saving £45K in Q2 2024". Numbers make your CV scannable and credible.\n\n2.  Tailor every application\n    A generic CV is spotted immediately. Mirror the language in the job description — ATS systems filter on keywords before a human ever sees your application.\n\n3.  Keep it to two pages maximum\n    Hiring managers spend an average of 7 seconds on a first scan. Front-load your strongest points.\n\n4.  The summary section is your elevator pitch\n    Three to four lines at the top should answer: who you are, what you bring, and what you are looking for. Make it specific.\n\n5.  LinkedIn must match your CV\n    Recruiters cross-reference constantly. Inconsistencies raise red flags.\n\nLet me know if you ever want to run your CV past me — happy to give informal feedback.\n\nLucy`, atts:[] },

  { from:'Sarah Mitchell',   subject:'Career Corner Newsletter — Finding Your Next Role',
    preview:"This month we cover job search strategy, working with recruiters, and how to get your application noticed...",
    text:`Hi David,\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📰  MERIDIAN CAREER CORNER — November 2024\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\nWHETHER YOU ARE EXPLORING OPTIONS OR ACTIVELY LOOKING\n\nAt Meridian, we believe in supporting your whole career — including your journey beyond us if the time comes. This month's Career Corner covers practical job search advice.\n\n─────────────────────────────\n🔍  JOB SEARCH STRATEGY\n─────────────────────────────\n• Only 30% of jobs are ever advertised. The hidden job market is accessed through networking — LinkedIn, industry events, and warm introductions.\n• Set up job alerts on LinkedIn, Indeed, and Reed with specific titles and locations. Review and refine them weekly.\n• Target companies proactively. Identify 10–15 organisations you would love to work for and follow them closely. Reach out to hiring managers directly on LinkedIn — a personalised connection request can open doors.\n\n─────────────────────────────\n🤝  WORKING WITH RECRUITERS\n─────────────────────────────\n• Specialist recruiters in your sector are worth cultivating. They often fill roles before they are advertised.\n• Be honest about your motivations and salary expectations — mismatches waste everyone's time.\n• Ask recruiters for feedback after every interview, whether successful or not.\n\n─────────────────────────────\n📝  MAKING YOUR APPLICATION STAND OUT\n─────────────────────────────\n• Your cover letter should answer three things: why this role, why this company, and why you.\n• Research the company thoroughly — their latest news, annual report, leadership team, and competitors.\n• If applying online, optimise your CV for Applicant Tracking Systems (ATS) by using keywords directly from the job description.\n\n─────────────────────────────\n💬  INTERVIEW PREPARATION\n─────────────────────────────\n• Prepare 8–10 STAR examples covering: leadership, problem-solving, conflict, failure, achievement, collaboration, and change.\n• Research your interviewers on LinkedIn beforehand.\n• Prepare three thoughtful questions that show strategic thinking — not just "what does the role involve?".\n• Send a thank-you email within 24 hours of every interview.\n\nGood luck to everyone on their journey!\n\nSarah Mitchell\nHR Business Partner`, atts:[] },

  { from:'Lucy Harrison',    subject:'Re: Your PDP — a few thoughts before our next 1:1',
    preview:"I've been thinking about where you could take your career from here. Wanted to share some honest thoughts...",
    text:`Hi David,\n\nAhead of our next 1:1, I wanted to share some honest thoughts on your development — I think you are underselling yourself and it is time to change that.\n\nYou have had a strong year. The TechNova delivery, the data migration work, the way you have supported the junior team members — all of that is real, tangible impact. But I am not sure it is visible enough on your profile or in how you talk about yourself.\n\nA few things I would encourage you to do:\n\n1.  Update your LinkedIn right now\n    Add the TechNova project with specific outcomes. "Led data workstream for digital transformation engagement, delivering migration 2 weeks ahead of schedule and within budget" is far more powerful than a blank endorsement section.\n\n2.  Start keeping an achievements log\n    Every time you do something noteworthy — a positive client comment, a problem solved, a process improved — write it down with numbers if possible. This becomes the raw material for your CV, your next appraisal, and your next job application.\n\n3.  Think about where you want to be in three years\n    Not in vague terms — specifically. What title? What salary range? What type of work? Working backwards from that gives your development a direction.\n\n4.  Consider getting interview-ready even if you are not looking\n    The best time to practise is when there is no pressure. I am happy to do a mock interview with you if useful.\n\nYou have a lot to offer, David. Let us make sure the right people know it.\n\nSee you Thursday.\n\nLucy`, atts:[] },

  { from:'Sarah Mitchell',   subject:'You are not behind — a note on career gaps',
    preview:"I wanted to reach out personally. Career breaks are more common than you think, and they do not have to hold you back...",
    text:`Hi David,\n\nI wanted to send a personal note on something I speak to a lot of people about: career gaps and the anxiety that can come with them.\n\nFirst — a career gap is not a red flag. Employers in 2024 understand this more than ever. Redundancy, health, caring responsibilities, personal circumstances — life happens to everyone. What matters is how you frame it and what you bring to the table.\n\n─────────────────────────────\nHOW TO ADDRESS A CAREER GAP CONFIDENTLY\n─────────────────────────────\n\n• Own it, do not apologise for it\n  "I took time out to [reason] and used the period to [skill/reflection/project]" is a completely acceptable answer. Fudging or hiding it almost always backfires.\n\n• Show what you did with the time\n  Even small things count — freelance work, volunteering, online courses, personal projects. They demonstrate drive and a growth mindset.\n\n• Reframe it as a strength\n  Many people come back from a break with renewed clarity about what they want, better self-awareness, and perspective that colleagues who have never stepped away simply do not have.\n\n• Update your skills during the gap\n  If you have time now, even a short Coursera or LinkedIn Learning course shows forward momentum on your CV.\n\n─────────────────────────────\nREMEMBER\n─────────────────────────────\n\n"The gap on your CV is a chapter, not the whole story."\n\nYou are defined by the quality of your work and the impact you have made — not by whether your employment was continuous.\n\nMy door is always open if you want to talk anything through.\n\nWarm regards,\nSarah Mitchell\nHR Business Partner`, atts:[{ name:'Returning_to_Work_Guide.pdf', size:'980 KB', type:'pdf' }] },

  { from:'Lucy Harrison',    subject:'Some words for anyone having a tough week',
    preview:"I know not everyone is in the same place right now. Wanted to share something that helped me during a difficult period...",
    text:`Hi David,\n\nI do not usually send emails like this but I have been thinking about the team and I wanted to reach out.\n\nJob searching — and particularly a period between roles — can be genuinely hard. It is easy to tie your sense of worth to your employment status, and when that changes it can knock you more than people expect.\n\nA few things I believe strongly:\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n💬  QUOTES THAT HAVE HELPED ME\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"You are not a failure because you are between jobs. You are a person navigating a difficult system, and that takes real courage."\n\n"Rejection is redirection. Every no is clearing the path to the right yes."\n— Unknown\n\n"It always seems impossible until it is done."\n— Nelson Mandela\n\n"You do not have to have it all figured out to move forward. Just take the next step."\n\n"The job does not make the person. The person makes the job."\n\n"Hard times never last. Hard people do."\n— Robert H. Schuller\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nA PRACTICAL REMINDER\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\nIf you are currently job hunting:\n\n• Treat the search as a job in itself — dedicate time each day but also protect time to switch off\n• Celebrate small wins (a call back, a positive recruiter conversation, a strong application sent)\n• Lean on your network — most people genuinely want to help if you ask\n• Look after yourself physically — sleep, movement, and fresh air make a measurable difference to your mindset\n\nYou are more capable than you feel on the hard days.\n\nLucy`, atts:[] },

  { from:'Sarah Mitchell',   subject:'Career Corner: Rebuilding confidence after redundancy',
    preview:"Redundancy affects more people than talk about it. This month we focus on bouncing back stronger...",
    text:`Hi David,\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📰  MERIDIAN CAREER CORNER — Special Edition\n     Rebuilding After Redundancy\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\nRedundancy is one of the most unsettling things that can happen in a career — yet it affects around 1 in 5 workers at some point in their professional life. If you are going through it now, or supporting someone who is, this edition is for you.\n\n─────────────────────────────\n🧠  THE EMOTIONAL SIDE (AND WHY IT MATTERS)\n─────────────────────────────\n\nRedundancy triggers real grief. The loss of routine, identity, colleagues, and purpose can manifest as anxiety, low motivation, or a persistent sense of inadequacy. This is completely normal.\n\nWhat helps:\n• Talk about it — to friends, family, a coach, or a professional\n• Maintain a daily structure even without a job to go to\n• Separate your professional setback from your personal worth — they are not the same thing\n\n─────────────────────────────\n📋  PRACTICAL STEPS IN THE FIRST MONTH\n─────────────────────────────\n\n1.  Claim what you are entitled to\n    Check your redundancy pay calculation. If you have been employed for 2+ years, you are entitled to statutory redundancy pay. Many employers offer more.\n\n2.  Register with HMRC and check Universal Credit eligibility\n    Even if you expect to find something quickly, it is worth knowing your options.\n\n3.  Refresh your CV while your achievements are fresh\n    This is the best time to write your CV — everything is current and your confidence in what you delivered is highest.\n\n4.  Tell your network immediately\n    A simple LinkedIn post or message to contacts that you are "open to opportunities" frequently leads to introductions before your CV reaches a recruiter.\n\n5.  Set a daily job search routine\n    Aim for: 2–3 targeted applications per day (quality over quantity), 1 networking message or call, and 30 minutes of skills development.\n\n─────────────────────────────\n💪  GETTING BACK YOUR CONFIDENCE\n─────────────────────────────\n\n• List everything you have delivered in the past 3 years — you will surprise yourself\n• Ask 3 former colleagues for a LinkedIn recommendation right now, while the relationship is warm\n• Do a mock interview — out loud, not just in your head. The physical rehearsal matters.\n• Remember: the company made the decision about the role, not about you as a person\n\n─────────────────────────────\n🌟  A FINAL THOUGHT\n─────────────────────────────\n\n"Many of the most successful people alive have been made redundant. It is not the end of your story — in many cases it is the moment that forced the change they needed."\n\nYou have got this.\n\nSarah Mitchell\nHR Business Partner\nMeridian Group`, atts:[{ name:'Redundancy_Support_Checklist.pdf', size:'720 KB', type:'pdf' }] },

  { from:'Lucy Harrison',    subject:'LinkedIn tips for when you are job hunting (the honest version)',
    preview:"LinkedIn is strange and nobody really tells you how to use it properly when you need it most. Here is what actually works...",
    text:`Hi David,\n\nLinkedIn is one of those platforms that feels more intimidating when you are actively job hunting — the pressure to perform, the visible connection counts, the seemingly constant stream of other people announcing new roles.\n\nHere is the honest, practical guide:\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n✅  THINGS THAT ACTUALLY WORK\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n1.  Turn on "Open to Work" (the private version)\n    You can signal to recruiters only — not your current employer — that you are open to opportunities. Go to your profile > Open to > Finding a new job > select "Recruiters only". This generates 40% more recruiter messages.\n\n2.  Your headline is not your job title\n    It is prime real estate. Instead of "Project Manager at Meridian Group", try "Project Manager | Digital Transformation | Stakeholder Engagement | Open to New Opportunities". Recruiters search by keywords.\n\n3.  The About section should sound like a person\n    Write in first person. Three to four sentences covering what you do, what you are proud of, and what you are looking for. Do not copy your CV.\n\n4.  Post one thing a week\n    It does not need to be profound. Share an article with a one-sentence thought. Comment meaningfully on someone else\'s post. Consistency builds visibility — your name starts appearing in feeds.\n\n5.  Message people directly — it works\n    "Hi [name], I noticed you work at [company] — I have been following your work in [area] and I am exploring opportunities in this space. Would you be open to a brief 15-minute call?" A genuine, short message gets a surprisingly high response rate.\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n❌  WHAT NOT TO DO\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n• Do not mass-connect with strangers without a note\n• Do not share anything that sounds desperate or anxious — keep your tone confident and curious\n• Do not update your profile once and forget it — treat it as a living document\n\nHope this helps. You have a strong background — make sure it is visible.\n\nLucy`, atts:[] },

  { from:'Sarah Mitchell',   subject:'Daily motivation — for anyone who needs it today',
    preview:"Some days the job search feels relentless. Here are a few things to hold onto when it gets hard...",
    text:`Hi David,\n\nJust a short one today.\n\nJob searching is one of the most psychologically demanding things a person can do. You are putting yourself out there repeatedly, facing silence or rejection, often without any feedback, while trying to maintain confidence and momentum.\n\nIf today is one of the harder days, here are some things worth remembering:\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n✨  FOR THE DIFFICULT DAYS\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"Your value does not decrease based on someone\'s inability to see your worth."\n\n"Every expert was once a beginner. Every professional was once an amateur. Every master was once a disaster. Keep going."\n\n"You have survived every hard day so far. Your track record is 100%."\n\n"The right opportunity is not always the next one. Sometimes it is the one after several no\'s."\n\n"Being unemployed does not make you unemployable. It makes you available."\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📋  THREE THINGS TO DO TODAY\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n1.  Send one application — just one, but make it a good one\n2.  Reach out to one person in your network, even just to say hello\n3.  Do one thing that is completely unrelated to job hunting — protect your mental space\n\nTomorrow will feel different.\n\nSarah\nHR Business Partner`, atts:[] },

  { from:'Emma Clarke',      subject:'Re: Transferable skills — you have more than you think',
    preview:"I was chatting to Sarah about career development and wanted to share something that changed how I think about my own experience...",
    text:`Hi David,\n\nI was having a conversation with Sarah earlier about professional development and it made me want to share something.\n\nWhen I was between roles a few years ago, I massively undersold myself because I could not see how my experience mapped across to what I wanted to do next. I kept thinking "I have not done exactly that before." It took an honest conversation with a mentor to help me see that transferable skills are often more valuable than direct experience.\n\nHere is a framework that helped me — and that I now share regularly:\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nTHE TRANSFERABLE SKILLS AUDIT\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\nFor each of the following, write down 1–2 specific examples from your career:\n\n📌  Communication — presenting to clients, writing reports, difficult conversations\n📌  Problem solving — a time a project went wrong and how you fixed it\n📌  Leadership — influencing without authority, mentoring, driving alignment\n📌  Delivery — managing competing priorities, hitting deadlines under pressure\n📌  Adaptability — a time the goalposts changed and you adjusted quickly\n📌  Commercial awareness — understanding budgets, client relationships, business impact\n\nOnce you have that list, you will realise that most job descriptions are just different combinations of these same things.\n\nThe person who gets the job is rarely the one who has done it before. It is usually the one who can most convincingly show they have done the underlying things that make the job possible.\n\nHope this is useful.\n\nEmma`, atts:[{ name:'Transferable_Skills_Worksheet.docx', size:'88 KB', type:'word' }] },
];

let poolIndex = 0;
function getNextPoolEmail() {
  const item = NEW_EMAIL_POOL[poolIndex % NEW_EMAIL_POOL.length];
  poolIndex++;
  const s = sender(item.from);
  const now = Date.now();
  return {
    id: uid(),
    folder: 'inbox',
    read: false,
    flagged: false,
    from: s,
    subject: item.subject,
    preview: item.preview,
    ts: now,
    attachments: item.atts || [],
    body: {
      from: s,
      text: item.text,
      quoted: item.quoted || null
    }
  };
}

// ── State ──────────────────────────────────────
let currentEmailId = null;
let currentFilter  = 'all';

// ── Render email list ─────────────────────────
function renderList() {
  const ul     = document.getElementById('emails');
  const search = (document.getElementById('searchInput').value || '').toLowerCase();
  ul.innerHTML = '';

  let emails = Store._emails.filter(e => e.folder === 'inbox');

  if (currentFilter === 'unread')  emails = emails.filter(e => !e.read);
  if (currentFilter === 'flagged') emails = emails.filter(e => e.flagged);

  if (search) {
    emails = emails.filter(e =>
      e.from.name.toLowerCase().includes(search) ||
      e.subject.toLowerCase().includes(search) ||
      e.preview.toLowerCase().includes(search)
    );
  }

  emails.forEach(email => {
    const li = document.createElement('li');
    li.className = 'email-item' +
      (email.read ? '' : ' unread') +
      (email.flagged ? ' flagged' : '') +
      (email.id === currentEmailId ? ' selected' : '');
    li.dataset.id = email.id;
    li.innerHTML = `
      <div class="ei-top">
        <span class="ei-sender">${htmlEscape(email.from.name)}</span>
        <span class="ei-time">${formatTime(email.ts)}</span>
      </div>
      <div class="ei-subject">${htmlEscape(email.subject)}</div>
      <div class="ei-preview">${htmlEscape(email.preview)}${email.attachments && email.attachments.length ? ' <i class="fas fa-paperclip" style="font-size:10px;color:#a19f9d"></i>' : ''}</div>
    `;
    li.addEventListener('click', () => openEmail(email.id));
    li.addEventListener('contextmenu', e => { e.preventDefault(); showCtxMenu(e, email.id); });
    ul.appendChild(li);
  });

  // Update badge
  const unread = Store.unreadCount();
  document.getElementById('inbox-count').textContent = unread || '';
  document.getElementById('inbox-count').style.display = unread ? '' : 'none';
  document.getElementById('statusMsg').textContent =
    emails.length + ' message' + (emails.length !== 1 ? 's' : '') + ', ' + unread + ' unread';
}

// ── Open and render an email ───────────────────
function openEmail(id) {
  currentEmailId = id;
  Store.markRead(id);

  const email = Store.get(id);
  if (!email) return;

  renderList(); // refresh to remove unread dot

  document.getElementById('reading-placeholder').style.display = 'none';
  document.getElementById('email-header').innerHTML = renderEmailHeader(email);
  document.getElementById('email-body-content').innerHTML = renderEmailBody(email);
}

function renderEmailHeader(email) {
  const ini   = initials(email.from.name);
  const color = avatarColor(ini);
  const date  = new Date(email.ts);
  const dateStr = date.toLocaleDateString('en-GB', { weekday:'short', day:'2-digit', month:'short', year:'numeric' }) +
                  ', ' + date.toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit' });
  const atts = (email.attachments || []).map(a => `
    <div class="re-att" title="${htmlEscape(a.name)}">
      <i class="far fa-file-${a.type === 'word' ? 'word' : a.type === 'excel' ? 'excel' : 'pdf'} re-att-icon ${a.type}"></i>
      <div class="re-att-info">
        <div class="re-att-name">${htmlEscape(a.name)}</div>
        <div class="re-att-size">${a.size}</div>
      </div>
    </div>`).join('');
  return `
    <div class="re-header">
      <div class="re-subject">${htmlEscape(email.subject)}</div>
      <div class="re-meta-row">
        <div class="re-avatar" style="background:${color}">${ini}</div>
        <div class="re-sender-info">
          <div class="re-sender-name">${htmlEscape(email.from.name)}</div>
          <div class="re-sender-email">${htmlEscape(email.from.email)}</div>
          <div class="re-to-line">To: David Jenner &lt;d.jenner@meridiangroup.co.uk&gt;</div>
        </div>
        <div class="re-datetime">${dateStr}</div>
      </div>
      <div class="re-actions">
        <button class="re-act-btn" onclick="replyTo('${email.id}')"><i class="fas fa-reply"></i> Reply</button>
        <button class="re-act-btn" onclick="replyAllTo('${email.id}')"><i class="fas fa-reply-all"></i> Reply all</button>
        <button class="re-act-btn" onclick="forwardEmail('${email.id}')"><i class="fas fa-share"></i> Forward</button>
        <div class="re-divider"></div>
        <button class="re-act-icon" title="Delete" onclick="deleteEmail('${email.id}')"><i class="far fa-trash-alt"></i></button>
        <button class="re-act-icon" title="Archive" onclick="archiveEmail('${email.id}')"><i class="far fa-folder"></i></button>
        <button class="re-act-icon" title="Flag" onclick="flagEmail('${email.id}')"><i class="far fa-flag"></i></button>
        <button class="re-act-icon" title="Mark unread" onclick="markUnread('${email.id}')"><i class="far fa-envelope"></i></button>
        <button class="re-act-icon" title="Print"><i class="fas fa-print"></i></button>
        <button class="re-act-icon" title="More"><i class="fas fa-ellipsis-h"></i></button>
      </div>
    </div>
    ${atts ? '<div class="re-attachments">' + atts + '</div>' : ''}
  `;
}

function renderEmailBody(email) {
  const b = email.body;
  const bodyText   = personaliseText(b.text);
  const quotedText = b.quoted ? personaliseText(b.quoted) : null;
  let html = `<div class="email-body-text">${htmlEscape(bodyText)}</div>`;
  if (b.from) {
    const ini   = initials(b.from.name);
    const color = avatarColor(ini);
    html = `<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
      <div style="width:32px;height:32px;border-radius:50%;background:${color};display:flex;align-items:center;justify-content:center;color:#fff;font-size:12px;font-weight:700;flex-shrink:0">${ini}</div>
      <div>
        <div style="font-size:13px;font-weight:600;color:#323130">${htmlEscape(b.from.name)}</div>
        <div style="font-size:11px;color:#605e5c">${htmlEscape(b.from.role || b.from.email)}</div>
      </div>
    </div>
    <div class="email-body-text">${htmlEscape(bodyText)}</div>`;
  }
  if (quotedText) {
    html += `<div class="email-quoted">${htmlEscape(quotedText)}</div>`;
  }
  const sig = (b.from && b.from.name && b.from.name !== 'IT Support' && b.from.name !== 'Meridian HR')
    ? `<div class="email-sig">${htmlEscape(b.from.name)}<br>${htmlEscape(b.from.role || '')}<br>Meridian Group | meridiangroup.co.uk</div>`
    : '';
  return html + sig;
}

// ── Email actions ──────────────────────────────
function replyTo(id) {
  const e = Store.get(id || currentEmailId);
  if (!e) return;
  openCompose({
    title: 'Re: ' + e.subject,
    to: e.from.email,
    subject: 'Re: ' + e.subject,
    quoted: `\n\n\nFrom: ${e.from.name} <${e.from.email}>\nSent: ${new Date(e.ts).toLocaleString('en-GB')}\nTo: ${getUserFullName()}\nSubject: ${e.subject}\n\n${e.body.text}`
  });
}

function replyAllTo(id) {
  const e = Store.get(id || currentEmailId);
  if (!e) return;
  openCompose({
    title: 'Re: ' + e.subject,
    to: e.from.email + '; team@meridiangroup.co.uk',
    subject: 'Re: ' + e.subject,
    quoted: `\n\n\nFrom: ${e.from.name} <${e.from.email}>\nSent: ${new Date(e.ts).toLocaleString('en-GB')}\nTo: ${getUserFullName()}; Team\nSubject: ${e.subject}\n\n${e.body.text}`
  });
}

function forwardEmail(id) {
  const e = Store.get(id || currentEmailId);
  if (!e) return;
  openCompose({
    title: 'Fwd: ' + e.subject,
    to: '',
    subject: 'Fwd: ' + e.subject,
    quoted: `\n\n\n---------- Forwarded message ----------\nFrom: ${e.from.name} <${e.from.email}>\nSent: ${new Date(e.ts).toLocaleString('en-GB')}\nTo: ${getUserFullName()}\nSubject: ${e.subject}\n\n${e.body.text}`
  });
}

function deleteEmail(id) {
  const eid = id || currentEmailId;
  Store.remove(eid);
  if (currentEmailId === eid) {
    currentEmailId = null;
    document.getElementById('email-header').innerHTML = '';
    document.getElementById('email-body-content').innerHTML = '';
    document.getElementById('reading-placeholder').style.display = '';
  }
  renderList();
}

function archiveEmail(id) {
  const e = Store.get(id || currentEmailId);
  if (e) { e.folder = 'archive'; Store.save(); }
  deleteEmail(id); // clears view
}

function flagEmail(id) {
  Store.toggleFlag(id || currentEmailId);
  renderList();
}

function markUnread(id) {
  Store.markUnread(id || currentEmailId);
  renderList();
}

// Global aliases for onclick in HTML
window.replyTo     = replyTo;
window.replyAllTo  = replyAllTo;
window.forwardEmail= forwardEmail;
window.deleteEmail = deleteEmail;
window.archiveEmail= archiveEmail;
window.flagEmail   = flagEmail;
window.markUnread  = markUnread;

// ── Compose window ────────────────────────────
let composeCount = 0;
function openCompose(opts) {
  opts = opts || {};
  composeCount++;
  const id  = 'cwin' + composeCount;
  const win = document.createElement('div');
  win.className = 'compose-win';
  win.id = id;
  win.innerHTML = `
    <div class="compose-header" id="${id}-hdr">
      <span class="compose-title">${htmlEscape(opts.title || 'New Message')}</span>
      <button class="compose-hbtn" id="${id}-min" title="Minimise">&#8722;</button>
      <button class="compose-hbtn" id="${id}-cls" title="Close">&#10005;</button>
    </div>
    <div class="compose-field">
      <label>To</label>
      <input type="text" id="${id}-to" value="${htmlEscape(opts.to || '')}">
    </div>
    <div class="compose-field">
      <label>Subject</label>
      <input type="text" id="${id}-subj" value="${htmlEscape(opts.subject || '')}">
    </div>
    <textarea class="compose-body" id="${id}-body" placeholder="Write your message here…">${htmlEscape(opts.quoted || '')}</textarea>
    <div class="compose-footer">
      <button class="compose-send" id="${id}-send"><i class="fas fa-paper-plane"></i> Send</button>
      <button class="compose-foot-btn" title="Attach file"><i class="fas fa-paperclip"></i></button>
      <button class="compose-foot-btn" title="Insert image"><i class="far fa-image"></i></button>
      <button class="compose-foot-btn" title="Formatting"><i class="fas fa-font"></i></button>
      <div style="margin-left:auto;display:flex;gap:4px">
        <button class="compose-foot-btn" title="Discard" id="${id}-discard"><i class="far fa-trash-alt"></i></button>
      </div>
    </div>
  `;
  document.getElementById('composeStack').appendChild(win);

  document.getElementById(id+'-min').addEventListener('click', e => {
    e.stopPropagation();
    win.classList.toggle('cmin');
  });
  document.getElementById(id+'-hdr').addEventListener('click', () => {
    win.classList.toggle('cmin');
  });
  document.getElementById(id+'-cls').addEventListener('click', e => {
    e.stopPropagation();
    win.remove();
  });
  document.getElementById(id+'-discard').addEventListener('click', () => win.remove());
  document.getElementById(id+'-send').addEventListener('click', () => {
    showToast('Message Sent', 'Your message has been sent successfully.', 'fa-check-circle');
    win.remove();
  });

  // Focus body
  setTimeout(() => {
    const body = document.getElementById(id+'-body');
    if (body) { body.focus(); body.setSelectionRange(0, 0); }
  }, 100);
}

// ── Toast ──────────────────────────────────────
function showToast(sender, subject, icon) {
  const container = document.getElementById('toastContainer');
  const t = document.createElement('div');
  t.className = 'toast';
  t.innerHTML = `
    <i class="fas ${icon || 'fa-envelope'} toast-icon"></i>
    <div class="toast-body">
      <div class="toast-sender">${htmlEscape(sender)}</div>
      <div class="toast-subject">${htmlEscape(subject)}</div>
    </div>
    <button class="toast-x">&#10005;</button>
  `;
  t.querySelector('.toast-x').addEventListener('click', e => { e.stopPropagation(); t.remove(); });
  container.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transform = 'translateX(20px)'; setTimeout(() => t.remove(), 300); }, 5000);
}

// ── Context menu ──────────────────────────────
let ctxEmailId = null;
function showCtxMenu(e, id) {
  ctxEmailId = id;
  const menu = document.getElementById('contextMenu');
  menu.style.display = 'block';
  const x = Math.min(e.clientX, window.innerWidth - 200);
  const y = Math.min(e.clientY, window.innerHeight - 200);
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
}

// ── Window management ─────────────────────────
function initWindowManagement() {
  const win  = document.getElementById('outlookWindow');
  const tb   = document.getElementById('titleBar');
  const rh   = document.getElementById('resizeHandle');
  const minB = document.getElementById('minBtn');
  const maxB = document.getElementById('maxBtn');
  const clsB = document.getElementById('closeBtn');
  const tbarMailBtn = document.getElementById('taskbarMailBtn');

  // Start full-screen (maximised)
  win.classList.add('maximized');
  tb.style.cursor = 'default';

  let isMax = true;
  let savedPos = { top:'20px', left:'20px', width:'1100px', height:'720px' };

  // Minimise
  minB.addEventListener('click', () => {
    win.classList.add('minimized');
    tbarMailBtn.classList.remove('running');
  });
  tbarMailBtn.addEventListener('click', () => {
    win.classList.remove('minimized');
    tbarMailBtn.classList.add('running');
  });

  // Maximise / restore
  function toggleMax() {
    if (isMax) {
      win.classList.remove('maximized');
      win.style.top    = savedPos.top;
      win.style.left   = savedPos.left;
      win.style.width  = savedPos.width;
      win.style.height = savedPos.height;
      maxB.innerHTML   = '&#9633;';
      tb.style.cursor  = 'move';
      isMax = false;
    } else {
      savedPos = { top: win.style.top, left: win.style.left, width: win.style.width, height: win.style.height };
      win.classList.add('maximized');
      maxB.innerHTML   = '&#10064;';
      tb.style.cursor  = 'default';
      isMax = true;
    }
  }
  maxB.addEventListener('click', toggleMax);
  tb.addEventListener('dblclick', toggleMax);

  // Close (just minimise for demo)
  clsB.addEventListener('click', () => {
    win.classList.add('minimized');
    tbarMailBtn.classList.remove('running');
  });

  // Drag
  let dragging = false, dx = 0, dy = 0;
  tb.addEventListener('mousedown', e => {
    if (e.target.closest('.title-bar-controls')) return;
    if (isMax) return;
    dragging = true;
    dx = e.clientX - win.offsetLeft;
    dy = e.clientY - win.offsetTop;
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    let nx = e.clientX - dx;
    let ny = e.clientY - dy;
    nx = Math.max(-win.offsetWidth + 100, Math.min(window.innerWidth - 100, nx));
    ny = Math.max(0, Math.min(window.innerHeight - 80, ny));
    win.style.left = nx + 'px';
    win.style.top  = ny + 'px';
  });
  document.addEventListener('mouseup', () => { dragging = false; });

  // Resize
  let resizing = false, rStartX, rStartY, rStartW, rStartH;
  rh.addEventListener('mousedown', e => {
    resizing = true;
    rStartX  = e.clientX;
    rStartY  = e.clientY;
    rStartW  = win.offsetWidth;
    rStartH  = win.offsetHeight;
    e.preventDefault();
    e.stopPropagation();
  });
  document.addEventListener('mousemove', e => {
    if (!resizing) return;
    const nw = Math.max(820, rStartW + (e.clientX - rStartX));
    const nh = Math.max(500, rStartH + (e.clientY - rStartY));
    win.style.width  = nw + 'px';
    win.style.height = nh + 'px';
    win.classList.remove('maximized');
    isMax = false;
  });
  document.addEventListener('mouseup', () => { resizing = false; });
}

// ── Clocks ────────────────────────────────────
function initClocks() {
  function tick() {
    const now  = new Date();
    const h    = String(now.getHours()).padStart(2,'0');
    const m    = String(now.getMinutes()).padStart(2,'0');
    const s    = String(now.getSeconds()).padStart(2,'0');
    document.getElementById('current-time').textContent = h + ':' + m + ':' + s;
  }
  tick();
  setInterval(tick, 1000);

  function updateLock() {
    const now  = new Date();
    const h    = String(now.getHours()).padStart(2,'0');
    const m    = String(now.getMinutes()).padStart(2,'0');
    document.getElementById('lock-clock').textContent = h + ':' + m;
    const days   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    document.getElementById('lock-date').textContent =
      days[now.getDay()] + ', ' + now.getDate() + ' ' + months[now.getMonth()];
  }
  updateLock();
  setInterval(updateLock, 30000);
}

// ── Search ────────────────────────────────────
function initSearch() {
  document.getElementById('searchInput').addEventListener('input', () => renderList());
}

// ── Simulate new emails arriving at random intervals ──
function startEmailStream() {
  function scheduleNext() {
    // Random delay between 8 and 45 seconds to feel realistic
    const delay = (8 + Math.random() * 37) * 1000;
    setTimeout(function() {
      const email = getNextPoolEmail();
      Store.add(email);
      renderList();
      showToast(email.from.name, email.subject, 'fa-envelope');
      scheduleNext();
    }, delay);
  }
  // Send the first one after a short initial delay
  setTimeout(function() {
    const email = getNextPoolEmail();
    Store.add(email);
    renderList();
    showToast(email.from.name, email.subject, 'fa-envelope');
    scheduleNext();
  }, 5000);
}

// ── DOMContentLoaded ─────────────────────────
document.addEventListener('DOMContentLoaded', function () {

  // Load or seed emails — clear stale data from old versions
  const DATA_VERSION = '4';
  Store.load();
  const stale = Store._emails.length > 0 && Store._emails.some(e => !e.folder);
  if (Store._emails.length === 0 || stale || localStorage.getItem('outseek_v') !== DATA_VERSION) {
    localStorage.setItem('outseek_v', DATA_VERSION);
    Store._emails = [];
    Store.save();
    buildInitialEmails().forEach(e => Store.add(e));
  }

  updateUserNameUI();
  renderList();
  initWindowManagement();
  initClocks();
  initSearch();
  startEmailStream();

  // New mail buttons
  document.getElementById('newEmailButton').addEventListener('click', () => openCompose({ title:'New Message' }));
  document.getElementById('sideNewMail').addEventListener('click',     () => openCompose({ title:'New Message' }));

  // Action bar buttons
  document.getElementById('actReply').addEventListener('click',   () => replyTo(currentEmailId));
  document.getElementById('actReplyAll').addEventListener('click',() => replyAllTo(currentEmailId));
  document.getElementById('actForward').addEventListener('click', () => forwardEmail(currentEmailId));
  document.getElementById('actDelete').addEventListener('click',  () => deleteEmail(currentEmailId));
  document.getElementById('actArchive').addEventListener('click', () => archiveEmail(currentEmailId));
  document.getElementById('actFlag').addEventListener('click',    () => flagEmail(currentEmailId));
  document.getElementById('actRead').addEventListener('click',    () => markUnread(currentEmailId));

  // Filter tabs
  document.querySelectorAll('.ftab').forEach(tab => {
    tab.addEventListener('click', function () {
      document.querySelectorAll('.ftab').forEach(t => t.classList.remove('active'));
      this.classList.add('active');
      currentFilter = this.dataset.filter;
      renderList();
    });
  });

  // Folder items
  document.querySelectorAll('.folder-item').forEach(item => {
    item.addEventListener('click', function () {
      document.querySelectorAll('.folder-item').forEach(f => f.classList.remove('active'));
      this.classList.add('active');
      // Clear reading pane when switching folder
      currentEmailId = null;
      document.getElementById('email-header').innerHTML = '';
      document.getElementById('email-body-content').innerHTML = '';
      document.getElementById('reading-placeholder').style.display = '';
    });
  });

  // Nav icon buttons
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.addEventListener('click', function () {
      document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
      this.classList.add('active');
    });
  });

  // Context menu actions
  document.getElementById('ctx-reply').addEventListener('click',   () => replyTo(ctxEmailId));
  document.getElementById('ctx-replyall').addEventListener('click',() => replyAllTo(ctxEmailId));
  document.getElementById('ctx-forward').addEventListener('click', () => forwardEmail(ctxEmailId));
  document.getElementById('ctx-delete').addEventListener('click',  () => deleteEmail(ctxEmailId));
  document.getElementById('ctx-unread').addEventListener('click',  () => { markUnread(ctxEmailId); renderList(); });
  document.getElementById('ctx-flag').addEventListener('click',    () => { flagEmail(ctxEmailId); renderList(); });

  // Hide context menu on click
  document.addEventListener('click', () => {
    document.getElementById('contextMenu').style.display = 'none';
  });

  // ── Office ambience audio (Web Audio API) ──
  let ambienceCtx   = null;
  let ambienceGain  = null;
  let ambiencePlaying = false;

  function buildOfficeAmbience() {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const master = ctx.createGain();
    master.gain.value = 0;
    master.connect(ctx.destination);

    // Pre-bake a long stereo white-noise buffer (20 s, loops seamlessly)
    const SR  = ctx.sampleRate;
    const len = SR * 20;
    const noiseBuf = ctx.createBuffer(2, len, SR);
    for (let c = 0; c < 2; c++) {
      const d = noiseBuf.getChannelData(c);
      // Brown noise: integrate white noise
      let acc = 0;
      for (let i = 0; i < len; i++) {
        acc = (acc + (Math.random() * 2 - 1) * 0.08) * 0.998;
        d[i] = acc * 4;
      }
      // Normalise
      let peak = 0;
      for (let i = 0; i < len; i++) peak = Math.max(peak, Math.abs(d[i]));
      if (peak > 0) for (let i = 0; i < len; i++) d[i] /= peak;
    }

    function noiseSrc() {
      const s = ctx.createBufferSource();
      s.buffer = noiseBuf;
      s.loop = true;
      return s;
    }

    // ── HVAC rumble (deep, continuous) ──────────────
    const hvac = noiseSrc();
    const hvacLP = ctx.createBiquadFilter();
    hvacLP.type = 'lowpass'; hvacLP.frequency.value = 180; hvacLP.Q.value = 0.5;
    const hvacG = ctx.createGain(); hvacG.gain.value = 0.18;
    hvac.connect(hvacLP); hvacLP.connect(hvacG); hvacG.connect(master);
    hvac.start();

    // ── Distant office chatter (3 overlapping voices) ──
    [900, 1400, 2100].forEach(function(freq, i) {
      const src = noiseSrc();
      const bp  = ctx.createBiquadFilter();
      bp.type = 'bandpass'; bp.frequency.value = freq; bp.Q.value = 1.8;
      const g = ctx.createGain(); g.gain.value = 0.06 + i * 0.01;
      // LFO to make chatter feel animated
      const lfo = ctx.createOscillator();
      const lfoG = ctx.createGain(); lfoG.gain.value = 0.03;
      lfo.frequency.value = 0.4 + i * 0.3;
      lfo.connect(lfoG); lfoG.connect(g.gain);
      lfo.start();
      src.connect(bp); bp.connect(g); g.connect(master);
      src.start(ctx.currentTime + i * 0.3);
    });

    // ── Room hiss (air) ──────────────────────────────
    const hiss = noiseSrc();
    const hissHP = ctx.createBiquadFilter();
    hissHP.type = 'highpass'; hissHP.frequency.value = 4000;
    const hissG = ctx.createGain(); hissG.gain.value = 0.022;
    hiss.connect(hissHP); hissHP.connect(hissG); hissG.connect(master);
    hiss.start();

    // ── Keyboard typing bursts ───────────────────────
    function keyClick() {
      const t   = ctx.currentTime;
      const osc = ctx.createOscillator();
      const env = ctx.createGain();
      osc.type = 'triangle';
      osc.frequency.value = 1800 + Math.random() * 2200;
      env.gain.setValueAtTime(0, t);
      env.gain.linearRampToValueAtTime(0.018, t + 0.003);
      env.gain.exponentialRampToValueAtTime(0.0001, t + 0.05);
      osc.connect(env); env.connect(master);
      osc.start(t); osc.stop(t + 0.06);
    }

    function typingBurst() {
      if (!ambiencePlaying) { setTimeout(typingBurst, 2000); return; }
      // Burst of 4-12 keystrokes at realistic WPM spacing (~120-200 ms apart)
      const count = 4 + Math.floor(Math.random() * 9);
      let delay = 0;
      for (let i = 0; i < count; i++) {
        delay += 80 + Math.random() * 140;
        setTimeout(keyClick, delay);
      }
      // Next burst in 1.5 – 6 seconds
      setTimeout(typingBurst, 1500 + Math.random() * 4500);
    }
    setTimeout(typingBurst, 800);

    // ── Phone ringing ────────────────────────────────
    function phoneRing() {
      if (!ambiencePlaying) { setTimeout(phoneRing, 5000); return; }

      // British dual-tone ringing: 400 Hz + 450 Hz, 0.4 s on / 0.2 s off / 0.4 s on, pause
      function playRingPair(done) {
        [0, 0.6].forEach(function(offset) {
          [400, 450].forEach(function(freq) {
            const o = ctx.createOscillator();
            const g = ctx.createGain();
            const t = ctx.currentTime + offset;
            o.frequency.value = freq;
            o.type = 'sine';
            g.gain.setValueAtTime(0, t);
            g.gain.linearRampToValueAtTime(0.055, t + 0.02);
            g.gain.setValueAtTime(0.055, t + 0.38);
            g.gain.linearRampToValueAtTime(0, t + 0.4);
            o.connect(g); g.connect(master);
            o.start(t); o.stop(t + 0.41);
          });
        });
        setTimeout(done, 1400);
      }

      // 2–4 ring pairs then silence for a while
      const pairs = 2 + Math.floor(Math.random() * 3);
      let p = 0;
      function nextPair() {
        if (p++ < pairs) playRingPair(nextPair);
        else setTimeout(phoneRing, 15000 + Math.random() * 35000);
      }
      nextPair();
    }
    setTimeout(phoneRing, 4000 + Math.random() * 8000);

    // ── Occasional paper rustle / chair creak ────────
    function rustle() {
      if (!ambiencePlaying) { setTimeout(rustle, 5000); return; }
      const t   = ctx.currentTime;
      const src = noiseSrc();
      const bp  = ctx.createBiquadFilter();
      bp.type = 'bandpass'; bp.frequency.value = 600 + Math.random() * 800; bp.Q.value = 0.4;
      const env = ctx.createGain();
      env.gain.setValueAtTime(0, t);
      env.gain.linearRampToValueAtTime(0.04, t + 0.1);
      env.gain.exponentialRampToValueAtTime(0.0001, t + 0.6 + Math.random() * 0.4);
      src.connect(bp); bp.connect(env); env.connect(master);
      src.start(t); src.stop(t + 1.2);
      setTimeout(rustle, 8000 + Math.random() * 20000);
    }
    setTimeout(rustle, 3000);

    return { ctx, masterGain: master };
  }

  document.getElementById('playButton').addEventListener('click', function() {
    const icon = document.getElementById('playIcon');
    if (!ambiencePlaying) {
      // First play — build the audio graph
      if (!ambienceCtx) {
        const a = buildOfficeAmbience();
        ambienceCtx  = a.ctx;
        ambienceGain = a.masterGain;
      }
      // Fade in
      ambienceGain.gain.cancelScheduledValues(ambienceCtx.currentTime);
      ambienceGain.gain.setValueAtTime(ambienceGain.gain.value, ambienceCtx.currentTime);
      ambienceGain.gain.linearRampToValueAtTime(1, ambienceCtx.currentTime + 1.5);
      ambiencePlaying = true;
      icon.classList.remove('fa-play');
      icon.classList.add('fa-pause');
    } else {
      // Fade out
      ambienceGain.gain.cancelScheduledValues(ambienceCtx.currentTime);
      ambienceGain.gain.setValueAtTime(ambienceGain.gain.value, ambienceCtx.currentTime);
      ambienceGain.gain.linearRampToValueAtTime(0, ambienceCtx.currentTime + 1);
      ambiencePlaying = false;
      icon.classList.remove('fa-pause');
      icon.classList.add('fa-play');
    }
  });

});
