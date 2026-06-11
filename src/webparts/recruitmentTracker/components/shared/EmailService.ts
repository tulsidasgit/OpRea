import { GraphService } from './GraphService';
import { IJobOpening, ICandidate, IInterview } from './models';

const HR_EMAILS = ['tulsidas.rp@operative.com'];
const COMPANY_NAME = 'OpRea Recruitment';
const SITE_URL = 'https://sintecmedia365.sharepoint.com';

/**
 * Checks whether an email belongs to the HR list above.
 * Used to gate access to HR-only tabs (Track Progress, Completed Jobs, All Candidates)
 * in addition to driving notification recipients — keep HR_EMAILS as the single source of truth.
 */
export function isHREmail(email: string): boolean {
  const normalized = email.trim().toLowerCase();
  return HR_EMAILS.some(hr => hr.toLowerCase() === normalized);
}

function emailShell(title: string, bodyContent: string): string {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>${title}</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background:#f3f2f1; margin:0; padding:0; }
    .wrapper { max-width:620px; margin:32px auto; background:#ffffff; border-radius:8px;
               box-shadow:0 2px 8px rgba(0,0,0,.12); overflow:hidden; }
    .header  { background:#0078d4; padding:24px 32px; }
    .header h1 { color:#fff; margin:0; font-size:20px; font-weight:600; }
    .header p  { color:#c7e0f4; margin:4px 0 0; font-size:13px; }
    .body    { padding:28px 32px; color:#323130; }
    .body h2 { font-size:16px; color:#323130; margin-top:0; }
    .field   { margin-bottom:12px; }
    .label   { font-size:12px; color:#605e5c; text-transform:uppercase; letter-spacing:.5px; }
    .value   { font-size:14px; color:#323130; margin-top:2px; }
    .badge   { display:inline-block; padding:3px 10px; border-radius:12px;
               font-size:12px; font-weight:600; }
    .badge-green  { background:#dff6dd; color:#107c10; }
    .badge-blue   { background:#deecf9; color:#0078d4; }
    .badge-orange { background:#fff4ce; color:#8a8000; }
    .badge-red    { background:#fde7e9; color:#a80000; }
    .btn { display:inline-block; background:#0078d4; color:#fff; text-decoration:none;
           padding:10px 20px; border-radius:4px; font-size:14px; font-weight:600;
           margin-top:16px; }
    .divider { border:none; border-top:1px solid #edebe9; margin:20px 0; }
    .footer  { background:#f3f2f1; padding:16px 32px; font-size:12px; color:#605e5c; }
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="header">
      <h1>${COMPANY_NAME}</h1>
      <p>Recruitment &amp; Application Tracking System</p>
    </div>
    <div class="body">
      ${bodyContent}
    </div>
    <div class="footer">
      This is an automated notification from ${COMPANY_NAME}. Do not reply to this email.
    </div>
  </div>
</body>
</html>`;
}

export class EmailService {
  private _graph: GraphService;

  constructor(graphService: GraphService) {
    this._graph = graphService;
  }

  // ── 1. New job posted — notify HR ─────────────────────────────────────────

  public async notifyHRJobPosted(job: IJobOpening, postedByName: string): Promise<void> {
    const dueDate = job.dueDate ? new Date(job.dueDate).toLocaleDateString('en-GB') : '—';
    const body = emailShell(
      'New Job Opening Posted',
      `<h2>New Job Opening: ${job.title}</h2>
       <div class="field"><div class="label">Department</div><div class="value">${job.department}</div></div>
       <div class="field"><div class="label">Job Title</div><div class="value">${job.jobTitle}</div></div>
       <div class="field"><div class="label">Experience Required</div><div class="value">${job.experience}</div></div>
       <div class="field"><div class="label">Due Date</div><div class="value">${dueDate}</div></div>
       <div class="field"><div class="label">Required Skills</div><div class="value">${job.requiredSkills}</div></div>
       <div class="field"><div class="label">Posted By</div><div class="value">${postedByName}</div></div>
       ${job.linkedInUrl ? `<div class="field"><div class="label">LinkedIn</div><div class="value"><a href="${job.linkedInUrl}">${job.linkedInUrl}</a></div></div>` : ''}
       <hr class="divider"/>
       <p>Please review the job opening and begin shortlisting candidates.</p>
       <a class="btn" href="${SITE_URL}">Open Recruitment Tracker</a>`
    );

    await this._graph.sendEmail({
      to: HR_EMAILS,
      subject: `[OpRea] New Job Posted: ${job.title} — ${job.department}`,
      bodyHtml: body,
    });
  }

  // ── 2. Interview scheduled — notify interviewer ───────────────────────────

  public async notifyInterviewerScheduled(
    interview: IInterview,
    candidate: ICandidate,
    job: IJobOpening
  ): Promise<void> {
    const scheduledDate = interview.scheduledDate
      ? new Date(interview.scheduledDate).toLocaleString('en-GB')
      : '—';

    const skillChips = (skills: string, color: string): string =>
      skills
        ? skills.split(',').map(s => s.trim()).filter(Boolean)
            .map(s => `<span style="display:inline-block;background:${color};padding:2px 8px;border-radius:10px;font-size:12px;font-weight:500;margin:2px 3px 2px 0">${s}</span>`)
            .join('')
        : '<span style="color:#605e5c">None identified</span>';

    const resumeLink = candidate.resumeUrl
      ? `<a href="${SITE_URL}${candidate.resumeUrl}" style="color:#0078d4;font-weight:600">Download Resume</a>`
      : '<span style="color:#605e5c">Not uploaded</span>';

    const hrFeedbackSection = candidate.hrFeedback
      ? `<hr class="divider"/>
         <h2>HR Feedback</h2>
         <div style="background:#fff4ce;border-left:3px solid #f7630c;padding:10px 14px;border-radius:0 4px 4px 0;margin:8px 0">
           ${candidate.hrFeedback}
         </div>`
      : '';

    const body = emailShell(
      'Interview Scheduled',
      `<h2>Interview Invitation — Round ${interview.interviewRound}</h2>
       <p>You have been assigned to interview a candidate. Please review all details below before the interview.</p>

       <div class="field"><div class="label">Scheduled Date &amp; Time</div><div class="value"><strong>${scheduledDate}</strong></div></div>
       <div class="field"><div class="label">Interview Round</div><div class="value"><strong>Round ${interview.interviewRound}</strong></div></div>

       <hr class="divider"/>
       <h2>Position Details</h2>
       <div class="field"><div class="label">Job Title</div><div class="value">${job.title}</div></div>
       <div class="field"><div class="label">Department</div><div class="value">${job.department}</div></div>
       <div class="field"><div class="label">Location</div><div class="value">${job.jobLocation || '—'}</div></div>
       <div class="field"><div class="label">Job Type</div><div class="value">${job.jobType || '—'}</div></div>
       <div class="field"><div class="label">Experience Required</div><div class="value">${job.experience || '—'}</div></div>
       <div class="field"><div class="label">Required Skills</div><div class="value">${skillChips(job.requiredSkills, '#deecf9')}</div></div>
       ${job.goodToHaveSkills ? `<div class="field"><div class="label">Good to Have</div><div class="value">${skillChips(job.goodToHaveSkills, '#f3f2f1')}</div></div>` : ''}

       <hr class="divider"/>
       <h2>Candidate Details</h2>
       <div class="field"><div class="label">Name</div><div class="value"><strong>${candidate.candidateName}</strong></div></div>
       <div class="field"><div class="label">Email</div><div class="value">${candidate.email}</div></div>
       <div class="field"><div class="label">Phone</div><div class="value">${candidate.phone || '—'}</div></div>
       <div class="field"><div class="label">Resume</div><div class="value">${resumeLink}</div></div>

       ${hrFeedbackSection}

       <hr class="divider"/>
       <p>Please submit your interview feedback in the Recruitment Tracker immediately after the interview.</p>
       <a class="btn" href="${SITE_URL}">Open Recruitment Tracker</a>`
    );

    await this._graph.sendEmail({
      to: [interview.interviewerEmail],
      cc: HR_EMAILS,
      subject: `[OpRea] Interview Scheduled — ${candidate.candidateName} for ${job.title} (Round ${interview.interviewRound})`,
      bodyHtml: body,
    });
  }

  // ── 2b. Interview scheduled — notify candidate ────────────────────────────

  public async notifyCandidateInterviewScheduled(
    interview: IInterview,
    candidate: ICandidate,
    job: IJobOpening
  ): Promise<void> {
    if (!candidate.email) return;

    const scheduledDate = interview.scheduledDate
      ? new Date(interview.scheduledDate).toLocaleString('en-GB')
      : '—';

    const body = emailShell(
      'Your Interview Has Been Scheduled',
      `<h2>Interview Scheduled — Round ${interview.interviewRound}</h2>
       <p>Dear ${candidate.candidateName},</p>
       <p>Thank you for applying to <strong>${COMPANY_NAME}</strong>. Your interview for the position below has been scheduled. Please find the details:</p>
       <div class="field"><div class="label">Position</div><div class="value">${job.title}</div></div>
       <div class="field"><div class="label">Department</div><div class="value">${job.department}</div></div>
       <div class="field"><div class="label">Location</div><div class="value">${job.jobLocation || '—'}</div></div>
       <hr class="divider"/>
       <div class="field"><div class="label">Interview Round</div><div class="value">Round ${interview.interviewRound}</div></div>
       <div class="field"><div class="label">Date &amp; Time</div><div class="value"><strong>${scheduledDate}</strong></div></div>
       <hr class="divider"/>
       <p>Please be available a few minutes before the scheduled time. If you have any questions or need to reschedule, reply to this email and our HR team will assist you.</p>
       <p>We look forward to speaking with you.</p>
       <p>Best regards,<br/>${COMPANY_NAME} — HR Team</p>`
    );

    await this._graph.sendEmail({
      to: [candidate.email],
      cc: HR_EMAILS,
      subject: `[${COMPANY_NAME}] Interview Scheduled — ${job.title} (Round ${interview.interviewRound})`,
      bodyHtml: body,
    });
  }

  // ── 3. Missed feedback escalation — notify HR ─────────────────────────────

  public async notifyHREscalation(
    interview: IInterview,
    candidate: ICandidate,
    job: IJobOpening,
    hoursOverdue: number
  ): Promise<void> {
    const scheduledDate = interview.scheduledDate
      ? new Date(interview.scheduledDate).toLocaleString('en-GB')
      : '—';

    const body = emailShell(
      'Interview Feedback Overdue',
      `<h2>⚠ Interview Feedback Overdue</h2>
       <p>The following interview feedback has not been submitted and is <strong>${hoursOverdue} hours overdue</strong>.</p>
       <div class="field"><div class="label">Candidate</div><div class="value">${candidate.candidateName}</div></div>
       <div class="field"><div class="label">Job Title</div><div class="value">${job.title}</div></div>
       <div class="field"><div class="label">Department</div><div class="value">${job.department}</div></div>
       <div class="field"><div class="label">Interview Round</div><div class="value">${interview.interviewRound}</div></div>
       <div class="field"><div class="label">Interviewer</div><div class="value">${interview.interviewerEmail}</div></div>
       <div class="field"><div class="label">Scheduled At</div><div class="value">${scheduledDate}</div></div>
       <hr class="divider"/>
       <p>Please follow up with the interviewer to ensure feedback is submitted promptly.</p>
       <a class="btn" href="${SITE_URL}">Open Recruitment Tracker</a>`
    );

    await this._graph.sendEmail({
      to: HR_EMAILS,
      subject: `[OpRea] ESCALATION: Feedback Overdue — ${candidate.candidateName} / ${job.title}`,
      bodyHtml: body,
    });
  }

  // ── 5. New application received from Ongoing Positions portal ────────────

  public async notifyHRNewApplication(
    job: IJobOpening,
    candidateName: string,
    email: string,
    phone: string,
    resumeUrl: string,
    referral?: { referredBy: string; referrerEmail: string; employeeId: string; designation: string }
  ): Promise<void> {
    const dueDate = job.dueDate ? new Date(job.dueDate).toLocaleDateString('en-GB') : '—';
    const resumeLink = resumeUrl
      ? `<div class="field"><div class="label">Resume (${referral ? 'Referred' : 'Direct'})</div>
         <div class="value"><a href="https://sintecmedia365.sharepoint.com${resumeUrl}" style="color:#0078d4">Download Resume</a></div></div>`
      : '';

    const referralSection = referral
      ? `<hr class="divider"/>
         <h2 style="color:#fd800b">Referred By</h2>
         <div class="field"><div class="label">Employee Name</div><div class="value"><strong>${referral.referredBy}</strong></div></div>
         <div class="field"><div class="label">Employee Email</div><div class="value"><a href="mailto:${referral.referrerEmail}" style="color:#0078d4">${referral.referrerEmail}</a></div></div>
         <div class="field"><div class="label">Employee ID</div><div class="value">${referral.employeeId || '—'}</div></div>
         <div class="field"><div class="label">Designation</div><div class="value">${referral.designation || '—'}</div></div>`
      : '';

    const sourceLabel = referral
      ? `<span class="badge badge-orange">Referral</span>`
      : `<span class="badge badge-blue">Direct Application</span>`;

    const body = emailShell(
      'New Job Application Received',
      `<h2>New Application: ${job.title} &nbsp;${sourceLabel}</h2>
       <p>A candidate has been ${referral ? 'referred' : 'submitted an application'} via the <strong>Ongoing Positions</strong> portal.</p>
       <hr class="divider"/>
       <h2>Candidate Details</h2>
       <div class="field"><div class="label">Full Name</div><div class="value">${candidateName}</div></div>
       <div class="field"><div class="label">Email</div><div class="value"><a href="mailto:${email}" style="color:#0078d4">${email}</a></div></div>
       <div class="field"><div class="label">Phone</div><div class="value">${phone || '—'}</div></div>
       ${resumeLink}
       ${referralSection}
       <hr class="divider"/>
       <h2>Position Details</h2>
       <div class="field"><div class="label">Job Title</div><div class="value">${job.title}</div></div>
       <div class="field"><div class="label">Department</div><div class="value">${job.department}</div></div>
       <div class="field"><div class="label">Location</div><div class="value">${job.jobLocation || '—'}</div></div>
       <div class="field"><div class="label">Experience Required</div><div class="value">${job.experience}</div></div>
       <div class="field"><div class="label">Due Date</div><div class="value">${dueDate}</div></div>
       <hr class="divider"/>
       <p>Please review the application in the Recruitment Tracker.</p>
       <a class="btn" href="${SITE_URL}">Open Recruitment Tracker</a>`
    );

    await this._graph.sendEmail({
      to: HR_EMAILS,
      subject: referral
        ? `[OpRea] Referral: ${candidateName} → ${job.title} (by ${referral.referredBy})`
        : `[OpRea] New Application: ${candidateName} → ${job.title} (${job.department})`,
      bodyHtml: body,
    });
  }
}
