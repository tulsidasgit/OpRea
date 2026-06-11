# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**OpRea / OpATS** — an Operational Applicant Tracking System built as a SharePoint Framework (SPFx) web part, with two satellite projects for the public job-application flow. The repo has **three independent npm projects**, each with its own `package.json` and `node_modules`:

| Project | Path | Purpose |
|---|---|---|
| SPFx web part (main) | `/` (root) | The HR-facing React app embedded in SharePoint |
| Azure Function API | `azure-function/` | Anonymous HTTP API backing the public application form (`GetJobDetails`, `SubmitApplication`) |
| Azure Static Web App | `azure-static-web-app/` | The public-facing "Apply for this job" form (Vite + React 18) |

[docs/APPLICATION_OVERVIEW.md](docs/APPLICATION_OVERVIEW.md) is a narrative walkthrough but **predates the removal of all AI-driven features** (JD generation, resume fitment scoring/screening, screening questions, feedback polishing) and the "Ongoing Positions" / "My Interviews" tabs and Track Progress report panel — treat it as historical only; this file (CLAUDE.md) is the current source of truth for architecture. Other files under `docs/` (`ARCHITECTURE.md`, `DEVELOPER_GUIDE.md`, `DECISIONS.md`, `LEARNING.md`, `PROMPTS.md`, `USER_GUIDE.md`) are currently placeholders/empty.

## Commands

### SPFx web part (root)
```
npm install
heft start --clean        # npm start — local dev server (config/serve.json opens the SharePoint workbench page)
heft test --clean --production && heft package-solution --production   # npm run build — full build + .sppkg package
heft clean                 # npm run clean
```
There is no test suite (`heft test` runs but no `*.test.*`/`*.spec.*` files exist yet).

### Azure Function (`azure-function/`)
```
npm install
npm run build      # tsc
npm run watch      # tsc -w
npm start          # build then `func start` (requires Azure Functions Core Tools)
```
Requires `local.settings.json` (copy from `local.settings.json.example`) with `SP_SITE_URL`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` (SharePoint app-only/ACS credentials), `ANTHROPIC_API_KEY`, `CORS_ORIGIN`.

### Azure Static Web App (`azure-static-web-app/`)
```
npm install
npm run dev        # vite dev server
npm run build      # tsc && vite build
npm run preview
```
Reads `VITE_API_BASE_URL` to point at the Azure Function.

## Architecture (SPFx web part)

Entry point: [src/webparts/recruitmentTracker/RecruitmentTrackerWebPart.ts](src/webparts/recruitmentTracker/RecruitmentTrackerWebPart.ts) — initializes `SPFI` (PnP.js) via `spSPFx(this.context)` and a `GraphService` for the current user, then renders [RecruitmentTracker.tsx](src/webparts/recruitmentTracker/components/RecruitmentTracker.tsx). The only web part property is `applyBaseUrl` (URL of the deployed Azure Static Web App, set via the property pane), used to build per-job public application links.

`RecruitmentTracker.tsx` is a plain class component (no global state management — each tab component owns its own state) that gates tab visibility:
- **Add Job Description** (`PostJob/PostJobForm.tsx`) — only if `SpService.isAllowedPoster(email)` is true (checked against the `AllowedPosters` SharePoint list); HR writes the job description manually (no AI generation)
- **Track Progress** (`TrackProgress/TrackProgress.tsx`, `CandidateCard.tsx`, `InterviewScheduler.tsx`), **Completed Jobs** (`CompletedJobs/CompletedJobs.tsx`), **All Candidates** (`AllCandidates/AllCandidates.tsx`) — HR-only, gated by `isHREmail()`

Users who are neither an AllowedPoster nor HR see no tabs.

### Shared services (`src/webparts/recruitmentTracker/components/shared/`)

- **`SpService.ts`** — all SharePoint CRUD (PnP.js `SPFI`) against the lists/library described below; also resume upload/folder management.
- **`GraphService.ts`** — wraps `MSGraphClientFactory`: `/me` for current user, `/me/sendMail` for outgoing mail, user photo lookup.
- **`EmailService.ts`** — builds HTML emails for the 3 notification scenarios reachable from the current UI (job posted, interview scheduled — interviewer + candidate), all sent via `GraphService`. `HR_EMAILS` (HR distribution list) and `isHREmail()` live here — this is the single source of truth for who counts as "HR" for tab gating.
- **`candidateCategory.ts`** — `deriveCategory()` computes a candidate's pipeline stage (`Received → Round 1 → Round 2 → HR Discussion`, or `Rejected`) purely from `ICandidate` + `IInterview[]` — there is no stored "stage" field; it's always derived.
- **`models.ts`** — all domain interfaces (`IJobOpening`, `ICandidate`, `IInterview`, etc.) — the canonical reference for SharePoint field shapes as seen by the client.

### SharePoint data backbone

Provisioned by [setup/CreateLists.ps1](setup/CreateLists.ps1) (PnP.PowerShell) — run against a target site with `-SiteUrl`. Lists/library:

| List / Library | Key fields | Notes |
|---|---|---|
| `Departments` | `Title`, `IsActive` | Source for cascading department dropdown |
| `JobTitles` | `Title`, `DepartmentId` | Filtered by department |
| `JobOpenings` | `Title`, `Department`, `JobTitle`, `RequiredSkills`, `GoodToHaveSkills`, `JobLocation`, `JobType`, `Experience`, `DueDate`, `Status` (Open/In Progress/Closed), `PostedBy`, `LinkedInUrl`, `JobDescription`, `ApplicationFormUrl` | `ApplicationFormUrl` column is optional — `updateJobOpeningApplicationUrl` silently no-ops if missing |
| `Candidates` | `JobOpeningId`, `CandidateName`, `Email`, `Phone`, `ResumeUrl`, `FitmentScore`, `MatchingSkills`, `MissingSkills`, `AISummary`, `HRFeedback`, `Recommendation`, `ExperienceMatch`, `ApplicationStatus`, plus referral fields (`ReferredBy`, `ReferrerEmail`, `ReferrerEmployeeId`, `ReferrerDesignation`) | Referral fields are written best-effort (try/catch) since the columns may not exist on all tenants |
| `Interviews` | `CandidateId`, `JobOpeningId`, `InterviewRound` (1/2/3), `InterviewerEmail`, `ScheduledDate`, `FeedbackStatus` (Pending/Submitted), `Feedback`, `HRNotes` | |
| `AllowedPosters` | `UserEmail` | Gates the "Add Job Description" tab |
| `Resumes` (document library) | — | Auto-organised as `Resumes/<Department>/<JobTitle>/{Direct,Referred}/` |

`SpService` resolves list/library URLs dynamically (works on root site or subsite) and sanitizes department/job-title names for folder paths.

## Architecture (Azure Function + Static Web App)

These two projects implement the **public, unauthenticated apply flow** referenced from `JobOpenings.ApplicationFormUrl`:

1. `azure-static-web-app/src/Apply.tsx` — public form at `/apply?jobId=N`, calls `GetJobDetails` then `SubmitApplication` via [apiService.ts](azure-static-web-app/src/services/apiService.ts).
2. `azure-function/src/functions/getJobDetails.ts` — anonymous GET, returns job details (404 if not found, 410 if `Status = Closed`).
3. `azure-function/src/functions/submitApplication.ts` → [formWriter.ts](azure-function/src/functions/formWriter.ts) — validates payload (PDF/DOCX only, 5 MB limit), uploads the resume to the `Resumes/<Department>/<JobTitle>/` folder, creates a `Candidates` item (`ApplicationStatus = Received`, `Source = Direct`), and bumps `JobOpenings.Status` from `Open` to `In Progress`.
4. [azure-function/src/shared/spClient.ts](azure-function/src/shared/spClient.ts) — server-side PnP.js client authenticated via SharePoint **Add-In (ACS) tokens** (`SP_CLIENT_ID`/`SP_CLIENT_SECRET` from `/_layouts/15/appregnew.aspx`, no Entra admin consent needed), with token caching.

Note: the `Candidates` fields written here (`CurrentCtc`, `ExpectedCtc`, `NoticePeriod`, `Gender`, `ReasonForLeaving`, `WorkCultureExpectation`, `LinkedInUrl`, `Source`) are additional columns beyond what `setup/CreateLists.ps1` provisions — ensure they exist on the target `Candidates` list before relying on this flow.

## Power Automate flows

[flows/ResumeReceivedFlow.json](flows/ResumeReceivedFlow.json) and [flows/InterviewReminderFlow.json](flows/InterviewReminderFlow.json) are exported flow definitions complementing the Microsoft Forms intake path and the interview-feedback escalation mechanism described in the application overview.
