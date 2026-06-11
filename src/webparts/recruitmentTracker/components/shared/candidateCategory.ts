import { ICandidate, IInterview } from './models';

export type CandidateCategory = 'Received' | 'Round 1' | 'Round 2' | 'HR Discussion' | 'Rejected';

export const CATEGORY_ORDER: CandidateCategory[] = ['Received', 'Round 1', 'Round 2', 'HR Discussion', 'Rejected'];

export const CATEGORY_CONFIG: Record<CandidateCategory, { label: string; color: string }> = {
  'Received':          { label: 'Received',               color: '#605e5c' },
  'Round 1':           { label: 'Shortlisted — Round 1',   color: '#8764b8' },
  'Round 2':           { label: 'Shortlisted — Round 2',   color: '#498205' },
  'HR Discussion':     { label: 'HR Discussion',           color: '#ca5010' },
  'Rejected':          { label: 'Rejected',                color: '#a80000' },
};

export function deriveCategory(candidate: ICandidate, jobInterviews: IInterview[]): CandidateCategory {
  if (candidate.applicationStatus === 'Rejected') return 'Rejected';
  const mine = jobInterviews.filter(iv => iv.candidateId === candidate.id);
  const maxRound = mine.reduce((max, iv) => Math.max(max, parseInt(iv.interviewRound, 10)), 0);
  if (maxRound >= 3) return 'HR Discussion';
  if (maxRound === 2) return 'Round 2';
  if (maxRound === 1) return 'Round 1';
  return 'Received';
}
