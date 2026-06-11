import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  Panel,
  PanelType,
  TextField,
  ActionButton,
  Icon,
} from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import type { IInterview, ICandidate, IJobOpening, ICurrentUser } from '../shared/models';
import { SpService } from '../shared/SpService';

export interface IMyInterviewsProps {
  sp: SPFI;
  currentUser: ICurrentUser;
}

interface IMyInterviewsState {
  loading: boolean;
  loadError: string;
  interviews: IInterview[];
  candidateMap: Record<number, ICandidate>;
  jobMap: Record<number, IJobOpening>;
  expandedIds: Record<number, boolean>;
  // Feedback panel
  activeFeedbackInterview: IInterview | undefined;
  feedbackText: string;
  panelError: string;
  submittingFeedback: boolean;
  submitSuccess: boolean;
}

export class MyInterviews extends React.Component<IMyInterviewsProps, IMyInterviewsState> {
  private _spService: SpService;

  constructor(props: IMyInterviewsProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this.state = {
      loading: true,
      loadError: '',
      interviews: [],
      candidateMap: {},
      jobMap: {},
      expandedIds: {},
      activeFeedbackInterview: undefined,
      feedbackText: '',
      panelError: '',
      submittingFeedback: false,
      submitSuccess: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadData();
  }

  private _loadData = async (): Promise<void> => {
    const { currentUser } = this.props;
    if (!currentUser?.email) {
      this.setState({ loading: false });
      return;
    }
    this.setState({ loading: true, loadError: '' });
    try {
      const interviews = await this._spService.getInterviewsByInterviewer(currentUser.email);

      const jobIds = Array.from(new Set(interviews.map(iv => iv.jobOpeningId)));

      const jobMap: Record<number, IJobOpening> = {};
      const candidateMap: Record<number, ICandidate> = {};

      if (jobIds.length > 0) {
        const results = await Promise.all([
          Promise.all(jobIds.map(id => this._spService.getJobOpeningById(id))),
          Promise.all(jobIds.map(id => this._spService.getCandidatesByJobOpening(id))),
        ]);
        const jobs = results[0] as IJobOpening[];
        const candidateArrays = results[1] as ICandidate[][];

        jobs.forEach(j => { jobMap[j.id] = j; });
        candidateArrays.forEach(arr => arr.forEach(c => { candidateMap[c.id] = c; }));
      }

      this.setState({ interviews, jobMap, candidateMap, loading: false });
    } catch (err) {
      this.setState({ loadError: (err as Error).message, loading: false });
    }
  };

  private _toggleExpand = (id: number): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] },
    }));
  };

  private _openFeedbackPanel = (interview: IInterview): void => {
    this.setState({
      activeFeedbackInterview: interview,
      feedbackText: interview.feedback ?? '',
      panelError: '',
      submitSuccess: false,
    });
  };

  private _closeFeedbackPanel = (): void => {
    this.setState({
      activeFeedbackInterview: undefined,
      feedbackText: '',
      panelError: '',
      submitSuccess: false,
    });
  };

  private _submitFeedback = async (): Promise<void> => {
    const { activeFeedbackInterview, feedbackText } = this.state;
    if (!activeFeedbackInterview || !feedbackText.trim()) return;

    this.setState({ submittingFeedback: true, panelError: '' });
    try {
      await this._spService.submitInterviewFeedback(activeFeedbackInterview.id, feedbackText);
      this.setState(prev => ({
        interviews: prev.interviews.map(iv =>
          iv.id === activeFeedbackInterview.id
            ? { ...iv, feedback: feedbackText, feedbackStatus: 'Submitted' }
            : iv
        ),
        submittingFeedback: false,
        submitSuccess: true,
      }));
      setTimeout(() => { this._closeFeedbackPanel(); }, 1600);
    } catch (err) {
      this.setState({ submittingFeedback: false, panelError: (err as Error).message });
    }
  };

  private _renderInterviewCard(interview: IInterview): React.ReactNode {
    const { expandedIds, candidateMap, jobMap } = this.state;
    const candidate = candidateMap[interview.candidateId];
    const job = jobMap[interview.jobOpeningId];
    const isExpanded = !!expandedIds[interview.id];
    const now = new Date();
    const isPast = !!interview.scheduledDate && new Date(interview.scheduledDate) < now;
    const isSubmitted = interview.feedbackStatus === 'Submitted';

    const dateStr = interview.scheduledDate
      ? new Date(interview.scheduledDate).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
      : 'Date not set';

    return (
      <div
        key={interview.id}
        style={{
          background: '#fff',
          borderRadius: 8,
          border: `1px solid ${isPast ? '#edebe9' : '#0078d4'}`,
          marginBottom: 12,
          overflow: 'hidden',
          boxShadow: '0 1px 4px rgba(0,0,0,0.07)',
        }}
      >
        {/* Card header — clickable to expand */}
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 12 }}
          styles={{ root: { padding: '12px 16px', background: isPast ? '#f3f2f1' : '#f0f6ff', cursor: 'pointer' } }}
          onClick={() => this._toggleExpand(interview.id)}
        >
          <span style={{ background: '#0078d4', color: '#fff', padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 700, minWidth: 36, textAlign: 'center' }}>
            R{interview.interviewRound}
          </span>
          <Stack grow tokens={{ childrenGap: 2 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
              <Text styles={{ root: { fontWeight: 600, fontSize: 15 } }}>
                {candidate?.candidateName ?? `Candidate #${interview.candidateId}`}
              </Text>
              <Text styles={{ root: { color: '#605e5c', fontSize: 13 } }}>·</Text>
              <Text styles={{ root: { color: '#323130', fontSize: 13 } }}>
                {job?.title ?? `Job #${interview.jobOpeningId}`}
              </Text>
              {job && (
                <Text styles={{ root: { color: '#605e5c', fontSize: 13 } }}>({job.department})</Text>
              )}
            </Stack>
            <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>
              <Icon iconName="Calendar" styles={{ root: { marginRight: 4, fontSize: 11 } }} />
              {dateStr}
            </Text>
          </Stack>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <span style={{
              background: isSubmitted ? '#dff6dd' : '#fff4ce',
              color: isSubmitted ? '#107c10' : '#8a6914',
              padding: '2px 10px', borderRadius: 12, fontSize: 11, fontWeight: 600,
            }}>
              {isSubmitted ? '✓ Submitted' : '⏳ Pending'}
            </span>
            <Icon
              iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'}
              styles={{ root: { fontSize: 14, color: '#605e5c' } }}
            />
          </Stack>
        </Stack>

        {/* Expanded body */}
        {isExpanded && (
          <div style={{ padding: '16px 20px', borderTop: '1px solid #edebe9' }}>
            {/* Candidate / role / fitment row */}
            <Stack horizontal tokens={{ childrenGap: 32 }} wrap styles={{ root: { marginBottom: 14 } }}>
              {candidate && (
                <Stack tokens={{ childrenGap: 2 }}>
                  <Text styles={{ root: { fontSize: 11, color: '#a19f9d', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>Candidate</Text>
                  <Text styles={{ root: { fontWeight: 600 } }}>{candidate.candidateName}</Text>
                  <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>{candidate.email}</Text>
                  {candidate.phone && <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>{candidate.phone}</Text>}
                </Stack>
              )}
              {job && (
                <Stack tokens={{ childrenGap: 2 }}>
                  <Text styles={{ root: { fontSize: 11, color: '#a19f9d', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>Role</Text>
                  <Text styles={{ root: { fontWeight: 600 } }}>{job.title}</Text>
                  <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>{job.department} · {job.experience}</Text>
                </Stack>
              )}
              {candidate?.resumeUrl && (
                <Stack tokens={{ childrenGap: 2 }}>
                  <Text styles={{ root: { fontSize: 11, color: '#a19f9d', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>Resume</Text>
                  <a
                    href={candidate.resumeUrl}
                    target="_blank"
                    rel="noreferrer"
                    style={{ fontSize: 13, color: '#0078d4', textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 4 }}
                  >
                    <Icon iconName="OpenInNewWindow" styles={{ root: { fontSize: 12 } }} />
                    View Resume
                  </a>
                </Stack>
              )}
            </Stack>

            {/* HR notes (from interview record) */}
            {interview.hrNotes && (
              <div style={{ background: '#fff4ce', border: '1px solid #f7e28c', borderRadius: 6, padding: '10px 14px', marginBottom: 12 }}>
                <Text styles={{ root: { fontSize: 11, fontWeight: 600, color: '#8a6914', textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block', marginBottom: 4 } }}>
                  HR Notes
                </Text>
                <Text styles={{ root: { fontSize: 13 } }}>{interview.hrNotes}</Text>
              </div>
            )}

            {/* Submitted feedback preview */}
            {isSubmitted && interview.feedback && (
              <div style={{ background: '#dff6dd', border: '1px solid #92c353', borderRadius: 6, padding: '10px 14px', marginBottom: 12 }}>
                <Text styles={{ root: { fontSize: 11, fontWeight: 600, color: '#107c10', textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block', marginBottom: 4 } }}>
                  Your Submitted Feedback
                </Text>
                <Text styles={{ root: { fontSize: 13, whiteSpace: 'pre-wrap' } }}>{interview.feedback}</Text>
              </div>
            )}

            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
              <DefaultButton
                iconProps={{ iconName: isSubmitted ? 'Edit' : 'Send' }}
                text={isSubmitted ? 'Update Feedback' : 'Submit Feedback'}
                onClick={() => this._openFeedbackPanel(interview)}
                styles={{ root: { borderColor: '#0078d4', color: '#0078d4' } }}
              />
            </Stack>
          </div>
        )}
      </div>
    );
  }

  private _renderFeedbackPanel(): React.ReactNode {
    const {
      activeFeedbackInterview,
      feedbackText,
      panelError,
      submittingFeedback,
      submitSuccess,
      candidateMap,
      jobMap,
    } = this.state;

    if (!activeFeedbackInterview) return undefined;

    const candidate = candidateMap[activeFeedbackInterview.candidateId];
    const job = jobMap[activeFeedbackInterview.jobOpeningId];
    const isUpdate = activeFeedbackInterview.feedbackStatus === 'Submitted';

    return (
      <Panel
        isOpen
        type={PanelType.medium}
        headerText={`${isUpdate ? 'Update' : 'Submit'} Feedback — Round ${activeFeedbackInterview.interviewRound}`}
        onDismiss={this._closeFeedbackPanel}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={
                submitSuccess
                  ? '✓ Saved!'
                  : submittingFeedback
                  ? 'Saving…'
                  : isUpdate
                  ? 'Update Feedback'
                  : 'Submit Feedback'
              }
              disabled={submittingFeedback || !feedbackText.trim() || submitSuccess}
              onClick={() => { this._submitFeedback().catch(() => undefined); }}
              styles={{
                root: submitSuccess
                  ? { background: '#107c10', borderColor: '#107c10' }
                  : undefined,
              }}
            />
            <DefaultButton text="Cancel" onClick={this._closeFeedbackPanel} disabled={submittingFeedback} />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '4px 0 16px' } }}>
          {/* Candidate summary */}
          <div style={{ background: '#f3f2f1', borderRadius: 8, padding: '14px 16px' }}>
            <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="center" wrap>
              <Stack tokens={{ childrenGap: 2 }}>
                <Text styles={{ root: { fontWeight: 700, fontSize: 15 } }}>
                  {candidate?.candidateName ?? '—'}
                </Text>
                <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>{candidate?.email ?? ''}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 2 }}>
                <Text styles={{ root: { fontWeight: 600, fontSize: 13 } }}>{job?.title ?? '—'}</Text>
                <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>
                  {job?.department ?? ''}{job?.experience ? ` · ${job.experience}` : ''}
                </Text>
              </Stack>
              {activeFeedbackInterview.scheduledDate && (
                <Text styles={{ root: { fontSize: 12, color: '#605e5c' } }}>
                  <Icon iconName="Calendar" styles={{ root: { marginRight: 4, fontSize: 11 } }} />
                  {new Date(activeFeedbackInterview.scheduledDate).toLocaleString('en-GB', {
                    day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
                  })}
                </Text>
              )}
            </Stack>
            {candidate?.resumeUrl && (
              <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 10 } }}>
                <a
                  href={candidate.resumeUrl}
                  target="_blank"
                  rel="noreferrer"
                  style={{ fontSize: 12, color: '#0078d4', textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 4 }}
                >
                  <Icon iconName="OpenInNewWindow" styles={{ root: { fontSize: 11 } }} />
                  Resume
                </a>
              </Stack>
            )}
          </div>

          {/* Feedback textarea */}
          <Stack tokens={{ childrenGap: 6 }}>
            <Text styles={{ root: { fontWeight: 600, fontSize: 14 } }}>
              Your Feedback <span style={{ color: '#a80000' }}>*</span>
            </Text>
            <TextField
              multiline
              rows={10}
              value={feedbackText}
              onChange={(_e, v) => this.setState({ feedbackText: v ?? '' })}
              placeholder="Share your observations about the candidate's technical skills, communication, problem-solving approach, depth of knowledge, and overall suitability for the role…"
              disabled={submittingFeedback}
            />

            {panelError && (
              <MessageBar
                messageBarType={MessageBarType.error}
                onDismiss={() => this.setState({ panelError: '' })}
                dismissButtonAriaLabel="Close"
              >
                {panelError}
              </MessageBar>
            )}

            {submitSuccess && (
              <MessageBar messageBarType={MessageBarType.success}>
                Feedback saved successfully! Closing…
              </MessageBar>
            )}
          </Stack>

          {/* HR notes for reference */}
          {activeFeedbackInterview.hrNotes && (
            <div style={{ background: '#fff4ce', border: '1px solid #f7e28c', borderRadius: 6, padding: '10px 14px' }}>
              <Text styles={{ root: { fontSize: 11, fontWeight: 600, color: '#8a6914', textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block', marginBottom: 4 } }}>
                HR Notes (Reference)
              </Text>
              <Text styles={{ root: { fontSize: 12, color: '#323130' } }}>{activeFeedbackInterview.hrNotes}</Text>
            </div>
          )}
        </Stack>
      </Panel>
    );
  }

  public render(): React.ReactElement {
    const { loading, loadError, interviews } = this.state;
    const now = new Date();

    if (loading) {
      return (
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: 220 } }}>
          <Spinner size={SpinnerSize.large} label="Loading your interviews…" />
        </Stack>
      );
    }

    if (loadError) {
      return (
        <Stack styles={{ root: { padding: 24 } }}>
          <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>
        </Stack>
      );
    }

    const upcoming = interviews.filter(iv => !iv.scheduledDate || new Date(iv.scheduledDate) >= now);
    const past = interviews.filter(iv => !!iv.scheduledDate && new Date(iv.scheduledDate) < now);

    return (
      <Stack styles={{ root: { padding: 24, maxWidth: 920 } }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 20 } }}>
          <Text styles={{ root: { fontSize: 20, fontWeight: 600, color: '#201f1e' } }}>
            My Scheduled Interviews
          </Text>
          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            onClick={() => { this._loadData().catch(() => undefined); }}
          />
        </Stack>

        {interviews.length === 0 ? (
          <Stack horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { height: 220 } }}>
            <Icon iconName="Calendar" styles={{ root: { fontSize: 52, color: '#c8c6c4' } }} />
            <Text styles={{ root: { fontSize: 16, color: '#605e5c' } }}>No interviews scheduled for you</Text>
            <Text styles={{ root: { fontSize: 13, color: '#a19f9d' } }}>
              Interviews assigned to {this.props.currentUser.email} will appear here
            </Text>
          </Stack>
        ) : (
          <>
            {upcoming.length > 0 && (
              <Stack styles={{ root: { marginBottom: 28 } }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
                  <Text styles={{ root: { fontSize: 15, fontWeight: 700, color: '#0078d4' } }}>Upcoming</Text>
                  <span style={{ background: '#0078d4', color: '#fff', borderRadius: 10, padding: '1px 8px', fontSize: 11, fontWeight: 600 }}>
                    {upcoming.length}
                  </span>
                </Stack>
                {upcoming.map(iv => this._renderInterviewCard(iv))}
              </Stack>
            )}

            {past.length > 0 && (
              <Stack>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
                  <Text styles={{ root: { fontSize: 15, fontWeight: 700, color: '#605e5c' } }}>Past</Text>
                  <span style={{ background: '#8a8886', color: '#fff', borderRadius: 10, padding: '1px 8px', fontSize: 11, fontWeight: 600 }}>
                    {past.length}
                  </span>
                </Stack>
                {past.map(iv => this._renderInterviewCard(iv))}
              </Stack>
            )}
          </>
        )}

        {this._renderFeedbackPanel()}
      </Stack>
    );
  }
}
