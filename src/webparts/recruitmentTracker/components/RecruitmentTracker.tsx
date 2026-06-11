import * as React from 'react';
import {
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
} from '@fluentui/react';
import type { IRecruitmentTrackerProps } from './IRecruitmentTrackerProps';
import { PostJobForm } from './PostJob/PostJobForm';
import { TrackProgress } from './TrackProgress/TrackProgress';
import { CompletedJobs } from './CompletedJobs/CompletedJobs';
import { AllCandidates } from './AllCandidates/AllCandidates';
import { isHREmail } from './shared/EmailService';
import { SpService } from './shared/SpService';

interface IRecruitmentTrackerState {
  activeTab: string;
  isAllowedPoster: boolean;
  checkingPermissions: boolean;
}

export default class RecruitmentTracker extends React.Component<
  IRecruitmentTrackerProps,
  IRecruitmentTrackerState
> {
  private _spService: SpService;

  constructor(props: IRecruitmentTrackerProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this.state = {
      activeTab: 'postJob',
      isAllowedPoster: false,
      checkingPermissions: true,
    };
  }

  public async componentDidMount(): Promise<void> {
    const { currentUser } = this.props;
    if (currentUser?.email) {
      await this._checkPermissions(currentUser.email);
    }
  }

  public async componentDidUpdate(prevProps: IRecruitmentTrackerProps): Promise<void> {
    const prevEmail = prevProps.currentUser?.email ?? '';
    const currEmail = this.props.currentUser?.email ?? '';
    if (!prevEmail && currEmail) {
      await this._checkPermissions(currEmail);
    }
  }

  private async _checkPermissions(email: string): Promise<void> {
    const defaultTab = isHREmail(email) ? 'trackProgress' : '';
    try {
      const allowed = await this._spService.isAllowedPoster(email);
      this.setState({
        isAllowedPoster: allowed,
        checkingPermissions: false,
        activeTab: allowed ? 'postJob' : defaultTab,
      });
    } catch {
      this.setState({ isAllowedPoster: false, checkingPermissions: false, activeTab: defaultTab });
    }
  }

  public render(): React.ReactElement<IRecruitmentTrackerProps> {
    const { sp, graphService, currentUser, applyBaseUrl } = this.props;
    const { activeTab, isAllowedPoster, checkingPermissions } = this.state;
    const isHR = !!currentUser && isHREmail(currentUser.email);

    // Show spinner until both the user profile and permissions are resolved
    if (!currentUser?.email || checkingPermissions) {
      return (
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: 200 } }}>
          <Spinner size={SpinnerSize.large} label="Loading Recruitment Tracker…" />
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 0 }}>
        {/* Header bar */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { background: '#fd800b', padding: '12px 24px' } }}
        >
          <Text styles={{ root: { color: '#fff', fontSize: 18, fontWeight: 600 } }}>
            OpRea — Recruitment Tracker
          </Text>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
            <Text styles={{ root: { color: '#fff', fontSize: 13 } }}>
              {currentUser.displayName}
            </Text>
            {isHR && (
              <span style={{ background: 'rgba(255,255,255,0.25)', color: '#fff', padding: '2px 10px', borderRadius: 12, fontSize: 11, fontWeight: 600 }}>
                HR
              </span>
            )}
            {isAllowedPoster && (
              <span style={{ background: 'rgba(255,255,255,0.25)', color: '#fff', padding: '2px 10px', borderRadius: 12, fontSize: 11, fontWeight: 600 }}>
                Poster
              </span>
            )}
          </Stack>
        </Stack>

        {/* Tab navigation */}
        <Pivot
          selectedKey={activeTab}
          onLinkClick={item => item && this.setState({ activeTab: item.props.itemKey ?? '' })}
          styles={{
            root: { paddingLeft: 16 },
            link: { selectors: { '&.ms-Pivot-link.is-selected': { borderBottomColor: '#fd800b', color: '#fd800b' } } },
          }}
        >
          {/* Only AllowedPosters see the Add Job Description tab */}
          {isAllowedPoster && (
            <PivotItem headerText="Add Job Description" itemKey="postJob" itemIcon="Add">
              <PostJobForm
                sp={sp}
                graphService={graphService}
                currentUser={currentUser}
                applyBaseUrl={applyBaseUrl}
              />
            </PivotItem>
          )}

          {/* HR-only tabs */}
          {isHR && (
            <PivotItem headerText="Track Progress" itemKey="trackProgress" itemIcon="ClipboardList">
              <TrackProgress sp={sp} graphService={graphService} />
            </PivotItem>
          )}

          {isHR && (
            <PivotItem headerText="Completed Jobs" itemKey="completedJobs" itemIcon="CheckMark">
              <CompletedJobs sp={sp} />
            </PivotItem>
          )}

          {isHR && (
            <PivotItem headerText="All Candidates" itemKey="allCandidates" itemIcon="People">
              <AllCandidates sp={sp} />
            </PivotItem>
          )}
        </Pivot>
      </Stack>
    );
  }
}
