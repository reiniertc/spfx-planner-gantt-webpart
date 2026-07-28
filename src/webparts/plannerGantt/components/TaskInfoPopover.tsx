import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import * as strings from 'PlannerGanttWebPartStrings';
import { PlannerService } from '../services/PlannerService';
import { IPlannerTaskComment } from '../models/IPlannerModels';
import styles from './PlannerGantt.module.scss';

/** Which info sections the property pane has turned on. */
export interface ITaskInfoOptions {
  show: boolean;
  showDescription: boolean;
  showStartDate: boolean;
  showEndDate: boolean;
  showAssignee: boolean;
  showStatus: boolean;
  showComments: boolean;
}

export interface ITaskInfoData {
  taskId: string;
  planId: string;
  start: Date;
  end: Date;
  progress: number;
  assignees: string[];
  hasDescription: boolean;
  hasConversation: boolean;
}

function statusLabel(progress: number): string {
  if (progress >= 100) {
    return strings.StatusCompletedLabel;
  }
  if (progress <= 0) {
    return strings.StatusNotStartedLabel;
  }
  return strings.StatusInProgressLabel;
}

interface ILoadState<T> {
  status: 'idle' | 'loading' | 'loaded';
  value?: T;
}

export const TaskInfoButton: React.FC<{
  data: ITaskInfoData;
  options: ITaskInfoOptions;
  plannerService: PlannerService;
}> = ({ data, options, plannerService }) => {
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [description, setDescription] = React.useState<ILoadState<string>>({ status: 'idle' });
  const [comments, setComments] = React.useState<ILoadState<IPlannerTaskComment[]>>({ status: 'idle' });
  const buttonRef: React.RefObject<HTMLButtonElement> = React.useRef<HTMLButtonElement>(null);

  const openCallout = (): void => {
    setIsOpen(true);

    if (options.showDescription && data.hasDescription && description.status === 'idle') {
      setDescription({ status: 'loading' });
      plannerService.getTaskNotes(data.taskId)
        .then(text => setDescription({ status: 'loaded', value: text }))
        .catch(() => setDescription({ status: 'loaded', value: '' }));
    }

    if (options.showComments && data.hasConversation && comments.status === 'idle') {
      setComments({ status: 'loading' });
      plannerService.getTaskComments(data.planId, data.taskId)
        .then(list => setComments({ status: 'loaded', value: list }))
        .catch(() => setComments({ status: 'loaded', value: [] }));
    }
  };

  return (
    <>
      <IconButton
        elementRef={buttonRef}
        className={styles.taskInfoButton}
        iconProps={{ iconName: 'Info' }}
        title={strings.TaskInfoButtonTitle}
        ariaLabel={strings.TaskInfoButtonTitle}
        onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
          event.stopPropagation();
          openCallout();
        }}
      />
      {isOpen && (
        <Callout
          target={buttonRef}
          directionalHint={DirectionalHint.rightTopEdge}
          onDismiss={() => setIsOpen(false)}
          setInitialFocus={true}
        >
          <div className={styles.taskInfoCallout}>
            {options.showDescription && (
              <div className={styles.taskInfoSection}>
                {description.status === 'loading' ? (
                  <Spinner size={SpinnerSize.small} label={strings.TaskInfoLoadingLabel} />
                ) : (
                  <div className={styles.taskInfoDescription}>
                    {(data.hasDescription && description.value) || strings.TaskInfoNoDescriptionLabel}
                  </div>
                )}
              </div>
            )}
            {options.showStartDate && (
              <div className={styles.taskInfoSection}>
                <span className={styles.taskInfoLabel}>{strings.TaskInfoStartLabel}</span>
                <span>{data.start.toLocaleDateString()}</span>
              </div>
            )}
            {options.showEndDate && (
              <div className={styles.taskInfoSection}>
                <span className={styles.taskInfoLabel}>{strings.TaskInfoEndLabel}</span>
                <span>{data.end.toLocaleDateString()}</span>
              </div>
            )}
            {options.showAssignee && (
              <div className={styles.taskInfoSection}>
                <span className={styles.taskInfoLabel}>{strings.TaskInfoAssigneeLabel}</span>
                <span>{data.assignees.join(', ') || '–'}</span>
              </div>
            )}
            {options.showStatus && (
              <div className={styles.taskInfoSection}>
                <span className={styles.taskInfoLabel}>{strings.TaskInfoStatusLabel}</span>
                <span>{statusLabel(data.progress)}</span>
              </div>
            )}
            {options.showComments && (
              <div className={styles.taskInfoSection}>
                <span className={styles.taskInfoLabel}>{strings.TaskInfoCommentsHeading}</span>
                {comments.status === 'loading' ? (
                  <Spinner size={SpinnerSize.small} label={strings.TaskInfoLoadingLabel} />
                ) : (
                  <div className={styles.taskInfoChats}>
                    {!data.hasConversation || !comments.value || comments.value.length === 0 ? (
                      <div className={styles.taskInfoNoComments}>{strings.TaskInfoNoCommentsLabel}</div>
                    ) : (
                      comments.value.map((comment, index) => (
                        <div className={styles.taskInfoChatItem} key={index}>
                          <div>
                            <span className={styles.taskInfoChatFrom}>{comment.from}</span>
                            <span className={styles.taskInfoChatDate}>{new Date(comment.createdDateTime).toLocaleString()}</span>
                          </div>
                          <div>{comment.bodyText}</div>
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        </Callout>
      )}
    </>
  );
};
