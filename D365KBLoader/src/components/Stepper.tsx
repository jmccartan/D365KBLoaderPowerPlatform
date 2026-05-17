import { Fragment } from 'react';
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';
import { Checkmark16Filled } from '@fluentui/react-icons';

const useStyles = makeStyles({
  wrap: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
  step: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    border: 'none',
    background: 'transparent',
    cursor: 'pointer',
    borderRadius: tokens.borderRadiusCircular,
    color: tokens.colorNeutralForeground2,
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    transitionProperty: 'background-color, color',
    transitionDuration: tokens.durationNormal,
    ':hover:not(:disabled)': {
      backgroundColor: tokens.colorSubtleBackgroundHover,
    },
    ':disabled': {
      cursor: 'not-allowed',
      opacity: 0.45,
    },
  },
  stepActive: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  bubble: {
    width: '28px',
    height: '28px',
    borderRadius: '50%',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: tokens.colorNeutralBackground4,
    color: tokens.colorNeutralForeground2,
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    transitionProperty: 'background-color, color, border-color',
    transitionDuration: tokens.durationNormal,
  },
  bubbleActive: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderTopColor: tokens.colorBrandBackground,
    borderRightColor: tokens.colorBrandBackground,
    borderBottomColor: tokens.colorBrandBackground,
    borderLeftColor: tokens.colorBrandBackground,
    boxShadow: `0 0 0 4px ${tokens.colorBrandBackground2}`,
  },
  bubbleDone: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
    color: tokens.colorNeutralForegroundOnBrand,
    borderTopColor: tokens.colorPaletteGreenBackground3,
    borderRightColor: tokens.colorPaletteGreenBackground3,
    borderBottomColor: tokens.colorPaletteGreenBackground3,
    borderLeftColor: tokens.colorPaletteGreenBackground3,
  },
  connector: {
    width: '28px',
    height: '2px',
    backgroundColor: tokens.colorNeutralStroke2,
    borderRadius: '1px',
  },
  connectorDone: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
  },
  label: {
    whiteSpace: 'nowrap',
  },
  count: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginLeft: tokens.spacingHorizontalXS,
  },
});

export interface StepperStep {
  value: string;
  label: string;
  count?: number;
  disabled?: boolean;
}

export interface StepperProps {
  steps: StepperStep[];
  active: string;
  onSelect: (value: string) => void;
}

export function Stepper({ steps, active, onSelect }: StepperProps) {
  const s = useStyles();
  const activeIndex = steps.findIndex(st => st.value === active);
  return (
    <div className={s.wrap} role="tablist" aria-label="Workflow steps">
      {steps.map((st, i) => {
        const isActive = i === activeIndex;
        const isDone = i < activeIndex;
        return (
          <Fragment key={st.value}>
            <button
              role="tab"
              aria-selected={isActive}
              disabled={st.disabled}
              className={mergeClasses(s.step, isActive && s.stepActive)}
              onClick={() => !st.disabled && onSelect(st.value)}
            >
              <span
                className={mergeClasses(
                  s.bubble,
                  isActive && s.bubbleActive,
                  isDone && s.bubbleDone,
                )}
              >
                {isDone ? <Checkmark16Filled /> : i + 1}
              </span>
              <span className={s.label}>{st.label}</span>
              {typeof st.count === 'number' && (
                <span className={s.count}>({st.count})</span>
              )}
            </button>
            {i < steps.length - 1 && (
              <span
                className={mergeClasses(s.connector, isDone && s.connectorDone)}
                aria-hidden
              />
            )}
          </Fragment>
        );
      })}
    </div>
  );
}
