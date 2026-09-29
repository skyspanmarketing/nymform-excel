// A status message. With `live`, it sits in an aria-live region that exists before its text does,
// so screen readers announce the audit and gate results.
import { makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import {
  CheckmarkCircle20Filled,
  DismissCircle20Filled,
  Info20Filled,
  Warning20Filled,
} from "@fluentui/react-icons";
import { useEffect, useState, type ReactNode } from "react";

export type Intent = "success" | "error" | "warning" | "info";

const useStyles = makeStyles({
  box: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-start",
    padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    minWidth: 0,
  },
  success: {
    backgroundColor: tokens.colorStatusSuccessBackground1,
    border: `1px solid ${tokens.colorStatusSuccessBorder1}`,
  },
  error: {
    backgroundColor: tokens.colorStatusDangerBackground1,
    border: `1px solid ${tokens.colorStatusDangerBorder1}`,
  },
  warning: {
    backgroundColor: tokens.colorStatusWarningBackground1,
    border: `1px solid ${tokens.colorStatusWarningBorder1}`,
  },
  info: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  icon: { flexShrink: 0, marginTop: "1px" },
  iconSuccess: { color: tokens.colorStatusSuccessForeground1 },
  iconError: { color: tokens.colorStatusDangerForeground1 },
  iconWarning: { color: tokens.colorStatusWarningForeground3 },
  iconInfo: { color: tokens.colorNeutralForeground3 },
  body: { display: "flex", flexDirection: "column", gap: "4px", minWidth: 0, overflowWrap: "anywhere" },
  title: { fontWeight: tokens.fontWeightSemibold },
});

export function StatusLine(props: { intent: Intent; title: ReactNode; children?: ReactNode; live?: boolean }) {
  const s = useStyles();
  const { intent, title, children, live = false } = props;
  // A live region is announced when its content changes, so the text arrives just after mount.
  const [shown, setShown] = useState(!live);
  useEffect(() => {
    if (live) setShown(true);
  }, [live]);

  const Icon =
    intent === "success"
      ? CheckmarkCircle20Filled
      : intent === "error"
        ? DismissCircle20Filled
        : intent === "warning"
          ? Warning20Filled
          : Info20Filled;
  const iconClass =
    intent === "success" ? s.iconSuccess : intent === "error" ? s.iconError : intent === "warning" ? s.iconWarning : s.iconInfo;

  const content = (
    <div className={mergeClasses(s.box, s[intent])}>
      <Icon className={mergeClasses(s.icon, iconClass)} aria-hidden="true" />
      <div className={s.body}>
        <span className={s.title}>{title}</span>
        {children}
      </div>
    </div>
  );
  if (!live) return content;
  return (
    <div role="status" aria-live="polite" aria-atomic="true">
      {shown ? content : null}
    </div>
  );
}
