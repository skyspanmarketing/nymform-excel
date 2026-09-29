// Styles for the task pane. Native Fluent look; the one distinctive element is the offset-line
// pattern in a single accent color (indigo, as in the icon), marking private columns and stand-in
// chips. Two groups of lines, the second shifted: the model gets a representation, not the data.
import { makeStaticStyles, makeStyles, tokens } from "@fluentui/react-components";

/** The one accent color, and the colors derived from it. Text colors meet WCAG AA on their backgrounds. */
const ACCENT_LIGHT = {
  "--nym-accent": "#4b3fd1",
  "--nym-accent-text": "#3a2fb8",
  "--nym-accent-line": "rgba(75, 63, 209, 0.55)",
  "--nym-accent-faint": "rgba(75, 63, 209, 0.16)",
  "--nym-accent-bg": "#f4f3fd",
};
const ACCENT_DARK = {
  "--nym-accent": "#9d94ff",
  "--nym-accent-text": "#c4bfff",
  "--nym-accent-line": "rgba(157, 148, 255, 0.7)",
  "--nym-accent-faint": "rgba(157, 148, 255, 0.2)",
  "--nym-accent-bg": "#26233f",
};

/** Two groups of horizontal lines, the second shifted down: the offset-line motif. */
const OFFSET_LINES = (color: string, gap: number) =>
  `repeating-linear-gradient(to bottom, ${color} 0 1px, transparent 1px ${gap}px)`;

/** The pane fills the webview edge to edge. */
export const usePageStyles = makeStaticStyles({
  "html, body": { margin: 0, padding: 0 },
});

export const useAppStyles = makeStyles({
  light: ACCENT_LIGHT,
  dark: ACCENT_DARK,
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    padding: "8px 12px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  main: {
    flexGrow: 1,
    padding: "12px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    minWidth: 0,
  },
  footer: {
    padding: "8px 12px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  centered: {
    padding: "24px 16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    alignItems: "flex-start",
  },
});

export const useUi = makeStyles({
  stack: { display: "flex", flexDirection: "column", gap: "12px", minWidth: 0 },
  tight: { display: "flex", flexDirection: "column", gap: "4px", minWidth: 0 },
  row: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px" },
  spread: { display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: "8px" },
  heading: {
    margin: 0,
    fontSize: tokens.fontSizeBase400,
    lineHeight: tokens.lineHeightBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  subheading: {
    margin: 0,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  muted: {
    margin: 0,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  /** Inputs that fill their grid cell without spilling out of narrow panes. */
  fill: { width: "100%", minWidth: 0, maxWidth: "100%", boxSizing: "border-box" },
  field: { minWidth: 0 },
  text: { margin: 0, fontSize: tokens.fontSizeBase300, lineHeight: tokens.lineHeightBase300 },
  plainList: { listStyle: "none", margin: 0, padding: 0 },
  columnName: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    fontWeight: tokens.fontWeightSemibold,
    overflowWrap: "anywhere",
    minWidth: 0,
  },
  list: { margin: 0, paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "2px" },
  code: {
    margin: 0,
    padding: "8px 10px",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
    wordBreak: "break-word",
    overflowY: "auto",
    maxHeight: "360px",
    minWidth: 0,
    ":focus-visible": {
      outline: `2px solid ${tokens.colorStrokeFocus2}`,
      outlineOffset: "1px",
    },
  },
  formula: {
    fontSize: tokens.fontSizeBase300,
    maxHeight: "none",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(min(140px, 100%), 1fr))",
    alignItems: "start",
    gap: "8px",
  },
  exportBox: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
  },
});

export const usePattern = makeStyles({
  /** A column card. Private cards get the offset-line strip on the left. */
  card: {
    position: "relative",
    padding: "8px 10px 10px 14px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
  },
  cardPrivate: {
    border: "1px solid var(--nym-accent-line)",
    backgroundColor: "var(--nym-accent-bg)",
  },
  strip: {
    position: "absolute",
    left: "3px",
    top: "6px",
    bottom: "6px",
    width: "7px",
    backgroundImage: `${OFFSET_LINES("var(--nym-accent)", 4)}, ${OFFSET_LINES("var(--nym-accent)", 4)}`,
    backgroundSize: "3px 100%, 3px 100%",
    backgroundPosition: "0 0, 4px 2px",
    backgroundRepeat: "no-repeat",
  },
  /** A stand-in token in the request view. */
  chip: {
    display: "inline",
    padding: "0 3px",
    border: "1px solid var(--nym-accent-line)",
    borderRadius: tokens.borderRadiusSmall,
    color: "var(--nym-accent-text)",
    backgroundColor: "var(--nym-accent-bg)",
    backgroundImage: `${OFFSET_LINES("var(--nym-accent-faint)", 4)}, ${OFFSET_LINES("var(--nym-accent-faint)", 4)}`,
    backgroundSize: "50% 100%, 50% 100%",
    backgroundPosition: "left 0 top 0, right 0 top 2px",
    backgroundRepeat: "no-repeat",
    fontWeight: tokens.fontWeightSemibold,
  },
  /** A small legend swatch with the same motif. */
  swatch: {
    display: "inline-block",
    width: "14px",
    height: "10px",
    marginRight: "6px",
    verticalAlign: "middle",
    borderRadius: "2px",
    border: "1px solid var(--nym-accent-line)",
    backgroundImage: `${OFFSET_LINES("var(--nym-accent)", 3)}, ${OFFSET_LINES("var(--nym-accent)", 3)}`,
    backgroundSize: "50% 100%, 50% 100%",
    backgroundPosition: "left 0 top 0, right 0 top 1px",
    backgroundRepeat: "no-repeat",
  },
  letter: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    minWidth: "24px",
    height: "20px",
    padding: "0 4px",
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
});
