// Text with stand-in tokens shown as chips carrying the offset-line pattern.
import { makeStyles } from "@fluentui/react-components";
import { Fragment } from "react";
import type { Segment } from "../flow";
import { usePattern } from "../styles";

const useStyles = makeStyles({
  // Wrapped lines keep their indent: a hanging indent per line.
  line: { textIndent: "-2ch", whiteSpace: "pre-wrap" },
});

function Segments({ segments }: { segments: readonly Segment[] }) {
  const p = usePattern();
  return (
    <>
      {segments.map((seg, i) =>
        seg.kind === "token" ? (
          <span key={i} className={p.chip} title="Stand-in. The original stays in Excel.">
            {seg.text}
          </span>
        ) : (
          <Fragment key={i}>{seg.text}</Fragment>
        ),
      )}
    </>
  );
}

/** Splits segments into lines, each with its leading spaces counted and removed. */
export function toLines(segments: readonly Segment[]): { indent: number; segments: Segment[] }[] {
  const lines: Segment[][] = [[]];
  for (const seg of segments) {
    if (seg.kind === "token") {
      lines[lines.length - 1]!.push(seg);
      continue;
    }
    const parts = seg.text.split("\n");
    parts.forEach((part, i) => {
      if (i > 0) lines.push([]);
      if (part !== "") lines[lines.length - 1]!.push({ kind: "text", text: part });
    });
  }
  return lines.map((line) => {
    const first = line[0];
    if (first?.kind !== "text") return { indent: 0, segments: line };
    const indent = first.text.length - first.text.trimStart().length;
    const rest = first.text.slice(indent);
    return { indent, segments: rest === "" ? line.slice(1) : [{ kind: "text", text: rest }, ...line.slice(1)] };
  });
}

export function TokenText({ segments, lines = false }: { segments: readonly Segment[]; lines?: boolean }) {
  const s = useStyles();
  if (!lines) return <Segments segments={segments} />;
  return (
    <>
      {toLines(segments).map((line, i) => (
        <div key={i} className={s.line} style={{ paddingLeft: `${line.indent + 2}ch` }}>
          {line.segments.length === 0 ? " " : <Segments segments={line.segments} />}
        </div>
      ))}
    </>
  );
}
