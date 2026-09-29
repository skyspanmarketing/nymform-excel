// log.ts (spec §7.12): in-memory ring buffer of 200 entries, update by id, JSON export.
import { describe, expect, it } from "vitest";
import { LOG_CAPACITY, RequestLog, type LogLocal } from "../src/core/log";
import type { LogEntry } from "../src/core/types";

function entry(n: number, extra: Partial<LogEntry> = {}): LogEntry {
  return {
    at: new Date(Date.UTC(2026, 8, 24, 12, 0, n)).toISOString(),
    mode: "structure_only",
    model: "openai/gpt-6-luna",
    host: "openrouter.ai",
    bytes: 100 + n,
    body: `{"model":"openai/gpt-6-luna","n":${n}}`,
    audit: { ok: true, hits: [], structural: [] },
    inserted: false,
    ...extra,
  };
}

describe("RequestLog", () => {
  it("holds 200 entries and drops the oldest past that", () => {
    const log = new RequestLog();
    expect(LOG_CAPACITY).toBe(200);
    const ids: number[] = [];
    for (let i = 0; i < 205; i++) ids.push(log.add(entry(i)));
    expect(new Set(ids).size).toBe(205);
    expect(log.size).toBe(200);
    const all = log.entries();
    expect(all).toHaveLength(200);
    expect(all[0]!.bytes).toBe(105);
    expect(all[199]!.bytes).toBe(304);
    expect(log.get(ids[0]!)).toBeUndefined();
    expect(log.get(ids[204]!)?.bytes).toBe(304);
  });

  it("updates an entry by id and ignores ids that are gone", () => {
    const log = new RequestLog();
    const a = log.add(entry(1));
    const b = log.add(entry(2));
    log.update(b, {
      replyRaw: '{"choices":[]}',
      gate: { ok: false, reasons: [{ code: "G-FUNC", detail: "Uses WEBSERVICE, which can send data over the internet." }] },
      latencyMs: 812,
      tokens: { prompt: 900, completion: 40 },
      providerName: "Azure",
    });
    log.update(b, { inserted: true });
    log.update(999, { inserted: true });
    expect(log.get(a)).toEqual(entry(1));
    expect(log.get(b)).toMatchObject({ inserted: true, latencyMs: 812, providerName: "Azure", body: entry(2).body });
    expect(log.entries()).toHaveLength(2);
  });

  it("returns a new entry object on update, so views can tell it changed", () => {
    const log = new RequestLog();
    const id = log.add(entry(1));
    const before = log.get(id);
    log.update(id, { inserted: true });
    expect(log.get(id)).not.toBe(before);
    expect(before?.inserted).toBe(false);
  });

  it("exports a pretty JSON array of the entries, oldest first", () => {
    const log = new RequestLog();
    log.add(entry(1));
    log.add(entry(2, { error: "Rate limited. Wait a minute and try again." }));
    const json = log.exportJson();
    expect(json).toBe(JSON.stringify([entry(1), entry(2, { error: "Rate limited. Wait a minute and try again." })], null, 2));
    expect(JSON.parse(json)).toHaveLength(2);
    expect(json).toContain("\n  {");
  });

  it("exports an empty array when there is nothing logged", () => {
    expect(new RequestLog().exportJson()).toBe("[]");
  });

  it("keeps only LogEntry fields, so a key or the stand-in map can't ride along", () => {
    const log = new RequestLog();
    const sneaky = { ...entry(1), key: "sk-or-v1-secret", map: { PERSON_001: "Maria Lopez" } } as LogEntry;
    const id = log.add(sneaky);
    log.update(id, { apiKey: "sk-or-v1-secret" } as Partial<LogEntry>);
    const json = log.exportJson();
    expect(json).not.toContain("sk-or-v1-secret");
    expect(json).not.toContain("Maria Lopez");
    expect(Object.keys(log.get(id)!)).toEqual(Object.keys(entry(1)));
  });

  it("clear empties the log", () => {
    const log = new RequestLog();
    log.add(entry(1));
    log.clear();
    expect(log.entries()).toEqual([]);
  });
});

describe("RequestLog: pane-only notes beside an entry", () => {
  const blockedEntry = entry(2, {
    body: "",
    audit: { ok: false, hits: [{ column: "A", variant: "word" }], structural: [], error: "Text from private column A is in the request. Nothing was sent." },
    error: "Blocked: a value from column A is in the request.",
  });
  const local: LogLocal = {
    matches: [{ column: "A", variant: "word", origin: "value", text: "Maria", before: "Total for ", after: " this year", where: "your question" }],
  };

  it("items() pairs each entry with its notes, oldest first", () => {
    const log = new RequestLog();
    log.add(entry(1));
    const id = log.add(blockedEntry, local);
    expect(log.items()).toEqual([
      { entry: entry(1), local: undefined },
      { entry: blockedEntry, local },
    ]);
    // An update keeps the notes with their entry.
    log.update(id, { inserted: false });
    expect(log.items()[1]?.local).toEqual(local);
  });

  it("entries(), get() and exportJson() never hold the notes", () => {
    const log = new RequestLog();
    const id = log.add(blockedEntry, local);
    expect(log.entries()).toEqual([blockedEntry]);
    expect(log.get(id)).toEqual(blockedEntry);
    expect(Object.keys(log.get(id)!)).toEqual(Object.keys(blockedEntry));
    const json = log.exportJson();
    expect(json).toBe(JSON.stringify([blockedEntry], null, 2));
    for (const text of ["Maria", "Total for", "this year", "your question", "matches"]) expect(json).not.toContain(text);
  });

  it("keeps a copy, so the caller can't change the notes afterwards", () => {
    const log = new RequestLog();
    const mine = { matches: [{ ...local.matches![0]! }] };
    log.add(blockedEntry, mine);
    mine.matches[0]!.text = "changed";
    mine.matches.push({ ...local.matches![0]! });
    expect(log.items()[0]?.local).toEqual(local);
  });

  it("drops the notes with their entry past the capacity, and clear() empties both", () => {
    const log = new RequestLog();
    log.add(blockedEntry, local);
    for (let i = 0; i < LOG_CAPACITY; i++) log.add(entry(i));
    expect(log.items().some((i) => i.local !== undefined)).toBe(false);
    log.add(blockedEntry, local);
    log.clear();
    expect(log.items()).toEqual([]);
    expect(log.entries()).toEqual([]);
    expect(log.exportJson()).toBe("[]");
  });
});
