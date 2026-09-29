import { defineConfig } from "vitest/config";

// Build-time constants, as webpack's DefinePlugin sets them. Tests run the release shape
// (bench off) against the default OpenRouter endpoint.
export default defineConfig({
  define: {
    __NYMFORM_BENCH__: "false",
    __NYMFORM_CONFIG__: JSON.stringify({
      endpoint: "https://openrouter.ai/api/v1",
      defaultModel: "openai/gpt-6-luna",
      version: "0.1.0-alpha",
      commit: "test",
    }),
  },
  test: {
    include: ["tests/**/*.test.ts", "tests/**/*.test.tsx"],
    environment: "node",
  },
});
