import "@testing-library/jest-dom/vitest";

let counter = 0;

if (!globalThis.crypto?.randomUUID) {
  Object.defineProperty(globalThis, "crypto", {
    value: {
      randomUUID: () => `test-id-${(counter += 1)}`
    }
  });
}
