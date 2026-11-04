# ESM-only package emitted by tsc, no bundler

The package ships ESM only, with an `exports` map whose `types` condition comes first, `sideEffects: false`, declarations from `tsc`, and `dist/` as the only published code. Node 20.19+ and 22.12+ can `require()` it, so CommonJS consumers on those versions still work; a dual build would add a second artifact to test for little gain and can be added later without an API change. TypeScript targets ES2022 with the strict family of flags on.
