# Identical input gives identical bytes

Zip entry timestamps are fixed to the DOS epoch, parts are written in a fixed order, no document-properties parts (which would carry a creation time) are emitted, and STORE output has no implementation-dependent compressor. This makes golden fixtures byte-exact tests, makes output cacheable, and removes a class of flaky comparisons. Readers do not care about entry timestamps inside an `.xlsx`.
