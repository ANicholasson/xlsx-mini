# Date values use UTC semantics; Temporal values are accepted by structure

A `Date` becomes an Excel serial from its UTC fields with plain arithmetic, no `Intl` call per cell, and gets the built-in date-time format 22. Temporal values are accepted because their types carry the intent Excel needs: `PlainDate` is date-only (serial without time, format 14), `PlainDateTime` and `ZonedDateTime` use their wall-clock fields (format 22), `Instant` is UTC. They are detected by `Symbol.toStringTag`, which native engines and both polyfills set, so the library never references a `Temporal` global and needs no polyfill of its own; consumers type them against small structural interfaces. The 1900 date system is used with Excel's inherited leap-year quirk; out-of-range dates throw. A time-zone option is deferred.

## Considered options

- Local-time semantics for `Date`: matches what a browser user sees, but differs per server and is not reproducible.
- A per-cell `Intl.DateTimeFormat`: supports zones but costs a formatter call on the hot path.
