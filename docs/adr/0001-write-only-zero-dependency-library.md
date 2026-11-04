# Write-only .xlsx library with zero runtime dependencies

The existing npm options each fall short: one has had no stable release since 2023 and carries nine runtime dependencies, one left npm, one is lean but single-maintainer. This library writes `.xlsx` only, reading is out of scope, and it ships with no runtime dependencies and no install scripts so its supply-chain surface is the package itself. Every piece of the format it needs (zip container, CRC32, XML, compression through platform APIs) is implemented or reached in-house.
