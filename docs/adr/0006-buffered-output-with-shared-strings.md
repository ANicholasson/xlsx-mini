# Buffered output and a shared-strings table

Version 1 produces the whole file in memory; there is no streaming or row-at-a-time API. That choice allows a shared-strings table, which is what Excel itself writes, has the broadest reader compatibility, and shrinks files with repeated text. Sheet XML is still generated row by row internally so memory stays linear in the data. Streaming, if ever needed, is a later major version and would move to inline strings.
