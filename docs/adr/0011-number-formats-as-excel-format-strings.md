# Number formats are Excel format strings or built-in ids

`numberFormat` accepts an Excel format string, registered once in the styles part from id 164 upward and deduplicated, or a built-in id from 0 to 49. There is no preset object; recipes for currency, percent and ISO dates live in the README. This keeps the surface small and lets every format Excel supports be expressed.
