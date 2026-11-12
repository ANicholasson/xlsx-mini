# Version 1 scope and what is deliberately deferred

Version 1 covers multiple sheets, typed cells (text, number, boolean, Date, Temporal), bold and number format on any cell or column, column widths, freeze panes and autofilter. Deferred but designed for: PNG images as drawing parts (the largest piece, after 1.0), merged cells, borders, fills, font colour, alignment, wrap, row heights, hyperlinks, formulas. When formulas arrive they are an explicit cell type and are never inferred from text starting with `=`, so pasted user data can never become a formula.

Update: merged cells, borders, fills, font colour and size, italic, alignment, wrap, indent and row heights shipped in 0.2.0 as additive style and sheet fields. Still deferred: images, hyperlinks, formulas.

Update: hyperlinks and formulas shipped in 0.4.0 as explicit cell types, with data validation and print setup at sheet level (ADR 0023). Still deferred: images.

Update: images shipped in 0.6.0 as a cell type anchored to its cell (ADR 0025). Nothing from this list remains deferred.

Update: what the library will not do at all, as opposed to not yet, is recorded in ADR 0026.
