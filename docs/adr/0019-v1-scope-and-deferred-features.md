# Version 1 scope and what is deliberately deferred

Version 1 covers multiple sheets, typed cells (text, number, boolean, Date, Temporal), bold and number format on any cell or column, column widths, freeze panes and autofilter. Deferred but designed for: PNG images as drawing parts (the largest piece, after 1.0), merged cells, borders, fills, font colour, alignment, wrap, row heights, hyperlinks, formulas. When formulas arrive they are an explicit cell type and are never inferred from text starting with `=`, so pasted user data can never become a formula.
