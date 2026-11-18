---
"xlsx-mini": patch
---

Text reads back as written when it looks like the format's own escape or holds a carriage return. Excel, LibreOffice and Numbers read `_x0041_` in a cell as `A`, `part _x00e6_ no` as `part æ no` and `_x005F_` as `_`. Every XML parser also reads a carriage return as a line feed, so `"a\rb"` came back as `"a\nb"`. The writer now escapes such text in shared strings, inline strings, a formula's cached text, and validation titles and messages, and it writes a carriage return as a character reference everywhere. Other text keeps its bytes. Numbers does not read the escape in inline strings, so a text like `_x0041_` that `streamXlsx` writes inline shows escaped in Numbers. Stream with `strings: "shared"` to avoid that.
