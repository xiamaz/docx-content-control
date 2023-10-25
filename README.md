# docx-content-control
Rust package to manipulate content control/placeholders and nothing else.

## Who this package is for

Your users need word documents, you don't want to do that in an Windows environment. You don't want to enable or have macros in your word files. You want to support this use-case with as much flexibility as possible.

Features: 

- Find and fill all ContentControl types
- Find and fill all explicit Placeholders with specified syntax
- Support simple markup for placed content (font-size, bold/italic, super- and subscript)
- Retain all other docx structers

Non-Goals:

- Creation of word documents
- Placing text anywhere else
- Parsing word documents

## How it works

This tool parses the DOCX xml and will explicitly search for either placeholders or ContentControl type fields. These will be replaced with the replacement text.

## Further reading

The Office Open XML format has been published as an ISO standard, while normally
these cost an arm and a leg, they have gracefully decided for our relevant
standard to be [publicly available](https://standards.iso.org/ittf/PubliclyAvailableStandards/index.html).

The XML format itself is published as ISO/IEC 29500-1:2016. Our package aims to
be fully conformant with
