# Convert word processor files to InDesign tagged text

(WORK IN PROGRESS)

This little Python 3 script converts a word processor file (currently .docx,
.odt, .fodt, or Markdown) into an InDesign tagged text format.

Currently usable only if you know how to install Python 3, python-lxml, etc.
Not difficult at all, but I didn't get around to writing a HOWTO or a proper
installer.

Main features:

- You can specify an arbitrary mapping from Word (character/paragraph) styles
  to InDesign styles, including consistently creating style groups in InDesign.

- You can specify "contextual rules". You'll love them, trust me.

- Importing a tagged text file is _much_ faster in InDesign.

- The sibling InDesign script to this will rerun the conversion, so you'll need
  to do the manual conversion only once.


Missing features (highlights):
- Support to preserve manual formatting
- Tables
- Endnotes
- Linked styles


(C) 2018 Erez Volk
