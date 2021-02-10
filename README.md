# VBA macros

This repository is for various VBA macros.

# MathML macro
The MathML macro converts all equations into regular mathml text. It should be run on a word document before pasting into Dreamweaver.

After pasting into Dreamweaver, the angle brackets <> will be converted into html entities, so you will have to replace them manually using the following regex statements:
1. Replace &amp;lt;(math.&ast;?)&amp;gt; with &lt;$1&gt;
2. Replace &amp;lt;(m[ion].&ast;?)&amp;gt; with &lt;$1&gt;

# Textbox macro
The textbox macro moves text out of textboxes and into the main document (in their original location), and then deletes the now-empty textboxes. This does not preserve any formatting in those textboxes, so it should only be used for Beyond Comparing.
