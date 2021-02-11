# VBA macros

This repository is for various VBA macros.

# MathML macro
The MathML macro converts all equations into regular mathml text. It should be run on a word document before pasting into Dreamweaver.

To run the macro for the first time, you will have to add the MSForms library. You can do this with the following steps:
1. If you do not already have the Developer tab in the Word toolbar, go to File -> Options -> Customize Ribbon and check "Developer" under Main Tabs.
2. In the Developer tab, click "Visual Basic". This is where you run your macros.
3. In the Visual Basic window, go to Tools -> References -> Browse, go to the System32 folder (type System32 into the address), and select (a file - forgot its name).
4. Back in the References window, check Microsoft Forms 2.0 Object Library, which should be at the bottom of the list.
5. Restart Word.

To run macros in general:
1. When you open a word file, immediately select "Enable Content".
2. Open Visual Basic from the Developer tab.
3. Go to Insert -> Module and paste the macro code in.
4. Click the green "Run" button and select the macro you want to run.

If these steps don't work, you may have to restart word.

After pasting into Dreamweaver, the angle brackets <> will be converted into html entities, so you will have to replace them manually using the following regex statements:
1. Replace &amp;lt;(/&ast;math.&ast;?)&amp;gt; with &lt;$1&gt;
2. Replace &amp;lt;(/&ast;m[ion].&ast;?)&amp;gt; with &lt;$1&gt;

This can also be done with the [general dreamweaver formatting tool](https://commwebteam.github.io/gen_dw_format/basic_format.html).

# Textbox macro
The textbox macro moves text out of textboxes and into the main document (in their original location), and then deletes the now-empty textboxes. This does not preserve any formatting in those textboxes, so it should only be used for Beyond Comparing.
