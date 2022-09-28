# VBA Macros

This repository is for various VBA macros.

## MathML Macro
The MathML macro converts all equations into regular mathml text. It should be run on a word document before pasting into Dreamweaver.

To run the macro for the first time, you will have to add the MSForms library. You can do this with the following steps:
1. If you do not already have the Developer tab in the Word toolbar, go to File -> Options -> Customize Ribbon and check "Developer" under Main Tabs.
2. In the Developer tab, click "Visual Basic". This is where you run your macros.
3. In the Visual Basic window, go to Tools -> References -> Browse, go to the System32 folder (type System32 into the address), and select the "FM20.DLL" file.
4. Back in the References window, check Microsoft Forms 2.0 Object Library, which should be at the bottom of the list.
5. Restart Word.

To run macros in general:
1. When you open a Word file, immediately select "Enable Content".
2. Open Visual Basic from the Developer tab.
3. Go to Insert -> Module and paste the macro code in.
4. Click the green "Run" button and select the macro you want to run.
    - If this doesn't work and you get an error message that says "DataObject:GetFromClipboard OpenClipboard Failed", then in the Visual Basic editor, click to the left of every line in the macro (a red dot should appear) and hold f5 to run through each line of the macro step-by-step in debug mode.

If these steps don't work, you may have to restart Word and retry.

After pasting into Dreamweaver, the angle brackets <> will be converted into html entities, so you will have to replace them manually using the following regex statements to convert them back into tags. Converting them into tags can also be done with the [general dreamweaver formatting tool](https://commwebteam.github.io/gen_dw_format/dreamweaver_paste_formatter/dw_paste_format.html).
1. Replace &amp;lt;(/&ast;math.&ast;?)&amp;gt; with &lt;$1&gt;
2. Replace &amp;lt;(/&ast;m[ion].&ast;?)&amp;gt; with &lt;$1&gt;

## Textbox Macro
The textbox macro (sourced from [here](https://word.tips.net/T001690_Removing_All_Text_Boxes_In_a_Document.html)) moves text out of textboxes and into the main document (in their original location), and then deletes the now-empty textboxes. This does not preserve any formatting in those textboxes, so it should only be used for Beyond Comparing.

## Table Mover Macro
1. Paste the document for which you want to have the tables removed
 <img width="229" alt="Capturedddd" src="https://user-images.githubusercontent.com/56009508/192818277-7ad3850e-d8d6-46eb-a544-1e3cdf13aa04.PNG">

2. Ensure that Developer mode is enabled with design mode off. For further instructions please refer to this link ([https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word](https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792#:~:text=The%20Developer%20tab%20isn't,select%20the%20Developer%20check%20box))
<img width="263" alt="Captureppp" src="https://user-images.githubusercontent.com/56009508/192818785-f84edb11-dcaf-43e1-953b-7e8b998b382c.PNG">

3. Click the "Move Tables to New Document" Button
<img width="428" alt="Capturegggg" src="https://user-images.githubusercontent.com/56009508/192819533-06d89821-4331-4342-a205-791d9b4fa0ed.PNG">

4. New document should be created with only the tables from the document. Feel free to rename the new document to your liking!
