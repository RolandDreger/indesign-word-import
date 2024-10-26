# Word to InDesign
The script **importDoxc.jsx** provides an alternative import for Microsoft Word document into Adobe InDesign. 

# Script usage

Download the script via `Code` ‣ `Download ZIP`

<img width="922" alt="download_zip" src="https://user-images.githubusercontent.com/19747449/173639516-21f37b19-e104-4904-ba95-b74c64877275.png">

Put the unzipped folder with all files in into the script folder of InDesign and start the script **importDoxc.jsx** from the script panel via double click.

# What's the difference?

A crucial difference: No style properties are taken over from Word only the style names and their assignment to the text places. The style properties must be set in InDesign. 

The included images are placed and not embedded. Instead of local overrides, character styles are applied. Comments, table styles or functional references in scientific papers can be imported. And much more.

Here you can find some example videos:

- [Character Styles](https://vimeo.com/698193760)
- [Table Styles](https://vimeo.com/702916024)
- [Index](https://vimeo.com/719616490)
- [Track Changes](https://vimeo.com/681081988)

In some areas, the native import will definitely be better, for instance when it comes to performance or if you simply want to import preformatted content. However, it is by nature a very general approach and so are many of the design decisions behind it.

The way via a script, on the other hand, offers the possibility to configure the import individually, to treat the content differently than InDesign does or even to omit parts of it completely that might cause problems. 

As a user, you can decide to import plain text only and mark the text passages to edit them individually. As programmer you can hook into the different states of import, e.g. if the index entry cases a crash of indesign (because of special characters) you can clean up the entries.

# Document preparation
## Word

It is best to work with paragraph and character styles already in Word. This results in less rework in InDesign and a better XML structure - if you need it.

### Renaming in-built styles

Word does not let you rename the in-built styles such as headings, list styles or some character styles. The problem with this is that Word displays their language-specific names, but in the document English names are stored. And these names appear in InDesign after the import.

This is where alias names come into play: To create an alias or alternative name for the same style, enter a semicolon or comma after the in-built name and then the desired name. Alias names get the priority when importing with the script.

## InDesign

In fact, no special preparation is needed for the InDesign document. But you can create paragraph, character, table and object styles even before importing. If the names match those in Word, the desired formatting will appear immediately after the import.

You can work with or without a primary text frame. Try out what is most suitable for you.

a) Without primary text frame 

- [x] No primary text frame
- [x] Text flow disabled in document settings or limited to primary text frame
- [x] Script setting `isAutoflowing: true` (default)
	
Text flow is created by the script (via the placeXML method before the script then continues). This method is generally preferable.

b) With primary text frames

- [x] Primary text frame on the master page
- [x] text flow enabled in the document settings
	
Text flow is created by InDesign after the script is complete. (Endnotes cannot be inserted on the last page this way).

## Technical background

Technically, the import works via an XSL transformation (1.0). The Word document is unpacked, transformed and imported as XML into InDesign. As a benefit you get an XML structure in your InDesign document and fully tagged content.

The stylesheet is designed to not lose content. After the import, however, please always check the contents to make sure that everything is there. As someone once put it so aptly: While Microsoft Word manuscripts *»(sometimes) look nice to a human reader, a peek under the hood reveals a messy slurry of largely unstructured text, tags, and cruft«.*[^1] And it is also not too uncommon that the structure of the Word document is damaged.[^2]  

[^1]: [The XSweet Story](https://xsweet.org/docs/3-xsweet-story/)
[^2]: [Quotation mark problem in Word index](https://indesign.uservoice.com/forums/601180-adobe-indesign-bugs/suggestions/36062545-index)

## Hooks

For special cases you can hook into the import with JavaScript, e.g. to create your own bibliography with endnotes or similar. The corresponding ExtendScript files can be found in the order `hooks`.

|Hook|File Name|Description|
|---|---|---|
|Before Import|beforeImport.jsx|Hooks in before the import takes place.|
|Before Mount|beforeMount.jsx|Hooks in before the InDesign objects (hyperlinks, comments, index markers, ...) are created.|
|Before Placed|beforePlaced|Hooks in before the content (XML story) is placed in the InDesign document.|
|After Placed|afterPlaced.jsx|Hooks in after the content has been placed.|


# Global Settings

|Option|Property|Type|Default|Description|
|---|---|---|---|---| 
|Logging|isLogged|Boolean|false|Logging of info messages, e.g. which objects are created in InDesign in the course of the import. Warning messages will always be output.|
|Dialog|isDialogShown|Boolean|true|Whether dialog is displayed or not. With `false` the dialog can be displayed by pressing and holding the **Shift** key when opening the document (click on the **Open** button).|

# Document Settings

|Option|Property|Type|Default|Description|
|---|---|---|---|---| 
|Autoflow|isAutoflow|Boolean|true|Controls automatic flow when no primary text frame is used.|
|XML Structure|isUntagged|Boolean|false|If true, then the XML structure will be removed out of the document after import.|
|Default Paragraph Style|defaultParagraphStyle|String|"Normal"|Name of the default paragraph style. This style is used for paragraphs that do not have a specific paragraph style applied in the Word document.|

## Metadata

The metadata entries in the Word document can be transferred to the InDesign document. Multiple values are separated by semicolons in the »Author(s)« and »Keywords« input fields in the Word document. Also the custom metadata from Word can be imported.

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Merge|areMerged|Boolean|Append the metadata of the imported Word document to the metadata of the InDesign document.|
|Replace|areReplaced|Boolean|Replace the metadata of the InDesign document with the metadata from the imported Word document.|
|Ignore|areIgnored|Boolean|Ignore metadata from the imported Word document.|

# Import Settings

|Option|Property|Type|Value|Default|Description|
|---|---|---|---|---|---| 
|Special Character Styles|styleMode|String|"extended" or "minimized"|"extended"|If minimized, all local overrides are ignored except the following: b (Bold), i (Italic), em (Emphasis), u (Underline), superscript, subscript, smallCaps, caps, highlight.|

# Tables

|Option|Property|Type|Value|Default|Description|
|---|---|---|---|---|---| 
|Table Mode|tableMode|String|"table" or "tabbedlist"|"table"|If 'tabbedlist', import Word tables as tab separated text to InDesign.|

# Images

A folder »Links« is created next to the InDesign file if document path is avaliable (for saved document). Otherwise the image will be embedded in the document.

|Option|Property|Type|Default|Description|
|---|---|---|---|---| 
|Remove|isRemoved|Boolean|false|Remove image.|
|Mark|isMarked|Boolean|false|Insert textbox content as plain text and highlighted with a condition.|
|Create|isPlaced|Boolean|true|Create anchored textframe in story.|
|Width|width|String|100|Default maximum image width in mm. Fitting into an (imaginary) rectangle with the values from height and width, i.e. the image frame will not be wider than this value.|
|Height|height|String|100|Default maximum image height in mm. Fitting into an (imaginary) rectangle with the values from height and width, i.e. the image frame will not be higher than this value.|
|Object Style|objectStyle|Object||Properties for the applied object style.|

# Textboxes

Text boxes from Word are inserted into the story as anchored text frames.

|Option|Property|Type|Default|Description|
|---|---|---|---|---| 
|Remove|isRemoved|Boolean|false|Remove textbox.|
|Mark|isMarked|Boolean|false|Insert textbox constent as plain text and highlighted with a condition.|
|Insert|isCreated|Boolean|true|Create anchored text frame in story.|
|Width|width|String|100|Default maximum text frame width in mm. Fitting into an (imaginary) rectangle with the values from height and width, i.e. the text frame will not be wider than this value.|
|Height|height|String|40|Default maximum text frame width in mm. Fitting into an (imaginary) rectangle with the values from height and width, i.e. the text frame will not be higher than this value.|
|Object Style|objectStyle|Object||Properties for the applied object style.|

# Page Breaks

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Insert|isInserted|Boolean|true|Insert as special character.|

# Column Breaks

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Insert|isInserted|Boolean|true|Insert as special character.|

# Forced Line Breaks

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Insert|isInserted|Boolean|true|Insert as special character.|

# Section Breaks

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Insert|isInserted|Boolean|true|Insert as special character.|

# Comments

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Remove|isRemoved|Boolean|Remove comment.|
|Mark|isMarked|Boolean|Insert comment as plain text and mark with condition.|
|Create|isCreated|Boolean|Insert as InDesign comment.|
|Metadata|isAdded|Boolean|Add metadata to the comment. (author, date)|

# Index

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Remove|isRemoved|Boolean|false|Remove index entries.|
|Create|isCreated|Boolean|true|Insert as InDesign index markers.|
|Topic Separator|topicSeparator|String|:|Topcis separator in the Word cross-reference field.|

## Cross-reference
### Präfix
#### Predefined prefixes

|Deutsch|English|Français|
|---|---|---|
|Siehe \[auch\]|See \[also\]|Voir \[aussi\]|
|Siehe auch hier|See also herein|Voir aussi ici|
|Siehe auch|See also|Voir aussi|
|Siehe hier|See herein|Voir ici|
|Siehe|See|Voir|

#### Custom prefixes

Furthermore, **custom prefixes** can be defined in the script settings, e.g.: `{"de":"→", "en":"→", "fr":"→"}` The entry in the Word cross-reference field then looks like this: `→ Topic0: Topic1`

If the prefix is not found in the predefined or user-defined prefixes, a non-joiner whitespace `\x{200B}` is used for the custom text string. In the input field for the index entry, the character combination `^k` appears in the custom string field in InDesign after the import.  (InDesign sets a unvisible `\uFEFF` character in case of the native Word import, but this breaks the XML structure when assigned by JavaScript).

### Topic
Nested topics can be input in Word in the Cross Reference field (Select index entry ▸ Options ▸ Cross Reference) with colon as separator, e.g. "See Animals: Cats".

# Hyperlinks

Hyperlinks are automatically named by InDesign and not renamed by the script. The tooltip text from Word is added as a label for later script editing. Unfortunately alternate text is not accessible via Scripting DOM.

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Mark|isMarked|Boolean|false|Insert hyperlink as plain text and mark with condition.|
|Create|isCreated|Boolean|true|Insert as InDesign hyperlink.|
|Ignore|isIgnored|Boolean|false|Ignore hyperlinks. Text content is imported as it is.|
|Character Style|characterStyleName|String|Hyperlink|Character style applied to the hyperlink text.|
|Add Character Style|isCharacterStyleAdded|Boolean|false| Add a character style to hyperlink text.|

# Cross-references

Be careful with cross-references: Some cross-reference types are not transferable 1:1 to InDesign, e.g. a reference to top/bottom, footnote/endnote number, or to bookmark content. 

Please check after the import if these correspond to your needs. Otherwise deactivate them during import. The information remains in the XML structure (except in footnote text, where no XML is allowed.) With this information, the cross-references can be adapted to your own needs.

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Mark|isMarked|Boolean|false|Insert cross-reference as plain text and mark with condition.|
|Create|isCreated|Boolean|true|Insert as InDesign cross-reference.|
|Ignore|isIgnored|Boolean|false|Ignore cross-references. Text content is imported as it is.|
|Character Style|characterStyleName|String|Cross_Reference|Character style applied to the cross-reference text.|
|Add Character Style|isCharacterStyleAdded|Boolean|false|Add a character style to cross-reference text.|

# Bookmarks

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Create|isCreated|Boolean|false|Insert as InDesign bookmark.|
|Marker|marker|String||Marker as a prefix of the bookmark name to identify the bookmarks to be included. (The underscore `_` is allowed as a special character in the bookmark name.) Example: Marker `indesign_`, Bookmark name in Word `indesign_my_bookmark_name]`. So Word bookmarks with prefix `indesign_` will be transferred to InDesign bookmarks. If the entry is an empty string and create bookmark is selected, all bookmarks in Word are created as InDesign bookmarks.|
|Remove Marker|isMarkerRemoved|String||The inserted marker (prefix) is removed from the bookmark content.|

# Track Changes

Word change tracking is currently implemented via conditional text in InDesign.

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Remove|isRemoved|Boolean|false|»Deleted« and »moved from« text is removed. »Inserted« and »moved to« Text is inserted as text.|
|Mark|isMarked|Boolean|true|Insert as text and mark with condition. »Deleted« and »moved from« text is hidden by default.|
|Create|isCreated|Boolean|false|Not yet implemented.|

# Footnotes

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Remove|isRemoved|Boolean|false|Remove footnote.|
|Mark|isMarked|Boolean|false|Insert as text and mark with condition.|
|Create|isCreated|Boolean|true|Insert as InDesign footnote|

# Endnotes

|Option|Property|Type|Default|Description|
|---|---|---|---|---|
|Remove|isRemoved|Boolean|false|Remove endnote.|
|Mark|isMarked|Boolean|false|Insert as text and mark with condition.|
|Create|isCreated|Boolean|true|Insert as InDesign endnote|

# Style Mapping

1. Click on Button `Load preset ...`
2. Select preset file (.smp, .xml or .txt)

You can use your .smp preset files in the InDesign preferences folder or create an XML file in the following form. (An example is also in the script preset folder.)

```
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Sangam-Import-Preset reader-type="Word/RTF">
	<Style-Mappings>
		<Paragraph-Style-Mappings>
			<Mapping style-name="1st-Paragraph-Style-Name-in-Word" mapped-to="1st-Paragraph-Style-Name-in-InDesign" />
			<Mapping style-name="2nd-Paragraph-Style-Name-in-Word" mapped-to="2nd-Paragraph-Style-Name-in-InDesign" />
		</Paragraph-Style-Mappings>
		<Character-Style-Mappings>
			<Mapping style-name="1st-Character-Style-Name-in-Word" mapped-to="1st-Character-Style-Name-in-InDesign" />
			<Mapping style-name="2nd-Character-Style-Name-in-Word" mapped-to="2nd-Character-Style-Name-in-InDesign" />
		</Character-Style-Mappings>
	</Style-Mappings>
</Sangam-Import-Preset>
```

## InDesign Preferences Folder

e.g. 
German version, macOS: /Users/[Your User Name]/Library/Preferences/Adobe InDesign/Version 19.0/de_DE/Word-Importvorgaben
German version, Windows 10: %USERPROFILE%\AppData\Roaming\Adobe\InDesign\Version 19.0\de_DE\Word-Importvorgaben

## Script Helper

The following script can be used to create a mapping file for the active InDesign document. You then only need to enter the names of the styles from Microsoft Word. Thanks to Jean-Claude Tremblay for the great idea and the script!

<details>
	<summary>ExtendScript Code (save as jsx file)</summary>
	
	```js
	// Microsoft Word InDesign Style Mapping
	// 
	// Description: 
	// The script creates an XML file for mapping Microsoft Word and InDesign styles. 
	// The file can be found on the desktop after the script run. The names of the paragraph 
	// and character styles from Word can then be entered there. Mappings that are not 
	// required can be removed.
	// 
	// Author: Jean-Claude Tremblay
	// Date: October 25, 2024

	var doc = app.activeDocument;
	var paragraphStyles = doc.allParagraphStyles;
	var characterStyles = doc.allCharacterStyles;

	// Create the content string
	var content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
	content += '<Sangam-Import-Preset reader-type="Word/RTF">\n';
	content += '\t<Style-Mappings>\n';
	content += '\t\t<Paragraph-Style-Mappings>\n';

	for (var i = 1; i < paragraphStyles.length; i++) {
			var style = paragraphStyles[i];
			var styleName = style.name;
			var groupName = style.parent.constructor.name === 'ParagraphStyleGroup' ? style.parent.name + '\uE00B' : '';
			content += '\t\t\t<Mapping style-name="' + i + '-Paragraph-Style-Name-in-Word" mapped-to="' + groupName + styleName + '" />\n';
	}

	content += '\t\t</Paragraph-Style-Mappings>\n';
	content += '\t\t<Character-Style-Mappings>\n';

	for (var j = 1; j < characterStyles.length; j++) {
			var style = characterStyles[j];
			var styleName = style.name;
			var groupName = style.parent.constructor.name === 'CharacterStyleGroup' ? style.parent.name + '\uE00B' : '';
			content += '\t\t\t<Mapping style-name="' + j + '-Character-Style-Name-in-Word" mapped-to="' + groupName + styleName + '" />\n';
	}

	content += '\t\t</Character-Style-Mappings>\n';
	content += '\t</Style-Mappings>\n';
	content += '</Sangam-Import-Preset>';

	// Get document name without extension
	var docName = doc.name.replace(/\.[^\.]+$/, '');

	// Create and write to file on desktop
	var desktop = Folder.desktop;
	var outputFile = new File(desktop + "/" + docName + "_style_mappings.xml");
	outputFile.encoding = "UTF-8";
	outputFile.open("w");
	outputFile.write(content);
	outputFile.close();

	// Open the file
	// outputFile.execute();

	```
</details>

# Drawbacks with the native docx import

- Local style overrides
- Import images as embedded images
- Table styles are not imported
- Index: Number style override is not transferred, nothing other than "See" is identified as a custom cross-reference text, Index entries get lost.[^3]
- Hyperlinks are not imported (correctly).[^4] 

[^3]: [Bug report: Index number style override](https://indesign.uservoice.com/forums/601180-adobe-indesign-bugs/suggestions/38549830-index-entries-lost-when-importing-a-docx-file-wit)
[^4]: [Bug report: Hyperlinks import](https://indesign.uservoice.com/forums/601021-adobe-indesign-feature-requests/suggestions/32872021-hyperlinks-from-word)

# Known Issues

Hyperlinks across multiple paragraphs. Only the part in the first paragraph becomes an active hyperlink.

# ToDo 
- [ ] Remove special characters (text, index entries, ...)?
- [ ] Import functional references (Bibliography)? with cross-references to text anchors with name e.g. Newton, 1743
- [ ] Section break (Numbering & Section Options)?
- [ ] Symbols via Unicode
- [ ] Create lists for list paragraphs during import (If same paragraph format but different list, then new paragraph format based original with new list.)


# Support
If you want to support the development of the script: 

[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=roland%2edreger%40a1%2enet&lc=AT&item_name=Roland%20Dreger%20%2f%20Donation%20for%20script%20development%20Import-DOCX&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted)

# License

[MIT](http://www.opensource.org/licenses/mit-license.php)
