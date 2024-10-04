/* DESCRIPTION: Import Word Document (docx) */ 

/*
	
		+ Adobe InDesign Version: CC2021+
		+ Author: Roland Dreger 
		+ Date: January 24, 2022
		
		+ Last modified: October 4, 2024
		
		
		+ Descriptions
			
			Alternative import for Microsoft Word documents
			for clean and sematically structured content

*/


//@include "utils/i18n.jsx"
//@include "utils/classes.jsx"
//@include "utils/dialogs.jsx"
//@include "utils/helpers.jsx"

//@include "hooks/beforeImport.jsx"
//@include "hooks/beforeMount.jsx"
//@include "hooks/beforePlaced.jsx"
//@include "hooks/afterPlaced.jsx"


var _global = {
	"projectName":"Import_Docx",
	"version":"0.4.0",
	"mode":"release", /* Type: String. Value: "debug" or "release" */
	"isLogged":false,
	"isDialogShown":true,
	"log":[]
};

/* Document Settings */
_global["setups"] = {
	"user":$.getenv("USER"),
	"document":{
		"isAutoflowing": true, /* Type: Boolean. Description: If true, autoflows placed text. (Depends on document settings.)  */
		"isUntagged":false, /* Type: Boolean. Description: If true, then the XML structure will be removed out of the document after import. */
		"defaultParagraphStyle":"Normal" /* Type: String. Value: e.g. "Normal". Description: This style is used for paragraphs that do not have a specific paragraph style applied in the Word document. */
	},
	"preset": {
		"filePath": ""
	},
	"metadata":{
		"tag":"document",
		"attributes":{
			"description":"description",
			"category":"category",
			"keywords":"keywords",
			"author":"author",
			"subject":"subject",
			"title":"title"
		},
		"areMerged":false, 
		"areReplaced":false,
		"areIgnored":true
	},
	"linkFolder":{
		"name":"Links", /* Type: String. Value: e.g. "Links". Description: Folder name for placed images; */
		"path":"" /* Type: String. Value: e.g. "" (next to the InDesign Document) or "~/Desktop". Description: Folder path for placed images (optional). */
	},
	"xslt":{
		"name":"docx2Indesign.xsl"
	},
	"import":{
		"styleMode":"extended", /* Type: String. Value: 'extended' or 'minimized'. Description:  If minimized, all local overrides are ignored except the following: strong, i, em, u, superscript, subscript, small caps, caps, highlight. */
		"tableMode":"table" /* Type: String. Value: 'table' or 'tabbedlist'. Description: If 'tabbedlist', import table structure as tab separated text. */
	},
	"place":{},
	"mount":{},
	"paragraph":{
		"tag":"paragraph"
	},
	"pageBreak":{
		"tag":"pagebreak",
		"isInserted":true
	},
	"columnBreak":{
		"tag":"columnbreak",
		"isInserted":true
	},
	"forcedLineBreak":{
		"tag":"forcedlinebreak",
		"isInserted":true
	},
	"sectionBreak":{
		"tag":"sectionbreak",
		"isInserted":true
	},
	"comment":{
		"tag":"comment", 
		"color":[255,255,155],
		"metadata":{
			"isAdded":true
		}, 
		"isRemoved":false,
		"isMarked":false, 
		"isCreated":true
	},
	"indexmark":{
		"tag":"indexmark", 
		"attributes":{
			"type":"type",
			"format":"format",
			"entry":"entry",
			"target":"target"
		},
		"topicSeparator": ":",
		"isRemoved":false,
		"isCreated":true,
		"crossReference":{
			"prefixes":[
				{ "de":"Siehe [auch]", "en":"See [also]", "fr":"Voir [aussi]" }, /* English key "en" and value "See [also]" is required as a minimum. Do not modify the value. */
				{ "de":"Siehe auch hier", "en":"See also herein", "fr":"Voir aussi ici" }, /* English key "en" and value "See also herein" is required as a minimum. Do not modify the value. */
				{ "de":"Siehe auch", "en":"See also", "fr":"Voir aussi" }, /* English key "en" and value "See also" is required as a minimum. Do not modify the value. */
				{ "de":"Siehe hier", "en":"See herein", "fr":"Voir ici" }, /* English key "en" and value "See herein" is required as a minimum. Do not modify the value. */
				{ "de":"Siehe", "en":"See", "fr":"Voir" }, /* English key "en" and value "See" is required as a minimum. Do not modify the value. */
				/* more objects can be added -> results in cross-reference with type CrossReferenceType.CUSTOM_CROSS_REFERENCE */
				{"de":"→", "en":"→", "fr":"→" } /* Word cross-reference field: e.g. x Topic0: Topic1 */
			],
			"noMatchCustomTypeString": "\u200B" /* Default: zero-width whitespace; If an empty string, the prefix "See [also]" is used.  */
		}
	},
	"hyperlink":{
		"tag":"hyperlink", 
		"attributes":{
			"uri":"uri",
			"title":"title"
		}, 
		"color":[120,190,255],
		"characterStyleName":"Hyperlink",
		"isCharacterStyleAdded":false, 
		"isMarked":false, 
		"isCreated":true,
		"isIgnored":false
	},
	"crossReference":{
		"tag":"cross-reference", 
		"attributes":{
			"uri":"uri",
			"type":"type",
			"format":"format"
		}, 
		"color":[120,190,255],
		"characterStyleName":"Cross_Reference",
		"isAnchorHidden":true,
		"isCharacterStyleAdded":true, 
		"isMarked":false, 
		"isCreated":true,
		"isIgnored":false
	},
	"bookmark":{
		"tag":"bookmark",
		"attributes":{
			"id":"id",
			"index":"index",
			"content":"content"
		},
		"marker":"", /* Marker as a prefix of Word bookmark ID to be included as InDesign bookmark. Value: String. Example: indesign_my_bookmark_name -> Marker: indesign_ */
		"isAnchorHidden":true,
		"isCreated":false
	},
	"textbox":{ 
		"tag":"textbox", 
		"color":[155,155,255], 
		"width":"100", /* Default textbox width in mm; Value: String */
		"height":"40", /* Default textbox height in mm; Value: String */
		"objectStyleProperties":{
			"enableAnchoredObjectOptions":true,
			"anchoredObjectSettings": {
				"anchoredPosition":AnchorPosition.ANCHORED,
				"anchorPoint":AnchorPoint.TOP_LEFT_ANCHOR,
				"horizontalAlignment": HorizontalAlignment.LEFT_ALIGN,
				"horizontalReferencePoint":AnchoredRelativeTo.TEXT_FRAME,
				"spineRelative":false,
				"pinPosition": false,
				"verticalReferencePoint":VerticallyRelativeTo.LINE_BASELINE
			},
			"enableTextWrapAndOthers":true,
			"textWrapPreferences":{
				"textWrapMode":TextWrapModes.JUMP_OBJECT_TEXT_WRAP
			}
		},
		"isRemoved":false,
		"isMarked":false, 
		"isCreated":true
	},
	"image":{
		"tag":"image", 
		"attributes":{
			"source":"source",
			"description":"description"
		},
		"width":"100", /* Default image width in mm; Value: String */
		"height":"100", /* Default image height in mm; Value: String */
		"objectStyleProperties":{
			"strokeWeight":0,
			"enableAnchoredObjectOptions":true,
			"anchoredObjectSettings": {
				"anchoredPosition":AnchorPosition.ANCHORED,
				"anchorPoint":AnchorPoint.TOP_LEFT_ANCHOR,
				"horizontalAlignment": HorizontalAlignment.LEFT_ALIGN,
				"horizontalReferencePoint":AnchoredRelativeTo.TEXT_FRAME,
				"spineRelative":false,
				"pinPosition": false,
				"verticalReferencePoint":VerticallyRelativeTo.LINE_BASELINE
			},
			"enableFrameFittingOptions":true,
			"frameFittingOptions":{
				"fittingAlignment":AnchorPoint.TOP_LEFT_ANCHOR,
				"fittingOnEmptyFrame":EmptyFrameFittingOptions.PROPORTIONALLY
			},
			"enableTextWrapAndOthers":true,
			"textWrapPreferences":{
				"textWrapMode":TextWrapModes.JUMP_OBJECT_TEXT_WRAP
			}
		},
		"color":[155,255,255], 
		"isAltTextInserted":true,
		"isRemoved":false,
		"isMarked":false,
		"isPlaced":true
	},
	"trackChanges":{
		"insertedText":{
			"tag":"insertedtext", 
			"color":[0,255,0]
		},
		"deletedText":{
			"tag":"deletedtext", 
			"color":[255,0,0]
		},
		"movedFromText":{
			"tag":"deletedtext", 
			"color":[155,155,255]
		},
		"movedToText":{
			"tag":"movedtext", 
			"color":[0,255,255]
		},
		"isRemoved":false,
		"isMarked":true, 
		"isCreated":false
	},
	"footnote":{ 
		"tag":"footnote", 
		"color":[155,255,255], 
		"isRemoved":false,
		"isMarked":false, 
		"isCreated":true 
	},
	"endnote":{ 
		"tag":"endnote", 
		"color":[255,155,255], 
		"isRemoved":false,
		"isMarked":false, 
		"isCreated":true
	}
	
};

/* Check: Developer or User? */
if(_global["setups"]["user"] === "rolanddreger") {
	// _global["mode"] = "debug";
	_global["isLogged"] = true;
}


__start();


function __start() {
	
	if(!_global) { 
		throw new Error("Global object [_global] not defined.");
	}
	
	/* Deutsch-Englische Dialogtexte definieren */
	__defLocalizeStrings();
	
	/* Progressbar definieren */
	_global["progressbar"] = new ProgressBar();
	if(!_global["progressbar"]) {
		throw new Error(localize(_global.createProgessbarErrorMessage));
	}
	
	/* Active document */
	var _doc = app.documents.firstItem();
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid || _doc.layoutWindows.length === 0) {
		alert(localize(_global.noDocOpenAlert));
		return false; 
	}

	/* Script Preferences */
	var _userEnableRedraw = app.scriptPreferences.enableRedraw;
	app.scriptPreferences.enableRedraw = false;
	var _userInteractionLevel = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	
	/* Document Preferences */
	var _userShowStructure = _doc.xmlViewPreferences.showStructure;
	// _doc.xmlViewPreferences.showStructure = false;

	try {
		if(_global["mode"] !== "debug") {
			app.doScript(
				__runMainSequence, 
				ScriptLanguage.JAVASCRIPT, 
				[_doc], 
				UndoModes.ENTIRE_SCRIPT, 
				localize(_global.goBackLabel)
			);
		} else {
			__runMainSequence([_doc]);
		}
	} catch(_error) {
		if(_error instanceof Error) {
			alert(
				_error.name + " | " + _error.number + "\n" +
				localize(_global.errorMessageLabel) + " " + _error.message + ";\n" +
				localize(_global.lineLabel) + " " + _error.line + "\n" + 
				localize(_global.fileNameLabel) + " " + _error.fileName,
				"Error", true
			);
		} else {
			alert(localize(_global.processingErrorAlert) + "\n" + _error, "Error", true);
		}
	} finally {
		if(_doc && _doc.isValid) {
			_doc.xmlViewPreferences.showStructure = _userShowStructure;
		}
		app.scriptPreferences.enableRedraw = _userEnableRedraw;
		app.scriptPreferences.userInteractionLevel = _userInteractionLevel;
	}

	/* Close progressbar */
	if(_global.hasOwnProperty("progressbar")) {
		_global["progressbar"].close();
	}

	/* Check: Log messages? */
	if(_global["log"].length > 0) {
		__showLog(_global["log"]);
		return false;
	}
	
	return true;
} /* END function __start */


_global = null;




/**
 * Main Sequence
 * @param {Array} _doScriptParameterArray 
 * @returns Boolean
 */
function __runMainSequence(_doScriptParameterArray) {
	
	if(!_global.hasOwnProperty("setups")) { 
		throw new Error("Global object has no property [_setups]."); 
	}
	if(!_doScriptParameterArray || !(_doScriptParameterArray instanceof Array) || _doScriptParameterArray.length === 0) { 
		throw new Error("Array with length > 1 as parameter required."); 
	}
	
	var _setupObj = _global["setups"];
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required."); 
	}
	
	var _doc = _doScriptParameterArray[0];
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required."); 
	}

	var _hooks = new Hooks();

	/* Get docx file */
	var _docxFile = __getDocxFile();
	if(!_docxFile) {
		return false;
	}
	
	/* Show import dialog */
	if(_global["isDialogShown"] || (ScriptUI && ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.shiftKey)) {
		_setupObj = __showImportDialog(_setupObj);
		if(!_setupObj) {
			return false;
		}
	}

	/* Get package data */
	var _unpackObj = __getPackageData(_docxFile);
	if(!_unpackObj) {
		return false;
	}
	
	/* Hook: beforeImport */
	var _beforeImportResultObj = _hooks.beforeImport(_doc, _unpackObj, _setupObj);
	if(!_beforeImportResultObj) {
		return false;
	}

	/* Import XML from unpacked docx file */
	var _wordXMLElement = __importXML(_doc, _unpackObj, _setupObj);
	if(!_wordXMLElement) {
		return false;
	}

	/* Hook: beforeMount */
	var _beforeMountResultObj = _hooks.beforeMount(_doc, _unpackObj, _wordXMLElement, _setupObj);
	if(!_beforeMountResultObj) {
		return false;
	}

	/* Mount InDesign items before XML placed */
	var _mountBeforePlaceResultObj = __mountBeforePlaced(_doc, _unpackObj, _wordXMLElement, _setupObj);
	if(!_mountBeforePlaceResultObj) {
		return false;
	}
	
	/* Hook: beforePlaced */
	var _beforePlaceResultObj = _hooks.beforePlaced(_doc, _unpackObj, _wordXMLElement, _setupObj);
	if(!_beforePlaceResultObj) {
		return false;
	}

	/* Place imported XML structure */
	var _wordStory = __placeXML(_doc, _wordXMLElement, _setupObj);
	if(!_wordStory) {
		return false;
	}

	/* Hook: afterPlaced */
	var _afterPlaceResultObj = _hooks.afterPlaced(_doc, _unpackObj, _wordXMLElement, _wordStory, _setupObj);
	if(!_afterPlaceResultObj) {
		return false;
	}

	/* Mount InDesign items after XML placed  */
	var _mountAfterPlaceResultObj = __mountAfterPlaced(_doc, _unpackObj, _wordXMLElement, _wordStory, _setupObj);
	if(!_mountAfterPlaceResultObj) {
		return false;
	}

	/* Remove XML Structure */
	if(_setupObj["document"]["isUntagged"]) {
		if(_wordXMLElement && _wordXMLElement.hasOwnProperty("untag") && _wordXMLElement.isValid) {
			_wordXMLElement.untag();
		}
	}

	return true;
} /* END function __runMainSequence */




/**
 * Get docx file
 * @returns File
 */
function __getDocxFile() {
	
	const _wordExtRegExp = new RegExp("(\\.docx$|\\.xml$)","i");

	var _wordFile = File.openDialog(localize(_global.selectWordFile));
	if(!_wordFile || !_wordFile.exists) { 
		return null; 
	}

	var _wordFileName = _wordFile.name;
	if(!_wordExtRegExp.test(_wordFileName)) {
		_global["log"].push(localize(_global.fileExtensionValidationMessage));
		return null;
	}
	
	return _wordFile; 
} /* END function __getDocxFile */


/**
 * Get package data
 * (Unpack file to temp folder if docx)
 * @param {File} _packageFile
 * @returns Object
 */
function __getPackageData(_packageFile) {
	
	if(!_packageFile || !(_packageFile instanceof File) || !_packageFile.exists) { 
		throw new Error("Existing file as parameter required."); 
	}

	const TEMP_FOLDER_NAME = "InDesign_Word_Import";

	const _xmlExtRegExp = new RegExp("\\.xml$","i");
	const _docxExtRegExp = new RegExp("\\.docx$","i");

	var _packageFileName = _packageFile.name;
	var _packageFilePath = _packageFile.fullName;
	
	/* Check: Word-XML-Document (.xml)? */
	if(_xmlExtRegExp.test(_packageFileName)) {
		return { 
			"folder":null,
			"word": {
				"document":_packageFile
			}
		};
	}

	/* Check: Word Document (.docx)? */
	if(!_docxExtRegExp.test(_packageFileName)) {
		return null;
	}

	var _tempFolderPath = "";
	var _tempFolder;
	var _packageFolderPath = "";
	var _packageFolder;

	var _timestamp = __getTimestamp();

	/* Create temporary package folder */
	try {
		_tempFolderPath = Folder.temp.fullName + "/" + TEMP_FOLDER_NAME;
		_tempFolder = Folder(_tempFolderPath);
		if(!_tempFolder.exists) {
			_tempFolder.create();
		}
		if(!_tempFolder.exists) {
			_global["log"].push(localize(_global.createFolderErrorMessage, _tempFolderPath));
			return null;
		}
		_packageFolderPath = _tempFolder.fullName + "/package_" + _timestamp;
		_packageFolder = Folder(_packageFolderPath);
	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	}

	if(!_packageFolder || !(_packageFolder instanceof Folder)) {
		_global["log"].push(localize(_global.createFolderErrorMessage, _packageFolderPath));
		return null;
	}

	/* Unpack Word Document */
	try {
		app.unpackageUCF(_packageFile, _packageFolder);
	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	}
	
	if(!_packageFolder.exists) {
		_global["log"].push(localize(_global.unpackageFolderErrorMessage, _packageFolderPath));
		return null;
	}

	var _xmlDocFile = File(_packageFolder.fullName + "/word/document.xml");
	if(!_xmlDocFile.exists) {
		_global["log"].push(localize(_global.unpackageDocumentFileErrorMessage, _packageFilePath));
		return null;
	}

	return { 
		"folder":_packageFolder,
		"word": {
			"document":_xmlDocFile
		}
	}; 
} /* END function __getPackageData */


/**
 * Import Word document xml file
 * @param {Document} _doc InDesign document
 * @param {Objekt} _unpackObj Result of unpacking Word document file
 * @param {Objekt} _setupObj 
 * @returns XMLElement
 */
function __importXML(_doc, _unpackObj, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");  
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const _defaultPStyleName = _setupObj["document"]["defaultParagraphStyle"];
	const _xsltFileName = _setupObj["xslt"]["name"];
	const _styleMode = _setupObj["import"]["styleMode"];
	const _tableMode = _setupObj["import"]["tableMode"];

	_global["progressbar"].init(0, 2, "", localize(_global.importProgressLabel));
	_global["progressbar"].step();

	var _transformParams = [];

	_transformParams.push(["app", "indesign"]);
	_transformParams.push(["style-mode", _styleMode]);
	_transformParams.push(["table-mode", _tableMode]);
	_transformParams.push(["fallback-paragraph-style-name", _defaultPStyleName]);

	var _xsltFile = __getXSLTFile(_xsltFileName);
	if(!_xsltFile) { 
		return null; 
	}

	var _unpackFolderPath = "";
	var _unpackFolder = _unpackObj["folder"];
	if(_unpackFolder && _unpackFolder instanceof Folder && _unpackFolder.exists) {
		_unpackFolderPath = _unpackFolder.fullName;
		_transformParams.push(["package-base-uri", ".."]);
	}

	var _wordXMLFile = _unpackObj["word"]["document"];
	if(!_wordXMLFile || !_wordXMLFile.exists) {
		_global["log"].push(localize(_global.wordDocumentFileErrorMessage, _wordXMLFile));
		return null;
	}
	
	_transformParams.push(["document-file-name",_wordXMLFile.name]);

	var _rootXMLElement = _doc.xmlElements.firstItem();
	var _lastXMLElement = _rootXMLElement.xmlElements.lastItem();
	if(_lastXMLElement.isValid) {
		_lastXMLElement = _lastXMLElement.getElements()[0];
	} else {
		_lastXMLElement = null;
	}

	/* Preset */
	var _importPresetFilePath = _setupObj["preset"]["filePath"];
	if(!!_importPresetFilePath) {
		var _importPresetFile = File(_importPresetFilePath);
		if(_importPresetFile.exists) {
			_transformParams.push(["indesign-import-preset-file-path", _importPresetFile.fsName]);
		}
	}

	if(File(_unpackFolderPath + "/" + "word/comments.xml").exists) {
		_transformParams.push(["comments-file-path", "comments.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "docProps/app.xml").exists) {
		_transformParams.push(["app-props-file-path", "../docProps/app.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "docProps/core.xml").exists) {
		_transformParams.push(["core-props-file-path", "../docProps/core.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "docProps/custom.xml").exists) {
		_transformParams.push(["custom-props-file-path", "../docProps/custom.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "word/_rels/document.xml.rels").exists) {
		_transformParams.push(["document-relationships-file-path", "_rels/document.xml.rels"]);
	}
	if(File(_unpackFolderPath + "/" + "word/endnotes.xml").exists) {
		_transformParams.push(["endnotes-file-path", "endnotes.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "word/_rels/endnotes.xml.rels").exists) {
		_transformParams.push(["endnotes-relationships-file-path", "_rels/endnotes.xml.rels"]);
	}
	if(File(_unpackFolderPath + "/" + "word/footnotes.xml").exists) {
		_transformParams.push(["footnotes-file-path", "footnotes.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "word/_rels/footnotes.xml.rels").exists) {
		_transformParams.push(["footnotes-relationships-file-path", "_rels/footnotes.xml.rels"]);
	}
	if(File(_unpackFolderPath + "/" + "word/numbering.xml").exists) {
		_transformParams.push(["numbering-file-path", "numbering.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "word/styles.xml").exists) {
		_transformParams.push(["styles-file-path", "styles.xml"]);
	}

	var _userXMLImportPreferences = _doc.xmlImportPreferences.properties;

	try {
		/* XML Import Preferences */
		_doc.xmlImportPreferences.properties = {
			importStyle:XMLImportStyles.APPEND_IMPORT,
			allowTransform:true,
			transformFilename:_xsltFile,
			transformParameters:_transformParams,
			repeatTextElements:false,
			ignoreWhitespace:false,
			createLinkToXML:false,
			ignoreUnmatchedIncoming:false,
			importCALSTables:false,
			importTextIntoTables:false,
			importToSelected:false,
			removeUnmatchedExisting:false
		};

		/* Import XML File */
		_doc.importXML(_wordXMLFile);

	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	} finally {
		_doc.xmlImportPreferences.properties = _userXMLImportPreferences;
	}

	/* Check: XML import successful? */
	var _wordXMLElement = _rootXMLElement.xmlElements.lastItem();
	if(!_wordXMLElement.isValid) {
		_global["log"].push(localize(_global.xmlDataImportErrorMessage));
		return null; 
	}
	_wordXMLElement = _wordXMLElement.getElements()[0];
	if(_wordXMLElement === _lastXMLElement) {
		_global["log"].push(localize(_global.xmlDataImportErrorMessage));
		return null; 
	}

	_global["progressbar"].step();

	return _wordXMLElement;
} /* END function __importXML */


/**
 * Get XSL transformation file
 * @param {String} _xsltFileName 
 * @returns File
 */
function __getXSLTFile(_xsltFileName) {
		 
	if(!_xsltFileName || _xsltFileName.constructor !== String) { 
		throw new Error("Object as string required.");
	}

	const _xslFileExtRegExp = new RegExp("\\.xsl$", "i");

	var _xsltFolder = __getScriptFolder();
	if(!_xsltFolder || !_xsltFolder.exists) { 
		_global["log"].push(localize(_global.scriptFolderErrorMessage));
		return false; 
	}

	var _xsltFile = _xsltFolder.getFiles(_xsltFileName)[0];
	if(!_xsltFile || !_xsltFile.exists) {
		_xsltFile = File.openDialog(localize(_global.selectXSLFile, _xsltFileName), null, false);
		if(!_xsltFile) {
			return null;
		}
		
	}

	if(!_xsltFile.exists || !_xslFileExtRegExp.test(_xsltFile.name)) {
		_global["log"].push(localize(_global.noXSLFileErrorMessage));
		return null;
	}

	return _xsltFile;
} /* END function __getXSLTFile */


/**
 * Get path for current script
 * @returns String
 */
function __getScriptFolder() {
	
	var _skriptFolder;
	
	try {
		_skriptFolder  = app.activeScript.parent;
	} catch(_error) { 
		_skriptFolder = File(_error.fileName).parent;
	}
	
	if(!_skriptFolder || !_skriptFolder.exists) { 
		return null; 
	}

	return _skriptFolder;
} /* END function __getScriptFolder */




/**
 * Mount InDesign items before placing XML
 * @param {Document} _doc 
 * @param {Object} _unpackObj 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Object
 */
function __mountBeforePlaced(_doc, _unpackObj, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	/* 01 Metadata */
	__handleMetadata(_doc, _wordXMLElement, _setupObj);

	/* 02 Breaks */
	__insertBreaks(_doc, _wordXMLElement, _setupObj);

	/* 03 Comments */
	__handleComments(_doc, _wordXMLElement, _setupObj);

	/* 04 Index */
	__handleIndexmarks(_doc, _wordXMLElement, _setupObj);
	
	/* 05Hyperlinks */
	__handleHyperlinks(_doc, _wordXMLElement, _setupObj);

	/* 06 Cross-references */
	__handleCrossReferences(_doc, _wordXMLElement, _setupObj);

	/* 07Bookmarks */
	__handleBookmarks(_doc, _wordXMLElement, _setupObj);
	
	/* 08 Textboxes */
	__handleTextboxes(_doc, _wordXMLElement, _setupObj);

	/* 09 Images */
	__handleImages(_doc, _wordXMLElement, _unpackObj, _setupObj);

	/* 10 Track Changes */
	__handleTrackChanges(_doc, _wordXMLElement, _setupObj);

	/* 
		Last in chain: Footnotes and Endnotes 
		(After all XML manipulations, since XML elements must be removed from footnotes and endnotes)
	*/

	/* 11 Footnotes */ 
	__handleFootnotes(_doc, _wordXMLElement, _setupObj);
	
	/* 12 Endnotes */
	__handleEndnotes(_doc, _wordXMLElement, _setupObj);	

	
	return {};
} /* END function __mountBeforePlaced */


/**
 * Handle Document Metadata
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean 
 */
function __handleMetadata(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	
	const METADATA_TAG_NAME = _setupObj["metadata"]["tag"];
	const ARE_METADATA_MERGED = _setupObj["metadata"]["areMerged"];
	const ARE_METADATA_REPLACED = _setupObj["metadata"]["areReplaced"];
	const ARE_METADATA_IGNORED = _setupObj["metadata"]["areIgnored"];

	var _metadataXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + METADATA_TAG_NAME);
	if(_metadataXMLElementArray.length === 0) {
		return true;
	}

	if(ARE_METADATA_IGNORED) {
		return true;
	}

	if(ARE_METADATA_MERGED) {
		__insertMetadata(_doc, _metadataXMLElementArray[0], "merge", _setupObj);
		return true;
	}

	if(ARE_METADATA_REPLACED) {
		__insertMetadata(_doc, _metadataXMLElementArray[0], "replace", _setupObj);
		return true;
	}


	return true;
} /* END function __handleMetadata */


/**
 * Create Comments
 * @param {Document} _doc 
 * @param {XMLElement} _metadataXMLElement 
 * @param {String} _mode Values: merge | replace 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __insertMetadata(_doc, _metadataXMLElement, _mode, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_metadataXMLElement || !(_metadataXMLElement instanceof XMLElement) || !_metadataXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_mode || _mode.constructor !== String || (_mode !== "merge" && _mode !== "replace")) { 
		throw new Error("String as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	
	const TITLE_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["title"];
	const AUTHOR_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["author"];
	const DESCRIPTION_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["description"];
	const CATEGORY_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["category"];
	const KEYWORDS_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["keywords"];
	const SUBJECT_ATTRIBUTE_NAME = _setupObj["metadata"]["attributes"]["subject"];

	const DC_NS = "http://purl.org/dc/elements/1.1/";
	const METADATA_VALUE_SEPARATOR = "; ";
	
	var _metadataCounter = 0;
	var _isSet = false;

	var _metadataPrefs = _doc.metadataPreferences;

	/* Document title */
	var _titleAttribute = _metadataXMLElement.xmlAttributes.itemByName(TITLE_ATTRIBUTE_NAME);
	if(_titleAttribute.isValid && _titleAttribute.value) {
		var _docTitle = "";
		if(_mode === "merge") {
			_docTitle = (_metadataPrefs.documentTitle && _metadataPrefs.documentTitle + METADATA_VALUE_SEPARATOR) || "";
		}
		_docTitle += _titleAttribute.value;
		var _titleObjArray = __getItemObjArray(_docTitle, [{ "name":"xml:lang", "value":"x-default" }]);
		_isSet = __setXMPContainerProperties(_doc, "ALT", DC_NS, "dc:title", _titleObjArray, "replace");
		if(_isSet) { 
			_metadataCounter += 1; 
		}
	}
	
	/* Authors */
	var _authorAttribute = _metadataXMLElement.xmlAttributes.itemByName(AUTHOR_ATTRIBUTE_NAME);
	if(_authorAttribute.isValid && _authorAttribute.value) {
		var _creatorObjArray = __getItemObjArray(_authorAttribute.value, [], "split");
		_isSet = __setXMPContainerProperties(_doc, "SEQ", DC_NS, "dc:creator", _creatorObjArray, _mode);
		if(_isSet) { 
			_metadataCounter += 1; 
		}
	}

	/* Description */
	var _descriptionAttribute = _metadataXMLElement.xmlAttributes.itemByName(DESCRIPTION_ATTRIBUTE_NAME);
	if(_descriptionAttribute.isValid && _descriptionAttribute.value) {
		var _description = "";
		if(_mode === "merge") {
			_description = (_metadataPrefs.description && _metadataPrefs.description + METADATA_VALUE_SEPARATOR) || "";
		}
		_description += _descriptionAttribute.value;
		var _descObjArray = __getItemObjArray(_description, [{ "name":"xml:lang", "value":"x-default" }]);
		_isSet = __setXMPContainerProperties(_doc, "ALT", DC_NS, "dc:description", _descObjArray, "replace");
		if(_isSet) { 
			_metadataCounter += 1; 
		}
	}

	/* Keywords */
	var _keywordsAttribute = _metadataXMLElement.xmlAttributes.itemByName(KEYWORDS_ATTRIBUTE_NAME);
	if(_keywordsAttribute.isValid && _keywordsAttribute.value) {
		var _subjectObjArray = __getItemObjArray(_keywordsAttribute.value, [], "split");
		_isSet = __setXMPContainerProperties(_doc, "BAG", DC_NS, "dc:subject", _subjectObjArray, _mode);
		if(_isSet) {
			_metadataCounter += 1;
		}
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.insertMetadataMessage, _metadataCounter));
	}

	return true;
} /* END function __insertMetadata */


/**
 * Get item object array for setting XMP property
 * @param {String} _inputValue 
 * @param {Array} _attrArray 
 * @param {String} _mode optional
 * @returns Array
 */
function __getItemObjArray(_inputValue, _attrArray, _mode) {

	if(!_inputValue || _inputValue.constructor !== String) {
		return [];
	}
	if(!_attrArray || !(_attrArray instanceof Array)) {
		return [];
	}

	const _valueSeparatorRegExp = new RegExp("[;]","g");
	const _trimRegExp = new RegExp("^\\s+|\\s+$","g");

	_inputValue = _inputValue.replace(_trimRegExp, "");
	if(!_inputValue) {
		return [];
	}

	var _valueArray = [];
	
	if(_mode === "split") {
		_valueArray = _inputValue.replace(_trimRegExp, "").split(_valueSeparatorRegExp);
	} else {
		_valueArray.push(_inputValue);
	}
	
	var _outputArray = [];

	for(var i=0; i<_valueArray.length; i+=1) {

		var _splitedValue = _valueArray[i].replace(_trimRegExp, "");
		if(!_splitedValue) {
			continue;
		}

		_outputArray.push({
			"value": _splitedValue,
			"attributes":_attrArray
		});
	}

	return _outputArray;
} /* END function __getItemObjArray */


/**
 * Merge/Replace metadata value in document
 * If _itemObjArray is an empty array, the metadata property is not changed. 
 * (This applies to both modes.)
 * Example _itemObjArray
 * [
 *		{
 *			"value": _description,
 *			"attributes": [
 *				{ "name":"xml:lang", "value":"x-default" },
 *				...
 *			]
 *		},
 *  	...
 *	]
 * @param {Document} _doc
 * @param {String} _containerType Values: "ALT", "BAG" or "SEQ"
 * @param {String} _namespace e.g. "http://purl.org/dc/elements/1.1/"
 * @param {String} _name e.g. dc:title
 * @param {Array} _itemObjArray 
 * @param {String} _mode Values: "merge" or "replace"
 * @returns Boolean
 */
function __setXMPContainerProperties (_doc, _containerType, _namespace, _name, _itemObjArray, _mode) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) {
		return false;
	}
	if(!_containerType  || _containerType.constructor !== String || !ContainerType.hasOwnProperty(_containerType)) {
		return false;
	}
	if(_namespace === null || _namespace === undefined || _namespace.constructor !== String) {
		return false;
	}
	if(!_name || _name.constructor !== String) {
		return false;
	}
	if(!_itemObjArray || !(_itemObjArray instanceof Array) || _itemObjArray.length === 0) {
		return false;
	}
	if(_mode !== "merge" && _mode !== "replace") {
		return false;
	}

	const _trimRegExp = new RegExp("^\\s+|\\s+$","g");

	try {

		var _metadataPrefs = _doc.metadataPreferences;
		var _numOfContainerItems = _metadataPrefs.countContainer(_namespace, _name);

		/* Check: Does the container exists? */
		if(_numOfContainerItems === 0) {
			/* Create new container */
			_metadataPrefs.createContainerItem(_namespace, _name, 0, ContainerType[_containerType]);
		}
		/* Check: Replace mode? */
		else if(_mode === "replace") {
			/* Remove all items from container */
			for(var n=_numOfContainerItems; n>0; n-=1) {
				_metadataPrefs.setProperty(_namespace, _name + "[" + n + "]", "");
			}
		}

		var _loopOffset = 0;

		/* Check: Merge mode? */
		if(_mode === "merge") {
			_loopOffset = _numOfContainerItems;
		}
		
		for(var i=0; i<_itemObjArray.length; i+=1) {
	
			var _itemObj = _itemObjArray[i];
			if(!_itemObj || !(_itemObj instanceof Object)) {
				_loopOffset -= 1;
				continue;
			}

			var _itemValue = _itemObj["value"];
			if(_itemValue === null || _itemValue === undefined || _itemValue.constructor !== String) {
				_loopOffset -= 1;
				continue;
			}

			_itemValue = _itemValue.replace(_trimRegExp, "");
			if(!_itemValue) {
				_loopOffset -= 1;
				continue;
			}

			var _itemIndex = _loopOffset + i + 1;
			if(_itemIndex <= 0) {
				continue;
			}

			var _itemPath = _name + "[" + _itemIndex + "]";

			/* Metadata property */
			_metadataPrefs.setProperty(_namespace, _itemPath, _itemValue);

			/* Attributes */
			var _attrObjArray = _itemObj["attributes"] || [];

			for(var a=0; a<_attrObjArray.length; a+=1) {

				var _attrObj = _attrObjArray[a];
				if(!_attrObj || !_attrObj.hasOwnProperty("name") || !_attrObj.hasOwnProperty("value")) {
					continue;
				}

				var _attrName = _attrObj["name"];
				var _attrValue = _attrObj["value"];
				if(!_attrName || _attrName.constructor !== String || !_attrValue || _attrValue.constructor !== String) {
					continue;
				}

				_metadataPrefs.setProperty(_namespace, _itemPath + "/@" + _attrName, _attrValue);
			}
		}
	} catch(_error) {
		return false;
	}
	
	return true;
} /* END function __setXMPContainerProperties */


/**
 * Insert Breaks
 * - Page Break
 * - Column Break
 * - Forced Line Break
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __insertBreaks(_doc, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	
	const IS_PAGE_BREAK_INSERTED = _setupObj["pageBreak"]["isInserted"];
	const PAGE_BREAK_TAG = _setupObj["pageBreak"]["tag"];
	const IS_COLUMN_BREAK_INSERTED = _setupObj["columnBreak"]["isInserted"];
	const COLUMN_BREAK_TAG = _setupObj["columnBreak"]["tag"];
	const IS_FORCED_LINE_BREAK_INSERTED = _setupObj["forcedLineBreak"]["isInserted"];
	const FORCED_LINE_BREAK_TAG = _setupObj["forcedLineBreak"]["tag"];

	if(IS_PAGE_BREAK_INSERTED) {
		var _pageBreakXMLElementsArray = _wordXMLElement.evaluateXPathExpression("//" + PAGE_BREAK_TAG);
		__insertSpecialCharacter(_pageBreakXMLElementsArray, "PAGE_BREAK");
	}
	
	if(IS_COLUMN_BREAK_INSERTED) {
		var _columnBreakXMLElementsArray = _wordXMLElement.evaluateXPathExpression("//" + COLUMN_BREAK_TAG);
		__insertSpecialCharacter(_columnBreakXMLElementsArray, "COLUMN_BREAK");
	}
	
	if(IS_FORCED_LINE_BREAK_INSERTED) {
		var _forcedLineBreakXMLElementsArray = _wordXMLElement.evaluateXPathExpression("//" + FORCED_LINE_BREAK_TAG);
		__insertSpecialCharacter(_forcedLineBreakXMLElementsArray, "FORCED_LINE_BREAK");
	}

	return true;
} /* END function __insertBreaks */

/**
 * Inserting special characters into empty XML elements
 * @param {Array} _xmlElementArray 
 * @param {String} _specialCharName 
 * @returns Boolean
 */
function __insertSpecialCharacter(_xmlElementArray, _specialCharName) {
	
	if(!_xmlElementArray || !(_xmlElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_specialCharName || _specialCharName.constructor !== String) { 
		throw new Error("String as parameter required.");
	}

	if(_xmlElementArray.length === 0) {
		return true;
	}

	var _counter = 0;

	for(var i=_xmlElementArray.length-1; i>=0; i-=1) { 

		var _xmlElement = _xmlElementArray[i];
		if(!_xmlElement || !_xmlElement.hasOwnProperty("contents") || !_xmlElement.isValid) {
			continue;
		}

		if(!SpecialCharacters.hasOwnProperty(_specialCharName)) {
			_global["log"].push(localize(_global.specialCharacterNotAvailableErrorMessage, _specialCharName));
			continue;
		}
		
		var _content = _xmlElement.contents;
		if(_content !== "") {
			_global["log"].push(localize(_global.xmlElementNotEmptyErrorMessage, _xmlElement.markupTag.name, _content));
			continue;
		}

		try {
			_xmlElement.contents = SpecialCharacters[_specialCharName]; 
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}
		
		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.insertSpecialCharactersMessage, _counter, _specialCharName));
	}

	return true;
} /* END function __insertSpecialCharacters */


/**
 * Handle Comments
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns 
 */
function __handleComments(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const COMMENT_TAG_NAME = _setupObj["comment"]["tag"];
	const COLOR_ARRAY = _setupObj["comment"]["color"];
	const IS_COMMENT_REMOVED = _setupObj["comment"]["isRemoved"];
	const IS_COMMENT_MARKED = _setupObj["comment"]["isMarked"];
	const IS_COMMENT_CREATED = _setupObj["comment"]["isCreated"];

	var _commentXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + COMMENT_TAG_NAME);
	if(_commentXMLElementArray.length === 0) {
		return true;
	}

	if(IS_COMMENT_REMOVED) {
		__removeXMLElements(_commentXMLElementArray, localize(_global.commentsLabel));
		return true;
	}

	if(IS_COMMENT_MARKED) {
		__markXMLElements(_doc, _commentXMLElementArray, localize(_global.commentsLabel), COLOR_ARRAY);
		return true;
	}

	if(IS_COMMENT_CREATED) {
		__createComments(_doc, _wordXMLElement, _commentXMLElementArray, _setupObj);
		return true;
	}
	
	return true;
} /* END function __handleComments */


/**
 * Create Comments
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _commentXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createComments(_doc, _wordXMLElement, _commentXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_commentXMLElementArray || !(_commentXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	var _wordXMLStory = _wordXMLElement.parentStory;
	if(!_wordXMLStory || !_wordXMLStory.isValid) {
		_global["log"].push(localize(_global.xmlStoryValidationError));
		return false;
	}

	var _counter = 0;
	
	for(var i=_commentXMLElementArray.length-1; i>=0; i-=1) {

		var _commentXMLElement = _commentXMLElementArray[i];
		if(!_commentXMLElement || !_commentXMLElement.isValid) {
			continue;
		}

		var _commentText = __getCommentText(_commentXMLElement, _setupObj);
		var _targetIP = _commentXMLElement.storyOffset;

		try {

			/* Add comment */
			var _comment = _wordXMLStory.notes.add(LocationOptions.BEFORE, _targetIP); /* -> DOC */
			if(!_comment || !_comment.isValid) {
				_global["log"].push(localize(_global.commentValidationErrorMessage));
				continue;
			}

			/* Insert text */
			_comment.texts[0].contents = _commentText;

			/* Remove XML container element */
			_commentXMLElement.remove();

		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.commentsLabel)));
	}

	return true;
} /* END function __createComments */


/**
 * Get content text of comment
 * @param {XMLElement} _xmlElement 
 * @param {Object} _setupObj 
 * @returns String
 */
function __getCommentText(_xmlElement, _setupObj) {

	if(!_xmlElement || !(_xmlElement instanceof XMLElement) || !_xmlElement.isValid) { return ""; }
	if(!_setupObj || !(_setupObj instanceof Object)) { return ""; }

	const PARAGRAPH_TAG_NAME = _setupObj["paragraph"]["tag"];
	const IS_METADATA_ADDED = _setupObj["comment"]["metadata"]["isAdded"];

	var _textArray = [];

	if(IS_METADATA_ADDED) {

		var _metadataArray = [];

		var _authorXMLAttribute = _xmlElement.xmlAttributes.itemByName("author");
		if(_authorXMLAttribute.isValid) {
			_metadataArray.push(_authorXMLAttribute.value);
		}

		var _dateXMLAttribute = _xmlElement.xmlAttributes.itemByName("date");
		if(_dateXMLAttribute.isValid) {
			_metadataArray.push(_dateXMLAttribute.value);
		}

		_textArray.push(_metadataArray.join(" | ") + ":");
	}

	var _paragraphXMLElementArray = _xmlElement.evaluateXPathExpression("//" + PARAGRAPH_TAG_NAME);
	
	for(var i=0; i<_paragraphXMLElementArray.length; i+=1) {
		
		var _paragraphXMLElement = _paragraphXMLElementArray[i];
		if(!_paragraphXMLElement || !_paragraphXMLElement.isValid) {
			continue;
		}

		_textArray.push(_paragraphXMLElement.contents);
	}

	return _textArray.join("\r");
} /* END function __getCommentText */


/**
 * Handle Indexmarks
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns 
 */
function __handleIndexmarks(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const INDEXMARK_TAG_NAME = _setupObj["indexmark"]["tag"];
	const IS_INDEXMARK_REMOVED = _setupObj["indexmark"]["isRemoved"];
	const IS_INDEXMARK_CREATED = _setupObj["indexmark"]["isCreated"];

	var _indexmarkXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + INDEXMARK_TAG_NAME);
	if(_indexmarkXMLElementArray.length === 0) {
		return true;
	}

	if(IS_INDEXMARK_REMOVED) {
		__removeXMLElements(_indexmarkXMLElementArray, localize(_global.indexmarksLabel));
		return true;
	}

	if(IS_INDEXMARK_CREATED) {
		__createIndexmarks(_doc, _wordXMLElement, _indexmarkXMLElementArray, _setupObj);
		return true;
	}
	
	return true;
} /* END function __handleIndexmarks */


/**
 * Create Indexmarks
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _indexmarkXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createIndexmarks(_doc, _wordXMLElement, _indexmarkXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_indexmarkXMLElementArray || !(_indexmarkXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const TYPE_ATTRIBUTE_NAME = _setupObj["indexmark"]["attributes"]["type"];
	const FORMAT_ATTRIBUTE_NAME = _setupObj["indexmark"]["attributes"]["format"];
	const ENTRY_ATTRIBUTE_NAME = _setupObj["indexmark"]["attributes"]["entry"];
	const TARGET_ATTRIBUTE_NAME = _setupObj["indexmark"]["attributes"]["target"];
	const ENTRY_SEPARATOR = _setupObj["indexmark"]["topicSeparator"];

	var _wordXMLStory = _wordXMLElement.parentStory;
	if(!_wordXMLStory || !_wordXMLStory.isValid) {
		_global["log"].push(localize(_global.xmlStoryValidationError));
		return false;
	}

	var _index = _doc.indexes.firstItem();
	if(!_index.isValid) {
		_index = _doc.indexes.add();
	}

	var _counter = 0;
	
	indexmarkLoop: 
	for(var i=_indexmarkXMLElementArray.length-1; i>=0; i-=1) {

		var _indexmarkXMLElement = _indexmarkXMLElementArray[i];
		if(!_indexmarkXMLElement || !_indexmarkXMLElement.isValid) {
			continue;
		}

		/* Type */
		var _typeAttribute = _indexmarkXMLElement.xmlAttributes.itemByName(TYPE_ATTRIBUTE_NAME);
		if(!_typeAttribute.isValid) {
			_global["log"].push(localize(_global.missingIndexmarkTypeMessage, TYPE_ATTRIBUTE_NAME));
			continue;
		}
		var _type = _typeAttribute.value;

		/* Format */
		var _formatAttribute = _indexmarkXMLElement.xmlAttributes.itemByName(FORMAT_ATTRIBUTE_NAME);
		if(!_formatAttribute.isValid) {
			_global["log"].push(localize(_global.missingIndexmarkFormatMessage, FORMAT_ATTRIBUTE_NAME));
			continue;
		}
		var _format = _formatAttribute.value;

		/* Entry */
		var _entryAttribute = _indexmarkXMLElement.xmlAttributes.itemByName(ENTRY_ATTRIBUTE_NAME);
		if(!_entryAttribute.isValid) {
			_global["log"].push(localize(_global.missingIndexmarkEntryMessage, ENTRY_ATTRIBUTE_NAME));
			continue;
		}
		var _entryValue = _entryAttribute.value;
		var _topicNameArray = _entryValue.split(ENTRY_SEPARATOR);
		if(_topicNameArray.length > 4) {
			_global["log"].push(localize(_global.maximumTopicLevelsErrorMessage, ENTRY_ATTRIBUTE_NAME, ENTRY_SEPARATOR));
			continue;
		}

		/* Target */
		var _targetAttribute = _indexmarkXMLElement.xmlAttributes.itemByName(TARGET_ATTRIBUTE_NAME);
		if(!_targetAttribute.isValid) {
			_global["log"].push(localize(_global.missingIndexmarkEntryMessage, TARGET_ATTRIBUTE_NAME));
			continue;
		}
		var _target = _targetAttribute.value;

		/* Style (overrides default number style) */
		var _numberOverrideStyle = undefined;
		if(_format !== "") {
			_numberOverrideStyle = _doc.characterStyles.itemByName(_format);
			if(!_numberOverrideStyle.isValid) {
				_numberOverrideStyle = _doc.characterStyles.add({ name:_format }); /* -> DOC */
			}
		}

		/* Create Topic */
		var _entryTopic = __createTopic(_index, _topicNameArray);
		if(!_entryTopic) {
			_global["log"].push(localize(_global.createTopicErrorMessage, _entryValue));
			continue;
		}

		var _pageRef;

		switch(_type) {
			case "r":
				/* Page range via bookmark */
				_global["log"].push(localize(_global.indexPageRangeOptionErrorMessage, _entryValue, _target));
				var _numOfParagraphs = __getNumberOfParagraphs(_wordXMLElement, _indexmarkXMLElement, _target, _setupObj);
				if(!_numOfParagraphs) {
					_global["log"].push(localize(_global.getNumberOfParagraphsErrorMessage, _entryValue, _target));
					_numOfParagraphs = 1;
				}
				_pageRef = __createPageReference(_doc, _entryTopic, _indexmarkXMLElement, "FOR_NEXT_N_PARAGRAPHS", _numOfParagraphs, _numberOverrideStyle);
				if(!_pageRef) {
					_global["log"].push(localize(_global.pageReferenceErrorMessage, _entryValue, _target));
					continue indexmarkLoop;
				}
				break;
			case "t":
				/* Add topic cross-reference */
				var _topicCrossRef = __createTopicCrossReference(_index, _entryTopic, _target, _setupObj);
				if(!_topicCrossRef || !_topicCrossRef.isValid) {
					_global["log"].push(localize(_global.topicCrossReferenceErrorMessage, _entryValue, _target));
					continue indexmarkLoop;
				}
				break;
			case "x":
				/* Add Page Reference */
				_pageRef = __createPageReference(_doc, _entryTopic, _indexmarkXMLElement, "CURRENT_PAGE", undefined, _numberOverrideStyle);
				if(!_pageRef) {
					_global["log"].push(localize(_global.pageReferenceErrorMessage, _entryValue, _target));
					continue indexmarkLoop;
				}
				break;
			default:
				_global["log"].push(localize(_global.indexmarkTypeErrorMessage, _type));
				continue indexmarkLoop;
		}
		
		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.indexmarksLabel)));
	}

	return true;
} /* END function __createIndexmarks */


/**
 * Create Page Reference for index topic
 * @param {Document} _doc
 * @param {Topic} _entryTopic
 * @param {XMLElement} _targetXMLElement 
 * @param {String} _pageReferenceType 
 * @param {ParagraphStyle|Number} _pageReferenceLimit (optional)
 * @param {CharacterStyle} _numberOverrideStyle (optional)
 * @returns PageReference
 */
function __createPageReference(_doc, _entryTopic, _targetXMLElement, _pageReferenceType, _pageReferenceLimit, _numberOverrideStyle) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return null; }
	if(!_entryTopic || !(_entryTopic instanceof Topic) || !_entryTopic.isValid) { return null; }
	if(!_targetXMLElement || !(_targetXMLElement instanceof XMLElement) || !_targetXMLElement.isValid) { return null; }
	if(!_pageReferenceType || _pageReferenceType.constructor !== String || !(PageReferenceType.hasOwnProperty(_pageReferenceType))) { return null; }

	var _tempTextframe;
	var _pageRef;
	var _isPageRefMoved = false;
	
	try {

		/* Backup character style at the target insertion point */
		var _targetIP = _targetXMLElement.insertionPoints.firstItem();
		if(!_targetIP.isValid) {
			return null;
		}
		var _targetIPCStyle = _targetIP.appliedCharacterStyle;

		/* 
			Add temporary text frame for bug fixing.
			Description: Insert a temporary text frame to work around the position shift when creating a page reference.
			Discussion: https://community.adobe.com/t5/indesign-discussions/crazy-bug-with-index-entry/m-p/10522748#M146836
		*/
		_tempTextframe = _doc.textFrames.add(); /* -> DOC */
		if(!_tempTextframe || !_tempTextframe.isValid) {
			return null;
		}

		var _tempStory = _tempTextframe.parentStory;

		/* Add page reference in temporary text frame */
		_pageRef = _entryTopic.pageReferences.add(_tempStory.texts[0], PageReferenceType[_pageReferenceType], _pageReferenceLimit, _numberOverrideStyle); /* -> DOC */
		if(!_pageRef || !_pageRef.isValid) {
			return null;
		}
		
		var _pageRefChar = _tempStory.characters.firstItem();
		if(!_pageRefChar.isValid || _pageRefChar.contents !== "\uFEFF") {
			return null;
		}

		/* Move page reference to correct position */
		_pageRefChar = _pageRefChar.move(LocationOptions.BEFORE, _targetXMLElement.texts[0]); /* -> DOC */
		if(!_pageRefChar || !_pageRefChar.isValid) {
			return null;
		}

		_isPageRefMoved = true;

		_pageRefChar.applyCharacterStyle(_targetIPCStyle);
		_pageRefChar.clearOverrides(OverrideType.CHARACTER_ONLY);

	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	} finally {
		/* Remove temporary text frame */
		if(!!_tempTextframe && _tempTextframe.hasOwnProperty("remove") && _tempTextframe.isValid) {
			_tempTextframe.remove();
		}
		/* Check: Has the page reference been moved correctly? */
		if(_isPageRefMoved === false) {
			_global["log"].push(localize(_global.movePageReferenceErrorMessage));
		}
	}
	
	/* Check: Is page reference valid? */
	if(!_pageRef || !_pageRef.isValid) {
		return null;
	}

	return _pageRef;
} /* END function __createPageReference */


/**
 * Add cross-reference for index topic
 * @param {Index} _index
 * @param {Topic} _entryTopic 
 * @param {String} _target 
 * @param {Object} _setupObj 
 * @returns CrossReference
 */
function __createTopicCrossReference(_index, _entryTopic, _target, _setupObj) {

	if(!_index || !(_index instanceof Index) || !_index.isValid) { return null; }
	if(!_entryTopic || !(_entryTopic instanceof Topic) || !_entryTopic.isValid) { return null; }
	if(_target ===  null || _target === undefined || _target.constructor !== String) { return null; }
	if(!_setupObj || !(_setupObj instanceof Object))  { return null; }

	const TOPIC_SEPARATOR = ":";
	
	const _crossRefPrefixObjArray = _setupObj["indexmark"]["crossReference"]["prefixes"];
	const _noMatchCustomTypeString = _setupObj["indexmark"]["crossReference"]["noMatchCustomTypeString"];
	const _horizontalWhitespaces = "[^\\S\\r\\n]";
	const _referencedTopicNameSplitRegExp = new RegExp(TOPIC_SEPARATOR + _horizontalWhitespaces + "*","");
	const _specialCharRegExp = new RegExp("([.*+?()[\\]{}\\^$|\\~\\\\])", "g");

	/* Get cross-reference type and cross-reference custom string */
	var _crossRefType;
	var _crossRefCustomString = "";
	var _referencedTopicName = _target;

	for(var i=0; i<_crossRefPrefixObjArray.length; i+=1) {

		var _crossRefPrefixObj = _crossRefPrefixObjArray[i];
		if(!_crossRefPrefixObj || !(_crossRefPrefixObj instanceof Object) || !_crossRefPrefixObj.hasOwnProperty("en")) {
			continue;
		}

		var _crossRefPrefix = localize(_crossRefPrefixObj);
		if(!_crossRefPrefix || _crossRefPrefix.constructor !== String) {
			continue;
		}

		var escapedCrossRefPrefix = _crossRefPrefix.replace(_specialCharRegExp, "\\$1");
		var _crossRefPrefixRegExp = new RegExp("^" + escapedCrossRefPrefix + _horizontalWhitespaces + "+","i");
		var _crossRefPrefixMatchArray = _target.match(_crossRefPrefixRegExp);
		if(!_crossRefPrefixMatchArray || _crossRefPrefixMatchArray.length === 0) {
			continue;
		}

		var _crossRefKey = _crossRefPrefixObj["en"];
		switch(_crossRefKey) {
			case "See [also]":
				_crossRefType = CrossReferenceType.SEE_OR_ALSO_BRACKET;
				break;
			case "See also herein":
				_crossRefType = CrossReferenceType.SEE_ALSO_HEREIN;
				break;
			case "See also":
				_crossRefType = CrossReferenceType.SEE_ALSO;
				break;
			case "See herein":
				_crossRefType = CrossReferenceType.SEE_HEREIN;
				break;
			case "See":
				_crossRefType = CrossReferenceType.SEE;
				break;
			default:
				_crossRefType = CrossReferenceType.CUSTOM_CROSS_REFERENCE_BEFORE;
				_crossRefCustomString = _crossRefPrefixMatchArray[0];
		}

		_referencedTopicName = _target.replace(_crossRefPrefixRegExp, "");
		break;
	}

	/* Check: No match for cross-reference prefix? */
	if(!_crossRefType) {
		_crossRefType = CrossReferenceType.CUSTOM_CROSS_REFERENCE;
		_crossRefCustomString = _noMatchCustomTypeString;
	}

	var _referencedTopicNameArray = _referencedTopicName.split(_referencedTopicNameSplitRegExp);
	var _referencedTopic = __createTopic(_index, _referencedTopicNameArray);
	if(!_referencedTopic) {
		_global["log"].push(localize(_global.createTopicErrorMessage, _referencedTopicName));
		return null;
	}

	/* Check: Cross-reference already exists? */
	var _topicCrossRef = __getTopicCrossReference(_entryTopic, _referencedTopic, _crossRefCustomString);
	if(_topicCrossRef === null) {
		try {
			/* Add cross-reference */
			_topicCrossRef = _entryTopic.crossReferences.add(_referencedTopic, _crossRefType, _crossRefCustomString); /* -> DOC */
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			return null;
		}
	} 

	return _topicCrossRef;
} /* END function __createTopicCrossReference */


/**
 * Get index cross-reference
 * Check if cross-reference already exists. If yes, return cross-reference.
 * Reason: An exactly identical cross-reference cannot be created again.
 * @param {Topic} _parentTopic 
 * @param {Topic} _referencedTopic 
 * @param {String} _customTypeString 
 * @returns CrossRefference|Null
 */
function __getTopicCrossReference(_parentTopic, _referencedTopic, _customTypeString) {
	
	if(!_parentTopic || !(_parentTopic instanceof Topic) || !_parentTopic.isValid) { return false; }
	if(!_referencedTopic || !(_referencedTopic instanceof Topic) || !_referencedTopic.isValid) { return false; }
	if(_customTypeString === null || _customTypeString === undefined || _customTypeString.constructor !== String) { return false; }
	
	var _crossReferenceArray = _parentTopic.crossReferences.everyItem().getElements();
	
	for(var i=0; i<_crossReferenceArray.length; i+=1) {
		
		var _curCrossReference = _crossReferenceArray[i];
		if(!_curCrossReference || !_curCrossReference.isValid) {
			continue;
		}
		
		if(
			_curCrossReference.referencedTopic === _referencedTopic &&
			_curCrossReference.customTypeString === _customTypeString
		) {
			return _curCrossReference;
		}
	}
	
	return null;
} /* END function __getTopicCrossReference */


/**
 * Create index topic 
 * If the topic does not exist, one will be created.
 * @param {Index|Topic} _inputTopic Document index.
 * @param {Array} _inputTopicNameArray Array with names of topics.
 */
function __createTopic(_inputTopic, _inputTopicNameArray) {

	if(!_inputTopic || !(_inputTopic instanceof Index || _inputTopic instanceof Topic) || !_inputTopic.isValid) {
		return null;
	}
	if(!_inputTopicNameArray || !(_inputTopicNameArray instanceof Array) || _inputTopicNameArray.length === 0) {
		return null;
	}

	const _trimWhitespaceRegExp = new RegExp("(^\\s+)|(\\s+$)","g"); 

	var _rawCurTopicName = _inputTopicNameArray[0];
	if(!_rawCurTopicName) {
		return null;
	}

	var _curTopicName = _rawCurTopicName.toString().replace(_trimWhitespaceRegExp, "");
	if(_curTopicName === "") {
		return null;
	}

	var _curTopic = _inputTopic.topics.itemByName(_curTopicName);
	if(!_curTopic.isValid) {
		try {
			/* Add topic */
			_curTopic = _inputTopic.topics.add(_curTopicName); /* -> DOC */
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			return null;
		}
	}
	
	var _outputTopic = _curTopic;

	var _childTopicArray = _inputTopicNameArray.slice(1, _inputTopicNameArray.length);
	if(_childTopicArray.length !== 0) {
		_outputTopic = __createTopic(_curTopic, _childTopicArray);
	}

	return _outputTopic;
} /* END function __createTopic */


/**
 * Get number of paragraphs between two XML elements
 * (Index entry element and bookmark element)
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _indexmarkXMLElement 
 * @param {String} _bookmarkID 
 * @param {Object} _setupObj 
 * @returns Number | Null
 */
function __getNumberOfParagraphs(_wordXMLElement, _indexmarkXMLElement, _bookmarkID, _setupObj) {

	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { return null; }
	if(!_indexmarkXMLElement || !(_indexmarkXMLElement instanceof XMLElement) || !_indexmarkXMLElement.isValid) { return null;	}
	if(!_bookmarkID || _bookmarkID.constructor !== String) { return null;	}
	if(!_setupObj || !(_setupObj instanceof Object)) { return null;	}

	const BOOKMARK_TAG_NAME = _setupObj["bookmark"]["tag"];
	const BOOKMARK_ID_ATTRIBUTE_NAME = _setupObj["bookmark"]["attributes"]["id"];

	var _bookmarkXMLElement = _wordXMLElement.evaluateXPathExpression("//" + BOOKMARK_TAG_NAME + "[@" + BOOKMARK_ID_ATTRIBUTE_NAME + " = '" + _bookmarkID + "']")[0];
	if(!_bookmarkXMLElement || !_bookmarkXMLElement.isValid) {
		_global["log"].push(localize(_global.indexEntryBookmarkNotFoundMessage, _bookmarkID));
		return null;
	}

	var _indexmarkStory = _indexmarkXMLElement.parentStory;
	var _bookmarkStory = _bookmarkXMLElement.parentStory;
	if(!_indexmarkStory.isValid || !_bookmarkStory.isValid || _indexmarkStory !== _bookmarkStory) {
		return null;
	}

	var _paragraphRange = _indexmarkStory.paragraphs.itemByRange(_indexmarkXMLElement.paragraphs[0], _bookmarkXMLElement.paragraphs[0]);
	if(!_paragraphRange.isValid) {
		return null;
	}

	var _numOfParagraphs = _paragraphRange.paragraphs.count();
	if(_numOfParagraphs < 1) {
		return null;
	}

	return _numOfParagraphs;
} /* END function __getNumberOfParagraphs */


/**
 * Handle Hyperlinks
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleHyperlinks(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const HYPERLINK_TAG_NAME = _setupObj["hyperlink"]["tag"];
	const COLOR_ARRAY = _setupObj["hyperlink"]["color"];
	const IS_HYPERLINK_MARKED = _setupObj["hyperlink"]["isMarked"];
	const IS_HYPERLINK_CREATED = _setupObj["hyperlink"]["isCreated"];
	const IS_HYPERLINK_IGNORED = _setupObj["hyperlink"]["isIgnored"];

	if(IS_HYPERLINK_IGNORED) {
		return true;
	}

	var _hyperlinkXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + HYPERLINK_TAG_NAME);
	if(_hyperlinkXMLElementArray.length === 0) {
		return true;
	}

	if(IS_HYPERLINK_MARKED) {
		__markXMLElements(_doc, _hyperlinkXMLElementArray, localize(_global.hyperlinksLabel), COLOR_ARRAY);
		return true;
	}

	if(IS_HYPERLINK_CREATED) {
		__createHyperlinks(_doc, _wordXMLElement, _hyperlinkXMLElementArray, _setupObj);
		return true;
	}
	
	return true;
} /* END function __handleHyperlinks */


/**
 * Create Hyperlinks
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _hyperlinkXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createHyperlinks(_doc, _wordXMLElement, _hyperlinkXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_hyperlinkXMLElementArray || !(_hyperlinkXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const URI_ATTRIBUTE_NAME = _setupObj["hyperlink"]["attributes"]["uri"];
	const TITLE_ATTRIBUTE_NAME = _setupObj["hyperlink"]["attributes"]["title"];
	const CHARACTER_STYLE_NAME = _setupObj["hyperlink"]["characterStyleName"];
	const IS_CHARACTER_STYLE_ADDED = _setupObj["hyperlink"]["isCharacterStyleAdded"];

	const BOOKMARK_TAG_NAME = _setupObj["bookmark"]["tag"];
	const BOOKMARK_ID_ATTRIBUTE_NAME = _setupObj["bookmark"]["attributes"]["id"];

	const _anchorOnlyRegExp = new RegExp("^#","");
	const _urlRegExp = new RegExp("(https?|ftp|mailto):","i");
	const _clearFilePathRegExp = new RegExp("^(../)+","");

	var _counter = 0;

	var _cStyle;
	if(IS_CHARACTER_STYLE_ADDED) {
		_cStyle = _doc.characterStyles.itemByName(CHARACTER_STYLE_NAME);
		if(!_cStyle.isValid) {
			_cStyle = _doc.characterStyles.add({ name:CHARACTER_STYLE_NAME }); /* -> DOC */
		}
	}
	
	for(var i=_hyperlinkXMLElementArray.length-1; i>=0; i-=1) {

		var _hyperlinkXMLElement = _hyperlinkXMLElementArray[i];
		if(!_hyperlinkXMLElement || !_hyperlinkXMLElement.isValid) {
			continue;
		}

		/* URI */
		var _uriAttribute = _hyperlinkXMLElement.xmlAttributes.itemByName(URI_ATTRIBUTE_NAME);
		if(!_uriAttribute.isValid) {
			_global["log"].push(localize(_global.missingHyperlinkURIMessage, URI_ATTRIBUTE_NAME));
			continue;
		}
		var _uri = decodeURI(_uriAttribute.value);
		if(!_uri) {
			_global["log"].push(localize(_global.missingHyperlinkURIMessage, URI_ATTRIBUTE_NAME));
			continue;
		}

		/* Title */
		var _title = "";
		var _titleAttribute = _hyperlinkXMLElement.xmlAttributes.itemByName(TITLE_ATTRIBUTE_NAME);
		if(_titleAttribute.isValid) {
			_title = _titleAttribute.value;
		}

		var _hyperlinkSource;
		var _hyperlinkDestination;

		try {

			/* Add hyperlink */
			var _hyperlinkXMLElementText = _hyperlinkXMLElement.texts[0];
			var _hyperlinkXMLElementContent = _hyperlinkXMLElementText.contents;
			if(_hyperlinkXMLElementContent === "") {
				_global["log"].push(localize(_global.emptyHyperlinkSourceMessage, (i+1), _uri));
				continue;
			}
			_hyperlinkSource = _doc.hyperlinkTextSources.add(_hyperlinkXMLElementText); /* -> DOC */

			/* Character Style */
			if(IS_CHARACTER_STYLE_ADDED && _cStyle && _cStyle.isValid) {
				_hyperlinkSource.appliedCharacterStyle = _cStyle;
			}
			
			/* Check: Anchor as destination? */
			_hyperlinkDestination = null;
			if(_anchorOnlyRegExp.test(_uri)) {
				var _bookmarkID = _uri.replace(_anchorOnlyRegExp,'');
				var _bookmarkXMLElement = _wordXMLElement.evaluateXPathExpression("//" + BOOKMARK_TAG_NAME + "[@" + BOOKMARK_ID_ATTRIBUTE_NAME + " = '" + _bookmarkID + "']")[0];
				if(_bookmarkXMLElement && _bookmarkXMLElement.isValid) {
					_hyperlinkDestination = _doc.hyperlinkTextDestinations.add(_bookmarkXMLElement.texts[0], { hidden: true }); /* -> DOC */
				}
			}
			/* Check: URI/file as destination? */
			if(!_hyperlinkDestination || !_hyperlinkDestination.isValid) {
				if(!_urlRegExp.test(_uri)) {
					_uri = "file:" + _uri.replace(_clearFilePathRegExp,"/");
				}
				_hyperlinkDestination = _doc.hyperlinkURLDestinations.add(_uri, { hidden: true }); /* -> DOC */
			}
			
			var _hyperlink = _doc.hyperlinks.add(_hyperlinkSource, _hyperlinkDestination); /* -> DOC */
			_hyperlink.visible = false;

			/* Add label to hyperlink */
			if(_title !== "") {
				_hyperlink.label = _title;
			}
		} catch(_error) {

			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));

			/* Clean up */
			if(_hyperlinkSource && _hyperlinkSource.isValid) {
				_hyperlinkSource.remove(); /* This also automatically removes hyperlink and hyperlink destination */
			}
			if(_hyperlinkDestination && _hyperlinkDestination.isValid) {
				_hyperlinkDestination.remove();
			}
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.hyperlinksLabel)));
	}

	return true;
} /* END function __createHyperlinks */


/**
 * Handle Cross-references
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleCrossReferences(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const CROSS_REFERENCE_TAG_NAME = _setupObj["crossReference"]["tag"];
	const COLOR_ARRAY = _setupObj["crossReference"]["color"];
	const IS_CROSS_REFERENCE_MARKED = _setupObj["crossReference"]["isMarked"];
	const IS_CROSS_REFERENCE_CREATED = _setupObj["crossReference"]["isCreated"];
	const IS_CROSS_REFERENCE_IGNORED = _setupObj["crossReference"]["isIgnored"];

	if(IS_CROSS_REFERENCE_IGNORED) {
		return true;
	}

	var _crossRefXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + CROSS_REFERENCE_TAG_NAME);
	if(_crossRefXMLElementArray.length === 0) {
		return true;
	}

	if(IS_CROSS_REFERENCE_MARKED) {
		__markXMLElements(_doc, _crossRefXMLElementArray, localize(_global.crossReferencesLabel), COLOR_ARRAY);
		return true;
	}

	if(IS_CROSS_REFERENCE_CREATED) {
		__createCrossReferences(_doc, _wordXMLElement, _crossRefXMLElementArray, _setupObj);
		return true;
	}
	
	return true;
} /* END function __handleCrossReferences */


/**
 * Create Cross-references
 * (possible helper function __getUniqueHyperlinkName)
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _crossRefXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createCrossReferences(_doc, _wordXMLElement, _crossRefXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_crossRefXMLElementArray || !(_crossRefXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const URI_ATTRIBUTE_NAME = _setupObj["crossReference"]["attributes"]["uri"];
	const TYPE_ATTRIBUTE_NAME = _setupObj["crossReference"]["attributes"]["type"];
	const FORMAT_ATTRIBUTE_NAME = _setupObj["crossReference"]["attributes"]["format"];
	const CHARACTER_STYLE_NAME = _setupObj["crossReference"]["characterStyleName"];
	const IS_ANCHOR_HIDDEN = _setupObj["crossReference"]["isAnchorHidden"];
	const IS_CHARACTER_STYLE_ADDED = _setupObj["crossReference"]["isCharacterStyleAdded"];

	const BOOKMARK_TAG_NAME = _setupObj["bookmark"]["tag"];
	const BOOKMARK_ID_ATTRIBUTE_NAME = _setupObj["bookmark"]["attributes"]["id"];

	const _anchorRegExp = new RegExp("^#","");
	const _trimWhitespaceRegExp = new RegExp("(^\\s+)|(\\s+$)","g");

	var _cStyle;
	if(IS_CHARACTER_STYLE_ADDED) {
		_cStyle = _doc.characterStyles.itemByName(CHARACTER_STYLE_NAME);
		if(!_cStyle.isValid) {
			_cStyle = _doc.characterStyles.add({ name:CHARACTER_STYLE_NAME }); /* -> DOC */
		}
	}

	var _counter = 0;
	
	xmlElementLoop: for(var i=_crossRefXMLElementArray.length-1; i>=0; i-=1) {

		var _crossRefXMLElement = _crossRefXMLElementArray[i];
		if(!_crossRefXMLElement || !_crossRefXMLElement.isValid) {
			continue;
		}

		/* Cross-reference content */
		var _rawCrossRefContent = _crossRefXMLElement.texts[0].contents;
		var _crossRefContent =  _rawCrossRefContent.replace(_trimWhitespaceRegExp,"");

		/* URI */
		var _uriAttribute = _crossRefXMLElement.xmlAttributes.itemByName(URI_ATTRIBUTE_NAME);
		if(!_uriAttribute.isValid) {
			_global["log"].push(localize(_global.missingCrossReferenceURIMessage, URI_ATTRIBUTE_NAME));
			continue;
		}
		var _uri = decodeURI(_uriAttribute.value);
		if(!_uri) {
			_global["log"].push(localize(_global.missingCrossReferenceURIMessage, URI_ATTRIBUTE_NAME));
			continue;
		}
		var _bookmarkID = _uri.replace(_anchorRegExp, "");

		/* Type */
		var _typeAttribute = _crossRefXMLElement.xmlAttributes.itemByName(TYPE_ATTRIBUTE_NAME);
		if(!_typeAttribute.isValid) {
			_global["log"].push(localize(_global.missingCrossReferenceTypeMessage, TYPE_ATTRIBUTE_NAME));
			continue;
		}
		var _type = _typeAttribute.value;

		/* Format */
		var _formatAttribute = _crossRefXMLElement.xmlAttributes.itemByName(FORMAT_ATTRIBUTE_NAME);
		if(!_formatAttribute.isValid) {
			_global["log"].push(localize(_global.missingCrossReferenceFormatMessage, FORMAT_ATTRIBUTE_NAME));
			continue;
		}
		var _format = _formatAttribute.value;

		/* Bookmark XML element (Cross-reference destination) */
		var _bookmarkXMLElement = _wordXMLElement.evaluateXPathExpression("//" + BOOKMARK_TAG_NAME + "[@" + BOOKMARK_ID_ATTRIBUTE_NAME + " = '" + _bookmarkID + "']")[0];
		if(!_bookmarkXMLElement || !_bookmarkXMLElement.isValid) {
			_global["log"].push(localize(_global.crossReferenceDestinationNotFoundMessage, _bookmarkID));
			continue;
		}

		var _formatID = ""; /* String (Format name) or Number (Format index) */
		var _blockTypeArray = [];

		switch(_type) {
			/* Page */
			case "PAGEREF":
				/* Page label + Page number */
				if(/\bp\b/i.test(_format)) {
					_formatID = localize(_global.pageLabel) + " " + localize(_global.pageNumberCrossReferenceFormatName) + localize(_global.crossReferenceFormatWordImportLabel);
					_blockTypeArray.push({"type":BuildingBlockTypes.CUSTOM_STRING_BUILDING_BLOCK, "text":(localize(_global.pageLabel) + " ")});
					_blockTypeArray.push({"type":BuildingBlockTypes.PAGE_NUMBER_BUILDING_BLOCK});
				} 
				/* Page number */
				else {
					_formatID = localize(_global.pageNumberCrossReferenceFormatName) + localize(_global.crossReferenceFormatWordImportLabel);
					_blockTypeArray.push({"type":BuildingBlockTypes.PAGE_NUMBER_BUILDING_BLOCK});
				}
				break;
			/* Footnote/Endnote */
			case "NOTEREF":
				/* Footnote/Endnote number */
				/* (Unfortunately, InDesign does not support this type of cross-references to footnotes or endnotes.) */
				continue xmlElementLoop;
				break;
			/* Paragraph */
			case "REF":
				/* r: Paragraph number, n: Paragraph number without context, w: Paragraph number with full context */
				if(/\b(r|n|w)\b/i.test(_format)) {
					_formatID = localize(_global.paragraphNumberCrossReferenceFormatName) + localize(_global.crossReferenceFormatWordImportLabel);
					_blockTypeArray.push({"type":BuildingBlockTypes.PARAGRAPH_NUMBER_BUILDING_BLOCK});
				}
				/* Custom text (above/below) */
				else if(/\bp\b/i.test(_format)) {
					_formatID = _crossRefContent + localize(_global.crossReferenceFormatWordImportLabel);
					_blockTypeArray.push({"type":BuildingBlockTypes.CUSTOM_STRING_BUILDING_BLOCK, "text":_crossRefContent});
				}
				/* Paragraph Number + Text */
				else {
					_formatID = localize(_global.paragraphTextCrossReferenceFormatName) + localize(_global.crossReferenceFormatWordImportLabel);
					_blockTypeArray.push({"type":BuildingBlockTypes.PARAGRAPH_TEXT_BUILDING_BLOCK});
				}
				break;
			default:
				_global["log"].push(localize(_global.noMatchingCrossReferenceTypeMessage));
				continue xmlElementLoop;
		}
		
		/* Cross-reference Format */
		var _crossRefFormat = __createCrossReferenceFormat(_doc, _formatID, _blockTypeArray, _cStyle);
		if(!_crossRefFormat || !_crossRefFormat.isValid) {
			_global["log"].push(localize(_global.crossReferenceValidationMessage, _type, _format));
			continue;
		}

		/* Cross-references destination properties */
		var _crossRefDestinationProps = { 
			hidden: IS_ANCHOR_HIDDEN 
		};

		var _crossRefDestination;
		var _crossRefSource;

		try {
			/* Cross-reference Destination */
			_crossRefDestination = _doc.hyperlinkTextDestinations.add(_bookmarkXMLElement.texts[0], _crossRefDestinationProps); /* -> DOC */
			/* Cross-reference source */
			_crossRefSource = _doc.crossReferenceSources.add(_crossRefXMLElement.texts[0],_crossRefFormat); /* -> DOC */
			/* Add Cross-reference */
			_doc.hyperlinks.add(_crossRefSource, _crossRefDestination); /* -> DOC */
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			/* Clean up */
			if(_crossRefSource && _crossRefSource.isValid) {
				_crossRefSource.remove(); /* This also automatically removes hyperlink and hyperlink destination */
			}
			if(_crossRefDestination && _crossRefDestination.isValid) {
				_crossRefDestination.remove();
			}
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.crossReferencesLabel)));
	}

	return true;
} /* END function __createHyperlinks */


/**
 * Create Cross-reference Format
 * @param {Document} _doc 
 * @param {Number|String} _crossRefFormatId 
 * @param {Array} _buildingBlockTypeArray 
 * @param {CharacterStyle} _cStyle (optional)
 * @returns CrossReferenceFormat
 */
function __createCrossReferenceFormat(_doc, _crossRefFormatId, _buildingBlockTypeArray, _cStyle) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return null; }
	if(_crossRefFormatId === null || _crossRefFormatId === undefined || !(_crossRefFormatId.constructor === Number || _crossRefFormatId.constructor === String) || _crossRefFormatId === "") { return null; }
	if(!_buildingBlockTypeArray || !(_buildingBlockTypeArray instanceof Array) || _buildingBlockTypeArray.length === 0) { return null; }
	
	var _crossRefFormat;

	if(_crossRefFormatId.constructor === Number) {
		_crossRefFormat = _doc.crossReferenceFormats.item(_crossRefFormatId);
	} 
	else if(_crossRefFormatId.constructor === String) {
		_crossRefFormat = _doc.crossReferenceFormats.itemByName(_crossRefFormatId);
	} 
	else {
		return null;
	}

	if(_crossRefFormat && _crossRefFormat.isValid) {
		return _crossRefFormat;
	}

	try {
		/* Add cross-reference format */
		_crossRefFormat = _doc.crossReferenceFormats.add(_crossRefFormatId.toString()); /* -> DOC */
		for(var i=0; i<_buildingBlockTypeArray.length; i+=1) {
			var _buildingBlockTypeObj = _buildingBlockTypeArray[i];
			if(!_buildingBlockTypeObj.hasOwnProperty("type")) {
				continue;
			}
			var _buildingBlockType = _buildingBlockTypeObj["type"];
			if(!BuildingBlockTypes.hasOwnProperty(_buildingBlockType)) {
				continue;
			}
			var _buildingBlock = _crossRefFormat.buildingBlocks.add(_buildingBlockType);
			if(_buildingBlock.blockType !== BuildingBlockTypes.CUSTOM_STRING_BUILDING_BLOCK) {
				continue;
			}
			var _customText = _buildingBlockTypeObj["text"];
			if(!_customText || _customText.constructor !== String) {
				continue;
			}
			_buildingBlock.customText = _customText;
		}
		/* Add character style */
		if(_cStyle && _cStyle instanceof CharacterStyle && _cStyle.isValid) {
			_crossRefFormat.appliedCharacterStyle = _cStyle;
		}
	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	}

	return _crossRefFormat;
} /* END function __createCrossReferenceFormat */


/**
 * Handle Bookmarks 
 * (Text anchors with content in Word document)
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleBookmarks(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const BOOKMARK_TAG_NAME = _setupObj["bookmark"]["tag"];
	const IS_BOOKMARK_CREATED = _setupObj["bookmark"]["isCreated"];

	var _bookmarkXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + BOOKMARK_TAG_NAME);
	if(_bookmarkXMLElementArray.length === 0) {
		return true;
	}

	if(IS_BOOKMARK_CREATED) {
		__createBookmarks(_doc, _wordXMLElement, _bookmarkXMLElementArray, _setupObj);
		return true;
	}

	return true;
} /* END function __handleBookmarks */


/**
 * Create Bookmarks
 * (possible helper function __getUniqueHyperlinkName)
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {XMLElement} _bookmarkXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createBookmarks(_doc, _wordXMLElement, _bookmarkXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_bookmarkXMLElementArray || !(_bookmarkXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const CONTENT_ATTRIBUTE_NAME = _setupObj["bookmark"]["attributes"]["content"];
	const ID_ATTRIBUTE_NAME = _setupObj["bookmark"]["attributes"]["id"];
	const IS_ANCHOR_HIDDEN = _setupObj["bookmark"]["isAnchorHidden"];
	
	const MAX_NAME_LENGTH = 100;

	var _marker = _setupObj["bookmark"]["marker"];
	var _markerRegExp;

	try {
		_markerRegExp = new RegExp("^" + _marker, "");
	} catch(_error) {
		_global["log"].push(localize(_global.bookmarksLabel) + " " + localize(_global.bookmarkMarkerLabel) + " [" + _marker + "]: " + _error.message);
		_marker = "";
	}

	var _counter = 0;
	
	for(var i=_bookmarkXMLElementArray.length-1; i>=0; i-=1) {

		var _bookmarkXMLElement = _bookmarkXMLElementArray[i];
		if(!_bookmarkXMLElement || !_bookmarkXMLElement.isValid) {
			continue;
		}

		/* Content */
		var _contentAttribute = _bookmarkXMLElement.xmlAttributes.itemByName(CONTENT_ATTRIBUTE_NAME);
		if(!_contentAttribute.isValid) {
			continue;
		}

		var _content = __cleanUpString(_contentAttribute.value, true, true);
		if(!_content) {
			continue;
		}

		/* ID */
		var _idAttribute = _bookmarkXMLElement.xmlAttributes.itemByName(ID_ATTRIBUTE_NAME);
		if(!_idAttribute.isValid) {
			continue;
		}

		var _id = _idAttribute.value;
		if(!_id) {
			continue;
		}

		/* Check: Marker available and included in ID? */
		if(_marker && !_markerRegExp.test(_id)) {
			continue;
		}

		var _bookmarkName = _content.substring(0, MAX_NAME_LENGTH);

		/* Bookmark destination properties */
		var _bookmarkDestinationProps = { 
			hidden: IS_ANCHOR_HIDDEN 
		};

		var _bookmarkDestination;
		var _bookmark;

		try {

			/* Bookmark Destination */
			_bookmarkDestination = _doc.hyperlinkTextDestinations.add(_bookmarkXMLElement.texts[0], _bookmarkDestinationProps); /* -> DOC */

			/* Add Bookmark */
			_bookmark = _doc.bookmarks.add(_bookmarkDestination); /* -> DOC */

			_bookmark.move(LocationOptions.AT_BEGINNING, _doc);
			_bookmark.name = _bookmarkName;
			
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			/* Clean up */
			if(_bookmark && _bookmark.isValid) {
				_bookmark.remove();
			}
			if(_bookmarkDestination && _bookmarkDestination.isValid) {
				_bookmarkDestination.remove();
			}
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.bookmarksLabel)));
	}

	return true;
} /* END function __createBookmarks */


/**
 * Handle Textboxes
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleTextboxes(_doc, _wordXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const TEXTBOX_TAG_NAME = _setupObj["textbox"]["tag"];
	const TEXTBOX_COLOR_ARRAY = _setupObj["textbox"]["color"];
	const IS_TEXTBOX_REMOVED = _setupObj["textbox"]["isRemoved"];
	const IS_TEXTBOX_MARKED = _setupObj["textbox"]["isMarked"];
	const IS_TEXTBOX_CREATED = _setupObj["textbox"]["isCreated"];

	var _textboxXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + TEXTBOX_TAG_NAME);
	if(_textboxXMLElementArray.length === 0) {
		return true;
	}

	if(IS_TEXTBOX_REMOVED) {
		__removeXMLElements(_textboxXMLElementArray, localize(_global.textboxesLabel));
		return true;
	}

	if(IS_TEXTBOX_MARKED) {
		__markXMLElements(_doc, _textboxXMLElementArray, localize(_global.textboxesLabel), TEXTBOX_COLOR_ARRAY);
		return true;
	}

	if(IS_TEXTBOX_CREATED) {
		__createTextboxes(_doc, _textboxXMLElementArray, _setupObj);
		return true;
	}
	
	return true;
} /* END function __handleTextboxes */


/**
 * Create Textboxes
 * @param {Document} _doc 
 * @param {XMLElement} _textboxXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createTextboxes(_doc, _textboxXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_textboxXMLElementArray || !(_textboxXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const OBJECT_STYLE_ATTRIBUTE_NAME = "ostyle";
	const OBJECT_STYLE_PROPERTIES = _setupObj["textbox"]["objectStyleProperties"];
	const TEXTBOX_WIDTH_IN_MM = _setupObj["textbox"]["width"];
	const TEXTBOX_HEIGHT_IN_MM = _setupObj["textbox"]["height"];

	const TEXTBOX_WIDTH = UnitValue(TEXTBOX_WIDTH_IN_MM, MeasurementUnits.MILLIMETERS).as(_doc.viewPreferences.verticalMeasurementUnits);
	const TEXTBOX_HEIGHT = UnitValue(TEXTBOX_HEIGHT_IN_MM, MeasurementUnits.MILLIMETERS).as(_doc.viewPreferences.verticalMeasurementUnits);

	var _counter = 0;
	
	for(var i=_textboxXMLElementArray.length-1; i>=0; i-=1) {

		var _textboxXMLElement = _textboxXMLElementArray[i];
		if(!_textboxXMLElement || !_textboxXMLElement.isValid) {
			continue;
		}

		/* Textbox Object Style */
		var _oStyleName = "";
		var _oStyleXMLAttribute = _textboxXMLElement.xmlAttributes.itemByName(OBJECT_STYLE_ATTRIBUTE_NAME);
		if(_oStyleXMLAttribute.isValid) {
			_oStyleName = _oStyleXMLAttribute.value;
		}

		try {

			/* Add comment */
			var _textbox = _textboxXMLElement.placeIntoInlineFrame([TEXTBOX_WIDTH,TEXTBOX_HEIGHT]);
			
			/* Apply object style */
			if(_oStyleName !== "") {
				var _oStyle = _doc.objectStyles.itemByName(_oStyleName);
				if(!_oStyle.isValid) {
					_oStyle = _doc.objectStyles.add({ name:_oStyleName }); /* -> DOC */
					_oStyle.properties = OBJECT_STYLE_PROPERTIES;
				}
				_textbox.applyObjectStyle(_oStyle, true, true);
			}

			/* Apply paragraph styles */
			__applyStylesToTextboxParagraphs(_doc, _textboxXMLElement, _setupObj);

		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.textboxesLabel)));
	}

	return true;
} /* END function __createTextboxes */


/**
 * Apply styles to textbox paragraphps
 * @param {Document} _doc 
 * @param {XMLElement} _textboxXMLElement 
 * @param {Object} _unpackObj
 * @param {Object} _setupObj
 * @returns Boolean
 */
function __applyStylesToTextboxParagraphs(_doc, _textboxXMLElement, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return false; }
	if(!_textboxXMLElement || !(_textboxXMLElement instanceof XMLElement) || !_textboxXMLElement.isValid) { return false; }
	if(!_setupObj || !(_setupObj instanceof Object)) { return false; }

	const PARAGRAPH_TAG_NAME = _setupObj["paragraph"]["tag"];
	const STYLE_ATTRIBUTE_NAME = "pstyle";

	var _isAssignmentCorrect = true;

	var _paragraphXMLElementArray = _textboxXMLElement.evaluateXPathExpression(PARAGRAPH_TAG_NAME);
	
	for(var i=0; i<_paragraphXMLElementArray.length; i+=1) {

		var _paragraphXMLElement = _paragraphXMLElementArray[i];
		if(!_paragraphXMLElement || !_paragraphXMLElement.isValid) {
			continue;
		}

		var _pStyleAttribute = _paragraphXMLElement.xmlAttributes.itemByName(STYLE_ATTRIBUTE_NAME);
		if(!_pStyleAttribute.isValid) {
			continue;
		}

		var _pStyleName = _pStyleAttribute.value;
		if(!_pStyleName) {
			continue;
		}

		var _pStyle = _doc.paragraphStyles.itemByName(_pStyleName);
		if(!_pStyle.isValid) {
			_pStyle = _doc.paragraphStyles.add({ name:_pStyleName }); /* -> DOC */
		}

		try {
			_paragraphXMLElement.applyParagraphStyle(_pStyle, true);
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			_isAssignmentCorrect = false;
			continue;
		}
	}

	if(_isAssignmentCorrect === false) {
		return false;
	}

	return true;
} /* END function __applyStylesToNoteParagraphs */


/**
 * Handle Images
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _unpackObj 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleImages(_doc, _wordXMLElement, _unpackObj, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const IMAGE_TAG_NAME = _setupObj["image"]["tag"];
	const IMAGE_COLOR_ARRAY = _setupObj["image"]["color"];
	const IS_IMAGE_REMOVED = _setupObj["image"]["isRemoved"];
	const IS_IMAGE_MARKED = _setupObj["image"]["isMarked"];
	const IS_IMAGE_PLACED = _setupObj["image"]["isPlaced"];

	var _imageXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + IMAGE_TAG_NAME);
	if(_imageXMLElementArray.length === 0) {
		return true;
	}

	if(IS_IMAGE_REMOVED) {
		__removeXMLElements(_imageXMLElementArray, localize(_global.imagesLabel));
		return true;
	}

	if(IS_IMAGE_MARKED) {
		__insertImageSources(_doc, _imageXMLElementArray, _setupObj);
		__markXMLElements(_doc, _imageXMLElementArray, localize(_global.imagesLabel), IMAGE_COLOR_ARRAY);
		return true;
	}

	if(IS_IMAGE_PLACED) {
		__placeImages(_doc, _imageXMLElementArray, _unpackObj, _setupObj);
		return true;
	}

	return true;
} /* END function __handleImages */


/**
 * Insert image sources as plain text
 * e.g. {media/image1.jpg}
 * @param {Document} _doc 
 * @param {XMLElement} _imageXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __insertImageSources(_doc, _imageXMLElementArray, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_imageXMLElementArray || !(_imageXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const SOURCE_ATTRIBUTE_NAME = _setupObj["image"]["attributes"]["source"];

	var _counter = 0;
	
	for(var i=_imageXMLElementArray.length-1; i>=0; i-=1) {

		var _imageXMLElement = _imageXMLElementArray[i];
		if(!_imageXMLElement || !_imageXMLElement.isValid) {
			continue;
		}

		var _sourceAttribute = _imageXMLElement.xmlAttributes.itemByName(SOURCE_ATTRIBUTE_NAME);
		if(!_sourceAttribute.isValid) {
			continue;
		}

		try {
			_imageXMLElement.insertTextAsContent("{" + _sourceAttribute.value + "}", XMLElementPosition.ELEMENT_START);
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.insertImageSourcesMessage, _counter, localize(_global.imageSourcesLabel)));
	}

	return true;
} /* END function __insertImageSources */


/**
 * Place images in document
 * (A folder »Links« is created next to the InDesign file if document path is avaliable.)
 * @param {Document} _doc 
 * @param {XMLElement} _imageXMLElementArray 
 * @param {Object} _unpackObj 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __placeImages(_doc, _imageXMLElementArray, _unpackObj, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_imageXMLElementArray || !(_imageXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const LINK_FOLDER_PATH = _setupObj["linkFolder"]["path"];
	const SOURCE_ATTRIBUTE_NAME = _setupObj["image"]["attributes"]["source"];
	const OBJECT_STYLE_ATTRIBUTE_NAME = "ostyle";
	const OBJECT_STYLE_PROPERTIES = _setupObj["image"]["objectStyleProperties"];
	const LINK_FOLDER_NAME = _setupObj["linkFolder"]["name"] || "Links";
	const IMAGE_WIDTH_IN_MM = _setupObj["image"]["width"];
	const IMAGE_HEIGHT_IN_MM = _setupObj["image"]["height"];
	const ALT_TEXT_ATTRIBUTE_NAME = _setupObj["image"]["attributes"]["description"];
	const IS_ALT_TEXT_INSERTED = _setupObj["image"]["isAltTextInserted"];

	const IMAGE_WIDTH = UnitValue(IMAGE_WIDTH_IN_MM, MeasurementUnits.MILLIMETERS).as(_doc.viewPreferences.verticalMeasurementUnits);
	const IMAGE_HEIGHT = UnitValue(IMAGE_HEIGHT_IN_MM, MeasurementUnits.MILLIMETERS).as(_doc.viewPreferences.verticalMeasurementUnits);

	/* Links folder */
	var _linkFolder = Folder(LINK_FOLDER_PATH);
	if(!_linkFolder.exists) {
		var _docFilePath = app.activeDocument.properties.fullName;
		if(_docFilePath !== undefined) {
			_linkFolder = Folder(_docFilePath.parent.fullName + "/" + LINK_FOLDER_NAME);
			if(!_linkFolder.exists) {
				_linkFolder.create(); /* -> hard disk */
			}
		}
	}
	
	/* Word folder */
	var _wordFolder;
	var _unpackFolder = _unpackObj["folder"];
	if(_unpackFolder && _unpackFolder instanceof Folder && _unpackFolder.exists) {
		_wordFolder = Folder(_unpackFolder.fullName + '/word');
	}
	if(!_wordFolder || !_wordFolder.exists) {
		_global["log"].push(localize(_global.wordFolderValidationMessage));
		return false;
	}

	var _counter = 0;
	
	for(var i=_imageXMLElementArray.length-1; i>=0; i-=1) {

		var _imageXMLElement = _imageXMLElementArray[i];
		if(!_imageXMLElement || !_imageXMLElement.isValid) {
			continue;
		}

		/* Image source */
		var _sourceAttribute = _imageXMLElement.xmlAttributes.itemByName(SOURCE_ATTRIBUTE_NAME);
		if(!_sourceAttribute.isValid) {
			_global["log"].push(localize(_global.missingImageSourceMessage, SOURCE_ATTRIBUTE_NAME));
			continue;
		}
		var _imageSource = _sourceAttribute.value;
		if(!_imageSource) {
			_global["log"].push(localize(_global.missingImageSourceMessage, SOURCE_ATTRIBUTE_NAME));
			continue;
		}

		/* Image object style */
		var _oStyleName;
		var _oStyleXMLAttribute = _imageXMLElement.xmlAttributes.itemByName(OBJECT_STYLE_ATTRIBUTE_NAME);
		if(_oStyleXMLAttribute.isValid) {
			_oStyleName = _oStyleXMLAttribute.value;
		}

		/* Alternativ text */
		var _altText;
		var _altTextXMLAttribute = _imageXMLElement.xmlAttributes.itemByName(ALT_TEXT_ATTRIBUTE_NAME);
		if(_altTextXMLAttribute.isValid) {
			_altText = _altTextXMLAttribute.value;
		}

		/* Image file */
		var _sourceImageFile = File(_wordFolder.fullName + '/' + _imageSource); /* in Word document embedded images */
		if(!_sourceImageFile.exists) {
			_sourceImageFile = File(_imageSource); /* external (linked) images on hard disk */
		}
		if(!_sourceImageFile.exists) {
			_global["log"].push(localize(_global.imageFileValidationMessage, _imageSource));
			continue;
		}

		try {

			/* Create anchored frame */
			_imageXMLElement.placeIntoInlineFrame([IMAGE_WIDTH,IMAGE_HEIGHT]);

			/* Place linked */
			if(_linkFolder.exists) {
				var _embedImageFile = File(_linkFolder.fullName + "/" + _sourceImageFile.name);
				var _isCopied = _sourceImageFile.copy(_embedImageFile); /* -> hard disk */
				if(_isCopied === true) {
					_imageXMLElement.setContent(_embedImageFile);
				}
			}
			/* Place embeded (for unsaved documents) */
			else {
				_imageXMLElement.setContent(_sourceImageFile);
			}

			/* Apply object style */
			if(!!_oStyleName) {
				var _oStyle = _doc.objectStyles.itemByName(_oStyleName);
				if(!_oStyle.isValid) {
					_oStyle = _doc.objectStyles.add({ name:_oStyleName }); /* -> DOC */
					_oStyle.properties = OBJECT_STYLE_PROPERTIES;
				}
				_imageXMLElement.applyObjectStyle(_oStyle, true, true);
			}


			var _placedImage = _imageXMLElement.xmlContent;
			if(!_placedImage || !_placedImage.isValid) {
				continue;
			}
			/* Fit frame to content */
			if(_placedImage.hasOwnProperty("fit")) {
				_placedImage.fit(FitOptions.FRAME_TO_CONTENT);
			}

			/* Embed image (for unsaved documents) */
			if(!_linkFolder.exists && _placedImage.hasOwnProperty("itemLink")) {
				_placedImage.itemLink.unlink();
			}

			/* Insert alternativ text */
			if(IS_ALT_TEXT_INSERTED && !!_altText) {
				var _imageFrame = _placedImage.parent;
				if(_imageFrame.hasOwnProperty("objectExportOptions")) {
					_imageFrame.objectExportOptions.altTextSourceType = SourceType.SOURCE_CUSTOM;
					_imageFrame.objectExportOptions.customAltText = _altText;
				}
			}
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.placeImageMessage, _counter, localize(_global.imageLabel)));
	}

	return true;
} /* END function __placeImages */


/**
 * Handle Track Changes
 * @param {Document} _doc  
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleTrackChanges(_doc, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const INSERTED_TEXT_TAG_NAME = _setupObj["trackChanges"]["insertedText"]["tag"];
	const INSERTED_TEXT_COLOR_ARRAY = _setupObj["trackChanges"]["insertedText"]["color"];
	const DELETED_TEXT_TAG_NAME = _setupObj["trackChanges"]["deletedText"]["tag"];
	const DELETED_TEXT_COLOR_ARRAY = _setupObj["trackChanges"]["deletedText"]["color"];
	const MOVED_FROM_TEXT_TAG_NAME = _setupObj["trackChanges"]["movedFromText"]["tag"];
	const MOVED_FROM_TEXT_COLOR_ARRAY = _setupObj["trackChanges"]["movedFromText"]["color"];
	const MOVED_TO_TEXT_TAG_NAME = _setupObj["trackChanges"]["movedToText"]["tag"];
	const MOVED_TO_TEXT_COLOR_ARRAY = _setupObj["trackChanges"]["movedToText"]["color"];

	const IS_TRACK_CHANGE_CREATED = _setupObj["trackChanges"]["isCreated"];
	const IS_TRACK_CHANGE_MARKED = _setupObj["trackChanges"]["isMarked"];
	const IS_TRACK_CHANGE_REMOVED = _setupObj["trackChanges"]["isRemoved"];

	var _insertedTextXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + INSERTED_TEXT_TAG_NAME);
	var _deletedTextXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + DELETED_TEXT_TAG_NAME);
	var _movedFromTextXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + MOVED_FROM_TEXT_TAG_NAME);
	var _movedToTextXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + MOVED_TO_TEXT_TAG_NAME);
	
	if(
		_insertedTextXMLElementArray.length === 0 && 
		_deletedTextXMLElementArray.length === 0 && 
		_movedFromTextXMLElementArray.length === 0 && 
		_movedToTextXMLElementArray.length === 0
	) {
		return true;
	}

	if(IS_TRACK_CHANGE_REMOVED) {
		__untagXMLElements(_insertedTextXMLElementArray, localize(_global.insertedTextLabel));
		__removeXMLElements(_deletedTextXMLElementArray, localize(_global.deletedTextLabel));
		__removeXMLElements(_movedFromTextXMLElementArray, localize(_global.movedFromTextLabel));
		__untagXMLElements(_movedToTextXMLElementArray, localize(_global.movedToTextLabel));
		return true;
	}

	if(IS_TRACK_CHANGE_MARKED || !_doc.hasOwnProperty("endnoteOptions")) {
		__markXMLElements(_doc, _insertedTextXMLElementArray, localize(_global.insertedTextLabel), INSERTED_TEXT_COLOR_ARRAY, "USE_UNDERLINE");
		__markXMLElements(_doc, _deletedTextXMLElementArray, localize(_global.deletedTextLabel), DELETED_TEXT_COLOR_ARRAY, "USE_UNDERLINE");
		__markXMLElements(_doc, _movedFromTextXMLElementArray, localize(_global.movedFromTextLabel), MOVED_FROM_TEXT_COLOR_ARRAY, "USE_UNDERLINE");
		__markXMLElements(_doc, _movedToTextXMLElementArray, localize(_global.movedToTextLabel), MOVED_TO_TEXT_COLOR_ARRAY, "USE_UNDERLINE");
		var _deletedTextCondition = _doc.conditions.itemByName(localize(_global.deletedTextLabel));
		if(_deletedTextCondition.isValid) {
			_deletedTextCondition.visible = false;
		}
		return true;
	}

	if(IS_TRACK_CHANGE_CREATED) {
		/* ... */
		return true;
	} 

	return true;
} /* END function __handleEndnotes */


/**
 * Handle Footnotes
 * @param {Document} _doc  
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleFootnotes(_doc, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const FOOTNOTE_TAG_NAME = _setupObj["footnote"]["tag"];
	const COLOR_ARRAY = _setupObj["footnote"]["color"];
	const IS_FOOTNOTE_CREATED = _setupObj["footnote"]["isCreated"];
	const IS_FOOTNOTE_MARKED = _setupObj["footnote"]["isMarked"];
	const IS_FOOTNOTE_REMOVED = _setupObj["footnote"]["isRemoved"];

	var _footnoteXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + FOOTNOTE_TAG_NAME);
	if(_footnoteXMLElementArray.length === 0) {
		return true;
	}

	if(IS_FOOTNOTE_REMOVED) {
		__removeXMLElements(_footnoteXMLElementArray, localize(_global.footnotesLabel));
		return true;
	}

	if(IS_FOOTNOTE_MARKED) {
		__markXMLElements(_doc, _footnoteXMLElementArray, localize(_global.footnotesLabel), COLOR_ARRAY);
		return true;
	}

	if(IS_FOOTNOTE_CREATED) {
		__createFootnotes(_doc, _wordXMLElement, _footnoteXMLElementArray, _setupObj);
		return true;
	} 

	return true;
} /* END function __handleFootnotes */


/**
 * Create Footnotes
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Array} _footnoteXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createFootnotes(_doc, _wordXMLElement, _footnoteXMLElementArray, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_footnoteXMLElementArray || !(_footnoteXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	var _wordXMLStory = _wordXMLElement.parentStory;
	if(!_wordXMLStory || !_wordXMLStory.isValid) {
		_global["log"].push(localize(_global.xmlStoryValidationError));
		return false;
	}

	var _counter = 0;

	for(var i=_footnoteXMLElementArray.length-1; i>=0; i-=1) {

		var _footnoteXMLElement = _footnoteXMLElementArray[i];
		if(!_footnoteXMLElement || !_footnoteXMLElement.isValid) {
			continue;
		}

		/* Get style names of footnote paragraphs */
		var _pStyleNameArray = __getNoteParagraphStyles(_footnoteXMLElement, _setupObj);
		if(!_pStyleNameArray) {
			_pStyleNameArray = [];
		}

		var _targetIP = _footnoteXMLElement.storyOffset;
		var _footnote;

		try {

			/* Add footnote */
			_footnote = _wordXMLStory.footnotes.add(LocationOptions.BEFORE, _targetIP); /* -> DOC */
			if(!_footnote || !_footnote.isValid) {
				_global["log"].push(localize(_global.footnoteValidationErrorMessage));
				continue;
			}

			/* Untag foonote XML elemente (InDesign does not allow XML elements in footnotes) */
			_footnoteXMLElement.xmlElements.everyItem().untag();

			/* Add text to footnote */
			var _footnoteText = _footnoteXMLElement.texts[0];
			if(!_footnoteText || !_footnoteText.isValid) {
				continue;
			}
			_footnoteText.move(LocationOptions.AT_END, _footnote.texts[0]);

			/* Remove XML container element */
			_footnoteXMLElement.remove();

		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;

		/* Apply styles to footnote paragraphs */
		var _isAssignmentCorrect = __applyStylesToNoteParagraphs(_doc, _footnote, _pStyleNameArray);
		if(!_isAssignmentCorrect) {
			_global["log"].push(localize(_global.footnoteParagraphStyleErrorMessage, i+1));
		}
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.footnotesLabel)));
	}

	return true;
} /* END function __createFootnotes */


/**
 * Handle Endnotes
 * @param {Document} _doc  
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __handleEndnotes(_doc, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const ENDNOTE_TAG_NAME = _setupObj["endnote"]["tag"];
	const COLOR_ARRAY = _setupObj["endnote"]["color"];
	const IS_ENDNOTE_CREATED = _setupObj["endnote"]["isCreated"];
	const IS_ENDNOTE_MARKED = _setupObj["endnote"]["isMarked"];
	const IS_ENDNOTE_REMOVED = _setupObj["endnote"]["isRemoved"];

	var _endnoteXMLElementArray = _wordXMLElement.evaluateXPathExpression("//" + ENDNOTE_TAG_NAME);
	if(_endnoteXMLElementArray.length === 0) {
		return true;
	}

	if(IS_ENDNOTE_REMOVED) {
		__removeXMLElements(_endnoteXMLElementArray, localize(_global.endnotesLabel));
		return true;
	}

	if(IS_ENDNOTE_MARKED || !_doc.hasOwnProperty("endnoteOptions")) {
		__markXMLElements(_doc, _endnoteXMLElementArray, localize(_global.endnotesLabel), COLOR_ARRAY);
		return true;
	}

	if(IS_ENDNOTE_CREATED) {
		__createEndnotes(_doc, _endnoteXMLElementArray, _setupObj);
		return true;
	} 

	return true;
} /* END function __handleEndnotes */


/**
 * Create Endnotes
 * @param {Document} _doc  
 * @param {Array} _endnoteXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createEndnotes(_doc, _endnoteXMLElementArray, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_endnoteXMLElementArray || !(_endnoteXMLElementArray instanceof Array)) { 
		throw new Error("Array as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	var _counter = 0;

	for(var i=_endnoteXMLElementArray.length-1; i>=0; i-=1) {

		var _endnoteXMLElement = _endnoteXMLElementArray[i];
		if(!_endnoteXMLElement || !_endnoteXMLElement.isValid) {
			continue;
		}

		/* Get style names of endnote paragraphs */
		var _pStyleNameArray = __getNoteParagraphStyles(_endnoteXMLElement, _setupObj);
		if(!_pStyleNameArray) {
			_pStyleNameArray = [];
		}

		var _targetIP = _endnoteXMLElement.storyOffset;
		var _endnote;

		try {

			/* Add endtnote */
			_endnote = _targetIP.createEndnote();
			if(!_endnote || !_endnote.isValid) {
				_global["log"].push(localize(_global.endnoteValidationErrorMessage));
				continue;
			}
			
			/* Untag foonote XML elemente (InDesign does not allow XML elements in endnotes) */
			_endnoteXMLElement.xmlElements.everyItem().untag();

			/* Add text to endnote */
			var _endnoteText = _endnoteXMLElement.texts[0];
			if(!_endnoteText || !_endnoteText.isValid) {
				continue;
			}
			_endnoteText.move(LocationOptions.AT_END, _endnote.texts[0]);

			/* Remove XML container element */
			_endnoteXMLElement.remove();

		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			continue;
		}

		_counter += 1;

		/* Apply styles to endnote paragraphs */
		var _isAssignmentCorrect = __applyStylesToNoteParagraphs(_doc, _endnote, _pStyleNameArray);
		if(!_isAssignmentCorrect) {
			_global["log"].push(localize(_global.endnoteParagraphStyleErrorMessage, i+1));
		}
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.createXMLElementsMessage, _counter, localize(_global.endnotesLabel)));
	}

	return true;
} /* END function __createEndnotes */


/**
 * Get style names of footnote or endnote paragraphs 
 * (defined in attribute »pstyle«)
 * @param {XMLElement} _containerXMLElement 
 * @param {Object} _setupObj
 * @returns Array
 */
function __getNoteParagraphStyles(_containerXMLElement, _setupObj) {

	if(!_containerXMLElement || !(_containerXMLElement instanceof XMLElement) || !_containerXMLElement.isValid) { return false; }
	if(!_setupObj || !(_setupObj instanceof Object)) { return false; }

	const PARAGRAPH_TAG_NAME = _setupObj["paragraph"]["tag"];
	const STYLE_ATTRIBUTE_NAME = "pstyle";

	var _pStyleNameArray = [];

	var _paragraphXMLElementArray = _containerXMLElement.evaluateXPathExpression(PARAGRAPH_TAG_NAME);
	
	for(var i=0; i<_paragraphXMLElementArray.length; i+=1) {

		var _paragraphXMLElement = _paragraphXMLElementArray[i];
		if(!_paragraphXMLElement || !_paragraphXMLElement.isValid) {
			_pStyleNameArray.push("");
			continue;
		}

		if(i === _paragraphXMLElementArray.length - 1 && _paragraphXMLElement.contents === "") {
			continue;
		}

		var _pStyleAttribute = _paragraphXMLElement.xmlAttributes.itemByName(STYLE_ATTRIBUTE_NAME);
		if(!_pStyleAttribute.isValid) {
			_pStyleNameArray.push("");
			continue;
		}

		_pStyleNameArray.push(_pStyleAttribute.value);
	}

	return _pStyleNameArray;
} /* END function __getNoteParagraphStyles */


/**
 * Apply styles to footnote or endnote paragraphps
 * @param {Document} _doc 
 * @param {Footnote|Endnote} _note 
 * @param {Array} _pStyleNameArray 
 * @returns Boolean
 */
function __applyStylesToNoteParagraphs(_doc, _note, _pStyleNameArray) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return false; }
	if(!_note || !_note.hasOwnProperty("texts") || !_note.isValid) { return false; }
	if(!_pStyleNameArray || !(_pStyleNameArray instanceof Array)) { return false; }

	var _isAssignmentCorrect = true;

	var _noteTexts = _note.texts.everyItem().getElements();
	if(!_noteTexts || _noteTexts.length === 0) {
		return false;
	}

	var _noteParagraphArray = _note.texts[0].paragraphs.everyItem().getElements();

	for(var i=0; i<_noteParagraphArray.length; i+=1) {

		var _noteParagraph = _noteParagraphArray[i];
		if(!_noteParagraph || !_noteParagraph.isValid) {
			_isAssignmentCorrect = false;
			continue;
		}

		var _pStyleName = _pStyleNameArray[i];
		if(!_pStyleName) {
			_isAssignmentCorrect = false;
			continue;
		}

		var _pStyle = _doc.paragraphStyles.itemByName(_pStyleName);
		if(!_pStyle.isValid) {
			_pStyle = _doc.paragraphStyles.add({ name:_pStyleName }); /* -> DOC */
		}

		try {
			_noteParagraph.applyParagraphStyle(_pStyle, true);
		} catch(_error) {
			_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
			_isAssignmentCorrect = false;
			continue;
		}
	}
	
	if(!_isAssignmentCorrect || _pStyleNameArray.length !== _noteParagraphArray.length) {
		return false;
	}

	return true;
} /* END function __applyStylesToNoteParagraphs */





/**
 * Place imported XML structure 
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement
 * @param {Object} _setupObj
 * @returns Object
 */
function __placeXML(_doc, _wordXMLElement, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");  
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const IS_AUTOFLOWING = _setupObj["document"]["isAutoflowing"];

	_global["progressbar"].init(0, 2, "", localize(_global.placeProgressLabel));
	_global["progressbar"].step();

	var _targetPage = __getTargetPage(_doc, IS_AUTOFLOWING);
	if(!_targetPage) {
		_global["log"].push(localize(_global.noTargetPageErrorMessage));
		return null;
	}

	var _userZeroPoint = _doc.zeroPoint;
	_doc.zeroPoint = [0,0];

	var _userRulerOrigin = _doc.viewPreferences.rulerOrigin;
	_doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

	var _placePointTop = _targetPage.marginPreferences.top;
	var _placePointLeft = _targetPage.marginPreferences.left;

	var _wordTextFrame;

	try {
		_wordTextFrame = _targetPage.placeXML(_wordXMLElement, [_placePointTop, _placePointLeft], IS_AUTOFLOWING);
	} catch(_error) {
		_global["log"].push(localize(_global.indesignErrorMessage, _error.message, _error.line));
		return null;
	} finally {
		_doc.viewPreferences.rulerOrigin = _userRulerOrigin;
		_doc.zeroPoint = _userZeroPoint;
	}
	
	if(!_wordTextFrame || !_wordTextFrame.isValid) {
		_global["log"].push(localize(_global.wordTextFrameValidationErrorMessage));
		return null;
	}

	var _wordStory = _wordTextFrame.parentStory;
	if(!_wordStory || !_wordStory.isValid) {
		_global["log"].push(localize(_global.wordStoryValidationErrorMessage));
		return null;
	}

	_global["progressbar"].step();

	return _wordStory;
} /* END function __placeXML */


/**
 * Get target page for placing XML structure
 * @param {Document} _doc 
 * @param {Boolean} _isAutoflowing (optional)
 * @returns Page
 */
function __getTargetPage(_doc, _isAutoflowing) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return null; }
	
	var _lastPage = _doc.pages.lastItem().getElements()[0];
	var _targetPage = _lastPage;

	/* Check: Endnotes on last page? */
	if(_isAutoflowing === false) {
		var _firstTextFrame = _targetPage.textFrames.firstItem();
		if(_firstTextFrame.isValid && _firstTextFrame.parentStory.isEndnoteStory) {
			_targetPage = _doc.pages[_targetPage.documentOffset - 1];
		}
	}
	
	/* Check: Page items on target page? */
	if(_targetPage.allPageItems.length !== 0) {
		_targetPage = _doc.pages.add(LocationOptions.AFTER, _targetPage); /* -> DOC */
	}

	if(!_targetPage || !_targetPage.isValid) {
		return null;
	}
	
	return _targetPage;
} /* END function __getTargetPage */





/**
 * Mount InDesign items after placing XML
 * @param {Document} _doc 
 * @param {Object} _unpackObj 
 * @param {XMLElement} _wordXMLElement 
 * @param {Story} _wordStory
 * @param {Object} _setupObj 
 * @returns Object
 */
function __mountAfterPlaced(_doc, _unpackObj, _wordXMLElement, _wordStory, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_wordStory || !(_wordStory instanceof Story) || !_wordStory.isValid) { 
		throw new Error("Story as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}


	return {};
} /* END function __mountAfterPlaced */