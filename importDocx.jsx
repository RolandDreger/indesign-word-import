/* DESCRIPTION: Import Microsoft Word Document (docx) */ 

/*
	
		+ Adobe InDesign Version: CC2021+
		+ Author: Roland Dreger 
		+ Date: January 24, 2022
		
		+ Last modified: April 3, 2022
		
		
		+ Descriptions
			
			Alternative import for Microsoft Word documents
		

		+ Hints
		
		  Temp folder e.g. /private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/import

			
	// var _unpackObj = {
	// 	"folder": Folder("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/InDesign_Word_Import/package_20220030_161205732"),
	// 	"word":{
	// 		"document":File("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/InDesign_Word_Import/package_20220030_161205732" + "/word/document.xml")
	// 	}
	// };


			ToDo:
			Radio-Buttons for Footnotes, Index, ... 
			1) import content
			1a) mark with conditional Text
			2) create InDesign objects

			Sonderzeichen entfernen aus Text???

			ToDo

			– Abschnittsumbruch
			– delete files and foldes off zip package in temp folder???
			– Remove if __runMainSequence
    
    # Images
    
    A folder »Links« is created next to the InDesign file if document path is avaliable (saved document). 
		Otherwise the image will be embedded in the document.
    
		Option: Mark Images
		Image source is inserted as plain text and highlighted with condition.
    
    # Track Changes
    
    Inserted Text

    app.selection[0].parentStory.trackChanges = false
    app.selection[0].contents = ""
    app.selection[0].parentStory.trackChanges = true
    app.selection[0].contents = "verspielt"
    
    Delete Text
    
    app.selection[0].parentStory.trackChanges = true
    app.selection[0].contents = ""
    app.selection[0].parentStory.trackChanges = false
    
    
    # Symbol mit Unicode
    
    # Listen für Listenabsätze beim Import erstellen
      (Wenn gleiches Absatzformat aber unterschiedliche Liste, 
      dann neues Absatzformat basierend original mit neuer Liste)
      
    
    # Zitate 
        
      mit Querverweise auf Textanker mit Name z.B. Newton, 1743
			
*/


//@include "utils/classes.jsx"
//@include "utils/dialogs.jsx"
//@include "utils/helpers.jsx"

//@include "hooks/beforeImport.jsx"
//@include "hooks/beforeMount.jsx"
//@include "hooks/beforePlaced.jsx"
//@include "hooks/afterPlaced.jsx"


var _global = {
	"projectName":"Import_Docx",
	"version":"1.0",
	"mode":"release", /* Values: "debug", "release" */
	"isLogged":false,
	"log":[]
};

/* Document Settings */
_global["setups"] = {
	"user":$.getenv("USER"),
	"xslt":{
		"name":"docx2Indesign.xsl"
	},
	"import":{},
	"dialog":{
		"isShown":false
	},
	"place":{
		"isAutoflowing": true /* Description: If true, autoflows placed text. (Depends on document settings.); Value: Boolean; */
	},
	"linkFolder":{
		"name":"Links", /* Description: Folder name for placed images; Value: String; */
		"path":"" /* Description: Folder path for placed images (optional); Value: String; */
	},
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
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) {
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
// Remove if ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
if(_doc.xmlElements[0].xmlElements.length === 0) {
	var _docxFile = __getDocxFile();
	if(!_docxFile) {
		return false;
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
} else {
	var _wordXMLElement = _doc.xmlElements[0].xmlElements.lastItem();
	var _unpackObj = {
		"folder": Folder("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/InDesign_Word_Import/package_20220030_161205732"),
		"word":{
			"document":File("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/InDesign_Word_Import/package_20220030_161205732" + "/word/document.xml")
		}
	};
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

	return true;
} /* END function __runMainSequence */




/**
 * Get docx file
 * @returns File
 */
function __getDocxFile() {
	
	const _wordExtRegExp = new RegExp("(\\.docx$|\\.xml$)","i");

	var _wordFile = File.openDialog(localize(_global.selectWordFile), null, false);
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
		_packageFolderPath = _tempFolder.fullName + "/package_" + _timestamp;
		_packageFolder = Folder(_packageFolderPath);
	} catch(_error) {
		_global["log"].push(_error.message);
		return null;
	}

	if(!_tempFolder || !(_tempFolder instanceof Folder) || !_tempFolder.exists) {
		_global["log"].push(localize(_global.createFolderErrorMessage, _tempFolderPath));
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
		_global["log"].push(_error.message);
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

	_global["progressbar"].init(0, 1, "", localize(_global.importProgressLabel));

	var _transformParams = [];

	_transformParams.push(["app","indesign"]);

	var _xsltFileName = _setupObj["xslt"]["name"];
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

	if(File(_unpackFolderPath + "/" + "word/comments.xml").exists) {
		_transformParams.push(["comments-file-path", "comments.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "docProps/core.xml").exists) {
		_transformParams.push(["core-props-file-path", "../docProps/core.xml"]);
	}
	if(File(_unpackFolderPath + "/" + "word/_rels/document.xml.rels").exists) {
		_transformParams.push(["document-relationships-file-path", "../word/_rels/document.xml.rels"]);
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

		_doc.importXML(_wordXMLFile);

	} catch(_error) {
		_global["log"].push(localize(_global.xmlFileImportXMLErrorMessage) + " " + _error.message);
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
	} catch (_error) { 
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

	/* Breaks */
	__insertBreaks(_doc, _wordXMLElement, _setupObj);

	/* Comments */
	__handleComments(_doc, _wordXMLElement, _setupObj);

	/* Index */
	// _doc.indexes[0].topics[0].pageReferences.add(_xmlElement.texts[0], PageReferenceType.CURRENT_PAGE)


	/* Hyperlinks */


	/* Textboxes */
	__handleTextboxes(_doc, _wordXMLElement, _setupObj);

	/* Images */
	__handleImages(_doc, _wordXMLElement, _unpackObj, _setupObj);

	/* 
		Last in chain. 
		XML elements must be removed from footnotes and endnotes
	*/

	/* Footnotes */ 
	__handleFootnotes(_doc, _wordXMLElement, _setupObj);
	
	/* Endnotes */
	__handleEndnotes(_doc, _wordXMLElement, _setupObj);
	
	/* Track Changes */
	__handleTrackChanges(_doc, _wordXMLElement, _setupObj);






	return {};
} /* END function __mountBeforePlaced */



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
			_global["log"].push(_error.message);
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
	const IS_COMMENT_CREATED = _setupObj["comment"]["isCreated"];
	const IS_COMMENT_MARKED = _setupObj["comment"]["isMarked"];
	const IS_COMMENT_REMOVED = _setupObj["comment"]["isRemoved"];

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

		var _commmentXMLElement = _commentXMLElementArray[i];
		if(!_commmentXMLElement || !_commmentXMLElement.isValid) {
			continue;
		}

		var _commentText = __getCommentText(_commmentXMLElement, _setupObj);
		var _targetIP = _commmentXMLElement.storyOffset;

		try {

			/* Add comment */
			var _comment = _wordXMLStory.notes.add(LocationOptions.BEFORE, _targetIP);
			if(!_comment || !_comment.isValid) {
				_global["log"].push(localize(_global.commentValidationErrorMessage));
				continue;
			}

			/* Insert text */
			_comment.texts[0].contents = _commentText;

			/* Remove XML container element */
			_commmentXMLElement.remove();

		} catch(_error) {
			_global["log"].push(_error.message);
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
 * @param {Obejct} _setupObj 
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
		var _pStyleNameArray = __getStylesOfNoteParagraphs(_footnoteXMLElement, _setupObj);
		if(!_pStyleNameArray) {
			_pStyleNameArray = [];
		}

		var _targetIP = _footnoteXMLElement.storyOffset;
		var _footnote;

		try {

			/* Add footnote */
			_footnote = _wordXMLStory.footnotes.add(LocationOptions.BEFORE, _targetIP);
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
			_global["log"].push(_error.message);
			continue;
		}

		_counter += 1;

		/* Apply styles to footnote paragraphs */
		var _isAssignmentCorrect = __applyStylesToNoteParagraphs(_doc, _footnote, _pStyleNameArray);
		if(!_isAssignmentCorrect) {
			_global["log"].push(localize(_global.footnoteParagraphStyleErrorMessage, _counter));
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
		var _pStyleNameArray = __getStylesOfNoteParagraphs(_endnoteXMLElement, _setupObj);
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
			_global["log"].push(_error.message);
			continue;
		}

		_counter += 1;

		/* Apply styles to endnote paragraphs */
		var _isAssignmentCorrect = __applyStylesToNoteParagraphs(_doc, _endnote, _pStyleNameArray);
		if(!_isAssignmentCorrect) {
			_global["log"].push(localize(_global.endnoteParagraphStyleErrorMessage, _counter));
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
function __getStylesOfNoteParagraphs(_containerXMLElement, _setupObj) {

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

		var _pStyleAttribute = _paragraphXMLElement.xmlAttributes.itemByName(STYLE_ATTRIBUTE_NAME);
		if(!_pStyleAttribute.isValid) {
			_pStyleNameArray.push("");
			continue;
		}

		_pStyleNameArray.push(_pStyleAttribute.value);
	}

	return _pStyleNameArray;
} /* END function __getStylesOfNoteParagraphs */


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
			_pStyle = _doc.paragraphStyles.add({ 
				name:_pStyleName 
			});
		}

		try {
			_noteParagraph.applyParagraphStyle(_pStyle, true);
		} catch(_error) {
			_global["log"].push(_error.message);
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
					_oStyle = _doc.objectStyles.add({ name:_oStyleName });
					_oStyle.properties = OBJECT_STYLE_PROPERTIES;
				}
				_textbox.applyObjectStyle(_oStyle, true, true);
			}

			/* Apply paragraph styles */
			__applyStylesToTextboxParagraphs(_doc, _textboxXMLElement, _setupObj);

		} catch(_error) {
			_global["log"].push(_error.message);
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
			_pStyle = _doc.paragraphStyles.add({ 
				name:_pStyleName 
			});
		}

		try {
			_paragraphXMLElement.applyParagraphStyle(_pStyle, true);
		} catch(_error) {
			_global["log"].push(_error.message);
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
			_global["log"].push(_error.message);
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
	if(!!_wordFolder && !_wordFolder.exists) {
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
					_oStyle = _doc.objectStyles.add({ name:_oStyleName });
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
			_global["log"].push(_error.message);
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
		__removeXMLElements(_insertedTextXMLElementArray, localize(_global.insertedTextLabel));
		__removeXMLElements(_deletedTextXMLElementArray, localize(_global.deletedTextLabel));
		__removeXMLElements(_movedFromTextXMLElementArray, localize(_global.movedFromTextLabel));
		__removeXMLElements(_movedToTextXMLElementArray, localize(_global.movedToTextLabel));
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

	const IS_AUTOFLOWING = _setupObj["place"]["isAutoflowing"];

	_global["progressbar"].init(0, 1, "", localize(_global.placeProgressLabel));

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
		_global["log"].push(_error.message);
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
		_targetPage = _doc.pages.add(LocationOptions.AFTER, _targetPage);
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








/**
 * Define localize strings
 */
function __defLocalizeStrings() {
	
	_global.noDocOpenAlert = { 
		en:"A document must be open to execute the script!",
		de:"F\u00FCr die Ausf\u00FChrung des Skriptes ist ein ge\u00F6ffnetes Dokument erforderlich!" 
	};
	
	_global.goBackLabel = { 
		en:"Import Word Document",
		de:"Word-Dokument importieren" 
	};
	
	_global.processingErrorAlert = { 
		en:"Skript Error",
		de:"Skriptfehler" 
	};

	_global.errorMessageLabel = { 
		en:"Error message:",
		de:"Fehlermeldung:" 
	};

	_global.lineLabel = { 
		en:"Line:",
		de:"Zeile:" 
	};
	
	_global.fileNameLabel = { 
		en:"File:",
		de:"Datei:" 
	};

	_global.createProgessbarErrorMessage = { 
		en:"Progress bar could not be created.",
		de:"Fortschrittsbalken konnte nicht erstellt werden." 
	};

	_global.logDialogTitle = { 
		en: "Messages",
		de: "Meldungen" 
	};
	
	_global.okButtonLabel = { 
		en: "OK",
		de: "OK" 
	};
	
	_global.importProgressLabel = { 
		en: "Import Word Document ...",
		de: "Word-Document importieren ..." 
	};

	_global.mountProgressLabel = { 
		en: "Create items ...",
		de: "Objekte erstellen ..." 
	};

	_global.placeProgressLabel = { 
		en: "Place content ...",
		de: "Inhalt platzieren ..." 
	};

	_global.selectWordFile = { 
		en: "Please select Word (.docx) or Word XML Document (.xml) ...",
		de: "Bitte Word-Dokument (.docx) oder Word-XML-Dokument (.xml) ausw\u00E4hlen ..." 
	};
	
	_global.fileExtensionValidationMessage = { 
		en: "Import is available only for Word (.docx) or Word XML Document (.xml).",
		de: "Import ist nur für Word-Dokumente (.docx) oder Word-XML-Dokument (.xml) möglich." 
	};
	
	_global.createFolderErrorMessage = { 
		en: "Order could not be created: %1",
		de: "Order konnte nicht erstellt werden: %1" 
	};

	_global.unpackageFolderErrorMessage = { 
		en: "Destination folder for the unzipped file could not be created: %1",
		de: "Ziel-Ordner für die entpackte Datei konnte nicht erstellt werden: %1" 
	};
	
	_global.unpackageDocumentFileErrorMessage = { 
		en: "File could not be extracted: %1",
		de: "Datei konnte nicht entpackt werden: %1" 
	};
	
	_global.scriptFolderErrorMessage = { 
		en: "Script folder could not be determined.",
		de: "Skriptordner konnte nicht ermittelt werden." 
	};

	_global.selectXSLFile = { 
		en:"Please select the XSL transformation file [%1] ...", 
		de:"Bitte die XSL-Transformationsdatei [%1] ausw\u00E4hlen ..."
	};

	_global.noXSLFileErrorMessage = { 
		en:"The XSL transformation file (.xsl) could not be found. The import will be canceled.",
		de:"Die XSL-Transformationsdatei (.xsl) konnte nicht gefunden werden. Der Import wird abgebrochen." 
	};

	_global.xmlDataImportErrorMessage = { 
		en:"No XML data imported",
		de:"Keine XML-Daten importiert" 
	};

	_global.xmlFileImportXMLErrorMessage = { 
		en:"Unable to import selected XML file.", 
		de:"Die ausgew\u00E4hlte XML-Datei konnte nicht importiert werden." 
	};

	_global.wordDocumentFileErrorMessage = { 
		en: "File for import could not be found: %1",
		de: "Datei für Import konnte nicht gefunden werden: %1" 
	};

	_global.noTargetPageErrorMessage= { 
		en: "Target page could not be determined.",
		de: "Zielseite konnte nicht ermittelt werden." 
	};

	_global.wordTextFrameValidationErrorMessage = { 
		en: "Textframe with placed content not valid.",
		de: "Textrahmen mit platziertem Inhalt nicht valide." 
	};

	_global.wordStoryValidationErrorMessage = { 
		en: "Story with placed content not valid.",
		de: "Textabschnitt mit platziertem Inhalt nicht valide." 
	};

	_global.xmlStoryValidationError = { 
		en: "Story of XML element not valid.",
		de: "Textabschnitt des XML-Elements nicht valide." 
	};

	_global.removeXMLElementsMessage = { 
		en: "%1 %2 removed.",
		de: "%1 %2 gelöscht." 
	};

	_global.markXMLElementsMessage = { 
		en: "%1 %2 marked.",
		de: "%1 %2 markiert." 
	};

	_global.createXMLElementsMessage = { 
		en: "%1 %2 created.",
		de: "%1 %2 erstellt." 
	};

	_global.insertImageSourcesMessage = { 
		en: "%1 %2 inserted as plain text.",
		de: "%1 %2 als Text eingefügt." 
	};

	_global.imageSourcesLabel = { 
		en: "images sources",
		de: "Bildquellen" 
	};

	_global.placeImageMessage = { 
		en: "%1 %2 placed.",
		de: "%1 %2 plaziert." 
	};

	_global.imageLabel = { 
		en: "image",
		de: "Bild" 
	};

	_global.footnotesLabel = { 
		en: "footnotes",
		de: "Fußnoten" 
	};

	_global.footnoteValidationErrorMessage = { 
		en: "Footnote not valid.",
		de: "Fußnote nicht valide." 
	};

	_global.footnoteParagraphStyleErrorMessage = {
		en: "Footnote [%1]: Error applying paragraph styles.",
		de: "Fußnote [%1]: Fehler beim Zuweisen der Absatzformate."
	};

	_global.endnotesLabel = { 
		en: "endnotes",
		de: "Endnoten" 
	};

	_global.endnoteValidationErrorMessage = { 
		en: "Endnote not valid.",
		de: "Endnote nicht valide." 
	};

	_global.endnoteParagraphStyleErrorMessage = {
		en: "Endtnote [%1]: Error applying paragraph styles.",
		de: "Endnote [%1]: Fehler beim Zuweisen der Absatzformate."
	};

	_global.specialCharacterNotAvailableErrorMessage = { 
		en: "Special character not available: 1%",
		de: "Sonderzeichen nicht verfügbar: %1" 
	}; 
	
	_global.xmlElementNotEmptyErrorMessage = { 
		en: "XML element [%1] not empty: %2",
		de: "XML-Element [%1] nicht leer: %2" 
	};
		
	_global.insertSpecialCharactersMessage = { 
		en: "%1 special characters [%2] inserted.",
		de: "%1 Sonderzeichen [%2] eingefügt." 
	};

	_global.commentsLabel = { 
		en: "Comments",
		de: "Kommentare" 
	};

	_global.commentValidationErrorMessage = { 
		en: "Comment not valid.",
		de: "Kommentar nicht valide." 
	};

	_global.textboxesLabel = { 
		en: "Textboxes",
		de: "Textboxen" 
	};

	_global.commentValidationErrorMessage = { 
		en: "Textbox not valid.",
		de: "Textbox nicht valide." 
	};

	_global.imagesLabel = { 
		en: "Images",
		de: "Bilder" 
	};

	_global.insertedTextLabel = { 
		en: "Inserted Text",
		de: "Eingefügter Text" 
	};

	_global.deletedTextLabel = { 
		en: "Deleted Text",
		de: "Gelöschter Text" 
	};

	_global.movedFromTextLabel = { 
		en: "Deleted Text",
		de: "Gelöschter Text" 
	};

	_global.movedToTextLabel = { 
		en: "Moved Text",
		de: "Verschobener Text" 
	};

	_global.wordFolderValidationMessage = { 
		en: "Folder with unzipped Word files could not be found.",
		de: "Folder mit entpackten Word-Dateien konnte nicht gefunden werden." 
	};

	_global.imageFileValidationMessage = { 
		en: "Media file could not be found: [%1]",
		de: "Medien-Datei konnte nicht gefunden werden: [%1]" 
	};

	_global.missingImageSourceMessage = { 
		en: "Media element without source. Attribute: [%1]",
		de: "Medien-Element ohne Quelle. Attribute: [%1]" 
	};

} /* END function __defLocalizeStrings */