/* DESCRIPTION: Import Microsoft Word Document (docx) */ 

/*
	
		+ Adobe InDesign Version: CC2021+
		+ Author: Roland Dreger 
		+ Date: January 24, 2022
		
		+ Last modified: February 8, 2022
		
		
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

			Sonderzeichen entfernen aus Text

			ToDo

			Abschnittsumbruch
    
    
    
    # Images
    
    copy to Link folder or place it from there
    
    
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
	"place":{
		"isAutoflowing": false /* Value: Boolean; Description: If true, autoflows placed text. */
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
	"footnote":{ 
		"tag":"footnote", 
		"color":[155,255,255],
		"isCreated":true, 
		"isMarked":false, 
		"isRemoved":false 
	},
	"endnote":{ 
		"tag":"endnote", 
		"color":[255,155,255],
		"isCreated":true, 
		"isMarked":false, 
		"isRemoved":false
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


	/* Index */
	// _doc.indexes[0].topics[0].pageReferences.add(_xmlElement.texts[0], PageReferenceType.CURRENT_PAGE)

	/* Hyperlinks */

	/* Footnotes */ /* Last in chain, XML elements are removed from footnotes */
	__mountFootnotes(_doc, _wordXMLElement, _setupObj);
	
	/* Endnotes */
	__mountEndnotes(_doc, _wordXMLElement, _setupObj);
	








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
 * Mount Footnotes
 * @param {Document} _doc  
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __mountFootnotes(_doc, _wordXMLElement, _setupObj) {
	
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
} /* END function __mountFootnotes */


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
 * Mount Endnotes
 * @param {Document} _doc  
 * @param {XMLElement} _wordXMLElement 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __mountEndnotes(_doc, _wordXMLElement, _setupObj) {
	
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
		__createEndnotes(_doc, _wordXMLElement, _endnoteXMLElementArray, _setupObj);
		return true;
	} 

	return true;
} /* END function __mountEndnotes */


/**
 * Create Endnotes
 * @param {Document} _doc 
 * @param {XMLElement} _wordXMLElement 
 * @param {Array} _endnoteXMLElementArray 
 * @param {Object} _setupObj 
 * @returns Boolean
 */
function __createEndnotes(_doc, _wordXMLElement, _endnoteXMLElementArray, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");
	}
	if(!_wordXMLElement || !(_wordXMLElement instanceof XMLElement) || !_wordXMLElement.isValid) { 
		throw new Error("XMLElement as parameter required."); 
	}
	if(!_endnoteXMLElementArray || !(_endnoteXMLElementArray instanceof Array)) { 
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
 * Get style names of paragraphs (defined in attribute »pstyle«)
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
	if(!_note || !_note.hasOwnProperty("paragraphs") || !_note.isValid) { return false; }
	if(!_pStyleNameArray || !(_pStyleNameArray instanceof Array)) { return false; }

	var _isAssignmentCorrect = true;

	var _noteParagraphArray = _note.paragraphs.everyItem().getElements();

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

	var _page = _doc.pages.lastItem();
	if(_page.allPageItems.length !== 0) {
		_page = _doc.pages.add(LocationOptions.AFTER, _page);
	}

	var _userZeroPoint = _doc.zeroPoint;
	_doc.zeroPoint = [0,0];

	var _placePointTop = _page.marginPreferences.top;
	var _placePointLeft = _page.marginPreferences.left;

	var _wordTextFrame;

	try {
		_wordTextFrame = _page.placeXML(_wordXMLElement, [_placePointTop, _placePointLeft], IS_AUTOFLOWING);
	} catch(_error) {
		_global["log"].push(_error.message);
		return null;
	} finally {
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



} /* END function __defLocalizeStrings */