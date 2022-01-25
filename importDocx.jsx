/* DESCRIPTION: Import Microsoft Word Document (docx) */ 

/*
	
		+ Adobe InDesign Version: CC2021
		+ Author: Roland Dreger 
		+ Date: 24. January 2022
		
		+ Last modified: 24. January 2022
		
		
		+ Descriptions
			
			Alternative import for Microsoft Word documents
		

		+ Hints
		
		  Temp folder e.g. /private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/import
			
*/

var _global = {
	"projectName":"Import_Word",
	"version":"1.0",
	"debug":false,
	"log":[]
};

/* Document Settings */
_global["setups"] = {
	"xslt":{
		"name":"docx2Indesign.xsl"
	}
};

/* Check: Developer or User? */
var _user = $.getenv("USER");
if(_user === "rolanddreger") {
	_global["debug"] = true;
}




__start();


function __start() {
	
	if(!_global) { 
		throw new Error("Global object [_global] not defined.");
	}
	
	/* Deutsch-Englische Dialogtexte definieren */
	__defLocalizeStrings();
	
	/* Progressbar definieren */
	_global["progressbar"] = __createProgressbar();
	if(!_global["progressbar"]) {
		throw new Error(localize(_global.createProgessbarErrorMessage));
	}
	
	/* Active document */
	var _doc = app.documents.firstItem();
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) {
		alert(localize(_global.noDocOpenAlert));
		return false; 
	}

	/* Script Presets */
	var _userEnableRedraw = app.scriptPreferences.enableRedraw;
	app.scriptPreferences.enableRedraw = false;
	var _userInteractionLevel = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	
	try {
		if(app.scriptPreferences.version >= 6 && !_global["debug"]) {
			app.doScript(
				__runSequence, 
				ScriptLanguage.JAVASCRIPT, 
				[_doc], 
				UndoModes.ENTIRE_SCRIPT, 
				localize(_global.goBackLabel)
			);
		} else {
			__runSequence([_doc]);
		}
	} catch(_error) {
		if(_error instanceof Error) {
			alert(
				_error.name + " | " + _error.number + "\n" +
				localize(_global.errorMessageLabel) + " " + _error.message + ";\n" +
				localize(_global.lineLabel) + " " + _error.line + ";",
				"Error", true
			);
		} else {
			alert(localize(_global.processingErrorAlert) + "\n" + _error, "Error", true);
		}
	} finally {
		app.scriptPreferences.enableRedraw = _userEnableRedraw;
		app.scriptPreferences.userInteractionLevel = _userInteractionLevel;
	}
	
	/* Check: Log messages? */
	if(_global["log"].length > 0) {
		__showLog(_global["log"]);
		return false;
	}
	  
	return true;
} /* END function __start */


if(
	_global !== null && _global !== undefined && 
	_global.hasOwnProperty("progressbar") && 
	_global["progressbar"].hasOwnProperty("close")
) {
	_global["progressbar"].close();
}

_global = null;




/* +++++++++++++++++++++ */
/* +++ Main Sequence +++ */
/* +++++++++++++++++++++ */
function __runSequence(_doScriptParameterArray) {
	
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
	
	/* Get docx file */
	// var _docxFile = __getDocxFile();
	// if(!_docxFile) {
	// 	return false;
	// }
	
	/* Get package data */
	// var _unpackResultObj = __getPackageData(_docxFile);
	// if(!_unpackResultObj) {
	// 	return false;
	// }
	
	var _unpackResultObj = {
		"folder": Folder("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/import"),
		"word":{
			"document":File("/private/var/folders/s5/st5j74qj0wj2vmhjtwwh4_hr0000gn/T/TemporaryItems/import" + "/word/document.xml")
		}
	}


	/* Import XML from unpacked docx file */
	var _docxXMLElement = __importXML(_doc, _unpackResultObj, _setupObj);
	if(!_docxXMLElement) {
		return false;
	}



	/* ... */

	return true;
} /* END function __runSequence */



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
	
	const _xmlExtRegExp = new RegExp("\\.xml$","i");
	const _docxExtRegExp = new RegExp("\\.docx$","i");
	const _fileExtRegExp = new RegExp("\\..+$","");

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

	var _destFolderPath = "";
	var _destFolder;
	
	/* Unpack Word Document */
	try {
		_destFolderPath = Folder.temp.fullName + "/" + _packageFileName.replace(_fileExtRegExp, "");
		_destFolder = Folder(_destFolderPath);
		app.unpackageUCF(_packageFile, _destFolder);
	} catch(_error) {
		_global["log"].push(_error.message);
		return null;
	}
	
	if(!_destFolder || !_destFolder.exists) {
		_global["log"].push(localize(_global.unpackageFolderErrorMessage, _destFolderPath));
		return null;
	}

	var _xmlDocFile = File(_destFolder.fullName + "/word/document.xml");
	if(!_xmlDocFile.exists) {
		_global["log"].push(localize(_global.unpackageDocumentFileErrorMessage, _packageFilePath));
		return null;
	}

	return { 
		"folder":_destFolder,
		"word": {
			"document":_xmlDocFile
		}
	}; 
} /* END function __getPackageData */


/**
 * Import Word document xml file
 * @param {Document} _doc InDesign document
 * @param {Objekt} _unpackResultObj Result of unpacking Word document file
 * @param {Objekt} _setupObj 
 * @returns XMLElement
 */
function __importXML(_doc, _unpackResultObj, _setupObj) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");  
	}
	if(!_unpackResultObj || !(_unpackResultObj instanceof Object)) { 
		throw new Error("Object as parameter required."); 
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	var _xsltFileName = _setupObj["xslt"]["name"];
	var _xsltFile = __getXSLTFile(_xsltFileName);
	if(!_xsltFile) { 
		return null; 
	}

	var _unpackFolderPath = "";
	var _unpackFolder = _unpackResultObj["folder"];
	if(_unpackFolder && _unpackFolder instanceof Folder && _unpackFolder.exists) {
		_unpackFolderPath = _unpackFolder.fullName;
	}

	var _wordXMLFile = _unpackResultObj["word"]["document"];
	if(!_wordXMLFile || !_wordXMLFile.exists) {
		_global["log"].push(localize(_global.wordDocumentFileErrorMessage, _wordXMLFile));
		return null;
	}
	
	var _rootXMLElement = _doc.xmlElements.firstItem();
	var _lastXMLElement = _rootXMLElement.xmlElements.lastItem();
	if(_lastXMLElement.isValid) {
		_lastXMLElement = _lastXMLElement.getElements()[0];
	} else {
		_lastXMLElement = null;
	}

	var _userXMLImportPreferences = _doc.xmlImportPreferences.properties;

	try {

		_doc.xmlImportPreferences.properties = {
			importStyle:XMLImportStyles.APPEND_IMPORT,
			allowTransform:true,
			transformFilename:_xsltFile,
			transformParameters:[["base-uri", _unpackFolderPath]],
			repeatTextElements:false,
			ignoreWhitespace:false,
			createLinkToXML:false,
			ignoreUnmatchedIncoming:false,
			importCALSTables:false,
			importTextIntoTables:false,
			importToSelected:false,
			removeUnmatchedExisting:false
		};

		_rootXMLElement.importXML(_wordXMLFile);

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












/* +++++++++++++++++++++++++ */
/* +++ General functions +++ */
/* +++++++++++++++++++++++++ */
/**
 * Progress bar
 * @returns SUIWindow
 */
function __createProgressbar() {
	
	var _progressbar;
	var _labelText;
	var _progressWindow = new Window ("palette", undefined, undefined, { borderless:true });
	with(_progressWindow) {	
		spacing = 10;
		margins = [20,10,20,20];
		alignChildren = ["fill","center"];
		_labelText = add("statictext");
		with(_labelText) {
			characters = 30; /* Breitenvorgabe des Fensters */
			justify = "center";
		} /* END _labelText */
		_progressbar = add("progressbar", undefined, 0, 0);
		with(_progressbar) {
			minimumSize.width = 340;
			maximumSize.height = 6;
		}
	} /* END _progressWindow */

	_progressWindow.initialize = function(_title, _start, _stop, _visible) {
		_progressWindow.text = (_title && _title.toString()) || "";
		_progressbar.value = (_start && !isNaN(_start) && Number(_start)) || 0;
		_progressbar.maxvalue = (_stop && !isNaN(_stop) && Number(_stop)) || 0;
		_progressbar.visible = !!_visible;
		this.show();
	}; /* END function initialize */

	_progressWindow.push = function(_label, _step) {
		_labelText.text = (_label && _label.toString()) || "";
		_progressbar.value = (_step && !isNaN(_step) && Number(_step)) || _progressbar.value + 1;
		this.update();
	}; /* END function push */
	
	return _progressWindow;
} /* END function __createProgressbar */


/**
 * Show log messages
 * @param {Array} _logMessageArray 
 * @returns Boolean
 */
function __showLog(_logMessageArray) {
	
	if(!_global) { return false; }
	if(!_logMessageArray || !(_logMessageArray instanceof Array)) { return false; }
	
	if(_logMessageArray.length === 0) { 
		return true; 
	}

	var _logMessageEdittext;
	var _okButton;

	var _logDialog = new Window("dialog", localize(_global.logDialogTitle), undefined, { closeButton: true });
	with(_logDialog) {
		alignChildren = ["fill","fill"];
		spacing = 15;
		var _logMessageGroup = add("group");
		with(_logMessageGroup) {
			alignChildren = ["fill","fill"];
			margins = [0,0,0,0];
			_logMessageEdittext = add("edittext", undefined, "", { multiline:true });
			with(_logMessageEdittext) {
				minimumSize = [500,60];
				maximumSize = [500,400];
			} /* END _logMessageEdittext */
		} /* END _logMessageGroup */
		/* Action Buttons */
		var _buttonGroup = add("group");
		with(_buttonGroup) {
			alignChildren = ["fill","fill"];
			margins = [0,0,20,0];
			spacing = 8;
			_okButton = add("button", undefined, localize(_global.okButtonLabel), { name:"OK" });
			with(_okButton) {
				alignment = ["right","top"];
			} /* END _okButton */
		} /* END _buttonGroup */
	} /* END _logDialog */
	
	/* Callbacks */
	_okButton.onClick = function() {
		_logDialog.hide();
		_logDialog.close(1);
	};
	/* END Callbacks */
	
	/* Dialog initialisieren */
	var _logMessages = _logMessageArray.join("\r\r");
	_logMessageEdittext.text = _logMessages;
	/* END Dialog initialisieren */

	/* Dialog aufrufen */
	_logDialog.show();
	
	return true;
} /* END function __showLog */


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
		en:"Error processing the document!",
		de:"Fehler bei der Verarbeitung des Dokuments!" 
	};

	_global.errorMessageLabel = { 
		en:"Error message:",
		de:"Fehlermeldung:" 
	};

	_global.lineLabel = { 
		en:"Line:",
		de:"Zeile:" 
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
	
	_global.selectWordFile = { 
		en: "Please select Word (.docx) or Word XML Document (.xml) ...",
		de: "Bitte Word-Dokument (.docx) oder Word-XML-Dokument (.xml) ausw\u00E4hlen ..." 
	};
	
	_global.fileExtensionValidationMessage = { 
		en: "Import is available only for Word (.docx) or Word XML Document (.xml).",
		de: "Import ist nur für Word-Dokumente (.docx) oder Word-XML-Dokument (.xml) möglich." 
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
	}

	_global.wordDocumentFileErrorMessage = { 
		en: "File for import could not be found: %1",
		de: "Datei für Import konnte nicht gefunden werden: %1" 
	};
} /* END function __defLocalizeStrings */