﻿/* DESCRIPTION: Import Microsoft Word Document (docx) */ 

/*
	
		+ Adobe InDesign Version: CC2021
		+ Author: Roland Dreger 
		+ Date: 24. January 2022
		
		+ Last modified: 24. January 2022
		
		
		+ Descriptions
			
			Alternative import for Microsoft Word documents
		

		+ Hints
		
			 
			
*/

var _global = {
	"projectName":"Import_Word",
	"version":"1.0",
	"debug":false,
	"log":[]
};

/* Document Settings */
_global["setups"] = {
	
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
	
	if(!_global || !_global.hasOwnProperty("setups")) { 
		throw new Error("Global object [_global] not defined or has no property [_setups]."); 
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
	var _docxFile = __getDocxFile();
	if(!_docxFile) {
		return false;
	}
	
	/* Unpack docx file to temp folder  */
	var _unpackResultObj = __unpackFile(_docxFile);
	if(!_unpackResultObj) {
		return false;
	}
	
	/* Import XML from unpacked docx file */
	var _rootXMLElement = __importXML(_doc, _unpackResultObj, _setupObj);
	if(!_rootXMLElement) {
		return false;
	}



	
	return true;
} /* END function __runSequence */



/**
 * Get docx file
 * @returns File
 */
function __getDocxFile() {
	
	if(!_global) { return false; }
	
	const _fileExtRegExp = new RegExp("\\.docx$","i");

	var _wordFile = File.openDialog(localize(_global.selectWordFile), null, false);
	if(!_wordFile || !_wordFile.exists) { 
		return null; 
	}

	var _wordFileName = _wordFile.name;
	if(!_fileExtRegExp.test(_wordFileName)) {
		_global["log"].push(localize(_global.fileExtensionValidationMessage));
		return null;
	}
	
	return _wordFile; 
} /* END function __getDocxFile */


/**
 * Unpack file
 * @param {File} _file
 * @returns Object
 */
function __unpackFile(_file) {
	
	if(!_global) { return false; }
	if(!_file || !(_file instanceof File) || !_file.exists) { 
		throw new Error("File as parameter required."); 
	}
	
	const _fileExtRegExp = new RegExp("\\..{3,4}$","i");

	var _destFolderPath = "";
	var _destFolder;
	
	try {
		_destFolderPath = Folder.temp.fullName + "/" + _file.name.replace(_fileExtRegExp,"");
		_destFolder = Folder(_destFolderPath);
		app.unpackageUCF(_file, _destFolder);
	} catch(_error) {
		_global["log"].push(_error.message);
		return null;
	}
	
	if(!_destFolder || !_destFolder.exists) {
		_global["log"].push(localize(_global.unpackDestinationFolderErrorMessage, _destFolderPath));
		return null;
	}

	var _xmlDocument = File(_destFolder.fullName + "/word/document.xml");
	if(!_xmlDocument.exists) {
		_global["log"].push(localize(_global.unpackFileErrorMessage, _file.fullName));
		return null;
	}

	return { 
		"folder":_destFolder,
		"document":_xmlDocument
	}; 
} /* END function __unpackFile */


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

	var _docxXMLElement;



	return _docxXMLElement;
} /* END function __importXML */














/* +++++++++++++++++++++++++++++ */
/* +++ Allgemeine Funktionen +++ */
/* +++++++++++++++++++++++++++++ */
/**
 * Fortschrittsanzeige
 * @returns {SUIWindow}
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


/* Anzeige der Log-Meldungen */
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


/* Deutsch-Englische Dialogtexte definieren */
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
		en: "Please select the word document ...",
		de: "Bitte das gew\u00FCnschte Word-Dokument ausw\u00E4hlen ..." 
	};
	
	_global.fileExtensionValidationMessage = { 
		en: "Import is available only for Word documents (docx).",
		de: "Import ist nur für Word-Dokumente (docx) möglich." 
	};
	
	_global.unpackDestinationFolderErrorMessage = { 
		en: "Destination folder for the unzipped file could not be created: %1",
		de: "Ziel-Ordner für die entpackte Datei konnte nicht erstellt werden: %1" 
	};
	
	_global.unpackFileErrorMessage = { 
		en: "File could not be extracted: %1",
		de: "Datei konnte nicht entpackt werden: %1" 
	};
	
} /* END function __defLocalizeStrings */