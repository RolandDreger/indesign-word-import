/**
 * Progress Bar
 */
 function ProgressBar() {
	
	this.window = new Window ("palette", undefined, undefined, { 
		borderless:true 
	});
	this.window.spacing = 10;
	this.window.margins = [20,10,20,20];
	this.window.alignChildren = ["fill","center"];

	this.labelText = this.window.add("statictext", undefined);
	this.labelText.justify = "center";
	this.labelText.characters = 30;

	this.progressbar = this.window.add("progressbar", undefined, 0, 0);
	this.progressbar.minimumSize.width = 340;
	this.progressbar.maximumSize.height = 6;
}

ProgressBar.prototype.init = function(_start, _stop, _title, _label) {
	if(_title && _title.constructor === String) {
		this.window.text = _title;
	}
	if(_label && _label.constructor === String) {
		this.labelText.text = _label;
	}
	if(!isNaN(_start)) {
		this.progressbar.value = Number(_start);
	} else {
		this.progressbar.value = 0;
	}
	if(!isNaN(_stop)) {
		this.progressbar.maxvalue = Number(_stop);
	} else {
		this.progressbar.maxvalue = 0;
	}
	this.window.show();
};

ProgressBar.prototype.setLabel = function(_label) {
	if(_label !== null && _label !== undefined && _label.constructor === String) {
		this.labelText.text = _label;
	}
	this.window.update();
};

ProgressBar.prototype.step = function(_value, _label) {
	if(_label && _label.constructor === String) {
		this.labelText.text = _label;
	}
	if(!isNaN(_value)) {
		this.progressbar.value = Number(_value);
	} else {
		this.progressbar.value += 1;
	}
	this.window.update();
};

ProgressBar.prototype.close = function() {
	this.window.hide();
	this.window.close();
};


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
 * Import Dialog
 * @param {Object} _setupObj
 * @returns Object
 */
function __showImportDialog(_setupObj) {
	
	if(!_global) { return null; }
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	const GAP = 5;
	const MARGIN = 5;

	var _defaultParagraphStyleEdittext;
	var _isAutoflowingCheckbox;
	var _isUntaggedCheckbox;
	var _extendedStyleModeRadiobutton; 
	var _minimizedStyleModeRadiobutton;
	var _pageBreakCheckbox;
	var _columnBreakCheckbox;
	var _forcedLineBreakCheckbox;
	var _sectionBreakCheckbox;
	var _commentCreatedRadiobutton;
	var _commentMarkedRadiobutton;
	var _commentRemovedRadiobutton;
	var _indexmarkCreatedRadiobutton;
	var _indexmarkRemovedRadiobutton;
	var _hyperlinkCreatedRadiobutton;
	var _hyperlinkMarkedRadiobutton;
	var _crossReferenceCreatedRadiobutton;
	var _crossReferenceMarkedRadiobutton;
	var _bookmarkCreatedCheckbox;
	var _bookmarkMarkerGroup;
	var _bookmarkMarkerEdittext;
	var _trackChangesMarkedRadiobutton;
	var _trackChangesRemovedRadiobutton;
	var _footnoteCreatedRadiobutton;
	var _footnoteMarkedRadiobutton;
	var _footnoteRemovedRadiobutton;
	var _endnoteCreatedRadiobutton;
	var _endnoteMarkedRadiobutton;
	var _endnoteRemovedRadiobutton;
	var _imagePlacedRadiobutton;
	var _imageMarkedRadiobutton;
	var _imageRemovedRadiobutton;
	var _imageInputGroup;
	var _imageWidthEdittext;
	var _imageHeightEdittext;
	var _textboxCreatedRadiobutton;
	var _textboxMarkedRadiobutton;
	var _textboxRemovedRadiobutton;
	var _textboxInputGroup;
	var _textboxWidthEdittext;
	var _textboxHeightEdittext;

	var _okButton;
	var _cancelButton;
	
	var _importDialog = new Window("dialog", localize(_global.importDialogTitle), undefined, { closeButton: true });
	with(_importDialog) {
		alignChildren = ["fill","fill"];
		margins = [MARGIN*5,MARGIN*4,MARGIN*5,MARGIN*4];
		spacing = GAP*4;
		var _optionsGroup = add("group");
		with(_optionsGroup) {
			alignChildren = ["fill","fill"];
			spacing = GAP*4;
			var _columnOne = add("group");
			with(_columnOne) {
				orientation = "column";
				alignChildren = ["fill","fill"];
				spacing = GAP*5;
				var _documentPanel = add("panel", undefined, localize(_global.documentLabel));
				with(_documentPanel) {
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP;
					_isAutoflowingCheckbox = add("checkbox", undefined, localize(_global.isAutoflowingLabel));
					_isUntaggedCheckbox = add("checkbox", undefined, localize(_global.isUntaggedLabel));
				} /* END _documentPanel */
				var _stylePanel = add("panel", undefined, localize(_global.stylePanelLabel));
				with(_stylePanel) {
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*2];
					alignChildren = ["fill","top"];
					spacing = GAP*2;
					var _defaultParagraphStyleGroup = add("group");
					with(_defaultParagraphStyleGroup) {
						orientation = "column";
						alignChildren = ["fill","top"];
						spacing = GAP;
						add("statictext", undefined, localize(_global.defaultParagraphStyleLabel));
						_defaultParagraphStyleEdittext = add("edittext");
						with(_defaultParagraphStyleEdittext) {
							characters = 10;
						} /* END _defaultParagraphStyleEdittext */
					} /* END _defaultParagraphStyleGroup */
					var _styleModeGroup = add("group");
					with(_styleModeGroup) {
						orientation = "column";
						alignChildren = "left";
						spacing = GAP;
						add("statictext", undefined, localize(_global.styleModeGroupLabel));
						var _styleModeRadioButtonGroup = add("group");
						with(_styleModeRadioButtonGroup) {
							orientation = "column";
							alignChildren = "left";
							spacing = GAP;
							margins.top = GAP;
							_minimizedStyleModeRadiobutton = add("radiobutton", undefined, localize(_global.minimizedStyleModeRadiobutton));
							_extendedStyleModeRadiobutton = add("radiobutton", undefined, localize(_global.extendedStyleModeRadiobutton));
						} /* END _styleModeRadioButtonGroup */
					} /* END _styleModeGroup */
				} /* END _stylePanel */
				var _breaksPanel = add("panel", undefined, localize(_global.breaksLabel));
				with(_breaksPanel) {
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP;
					_pageBreakCheckbox = add("checkbox", undefined, localize(_global.pageBreakLabel));
					_columnBreakCheckbox = add("checkbox", undefined, localize(_global.columnBreakLabel));
					_forcedLineBreakCheckbox = add("checkbox", undefined, localize(_global.forcedLineLabel));
					_sectionBreakCheckbox = add("checkbox", undefined, localize(_global.sectionLabel));
				} /* END _breaksPanel */
			} /* END _columnOne */
			var _columnTwo = add("group");
			with(_columnTwo) {
				orientation = "column";
				alignChildren = ["fill","fill"];
				spacing = GAP*5;
				var _imagePanel = add("panel", undefined, localize(_global.imagesPanelLabel));
				with(_imagePanel) {
					orientation = "column";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*2];
					alignChildren = ["fill","top"];
					spacing = GAP*2;
					var _imageButtonGroup = add("group");
					with(_imageButtonGroup) {
						_imagePlacedRadiobutton = add("radiobutton", undefined, localize(_global.placeLabel));
						_imageMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
						_imageRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
					} /* END _imageButtonGroup */
					_imageInputGroup = add("group");
					with(_imageInputGroup) {
						alignChildren = ["fill","fill"];
						var _imageWidthGroup = add("group");
						with(_imageWidthGroup) {
							orientation = "column";
							alignChildren = ["fill","fill"];
							spacing = GAP;
							add("statictext", undefined, localize(_global.imageWidthLabel));
							_imageWidthEdittext = add("edittext", undefined, "");
							with(_imageWidthEdittext) {
								characters = 10;
							} /* END _imageWidthEdittext */
						} /* END _imageWidthGroup */
						var _imageHeightGroup = add("group");
						with(_imageHeightGroup) {
							orientation = "column";
							alignChildren = ["fill","fill"];
							spacing = GAP;
							add("statictext", undefined, localize(_global.imageHeightLabel));
							_imageHeightEdittext = add("edittext", undefined, "");
							with(_imageHeightEdittext) {
								characters = 10;
							} /* END _imageHeightEdittext */
						} /* END _imageHeightGroup */
					} /* END _imageInputGroup */
				} /* END _imagePanel */
				var _textboxPanel = add("panel", undefined, localize(_global.textboxesPanelLabel));
				with(_textboxPanel) {
					orientation = "column";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*2];
					alignChildren = ["fill","top"];
					spacing = GAP*2;
					var _textboxButtonGroup = add("group");
					with(_textboxButtonGroup) {
						_textboxCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
						_textboxMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
						_textboxRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
					} /* END _textboxButtonGroup */
					_textboxInputGroup = add("group");
					with(_textboxInputGroup) {
						alignChildren = ["fill","fill"];
						var _textboxWidthGroup = add("group");
						with(_textboxWidthGroup) {
							orientation = "column";
							alignChildren = ["fill","fill"];
							spacing = GAP;
							add("statictext", undefined, localize(_global.textboxWidthLabel));
							_textboxWidthEdittext = add("edittext", undefined, "");
							with(_textboxWidthEdittext) {
								characters = 10;
							} /* END _textboxWidthEdittext */
						} /* END _textboxWidthGroup */
						var _textboxHeightGroup = add("group");
						with(_textboxHeightGroup) {
							orientation = "column";
							alignChildren = ["fill","fill"];
							spacing = GAP;
							add("statictext", undefined, localize(_global.textboxHeightLabel));
							_textboxHeightEdittext = add("edittext", undefined, "");
							with(_textboxHeightEdittext) {
								characters = 10;
							} /* END _textboxHeightEdittext */
						} /* END _textboxHeightGroup */
					} /* END _textboxInputGroup */
				} /* END _textboxPanel */
				var _footnotePanel = add("panel", undefined, localize(_global.footnotesLabel));
				with(_footnotePanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_footnoteCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_footnoteMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
					_footnoteRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
				} /* END _footnotePanel */
				var _endnotePanel = add("panel", undefined, localize(_global.endnotesLabel));
				with(_endnotePanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_endnoteCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_endnoteMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
					_endnoteRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
				} /* END _endnotePanel */
			} /* END _columnTwo */
			var _columnThree = add("group");
			with(_columnThree) {
				orientation = "column";
				alignChildren = ["fill","fill"];
				spacing = GAP*4;
				var _commentPanel = add("panel", undefined, localize(_global.commentsLabel));
				with(_commentPanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_commentCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_commentMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
					_commentRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
				} /* END _commentPanel */
				var _hyperlinkPanel = add("panel", undefined, localize(_global.hyperlinksLabel));
				with(_hyperlinkPanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_hyperlinkCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_hyperlinkMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
				} /* END _hyperlinkPanel */
				var _crossReferencePanel = add("panel", undefined, localize(_global.crossReferencesLabel));
				with(_crossReferencePanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_crossReferenceCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_crossReferenceMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
				} /* END _crossReferencePanel */
				var _indexmarkPanel = add("panel", undefined, localize(_global.indexmarksLabel));
				with(_indexmarkPanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_indexmarkCreatedRadiobutton = add("radiobutton", undefined, localize(_global.createLabel));
					_indexmarkRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
				} /* END _indexmarkPanel */
				var _trackChangesPanel = add("panel", undefined, localize(_global.trackChangesLabel));
				with(_trackChangesPanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*3,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*2;
					_trackChangesMarkedRadiobutton = add("radiobutton", undefined, localize(_global.markLabel));
					_trackChangesRemovedRadiobutton = add("radiobutton", undefined, localize(_global.removeLabel));
				} /* END _trackChangesPanel */
				var _bookmarkPanel = add("panel", undefined, localize(_global.bookmarksPanelLabel));
				with(_bookmarkPanel) {
					orientation = "row";
					margins = [MARGIN*2,MARGIN*2,MARGIN*2,MARGIN*1];
					alignChildren = ["left","top"];
					spacing = GAP*3;
					var _bookmarkCreatedGroup = add("group");
					with(_bookmarkCreatedGroup) {
						margins.top = 4;
						_bookmarkCreatedCheckbox = add("checkbox", undefined, localize(_global.createLabel));
					} /* END _bookmarkCreatedGroup */
					_bookmarkMarkerGroup = add("group");
					with(_bookmarkMarkerGroup) {
						alignChildren = "left";
						margins = [0,0,0,0];
						spacing = GAP;
						add("statictext", undefined, localize(_global.bookmarkMarkerLabel) + ":");
						_bookmarkMarkerEdittext = add("edittext", undefined, "");
						with(_bookmarkMarkerEdittext) {
							characters = 6;
						} /* END _bookmarkMarkerEdittext */
					} /* END _bookmarkMarkerGroup */
				} /* END _bookmarkPanel */
			} /* END _columnThree */
		} /* END _optionsGroup */
		/* Action Buttons */
		var _actionButtonGroup = add("group");
		with(_actionButtonGroup) {
			alignChildren = ["fill","fill"];
			margins = [0,0,GAP*4,0];
			spacing = 8;
			_cancelButton = add("button", undefined, localize(_global.cancelButtonLabel), { name:"Cancel" });
			with(_cancelButton) {
				alignment = ["right","top"];
			} /* END _cancelButton */
			_okButton = add("button", undefined, localize(_global.okButtonLabel), { name:"OK" });
			with(_okButton) {
				alignment = ["right","top"];
			} /* END _okButton */
		} /* END _actionButtonGroup */
	} /* END _importDialog */
	

	/* Callbacks */
	_bookmarkCreatedCheckbox.onClick = function() {
		if(_bookmarkCreatedCheckbox.value === true) {
			_bookmarkMarkerGroup.enabled = true;
		} else {
			_bookmarkMarkerGroup.enabled = false;
		}
	};

	_imagePlacedRadiobutton.onClick = 
	_imageMarkedRadiobutton.onClick =
	_imageRemovedRadiobutton.onClick = function() {
		if(_imagePlacedRadiobutton.value === true) {
			_imageInputGroup.enabled = true;
		} else {
			_imageInputGroup.enabled = false;
		}
	};

	_textboxCreatedRadiobutton.onClick = 
	_textboxMarkedRadiobutton.onClick =
	_textboxRemovedRadiobutton.onClick = function() {
		if(_textboxCreatedRadiobutton.value === true) {
			_textboxInputGroup.enabled = true;
		} else {
			_textboxInputGroup.enabled = false;
		}
	};
	
	_cancelButton.onClick = function() {
		_importDialog.hide();
		_importDialog.close(2);
	};

	_okButton.onClick = function() {
		_importDialog.hide();
		_importDialog.close(1);
	};
	/* END Callbacks */
	

	/* Initialize Dialog */	
	_defaultParagraphStyleEdittext.text = _setupObj["document"]["defaultParagraphStyle"];
	_isAutoflowingCheckbox.value = _setupObj["document"]["isAutoflowing"];
	_isUntaggedCheckbox.value = _setupObj["document"]["isUntagged"];
	if(_setupObj["import"]["styleMode"] === "extended") {
		_extendedStyleModeRadiobutton.value = true; 
	} else {
		_minimizedStyleModeRadiobutton.value = true;
	}
	_pageBreakCheckbox.value = _setupObj["pageBreak"]["isInserted"];
	_columnBreakCheckbox.value = _setupObj["columnBreak"]["isInserted"];
	_forcedLineBreakCheckbox.value = _setupObj["forcedLineBreak"]["isInserted"];
	_sectionBreakCheckbox.value = _setupObj["sectionBreak"]["isInserted"];
	_commentCreatedRadiobutton.value = _setupObj["comment"]["isCreated"];
	_commentMarkedRadiobutton.value = _setupObj["comment"]["isMarked"];
	_commentRemovedRadiobutton.value = _setupObj["comment"]["isRemoved"];
	_indexmarkCreatedRadiobutton.value = _setupObj["indexmark"]["isCreated"];
	_indexmarkRemovedRadiobutton.value = _setupObj["indexmark"]["isRemoved"];
	_hyperlinkCreatedRadiobutton.value = _setupObj["hyperlink"]["isCreated"];
	_hyperlinkMarkedRadiobutton.value = _setupObj["hyperlink"]["isMarked"];
	_crossReferenceCreatedRadiobutton.value = _setupObj["crossReference"]["isCreated"];
	_crossReferenceMarkedRadiobutton.value = _setupObj["crossReference"]["isMarked"];
	_bookmarkCreatedCheckbox.value = _setupObj["bookmark"]["isCreated"];
	_bookmarkMarkerEdittext.text = _setupObj["bookmark"]["marker"];
	_trackChangesMarkedRadiobutton.value = _setupObj["trackChanges"]["isMarked"];
	_trackChangesRemovedRadiobutton.value = _setupObj["trackChanges"]["isRemoved"];
	_footnoteCreatedRadiobutton.value = _setupObj["footnote"]["isCreated"];
	_footnoteMarkedRadiobutton.value = _setupObj["footnote"]["isMarked"];
	_footnoteRemovedRadiobutton.value = _setupObj["footnote"]["isRemoved"];
	_endnoteCreatedRadiobutton.value = _setupObj["endnote"]["isCreated"];
	_endnoteMarkedRadiobutton.value = _setupObj["endnote"]["isMarked"];
	_endnoteRemovedRadiobutton.value = _setupObj["endnote"]["isRemoved"];
	_imagePlacedRadiobutton.value = _setupObj["image"]["isPlaced"];
	_imageMarkedRadiobutton.value = _setupObj["image"]["isMarked"];
	_imageRemovedRadiobutton.value = _setupObj["image"]["isRemoved"];
	_imageWidthEdittext.text = _setupObj["image"]["width"];
	_imageHeightEdittext.text = _setupObj["image"]["height"];
	_textboxCreatedRadiobutton.value = _setupObj["textbox"]["isCreated"];
	_textboxMarkedRadiobutton.value = _setupObj["textbox"]["isMarked"];
	_textboxRemovedRadiobutton.value = _setupObj["textbox"]["isRemoved"];
	_textboxWidthEdittext.text = _setupObj["textbox"]["width"];
	_textboxHeightEdittext.text = _setupObj["textbox"]["height"];

	if(_bookmarkCreatedCheckbox.value === true) {
		_bookmarkMarkerGroup.enabled = true;
	} else {
		_bookmarkMarkerGroup.enabled = false;
	}
	if(_imagePlacedRadiobutton.value === true) {
		_imageInputGroup.enabled = true;
	} else {
		_imageInputGroup.enabled = false;
	}
	if(_textboxCreatedRadiobutton.value === true) {
		_textboxInputGroup.enabled = true;
	} else {
		_textboxInputGroup.enabled = false;
	}


	/* Show Dialog */
	var _closeValue = _importDialog.show ();
	if(_closeValue === 2) { 
		return null; 
	}
	

	/* Evaluate inputs */
	_setupObj["document"]["defaultParagraphStyle"] = _defaultParagraphStyleEdittext.text;
	_setupObj["document"]["isAutoflowing"] = _isAutoflowingCheckbox.value;
	_setupObj["document"]["isUntagged"] = _isUntaggedCheckbox.value;
	_setupObj["import"]["styleMode"] = (_extendedStyleModeRadiobutton.value && "extended") || (_minimizedStyleModeRadiobutton.value && "minimized");
	_setupObj["pageBreak"]["isInserted"] = _pageBreakCheckbox.value;
	_setupObj["columnBreak"]["isInserted"] = _columnBreakCheckbox.value;
	_setupObj["forcedLineBreak"]["isInserted"] = _forcedLineBreakCheckbox.value;
	_setupObj["sectionBreak"]["isInserted"] = _sectionBreakCheckbox.value;
	_setupObj["comment"]["isCreated"] = _commentCreatedRadiobutton.value;
	_setupObj["comment"]["isMarked"] = _commentMarkedRadiobutton.value;
	_setupObj["comment"]["isRemoved"] = _commentRemovedRadiobutton.value;
	_setupObj["indexmark"]["isCreated"] = _indexmarkCreatedRadiobutton.value;
	_setupObj["indexmark"]["isRemoved"] = _indexmarkRemovedRadiobutton.value;
	_setupObj["hyperlink"]["isCreated"] = _hyperlinkCreatedRadiobutton.value;
	_setupObj["hyperlink"]["isMarked"] = _hyperlinkMarkedRadiobutton.value;
	_setupObj["crossReference"]["isCreated"] = _crossReferenceCreatedRadiobutton.value;
	_setupObj["crossReference"]["isMarked"] = _crossReferenceMarkedRadiobutton.value;
	_setupObj["bookmark"]["isCreated"] = _bookmarkCreatedCheckbox.value;
	_setupObj["bookmark"]["marker"] = _bookmarkMarkerEdittext.text;
	_setupObj["trackChanges"]["isMarked"] = _trackChangesMarkedRadiobutton.value;
	_setupObj["trackChanges"]["isRemoved"] = _trackChangesRemovedRadiobutton.value;
	_setupObj["footnote"]["isCreated"] = _footnoteCreatedRadiobutton.value;
	_setupObj["footnote"]["isMarked"] = _footnoteMarkedRadiobutton.value;
	_setupObj["footnote"]["isRemoved"] = _footnoteRemovedRadiobutton.value;
	_setupObj["endnote"]["isCreated"] = _endnoteCreatedRadiobutton.value;
	_setupObj["endnote"]["isMarked"] = _endnoteMarkedRadiobutton.value;
	_setupObj["endnote"]["isRemoved"] = _endnoteRemovedRadiobutton.value;
	_setupObj["image"]["isPlaced"] = _imagePlacedRadiobutton.value;
	_setupObj["image"]["isMarked"] = _imageMarkedRadiobutton.value;
	_setupObj["image"]["isRemoved"] = _imageRemovedRadiobutton.value;
	_setupObj["image"]["width"] = _imageWidthEdittext.text;
	_setupObj["image"]["height"] = _imageHeightEdittext.text;
	_setupObj["textbox"]["isCreated"] = _textboxCreatedRadiobutton.value;
	_setupObj["textbox"]["isMarked"] = _textboxMarkedRadiobutton.value;
	_setupObj["textbox"]["isRemoved"] = _textboxRemovedRadiobutton.value;
	_setupObj["textbox"]["width"] = _textboxWidthEdittext.text;
	_setupObj["textbox"]["height"] = _textboxHeightEdittext.text;

	return _setupObj;
} /* END function __showImportDialog */