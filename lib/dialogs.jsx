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