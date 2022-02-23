/**
 * Remove XML elements
 * @param {Array} _xmlElementArray 
 * @param {String} _label 
 * @returns Boolean
 */
function __removeXMLElements(_xmlElementArray, _label) {

	if(!_xmlElementArray || !(_xmlElementArray instanceof Array)) { 
		throw new Error("XMLElement as parameter required."); 
	}

	if(_label === null || _label === undefined || _label.constructor !== String) {
		_label = "";
	}

	var _counter = 0;

	for(var i=0; i<_xmlElementArray.length; i+=1) {

		var _xmlElement = _xmlElementArray[i];
		if(!_xmlElement || !_xmlElement.hasOwnProperty("remove") || !_xmlElement.isValid) {
			continue;
		}

		try {
			_xmlElement.remove();
		} catch(_error) {
			_global["log"].push(_error.message);
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.removeXMLElementsMessage, _counter, _label));
	}

	return true;
} /* END function __removeXMLElements */



/**
 * Mark XML elements with condition
 * @param {Document} _doc
 * @param {Array} _xmlElementArray 
 * @param {String} _label 
 * @param {Array} _colorArray 
 * @param {String} _indicatorMethod 
 * @returns Boolean
 */
function __markXMLElements(_doc, _xmlElementArray, _label, _colorArray, _indicatorMethod) {

	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");  
	}
	if(!_xmlElementArray || !(_xmlElementArray instanceof Array)) { 
		throw new Error("XMLElement as parameter required."); 
	}
	
	if(_label === null || _label === undefined || _label.constructor !== String) {
		_label = "</>";
	}
	if(!_colorArray || !(_colorArray instanceof Array) || _colorArray.length !== 3 || isNaN(_colorArray[0]) || isNaN(_colorArray[1]) || isNaN(_colorArray[2])) { 
		_colorArray = [255,255,80];  
	}

	_indicatorMethod = _indicatorMethod || "USE_HIGHLIGHT";

	var _markerCondition = __createCondition(_doc, _label, _colorArray, _indicatorMethod);

	var _counter = 0;

	for(var i=0; i<_xmlElementArray.length; i+=1) {

		var _xmlElement = _xmlElementArray[i];
		if(!_xmlElement || !_xmlElement.hasOwnProperty("texts") || !_xmlElement.isValid) {
			continue;
		}

		var _text = _xmlElement.texts[0];
		if(!_text || !_text.hasOwnProperty("applyConditions") || !_text.isValid) {
			continue;
		}

		try {
			_text.applyConditions([_markerCondition], false);
		} catch(_error) {
			_global["log"].push(_error.message);
			continue;
		}

		_counter += 1;
	}

	if(_global["isLogged"]) {
		_global["log"].push(localize(_global.markXMLElementsMessage, _counter, _label));
	}

	return true;
} /* END function __markXMLElements */


/**
 * Create text condition
 * @param {Document} _doc
 * @param {String} _name 
 * @param {Array} _color Array of 3 integers
 * @param {String} _conditionIndicatorMethod 
 * @returns {Condition}
 */
function __createCondition(_doc, _name, _colorArray, _indicatorMethod) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { return null; }
	if(!_name || _name.constructor !== String ) { return null; }
	if(!_colorArray || !(_colorArray instanceof Array) || _colorArray.length !== 3 || isNaN(_colorArray[0]) || isNaN(_colorArray[1]) || isNaN(_colorArray[2])) { return null; }
	if(!_indicatorMethod || _indicatorMethod.constructor !== String ) { return null; }
	
	var _condition = _doc.conditions.itemByName(_name);

	if(!_condition.isValid) {
		try {
			_condition = _doc.conditions.add ({ 
				name: _name, 
				indicatorMethod: ConditionIndicatorMethod[_indicatorMethod],
				indicatorColor: _colorArray
			});
		} catch(_error) {
			_global["log"].push(_error.message);
			return null;
		}
	}

	if(!_condition || !_condition.isValid) { 
		return null; 
	}
	
	return _condition;
} /* END function __createCondition */


/**
 * Generate Timestamp
 * @returns String
 */
function __getTimestamp() {
	
	var _date = new Date();

	var _timestamp = "";
	
	_timestamp += _date.getFullYear().toString();
	_timestamp += __padZeros(_date.getMonth(), 2);
	_timestamp += __padZeros(_date.getDate(), 2);
	_timestamp += "_";
	_timestamp += __padZeros(_date.getHours(), 2);
	_timestamp += __padZeros(_date.getMinutes(), 2);
	_timestamp += __padZeros(_date.getSeconds(), 2);
	_timestamp += __padZeros(_date.getMilliseconds(), 2);
	
	return _timestamp;
} /* END __convertToJSDate */


/**
 * Prepend Zeros to Number
 * @param {Number} _number 
 * @param {Number} _numOfPlaces 
 * @returns String
 */
function __padZeros(_number, _numOfPlaces) {
	
	if(
		isNaN(_number) || 
		!isFinite(_number) ||
		Math.floor(_number) !== _number ||
		_number < 0
	) { 
		return ""; 
	}
	if(
		isNaN(_numOfPlaces) || 
		!isFinite(_numOfPlaces) ||
		Math.floor(_numOfPlaces) !== _numOfPlaces ||
		_numOfPlaces < 1
	) { 
		return "";
	}
	
	var _string = _number.toString();
	
	while(_string.length < _numOfPlaces) {
			_string = '0' + _string;
	}

	return _string;
} /* END function __padZeros */