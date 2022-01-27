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