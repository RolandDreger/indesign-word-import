/**
 * Hook: beforeImport 
 * Handler runs before XML import 
 * @param {Document} _doc 
 * @param {Object} _unpackObj
 * @param {Object} _setupObj
 * @returns Object
 */
function __beforeImport(_doc, _unpackObj, _setupObj) {
	
	if(!_doc || !(_doc instanceof Document) || !_doc.isValid) { 
		throw new Error("Document as parameter required.");  
	}
	if(!_unpackObj || !(_unpackObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}
	if(!_setupObj || !(_setupObj instanceof Object)) { 
		throw new Error("Object as parameter required.");
	}

	$.writeln("beforeImport");

	return {};
} /* END function __beforeImport */