/**
 * Hook: beforePlace 
 * Handler runs before placing XML
 * @param {Document} _doc 
 * @param {Object} _unpackObj
 * @param {XMLElement} _wordXMLElement
 * @param {Object} _setupObj
 * @returns Object
 */
function __beforePlace(_doc, _unpackObj, _wordXMLElement, _setupObj) {
	
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

	// $.writeln("beforePlace");

	return {};
} /* END function __beforePlace */