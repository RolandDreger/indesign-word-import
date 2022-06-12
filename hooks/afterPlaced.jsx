/**
 * Hook: afterPlace 
 * Handler runs after placing XML 
 * @param {Document} _doc 
 * @param {Object} _unpackObj
 * @param {XMLElement} _wordXMLElement
 * @param {Story} _wordStory
 * @param {Object} _setupObj
 * @returns Object
 */
 Hooks.prototype.afterPlaced = function (_doc, _unpackObj, _wordXMLElement, _wordStory, _setupObj) {
	
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

	// $.writeln("afterPlaced");

	/* Update Cross-reference sources */
	if(_doc.crossReferenceSources.length > 0) {
		_doc.crossReferenceSources.everyItem().update();
	}

	/* Update index */
	if(_doc.indexes.length > 0) {
		_doc.indexes[0].update();
	}
	
	return {};
} /* END method afterPlaced */