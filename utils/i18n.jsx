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
	
	_global.indesignErrorMessage = { 
		en:"Error message [%1] Line [%2]",
		de:"Fehlermeldung [%1] Zeile [%2]" 
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
	
	_global.openProgressLabel = { 
		en: "Open Word Document ...",
		de: "Word-Dokument \u00f6ffnen ..." 
	};

	_global.importProgressLabel = { 
		en: "Import Word Document ...",
		de: "Word-Dokument importieren ..." 
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
		de: "Import ist nur f\u00fcr Word-Dokumente (.docx) oder Word-XML-Dokument (.xml) m\u00f6glich." 
	};
	
	_global.createFolderErrorMessage = { 
		en: "Order could not be created: %1",
		de: "Order konnte nicht erstellt werden: %1" 
	};

	_global.unpackageFolderErrorMessage = { 
		en: "Destination folder for the unzipped file could not be created: %1",
		de: "Ziel-Ordner f\u00fcr die entpackte Datei konnte nicht erstellt werden: %1" 
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
		en: "File for import could not be found: [%1]",
		de: "Datei f\u00fcr Import konnte nicht gefunden werden: [%1]" 
	};

	_global.noTargetPageErrorMessage= { 
		en: "Target page could not be determined.",
		de: "Zielseite konnte nicht ermittelt werden." 
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

	_global.untagXMLElementsMessage = { 
		en: "%1 %2 untaged.",
		de: "%1 %2 Tags entfernt." 
	};

	_global.removeXMLElementsMessage = { 
		en: "%1 %2 removed.",
		de: "%1 %2 gel\u00f6scht." 
	};

	_global.markXMLElementsMessage = { 
		en: "%1 %2 marked.",
		de: "%1 %2 markiert." 
	};

	_global.createXMLElementsMessage = { 
		en: "%1 %2 created.",
		de: "%1 %2 erstellt." 
	};

	_global.insertImageSourcesMessage = { 
		en: "%1 %2 inserted as plain text.",
		de: "%1 %2 als Text eingef\u00fcgt." 
	};

	_global.imageSourcesLabel = { 
		en: "images sources",
		de: "Bildquellen" 
	};

	_global.placeImageMessage = { 
		en: "%1 %2 placed.",
		de: "%1 %2 plaziert." 
	};

	_global.imageLabel = { 
		en: "image",
		de: "Bild" 
	};

	_global.footnotesLabel = { 
		en: "footnotes",
		de: "Fu\u00dfnoten" 
	};

	_global.footnoteValidationErrorMessage = { 
		en: "Footnote not valid.",
		de: "Fu\u00dfnote nicht valide." 
	};

	_global.footnoteParagraphStyleErrorMessage = {
		en: "Footnote [%1]: Error applying paragraph styles.",
		de: "Fu\u00dfnote [%1]: Fehler beim Zuweisen der Absatzformate."
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
		en: "Special character not available: [1%]",
		de: "Sonderzeichen nicht verf\u00fcgbar: [%1]" 
	}; 
	
	_global.xmlElementNotEmptyErrorMessage = { 
		en: "XML element [%1] not empty: [%2]",
		de: "XML-Element [%1] nicht leer: [%2]" 
	};
		
	_global.insertSpecialCharactersMessage = { 
		en: "%1 special characters [%2] inserted.",
		de: "%1 Sonderzeichen [%2] eingef\u00fcgt." 
	};

	_global.indexmarksLabel = { 
		en: "Indexmarks",
		de: "Indexmarken" 
	};

	_global.indexmarkValidationErrorMessage = { 
		en: "Indexmark not valid.",
		de: "Indexmarke nicht valide." 
	};

	_global.missingIndexmarkTypeMessage = { 
		en: "Indexmark element without type. Attribute [%1]",
		de: "Indexmarker-Element ohne Typ. Attribut [%1]" 
	};

	_global.missingIndexmarkEntryMessage = { 
		en: "Indexmark element without entry. Attribute [%1]",
		de: "Indexmarker-Element ohne Eintrag. Attribut [%1]" 
	};

	_global.missingIndexmarkFormatMessage = { 
		en: "Indexmark element without format. Attribute [%1]",
		de: "Indexmarker-Element ohne Format. Attribut [%1]" 
	};

	_global.missingIndexmarkTargetMessage = { 
		en: "Indexmark element without target. Attribute [%1]",
		de: "Indexmarker-Element ohne Ziel. Attribut [%1]" 
	};

	_global.createTopicErrorMessage = { 
		en: "Index entry [%1]. Topic for index could not be created (correctly).",
		de: "Indexeintrag [%1]. Thema f\u00fcr Index konnte nicht (korrekt) erstellt werden." 
	};

	_global.indexPageRangeOptionErrorMessage = {
		en: "Index entry [%1] Target [%2]. There is no direct equivalent in InDesign for the option [Page range → bookmark] in Word. Please check the entries in the index.",
		de: "Indexeintrag [%1] Ziel [%2]. F\u00fcr die Option [Seitenbereich → Textmarke] in Word gibt es keine direkte Entsprechung in InDesign. Bitte die Eintr\u00e4ge im Index kontrollieren."
	};

	_global.getNumberOfParagraphsErrorMessage = {
		en: "Index entry [%1] Target [%2]. The number of paragraphs for page reference could not be determined.",
		de: "Indexeintrag [%1] Ziel [%2]. Die Anzahl der Abs\u00e4tze f\u00fcr die Seitenreferenz konnte nicht ermittelt werden."
	};

	_global.maximumTopicLevelsErrorMessage = { 
		en: "Index entry [%1]. A maximum of 4 topic levels are allowed. Determined via topic separator [%2].",
		de: "Indexeintrag [%1]. Es sind maximal 4 Themenebenen erlaubt. Ermittelt \u00fcber Thementrenner [%2]." 
	};

	_global.indexmarkTypeErrorMessage = { 
		en: "Type for index entry not defined or incorrect. Type [%1]",
		de: "Typ f\u00fcr Indexeintrag nicht definiert oder fehlerhaft. Typ [%1]" 
	};

	_global.topicCrossReferenceErrorMessage = { 
		en: "Index entry [%1] Target [%2]. Cross-reference for index entry could not be created(correctly).",
		de: "Eintrag [%1] Ziel [%2]. Querverweis f\u00fcr Indexeintrag konnte nicht (korrekt) erstellt werden." 
	};

	_global.pageReferenceErrorMessage = { 
		en: "Index entry [%1] Target [%2]. Page reference for index entry could not be created(correctly).",
		de: "Eintrag [%1] Ziel [%2]. Seitenverweis f\u00fcr Indexeintrag konnte nicht (korrekt) erstellt werden." 
	};

	_global.movePageReferenceErrorMessage = { 
		en: "Page reference for index entry could not be inserted at the correct position.",
		de: "Seitenverweis f\u00fcr Indexeintrag konnte nicht an der korrekten Stelle eingef\u00fcgt werden." 
	};

	_global.indexEntryBookmarkNotFoundMessage = { 
		en: "Bookmark for index entry (page range) could not be found. Bookmark ID [%1]",
		de: "Textmarke f\u00fcr Indexeintrag (Seitenbereich) konnte nicht gefunden werden. ID Textmarke [%1]" 
	};

	_global.commentsLabel = { 
		en: "Comments",
		de: "Kommentare" 
	};

	_global.commentValidationErrorMessage = { 
		en: "Comment not valid.",
		de: "Kommentar nicht valide." 
	};

	_global.textboxesLabel = { 
		en: "Textboxes",
		de: "Textboxen" 
	};

	_global.commentValidationErrorMessage = { 
		en: "Textbox not valid.",
		de: "Textbox nicht valide." 
	};

	_global.imagesLabel = { 
		en: "Images",
		de: "Bilder" 
	};

	_global.insertedTextLabel = { 
		en: "Inserted Text",
		de: "Eingef\u00fcgter Text" 
	};

	_global.deletedTextLabel = { 
		en: "Deleted Text",
		de: "Gel\u00f6schter Text" 
	};

	_global.movedFromTextLabel = { 
		en: "Deleted Text",
		de: "Gel\u00f6schter Text" 
	};

	_global.movedToTextLabel = { 
		en: "Moved Text",
		de: "Verschobener Text" 
	};

	_global.wordFolderValidationMessage = { 
		en: "Folder with unzipped Word files could not be found.",
		de: "Folder mit entpackten Word-Dateien konnte nicht gefunden werden." 
	};

	_global.imageFileValidationMessage = { 
		en: "Media file could not be found: [%1]",
		de: "Medien-Datei konnte nicht gefunden werden: [%1]" 
	};

	_global.missingImageSourceMessage = { 
		en: "Media element without source. Attribute [%1]",
		de: "Medien-Element ohne Quelle. Attribut [%1]" 
	};

	_global.hyperlinksLabel = { 
		en: "hyperlinks",
		de: "Hyperlinks" 
	};

	_global.missingHyperlinkURIMessage = { 
		en: "Hyperlink element without URI. Attribute [%1]",
		de: "Hyperlink-Element ohne URI. Attribut [%1]" 
	};

	_global.emptyHyperlinkSourceMessage = { 
		en: "Hyperlink without text content. Index [%1] URI [%2]",
		de: "Hyperlink ohne Textinhalt. Index [%1] URI [%2]" 
	};

	_global.crossReferencesLabel = { 
		en: "cross-references",
		de: "Querverweise" 
	};

	_global.missingCrossReferenceURIMessage = { 
		en: "Cross-reference element without URI. Attribute [%1]",
		de: "Querverweis-Element ohne URI. Attribut [%1]" 
	};

	_global.missingCrossReferenceTypeMessage = { 
		en: "Cross-reference element without type. Attribute [%1]",
		de: "Querverweis-Element ohne Typ-Definition. Attribut [%1]" 
	};

	_global.missingCrossReferenceFormatMessage = { 
		en: "Cross-reference element without format. Attribute [%1]",
		de: "Querverweis-Element ohne Format-Definition. Attribut [%1]" 
	};

	_global.noMatchingCrossReferenceTypeMessage = { 
		en: "No matching cross reference type found.",
		de: "Kein passender Querverweistyp gefunden." 
	};

	_global.crossReferenceValidationMessage = { 
		en: "Cross-reference format not found. Type [%1] Format [%2]",
		de: "Querverweisformat nicht gefunden. Typ [%1] Format [%2]" 
	};

	_global.crossReferenceDestinationNotFoundMessage = { 
		en: "Cross-reference destination not found. ID [%1]",
		de: "Querverweisziel nicht gefunden. ID [%1]" 
	};

	_global.crossReferenceFormatWordImportLabel = {
		en: " (Word)",
		de: " (Word)"
	}	

	_global.pageNumberCrossReferenceFormatName = { 
		en: "Page number",
		de: "Seitenzahl" 
	};

	_global.paragraphTextCrossReferenceFormatName = { 
		en: "Paragraph Text",
		de: "Absatztext" 
	};

	_global.paragraphNumberCrossReferenceFormatName = { 
		en: "Paragraph number",
		de: "Absatznummer" 
	};

	_global.textAnchorNameCrossReferenceFormatName = { 
		en: "Text Anchor Name",
		de: "Name des Textankers" 
	};

	_global.pageLabel = { 
		en: "Page",
		de: "Seite",
		fr: "Page",
		es: "Página",
		it: "Pagina" 
	};

	_global.bookmarksLabel = { 
		en: "Bookmarks",
		de: "Lesezeichen" 
	};

	_global.importDialogTitle = { 
		en: "Microsoft Word Import Options",
		de: "Microsoft Word Importoptionen" 
	};
	
	_global.okButtonLabel = { 
		en: "OK",
		de: "OK" 
	};

	_global.cancelButtonLabel = { 
		en: "Cancel",
		de: "Abbrechen" 
	};

	_global.documentLabel = { 
		en: "Document",
		de: "Dokument" 
	};

	_global.metadataLabel = { 
		en: "Metadata",
		de: "Metadaten" 
	};

	_global.insertMetadataMessage = { 
		en: "[%1] Metadata interted",
		de: "[%1] Metadaten eingef\u00FCgt" 
	};

	_global.defaultParagraphStyleLabel = { 
		en: "Default paragraph style",
		de: "Standard-Absatzformat" 
	};

	_global.isAutoflowingLabel = { 
		en: "Automatic flow",
		de: "Automatischer Textfluss" 
	};

	_global.isUntaggedLabel = { 
		en: "Remove Tags",
		de: "Tags entfernen" 
	};

	_global.stylePanelLabel = { 
		en: "Styles",
		de: "Formate" 
	};

	_global.styleModeGroupLabel = { 
		en: "Character Styles",
		de: "Zeichenformate" 
	};

	_global.extendedStyleModeRadiobutton = { 
		en: "Extended Mode",
		de: "Erweiterter Modus" 
	};

	_global.minimizedStyleModeRadiobutton = { 
		en: "Minimized Mode",
		de: "Reduzierter Modus" 
	};

	_global.breaksHelpTip = { 
		en: "Inserting %1 as InDesign special characters",
		de: "%1 als InDesign-Sonderzeichen einf\u00FCgen" 
	};

	_global.breaksLabel = { 
		en: "Breaks",
		de: "Umbr\u00FCche" 
	};

	_global.pageBreakLabel = { 
		en: "Page break",
		de: "Seitenumbruch" 
	};

	_global.columnBreakLabel = { 
		en: "Column break",
		de: "Spaltenumbruch" 
	};

	_global.forcedLineLabel = { 
		en: "Forced line break",
		de: "Erzwungener Zeilenumbruch" 
	};

	_global.sectionLabel = { 
		en: "Section break",
		de: "Abschnittsumbruch" 
	};

	_global.commentsLabel = { 
		en: "Comments",
		de: "Kommentare" 
	};
	
	_global.removeLabel = { 
		en: "Remove",
		de: "Entfernen" 
	};

	_global.markLabel = { 
		en: "Mark",
		de: "Markieren" 
	};

	_global.ignoreLabel = { 
		en: "Ignore",
		de: "Ignorieren" 
	};

	_global.mergeLabel = { 
		en: "Merge",
		de: "Anf\u00FCgen" 
	};

	_global.replaceLabel = { 
		en: "Replace",
		de: "Ersetzen" 
	};

	_global.createLabel = { 
		en: "Create",
		de: "Erstellen" 
	};

	_global.placeLabel = { 
		en: "Place",
		de: "Platzieren" 
	};

	_global.indexmarksLabel = { 
		en: "Index marks",
		de: "Indexmarken" 
	};

	_global.hyperlinksLabel = { 
		en: "Hyperlinks",
		de: "Hyperlinks" 
	};

	_global.crossReferencesLabel = { 
		en: "Cross-references",
		de: "Querverweise" 
	};

	_global.bookmarksPanelLabel = { 
		en: "Bookmarks",
		de: "Lesezeichen" 
	};

	_global.bookmarkMarkerLabel = { 
		en: "Marker",
		de: "Marker" 
	};

	_global.trackChangesLabel = { 
		en: "Track changes",
		de: "\u00C4nderungsverfolgung" 
	};

	_global.footnotesLabel = { 
		en: "Footnotes",
		de: "Fu\u00DFnoten" 
	};

	_global.endnotesLabel = { 
		en: "Endnotes",
		de: "Endnoten" 
	};

	_global.imagesPanelLabel = { 
		en: "Images",
		de: "Bilder" 
	};

	_global.imageWidthLabel = { 
		en: "Width (mm)",
		de: "Breite (mm)" 
	};

	_global.imageHeightLabel = { 
		en: "Height (mm)",
		de: "H\u00F6he (mm)" 
	};

	_global.textboxesPanelLabel = { 
		en: "Textboxes",
		de: "Textfelder" 
	};

	_global.textboxWidthLabel = { 
		en: "Width (mm)",
		de: "Breite (mm)" 
	};

	_global.textboxHeightLabel = { 
		en: "Height (mm)",
		de: "H\u00F6he (mm)" 
	};
} /* END function __defLocalizeStrings */