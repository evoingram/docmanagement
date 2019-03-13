# docmanagement

These are sample vba codes for managing Office documents.  Some are dependent on an access database. 

All code provided as is, no warranty.

Short Function Description:

	 AcrobatGetNumPages:
		gets number of pages from PDF and confirms with you

	AddAQCExamStyle:  
		finds specific phrases in activedocument(transcript) and applies a specific style
	
	BuildRCWsAndRCWHyperlinks:
		Acquires RCWs and their hyperlinks, validates them, and adds an entry for each one to CitationHyperlinks table
		
	BuildUSCRulesandUSCHyperlinks:
		validates U.S.C. code citations and builds corresponding hyperlinks, no front matter, no appendices
		
	AutoCalculateInterest:
		add 1% interest cost after every 7 days payment not made
		
	CDLabelMergeF:  
		makes CD label and prompts for print or no
		
	CommunicationHistoryAdd:  
		adds entry to CommunicationHistory table in Access DB
		
	CreateIndexBMKs:  
		replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
	
	CreateWorkingCopy:
		creates "working copy" sent to client
		
	FPJurors:  
		does find/replacements of prospective juror shorthand in transcript
	
	FillInPDFfromDatabase:
		inserts page count & other transcript info into invoice PDF
	
	FindAndReplaceCitationHyperlinks:  
		adds citations and hyperlinks from CitationHyperlinks table in transcript
		
	GenerateInvoiceAndEmailWithPPButton:
		generates invoice and Outlook e-mail body to include a linked PP button
		
	GenerateInvoiceNumber:
		generates invoice number
		
	GenericExportandMailMerge:  
		exports to specified template from T:\Document Generator\Templates and saves in T:\In Progress\sCourtDatesID\
		
  	pfAutoCorrect:
		adds autocorrect entries as listed on form (from matching table row) to rough draft autocorrect in Word
		
  	ReadXML:  
		reads shipping XML and sends "Shipped" email to client
		
	SendWordDocAsEmail:  
		sends Word document as an e-mail body with optional attachments
	
	TCEntryReplacementPARENT:
	 	parent function that finds certain entries within a transcript & assigns TC entries to them for indexing purposes
  	
	TCEntryReplacementCHILD-SingleReplaceAll
		one replace TC entry function for ones with no field entry
  	
	TCEntryReplacementCHILD-FieldReplaceAll
		one replace TC entry function for ones with field entry
  	
	WordIndexBuilder:  
		builds word index in separate docx & PDF from transcript
		
  	~MySQLExamples
		samples of actual queries from my database.
  
