# docmanagement

These are sample vba codes for managing Office documents.  Some are dependent on an access database. 

All code provided as is, no warranty.

I'll be explaining how all this code works and more on my blog at https://transcription.aquoco.co/

Hackerrank:  https://www.hackerrank.com/evoingram
Pluralsight: https://app.pluralsight.com/profile/erica-ingram
Company:     https://www.aquoco.co

Working on a portfolio, but here's some of my VBA code for now.  There's a whole bunch of good stuff here.

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
		
  	AutoCorrect:
		adds autocorrect entries as listed on form (from matching table row) to rough draft autocorrect in Word
		
	CDLabelMergeF:  
		makes CD label and prompts for print or no
		
	CommunicationHistoryAdd:  
		adds entry to CommunicationHistory table in Access DB
		
	CreateIndexBMKs:  
		replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
		
	CreateIndexesTOAs:
		creates several TOCs that are marked differently and several sections of tables of authority.
		example: separate TOCs for exhibits, witnesses, and a general index
			 separate TOAs, three sections, one for cases, one for rules/regs/statutes/code/etc, one for other authority
		
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
		
	HeadersFooters:  
		programmatically adds headers and footers
		
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
  	
	WordIndexer:  
		builds word index in separate docx & PDF from transcript
	
	WunderlistAdd:
		arguments sTitle as string, sDueDate as string, due date format YYYY-MM-DD
		adds task to Wunderlist
		
	WunderlistGetAllLists:
		gets all lists from Wunderlist
		
	WunderlistGetList:
		gets tasks from an existing wunderlist list

  	~MySQLExamples
		samples of actual queries from my database.
  
