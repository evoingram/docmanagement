# docmanagement

These are sample vba codes for managing Office documents.  Some are dependent on an access database. 

All code provided as is, no warranty.

Short Function Description:

	AddAQCExamStyle:  
		finds specific phrases in activedocument(transcript) and applies a specific style
		
	AutoCalculateInterest:
		add 1% interest cost after every 7 days payment not made
		
	CDLabelMergeF:  
		makes CD label and prompts for print or no
		
	CommunicationHistoryAdd:  
		adds entry to CommunicationHistory table in Access DB
		
	ConcordanceBuilder:  
		builds word index in separate docx & PDF from transcript
		
	CreateIndexBMKs:  
		replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
		
	FPJurors:  
		does find/replacements of prospective juror shorthand in transcript
		
	FindAndReplaceCitationHyperlinks:  
		adds citations and hyperlinks from CitationHyperlinks table in transcript
		
	GenericExportandMailMerge:  
		exports to specified template from T:\Document Generator\Templates and saves in T:\In Progress\sCourtDatesID\
  
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
  
  	~MySQLExamples
		samples of actual queries from my database.
  
